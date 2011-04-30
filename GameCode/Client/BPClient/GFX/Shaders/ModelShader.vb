Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class ModelShader

	Private Enum RenderFlags As Int32
		RenderNormalMap = 1
		RenderIllumMap = 2
        'RenderPlayerCustom = 4
        RenderSpecMap = 4
	End Enum

	Private moEffect As Effect = Nothing

	'Texture Samplers
	Private mhndlDiffuseMap As EffectHandle
	Private mhndlNormalMap As EffectHandle
    Private mhndlIllumMap As EffectHandle
    Private mhndlSpecMap As EffectHandle

	'Set in the SetStaticData stuff
	Private mhndlEyePos As EffectHandle
	Private mhndlLightDir As EffectHandle
	Private mhndlLightClr As EffectHandle
	Private mhndlLightAmb As EffectHandle
	Private mhndlLightSpec As EffectHandle

	'Used for Bump Mapping Technique
	Private mhndlWorldInv As EffectHandle		'WorldInverse
	Private mhndlWVP As EffectHandle			'WorldViewProjection

	'Used for Per Pixel Only Technique
	Private mhndlWorld As EffectHandle			'World
	Private mhndlWorldInvTrans As EffectHandle	'WorldInverseTranspose

	'EntityColorAdjust
    'Private mhndlEntityColor As EffectHandle
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

    Private mvec4SelectedAmbient As Vector4
    Private mvec4NormalAmbient As Vector4
    Private Shared mlSelectedAmbient As Int32 = 0
    Private Shared mlSelectedAmbientChng As Int32 = 1

    Public Sub PrepareToRender(ByVal vecLight As Vector3, ByVal clrDiffuse As System.Drawing.Color, ByVal clrAmbient As System.Drawing.Color, ByVal clrSpecular As System.Drawing.Color)

        mlRenderFlags = 0

        If muSettings.LightQuality = EngineSettings.LightQualitySetting.BumpMap Then mlRenderFlags = mlRenderFlags Or RenderFlags.RenderNormalMap
        If muSettings.RenderSpecularMap = True Then mlRenderFlags = mlRenderFlags Or RenderFlags.RenderSpecMap
        If muSettings.IlluminationMap <> -1 Then mlRenderFlags = mlRenderFlags Or RenderFlags.RenderIllumMap
        Dim moDevice As Device = GFXEngine.moDevice
        'Now, determine what technique to use...
        Try
            If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                'NormalMap
                Select Case (mlRenderFlags And (RenderFlags.RenderIllumMap Or RenderFlags.RenderSpecMap))
                    Case RenderFlags.RenderIllumMap
                        moEffect.Technique = "NormalMapWithIllum"
                    Case RenderFlags.RenderSpecMap
                        moEffect.Technique = "NormalMapWithSpecularMap"
                    Case (RenderFlags.RenderIllumMap Or RenderFlags.RenderSpecMap)
                        moEffect.Technique = "NormalMapFull"
                    Case Else
                        moEffect.Technique = "NormalMapOnly"
                End Select
            Else
                'PerPixel
                Select Case (mlRenderFlags And (RenderFlags.RenderIllumMap Or RenderFlags.RenderSpecMap))
                    Case RenderFlags.RenderIllumMap
                        moEffect.Technique = "PerPixelWithIllum"
                    Case RenderFlags.RenderSpecMap
                        moEffect.Technique = "PerPixelWithSpecularMap"
                    Case (RenderFlags.RenderIllumMap Or RenderFlags.RenderSpecMap)
                        moEffect.Technique = "PerPixelFull"
                    Case Else
                        moEffect.Technique = "PerPixelOnly"
                End Select
            End If
        Catch
            If goUILib Is Nothing = False Then goUILib.AddNotification("Unable to set shader technique. Possible incompatibility. Auto-adjusted settings.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            muSettings.LightQuality = EngineSettings.LightQualitySetting.VSPS1
            muSettings.RenderSpecularMap = False
            muSettings.IlluminationMap = -1
            Return
        End Try

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
        'vecLight.Y = 0
        vecLight.Normalize()
        'Dim vecTemp As Vector3 = New Vector3(0, -1, 0)
        'vecTemp.Normalize()
        moEffect.SetValue(mhndlLightDir, New Vector4(vecLight.X, vecLight.Y, vecLight.Z, 1))
        'moEffect.SetValue(mhndlLightDir, New Vector4(vecTemp.X, vecTemp.Y, vecTemp.Z, 1))

        'lightColor
        Dim fLtR As Single = (clrDiffuse.R / 255.0F) * fTimeOfDayMult
        Dim fLtG As Single = (clrDiffuse.G / 255.0F) * fTimeOfDayMult
        Dim fLtB As Single = (clrDiffuse.B / 255.0F) * fTimeOfDayMult
        moEffect.SetValue(mhndlLightClr, New Vector4(fLtR, fLtG, fLtB, 1.0F))
        'moEffect.SetValue(mhndlLightClr, New Vector4(fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult))
        'lightAmbient
        Dim fVal As Single = muSettings.AmbientLevel * 0.01F
        'moEffect.SetValue(mhndlLightAmb, New Vector4(Math.Max(0.19F, clrAmbient.R / 255.0F), Math.Max(0.19F, clrAmbient.G / 255.0F), Math.Max(0.19F, clrAmbient.B / 255.0F), 1.0F))
        mvec4NormalAmbient = New Vector4(fVal, fVal, fVal, fVal)

        If muSettings.FlashSelections = True Then
            mlSelectedAmbient += mlSelectedAmbientChng
            If mlSelectedAmbient < Math.Max(muSettings.AmbientLevel, 0) Then
                mlSelectedAmbient = Math.Max(muSettings.AmbientLevel, 0)
                mlSelectedAmbientChng = muSettings.FlashRate
            ElseIf mlSelectedAmbient > 75 Then
                mlSelectedAmbient = 75
                mlSelectedAmbientChng = -muSettings.FlashRate
            End If
            Dim fSelAmb As Single = mlSelectedAmbient * 0.01F
            mvec4SelectedAmbient = New Vector4(fSelAmb, fSelAmb, fSelAmb, fSelAmb)
        Else
            mvec4SelectedAmbient = mvec4NormalAmbient
        End If

        moEffect.SetValue(mhndlLightAmb, mvec4NormalAmbient)

        'lightSpecular
        fLtR = (((clrSpecular.R / 255.0F) * 0.75F) + 0.25F) * fTimeOfDayMult
        fLtG = (((clrSpecular.G / 255.0F) * 0.75F) + 0.25F) * fTimeOfDayMult
        fLtB = (((clrSpecular.B / 255.0F) * 0.75F) + 0.25F) * fTimeOfDayMult
        moEffect.SetValue(mhndlLightSpec, New Vector4(fLtR, fLtG, fLtB, 1.0F))
        'moEffect.SetValue(mhndlLightSpec, New Vector4(fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult))

        mlPasses = moEffect.Begin(FX.None)
        'moEffect.BeginPass(0)
    End Sub

	Public Sub EndRender()
		'moEffect.EndPass()
		moEffect.End()

        GFXEngine.moDevice.RenderState.FogEnable = mbFinalFogEnable
	End Sub

    Private moTmpTx As Texture = Nothing
	Public Sub RenderMesh(ByRef oObj As BaseEntity, ByVal ySpecialType As GFXEngine.RenderModelType, ByVal lTexMod As Int32)
		If oObj Is Nothing Then Return
		If oObj.oMesh Is Nothing Then Return
		If oObj.oMesh.oMesh Is Nothing Then Return
        If oObj.ObjectID = 4958929 Or oObj.ObjectID = 4952071 Or oObj.ObjectID = 4958927 Or oObj.ObjectID = 2527432 Then
            '   Debug.Print(oObj.ObjectID.ToString + " " + oObj.oUnitDef.ModelID.ToString + " " + oObj.LocX.ToString + " " + oObj.LocY.ToString + " " + oObj.LocZ.ToString)
        End If
        Dim matWorld As Matrix
        Dim moDevice As Device = GFXEngine.moDevice

		If (ySpecialType And GFXEngine.RenderModelType.eOldData) <> 0 Then
			matWorld = oObj.CurrentWorldMatrix
		Else : matWorld = oObj.GetWorldMatrix
		End If
        moDevice.Transform.World = matWorld

        If oObj.bSelected = True Then
            moEffect.SetValue(mhndlLightAmb, mvec4SelectedAmbient)
        Else
            moEffect.SetValue(mhndlLightAmb, mvec4NormalAmbient)
        End If

        If ySpecialType = GFXEngine.RenderModelType.eOldData Then
            moEffect.SetValue(mhndlRelColor, muSettings.TacticalAssetColor)
        Else
            Select Case lTexMod
                Case 1          'player
                    moEffect.SetValue(mhndlRelColor, muSettings.MyAssetColor)
                Case 2          'ally
                    moEffect.SetValue(mhndlRelColor, muSettings.AllyAssetColor)
                Case 3          'war
                    moEffect.SetValue(mhndlRelColor, muSettings.EnemyAssetColor)
                Case 4          'guild
                    moEffect.SetValue(mhndlRelColor, muSettings.GuildAssetColor)
                Case Else       'neutral
                    moEffect.SetValue(mhndlRelColor, muSettings.NeutralAssetColor)
            End Select
        End If

        'NOTE: we don't render iron curtain through here...
        With oObj.oMesh
            'Now, let's set our variables
            Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection) ' Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
            moEffect.SetValue(mhndlWVP, matWVP)
            Dim matWorldInv As Matrix = matWorld
            matWorldInv.Invert()
            moEffect.SetValue(mhndlWorldInv, matWorldInv)

            If (mlRenderFlags And RenderFlags.RenderNormalMap) = 0 Then
                'ok, use the pixel shader, mhndlWorld and mhndlWorldInvTrans
                moEffect.SetValue(mhndlWorld, matWorld)
                Dim matWorldInvTrans As Matrix = matWorld
                matWorldInvTrans.Invert()
                matWorldInvTrans = Matrix.TransposeMatrix(matWorldInvTrans)
                moEffect.SetValue(mhndlWorldInvTrans, matWorldInvTrans)
            End If
            moEffect.CommitChanges()

            'For lPass As Int32 = 0 To mlPasses - 1
            'mlPasses = moEffect.Begin(FX.None)
            moEffect.BeginPass(0)

            Try
                If .bMTTexCapable = True AndAlso oObj.oUnitDef Is Nothing = False Then
                    'ok, this is mttex capable, so render it properly
                    If (.lModelID And 255) = 141 Then
                        With oObj.oUnitDef
                            moEffect.SetValue(mhndlDiffuseMap, oObj.oMesh.oSharedTex(.lTexNum).oDiffuseTexture)
                            If (mlRenderFlags And RenderFlags.RenderIllumMap) <> 0 Then
                                If .lIllumNum <> 0 Then
                                    moEffect.SetValue(mhndlIllumMap, goResMgr.moGHTex(.lIllumNum - 1))
                                Else
                                    moEffect.SetValue(mhndlIllumMap, oObj.oMesh.oSharedTex(.lTexNum).oIllumTexture(.lIllumNum))
                                End If
                            End If
                            If (mlRenderFlags And RenderFlags.RenderSpecMap) <> 0 Then
                                moEffect.SetValue(mhndlSpecMap, oObj.oMesh.oSharedTex(.lTexNum).oSpecTexture)
                            End If
                            If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                                moEffect.SetValue(mhndlNormalMap, oObj.oMesh.oSharedTex(.lTexNum).oNormalTexture(.lNormalNum))
                            End If
                        End With
                    Else
                        With oObj.oUnitDef
                            moEffect.SetValue(mhndlDiffuseMap, oObj.oMesh.oSharedTex(.lTexNum).oDiffuseTexture)
                            If (mlRenderFlags And RenderFlags.RenderIllumMap) <> 0 Then
                                moEffect.SetValue(mhndlIllumMap, oObj.oMesh.oSharedTex(.lTexNum).oIllumTexture(.lIllumNum))
                            End If
                            If (mlRenderFlags And RenderFlags.RenderSpecMap) <> 0 Then
                                moEffect.SetValue(mhndlSpecMap, oObj.oMesh.oSharedTex(.lTexNum).oSpecTexture)
                            End If
                            If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                                moEffect.SetValue(mhndlNormalMap, oObj.oMesh.oSharedTex(.lTexNum).oNormalTexture(.lNormalNum))
                            End If
                        End With
                        
                    End If
                   
                    moEffect.CommitChanges()
                    For X As Int32 = 0 To .NumOfMaterials - 1
                        .oMesh.DrawSubset(X)
                    Next X
                ElseIf .bMTTexCapable = True AndAlso glCurrentEnvirView = CurrentView.eStartupLogin Then
                    'ok, this is mttex capable, so render it properly
                    Dim lTexNum As Int32 = 0
                    Dim lIllumNum As Int32 = 0
                    Dim lNormalNum As Int32 = 0
                    moEffect.SetValue(mhndlDiffuseMap, oObj.oMesh.oSharedTex(lTexNum).oDiffuseTexture)
                    If (mlRenderFlags And RenderFlags.RenderIllumMap) <> 0 Then
                        moEffect.SetValue(mhndlIllumMap, oObj.oMesh.oSharedTex(lTexNum).oIllumTexture(lIllumNum))
                    End If
                    If (mlRenderFlags And RenderFlags.RenderSpecMap) <> 0 Then
                        moEffect.SetValue(mhndlSpecMap, oObj.oMesh.oSharedTex(lTexNum).oSpecTexture)
                    End If
                    If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                        moEffect.SetValue(mhndlNormalMap, oObj.oMesh.oSharedTex(lTexNum).oNormalTexture(lNormalNum))
                    End If
                    moEffect.CommitChanges()
                    For X As Int32 = 0 To .NumOfMaterials - 1
                        .oMesh.DrawSubset(X)
                    Next X
                Else
                    If (mlRenderFlags And RenderFlags.RenderIllumMap) <> 0 Then
                        moEffect.SetValue(mhndlIllumMap, .oIllumMap)
                    End If
                    If (mlRenderFlags And RenderFlags.RenderSpecMap) <> 0 Then
                        moEffect.SetValue(mhndlSpecMap, GFXResourceManager.oDefaultSpecularTex)
                    End If
                    If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                        moEffect.SetValue(mhndlNormalMap, .oNormalMap)
                    End If

                    For X As Int32 = 0 To .NumOfMaterials - 1
                        moEffect.SetValue(mhndlDiffuseMap, .Textures(X))
                        moEffect.CommitChanges()
                        .oMesh.DrawSubset(X)
                    Next X
                End If

                'Now, check for a turret...
                If .bTurretMesh = True Then
                    'NOTE: Only allowed one material for the turret...
                    If .bMTTexCapable = False OrElse oObj.oUnitDef Is Nothing = True Then moEffect.SetValue(mhndlDiffuseMap, .Textures(0))

                    'get the turret's matrix
                    If (ySpecialType And GFXEngine.RenderModelType.eOldData) <> 0 Then
                        matWorld = oObj.CurrentTurretMatrix
                    Else : matWorld = oObj.GetTurretMatrix
                    End If
                    moDevice.Transform.World = matWorld

                    matWVP = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection) ' Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
                    moEffect.SetValue(mhndlWVP, matWVP)
                    matWorldInv = matWorld
                    matWorldInv.Invert()
                    moEffect.SetValue(mhndlWorldInv, matWorldInv)

                    If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                        'Ok, use the Normal Map shader, that uses... mhndlWorldInv and mhndlWVP and the normal map
                        If .oNormalMap Is Nothing = True OrElse .oNormalMap.Disposed = True Then moEffect.SetValue(mhndlNormalMap, CType(Nothing, Texture)) Else moEffect.SetValue(mhndlNormalMap, .oNormalMap)
                    Else
                        'ok, use the pixel shader, mhndlWorld and mhndlWorldInvTrans
                        moEffect.SetValue(mhndlWorld, matWorld)
                        Dim matWorldInvTrans As Matrix = matWorld
                        matWorldInvTrans.Invert()
                        matWorldInvTrans = Matrix.TransposeMatrix(matWorldInvTrans)
                        moEffect.SetValue(mhndlWorldInvTrans, matWorldInvTrans)
                    End If
                    moEffect.CommitChanges()

                    .oTurretMesh.DrawSubset(0)
                End If
                Try
                    If isAdmin() = True Then
                        If .bHasLights = True Then
                            For x As Int32 = 0 To .Lights.GetUpperBound(0)
                                With .Lights(x)
                                    'If .oLight Is Nothing = True Then
                                    '    .oLight = New Light
                                    'End If
                                    'moDevice.Lights(x).FromLight(.oLight)
                                    '.Color                 ' Its Color
                                    '.Source                ' Offset of oObj
                                    '.Direction             ' The direction its pointing, offset from oObj?
                                    '.LightTex              ' LoadTexture(the dds file)

                                    'GoTo diaf
                                    moDevice.Lights(x).Type = LightType.Directional
                                    moDevice.Lights(x).Direction = .Direction
                                    moDevice.Lights(x).Enabled = True
                                    moDevice.Lights(x).Update()

                                    Dim fSizeMult As Single = 5.0F
                                    Dim fMult As Single = 1.0F
                                    Dim matTemp As Matrix = Matrix.Identity
                                    matWorld = Matrix.Identity
                                    matTemp = Matrix.Identity
                                    'matTemp.Scale(fMult * fSizeMult, fMult * fSizeMult, fMult * fSizeMult)
                                    'matTemp = Matrix.RotationY(1.0F)
                                    matWorld.Multiply(matTemp)
                                    'matTemp = Matrix.Identity
                                    'Dim lCamXMod As Int32 = goCamera.mlCameraAtX
                                    'matTemp.Translate(((-oObj.LocX) * fMult) + lCamXMod, (-oObj.LocY) * fMult, goCamera.mlCameraAtZ)
                                    'matWorld.Multiply(matTemp)
                                    'matTemp = Matrix.Identity
                                    matTemp = Nothing

                                    moDevice.Transform.World = matWorld



                                    moDevice.RenderState.SourceBlend = Blend.SourceColor
                                    moDevice.RenderState.DestinationBlend = Blend.DestinationColor

                                    If .TheSprite Is Nothing Then
                                        Device.IsUsingEventHandlers = False
                                        .TheSprite = New Sprite(moDevice)
                                        Device.IsUsingEventHandlers = True
                                    End If
                                    'moDevice.RenderState.Lighting = False
                                    'moDevice.RenderState.ZBufferWriteEnable = False

                                    .TheSprite.SetWorldViewLH(matWorld, moDevice.Transform.View)
                                    .TheSprite.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace Or SpriteFlags.SortDepthBackToFront)

                                    Dim lClr As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                                    Dim vLoc As New Vector3(oObj.LocX, oObj.LocY, oObj.LocZ)

                                    'Ship is at locY 2704
                                    'X is north and south
                                    'Y is up and down altitude
                                    'X is Left to Right on planets when facing north


                                    'vLoc.X = vLoc.X + 300
                                    'vLoc.Y = vLoc.Y
                                    vLoc.X = vLoc.X + oObj.oMesh.RangeOffset
                                    lClr = System.Drawing.Color.White
                                    'vLoc.Y = vLoc.Y + 300 'for debugging
                                    vLoc.Y = vLoc.Y + (oObj.oMesh.YMidPoint * 2.5F) + 64
                                    '.TheSprite.Draw(.LightTex, New System.Drawing.Rectangle(0, 0, 128, 128), New Vector3(128, 128, 0), vLoc, lClr)
                                    .TheSprite.Draw(.LightTex, System.Drawing.Rectangle.Empty, New Vector3(16, 16, 0), vLoc, System.Drawing.Color.White)

                                    .TheSprite.End()

                                    'moDevice.RenderState.Lighting = True
                                    'moDevice.RenderState.ZBufferWriteEnable = True

                                    'moDevice.RenderState.SourceBlend = Blend.SourceAlpha
                                    'moDevice.RenderState.DestinationBlend = Blend.InvSourceAlpha
                                    'moDevice.RenderState.BlendOperation = BlendOperation.Add
                                    'moDevice.RenderState.ZBufferEnable = True


                                End With
                            Next
                        End If
                    End If
                Catch ex As Exception
                End Try
            Catch
                oObj.oMesh = goResMgr.ClearAndGetMesh(.lModelID)
            End Try
            moEffect.EndPass()
            'moEffect.End()
            'Next lPass 

        End With
	End Sub

    Public Sub RenderHullBuilderMesh(ByRef oMesh As BaseMesh, ByVal matWorld As Matrix, ByVal iFullModelID As Int16)
        mlRenderFlags = 0

        If muSettings.LightQuality = EngineSettings.LightQualitySetting.BumpMap Then mlRenderFlags = mlRenderFlags Or RenderFlags.RenderNormalMap
        If muSettings.RenderSpecularMap = True Then mlRenderFlags = mlRenderFlags Or RenderFlags.RenderSpecMap
        If muSettings.IlluminationMap <> -1 Then mlRenderFlags = mlRenderFlags Or RenderFlags.RenderIllumMap

        Dim moDevice As Device = GFXEngine.moDevice

        'Now, determine what technique to use...
        If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
            'NormalMap
            Select Case (mlRenderFlags And (RenderFlags.RenderIllumMap Or RenderFlags.RenderSpecMap))
                Case RenderFlags.RenderIllumMap
                    moEffect.Technique = "NormalMapWithIllum"
                Case RenderFlags.RenderSpecMap
                    moEffect.Technique = "NormalMapWithEntityColor"
                Case (RenderFlags.RenderIllumMap Or RenderFlags.RenderSpecMap)
                    moEffect.Technique = "NormalMapFull"
                Case Else
                    moEffect.Technique = "NormalMapOnly"
            End Select
        Else
            'PerPixel
            Select Case (mlRenderFlags And (RenderFlags.RenderIllumMap Or RenderFlags.RenderSpecMap))
                Case RenderFlags.RenderIllumMap
                    moEffect.Technique = "PerPixelWithIllum"
                Case RenderFlags.RenderSpecMap
                    moEffect.Technique = "PerPixelWithEntityColor"
                Case (RenderFlags.RenderIllumMap Or RenderFlags.RenderSpecMap)
                    moEffect.Technique = "PerPixelFull"
                Case Else
                    moEffect.Technique = "PerPixelOnly"
            End Select
        End If

        moEffect.SetValue(mhndlSpecPwr, 30.0F)

        Dim fTimeOfDayMult As Single = 1.0F
        Dim vecTemp As Vector3 = New Vector3(-goCamera.mlCameraX, -goCamera.mlCameraY, -goCamera.mlCameraZ)

        Dim vecLight As Vector3 = vecTemp 'Vector3.Multiply(moDevice.Lights(0).Direction, -1.0F)
        Dim clrDiffuse As System.Drawing.Color = Color.White 'moDevice.Lights(0).Diffuse
        Dim clrAmbient As System.Drawing.Color = moDevice.Lights(0).Ambient
        Dim clrSpecular As System.Drawing.Color = Color.White ' moDevice.Lights(0).Specular

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
        'moEffect.SetValue(mhndlLightAmb, New Vector4(0.19F, 0.19F, 0.19F, 0.19F))
        'moEffect.SetValue(mhndlLightAmb, New Vector4(0.19F, 0.19F, 0.19F, 0.19F))
        'lightSpecular
        moEffect.SetValue(mhndlLightSpec, New Vector4((clrSpecular.R / 255.0F) * fTimeOfDayMult, (clrSpecular.G / 255.0F) * fTimeOfDayMult, (clrSpecular.B / 255.0F) * fTimeOfDayMult, 1.0F))
        'moEffect.SetValue(mhndlLightSpec, New Vector4(fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult))

        mlPasses = moEffect.Begin(FX.None)
        'moEffect.BeginPass(0)

        moDevice.Transform.World = matWorld

        moEffect.SetValue(mhndlRelColor, muSettings.MyAssetColor)

        'NOTE: we don't render iron curtain through here...
        With oMesh
            'Now, let's set our variables
            Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection) ' Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
            moEffect.SetValue(mhndlWVP, matWVP)
            Dim matWorldInv As Matrix = matWorld
            matWorldInv.Invert()
            moEffect.SetValue(mhndlWorldInv, matWorldInv)

            If (mlRenderFlags And RenderFlags.RenderNormalMap) = 0 Then
                'ok, use the pixel shader, mhndlWorld and mhndlWorldInvTrans
                moEffect.SetValue(mhndlWorld, matWorld)
                Dim matWorldInvTrans As Matrix = matWorld
                matWorldInvTrans.Invert()
                matWorldInvTrans = Matrix.TransposeMatrix(matWorldInvTrans)
                moEffect.SetValue(mhndlWorldInvTrans, matWorldInvTrans)
            End If
            moEffect.CommitChanges()

            'For lPass As Int32 = 0 To mlPasses - 1
            'mlPasses = moEffect.Begin(FX.None)
            moEffect.BeginPass(0)


            If .bMTTexCapable = True Then
                'ok, this is mttex capable, so render it properly
                If (iFullModelID And 255) = 141 Then
                    Dim iTexNum As Int16 = (iFullModelID And 7936S) \ 256S
                    Dim iIllumNum As Int16 = (iFullModelID And 24576S) \ 8192S
                    Dim iBumpNum As Int16 = 0
                    If (iFullModelID And -32768) <> 0 Then iBumpNum = 1
                    moEffect.SetValue(mhndlDiffuseMap, .oSharedTex(iTexNum).oDiffuseTexture)
                    If (mlRenderFlags And RenderFlags.RenderIllumMap) <> 0 Then
                        If iIllumNum <> 0 Then
                            moEffect.SetValue(mhndlIllumMap, goResMgr.moGHTex(iIllumNum - 1))                        'moEffect.SetValue(mhndlIllumMap, .oSharedTex(iTexNum).oIllumTexture(iIllumNum))
                        Else
                            moEffect.SetValue(mhndlIllumMap, .oSharedTex(iTexNum).oIllumTexture(iIllumNum))
                        End If
                    End If
                    If (mlRenderFlags And RenderFlags.RenderSpecMap) <> 0 Then
                        moEffect.SetValue(mhndlSpecMap, .oSharedTex(iTexNum).oSpecTexture)
                    End If
                    If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                        moEffect.SetValue(mhndlNormalMap, .oSharedTex(iTexNum).oNormalTexture(iBumpNum))
                    End If
                Else
                    Dim iTexNum As Int16 = (iFullModelID And 7936S) \ 256S
                    Dim iBumpNum As Int16 = (iFullModelID And 24576S) \ 8192S
                    Dim iIllumNum As Int16 = 0
                    If (iFullModelID And -32768) <> 0 Then iIllumNum = 1
                    moEffect.SetValue(mhndlDiffuseMap, .oSharedTex(iTexNum).oDiffuseTexture)
                    If (mlRenderFlags And RenderFlags.RenderIllumMap) <> 0 Then
                        moEffect.SetValue(mhndlIllumMap, .oSharedTex(iTexNum).oIllumTexture(iIllumNum))
                    End If
                    If (mlRenderFlags And RenderFlags.RenderSpecMap) <> 0 Then
                        moEffect.SetValue(mhndlSpecMap, .oSharedTex(iTexNum).oSpecTexture)
                    End If
                    If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                        moEffect.SetValue(mhndlNormalMap, .oSharedTex(iTexNum).oNormalTexture(iBumpNum))
                    End If
                End If
                moEffect.CommitChanges()
                For X As Int32 = 0 To .NumOfMaterials - 1
                    .oMesh.DrawSubset(X)
                Next X
            Else
                If (mlRenderFlags And RenderFlags.RenderIllumMap) <> 0 Then
                    moEffect.SetValue(mhndlIllumMap, .oIllumMap)
                End If
                If (mlRenderFlags And RenderFlags.RenderSpecMap) <> 0 Then
                    moEffect.SetValue(mhndlSpecMap, GFXResourceManager.oDefaultSpecularTex)
                End If
                If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                    moEffect.SetValue(mhndlNormalMap, .oNormalMap)
                End If

                For X As Int32 = 0 To .NumOfMaterials - 1
                    moEffect.SetValue(mhndlDiffuseMap, .Textures(X))
                    moEffect.CommitChanges()
                    .oMesh.DrawSubset(X)
                Next X
            End If

            'For X As Int32 = 0 To .NumOfMaterials - 1
            '    moEffect.SetValue(mhndlDiffuseMap, .Textures(X))
            '    moEffect.CommitChanges()
            '    .oMesh.DrawSubset(X)
            'Next X

            'Now, check for a turret...
            If .bTurretMesh = True Then
                'NOTE: Only allowed one material for the turret...
                If .bMTTexCapable = False Then moEffect.SetValue(mhndlDiffuseMap, .Textures(0))

                'get the turret's matrix
                'If (ySpecialType And GFXEngine.RenderModelType.eOldData) <> 0 Then
                '    matWorld = oObj.CurrentTurretMatrix
                'Else : matWorld = oObj.GetTurretMatrix
                'End If
                moDevice.Transform.World = matWorld

                matWVP = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection) ' Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
                moEffect.SetValue(mhndlWVP, matWVP)
                matWorldInv = matWorld
                matWorldInv.Invert()
                moEffect.SetValue(mhndlWorldInv, matWorldInv)

                If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                    'Ok, use the Normal Map shader, that uses... mhndlWorldInv and mhndlWVP and the normal map
                    moEffect.SetValue(mhndlNormalMap, .oNormalMap)
                Else
                    'ok, use the pixel shader, mhndlWorld and mhndlWorldInvTrans
                    moEffect.SetValue(mhndlWorld, matWorld)
                    Dim matWorldInvTrans As Matrix = matWorld
                    matWorldInvTrans.Invert()
                    matWorldInvTrans = Matrix.TransposeMatrix(matWorldInvTrans)
                    moEffect.SetValue(mhndlWorldInvTrans, matWorldInvTrans)
                End If
                moEffect.CommitChanges()

                .oTurretMesh.DrawSubset(0)
            End If
            moEffect.EndPass()
            'moEffect.End()
            'Next lPass 

        End With

        'moEffect.EndPass()
        moEffect.End()

        moDevice.RenderState.FogEnable = mbFinalFogEnable

    End Sub

    Public Sub New()
        Device.IsUsingEventHandlers = False

        Dim oAssembly As System.Reflection.Assembly
        oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
        Dim sBase As String = AppDomain.CurrentDomain.FriendlyName
        Dim lIdx As Int32 = sBase.IndexOf("."c)
        sBase = sBase.Substring(0, lIdx)
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.NormalMap.fx")
        Device.IsUsingEventHandlers = False
        moEffect = Effect.FromStream(GFXEngine.moDevice, oStream, Nothing, ShaderFlags.None, Nothing) 'Effect.FromStream(oDevice, oStream, Nothing, ShaderFlags.None, Nothing)

        mhndlDiffuseMap = moEffect.GetParameter(Nothing, "gTex")
        mhndlNormalMap = moEffect.GetParameter(Nothing, "gNormalMap")
        mhndlIllumMap = moEffect.GetParameter(Nothing, "gIllumMap")
        mhndlSpecMap = moEffect.GetParameter(Nothing, "gSpecMap")

        mhndlEyePos = moEffect.GetParameter(Nothing, "gEyePosW")
        mhndlLightDir = moEffect.GetParameter(Nothing, "lightDir")
        mhndlLightClr = moEffect.GetParameter(Nothing, "lightColor")
        mhndlLightAmb = moEffect.GetParameter(Nothing, "lightAmbient")
        mhndlLightSpec = moEffect.GetParameter(Nothing, "lightSpecular")

        mhndlWorldInv = moEffect.GetParameter(Nothing, "gWorldInv")
        mhndlWVP = moEffect.GetParameter(Nothing, "gWVP")

        mhndlWorld = moEffect.GetParameter(Nothing, "gWorld")
        mhndlWorldInvTrans = moEffect.GetParameter(Nothing, "worldInverseTranspose")

        'mhndlEntityColor = moEffect.GetParameter(Nothing, "EntityColorAdjust")

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
        mhndlSpecMap = Nothing
        mhndlEyePos = Nothing
        mhndlLightDir = Nothing
        mhndlLightClr = Nothing
        mhndlLightAmb = Nothing
        mhndlLightSpec = Nothing
        mhndlWorldInv = Nothing       'WorldInverse
        mhndlWVP = Nothing            'WorldViewProjection
        mhndlWorld = Nothing          'World
        mhndlWorldInvTrans = Nothing  'WorldInverseTranspose
        'mhndlEntityColor = Nothing
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
        MyBase.Finalize()
    End Sub

    Public Sub DisposeMe()
        If moEffect Is Nothing = False Then moEffect.Dispose()
        moEffect = Nothing
    End Sub
End Class
