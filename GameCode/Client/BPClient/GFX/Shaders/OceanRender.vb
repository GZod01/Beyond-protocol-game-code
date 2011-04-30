Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class OceanRender

	Private moEffect As Effect = Nothing

	Private moNormalMap As Texture = Nothing
	Public Shared bRefreshTexture As Boolean = False

	Private mhdlWorldViewProj As EffectHandle = Nothing
	Private mhdlWorld As EffectHandle = Nothing
    'Private mhdlInverseView As EffectHandle = Nothing
	Private mhdlTimer As EffectHandle = Nothing
	Private mhdlLightDir As EffectHandle = Nothing
	Private mhdlNormalTexture As EffectHandle = Nothing
	Private mhdlBumpHeight As EffectHandle = Nothing
	Private mhdlShininess As EffectHandle = Nothing
	Private mhdlTextReptX As EffectHandle = Nothing
	Private mhdlTextReptY As EffectHandle = Nothing
	Private mhdlBumpSpeedX As EffectHandle = Nothing
	Private mhdlBumpSpeedY As EffectHandle = Nothing

	Private mhdlSpecularColor As EffectHandle = Nothing
	Private mhdlDeepColor As EffectHandle = Nothing
	Private mhdlShallowColor As EffectHandle = Nothing
	Private mhdlWaterClrStr As EffectHandle = Nothing

	Private mhdlWaterAlpha As EffectHandle = Nothing

	Private mhdlFogEnable As EffectHandle = Nothing
	Private mhdlFogStart As EffectHandle = Nothing
	Private mhdlFogRange As EffectHandle = Nothing
    Private mhdlFogColor As EffectHandle = Nothing

    Private mhdlEyePos As EffectHandle = Nothing

    Private mhdlCycle As EffectHandle = Nothing


	Public yCurrentPlanetType As Byte = 255
	Public lCurrentCellSpacing As Int32 = 0
	Public fWaterHeight As Single

	Private mfSpecR As Single = 1.0F
	Private mfSpecG As Single = 1.0F
	Private mfSpecB As Single = 1.0F

	Private Shared moSW As Stopwatch

    Private mfBaseWaterAlpha As Single
    Private mfCurrentWaterAlpha As Single

    Private mbFogEnable As Boolean = False

    Public Function PrepareEffect(ByVal bLavaPlanet As Boolean) As Int32
        If moSW Is Nothing Then moSW = Stopwatch.StartNew

        Dim moDevice As Device = GFXEngine.moDevice

        'Validate our parameters
        ValidateTextures()

        Dim fTarget As Single = mfBaseWaterAlpha
        If goUILib Is Nothing = False Then
            If goUILib.BuildGhost Is Nothing = False AndAlso goUILib.bBuildGhostNaval = True Then
                fTarget = 0.25F
            End If
        End If
        If mfCurrentWaterAlpha > fTarget Then
            mfCurrentWaterAlpha -= 0.05F
            If mfCurrentWaterAlpha < fTarget Then mfCurrentWaterAlpha = fTarget
            moEffect.SetValue(mhdlWaterAlpha, mfCurrentWaterAlpha)
        ElseIf mfCurrentWaterAlpha < fTarget Then
            mfCurrentWaterAlpha += 0.05F
            If mfCurrentWaterAlpha > fTarget Then mfCurrentWaterAlpha = fTarget
            moEffect.SetValue(mhdlWaterAlpha, mfCurrentWaterAlpha)
        End If

        'set up our world to be at our camera
        Dim matWorld As Matrix = Matrix.Identity
        matWorld.Multiply(Matrix.Translation(0, fWaterHeight, 0))
        moDevice.Transform.World = matWorld

        'Now, set our parameters...
        With moEffect
            Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
            .SetValue(mhdlWorldViewProj, matWVP)
            .SetValue(mhdlWorld, moDevice.Transform.World)
            '.SetValue(mhdlInverseView, Matrix.Invert(moDevice.Transform.View))
            .SetValue(mhdlTimer, CSng(moSW.ElapsedTicks / Stopwatch.Frequency))
        End With

        moDevice.VertexFormat = CustomVertex.PositionTextured.Format
        moDevice.Indices = Nothing
        moDevice.SetTexture(0, Nothing)
        moDevice.RenderState.Lighting = False

        Dim vecEyePos As Vector4 = New Vector4(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ, 0.0F)
        'vecEyePos.Normalize()
        moEffect.SetValue(mhdlEyePos, vecEyePos)

        If bLavaPlanet = False Then
            moEffect.SetValue(mhdlFogEnable, moDevice.RenderState.FogEnable)
        Else
            moEffect.SetValue(mhdlFogEnable, False)
        End If
        moEffect.SetValue(mhdlFogStart, moDevice.RenderState.FogStart) ' * 0.0001F)
        moEffect.SetValue(mhdlFogRange, moDevice.RenderState.FogEnd) ' * 0.0001F)
        With moDevice.RenderState.FogColor
            moEffect.SetValue(mhdlFogColor, New Vector4(.R / 255.0F, .G / 255.0F, .B / 255.0F, 1.0F))
        End With

        Dim fCycle As Single = CSng(Microsoft.VisualBasic.Timer Mod 100.0F) ' CSng(((Microsoft.VisualBasic.Timer Mod 20) - 10))
        moEffect.SetValue(mhdlCycle, fCycle)

        mbFogEnable = moDevice.RenderState.FogEnable
        moDevice.RenderState.FogEnable = False

        Return moEffect.Begin(FX.None)
    End Function
    Public Sub StartPass(ByVal lIdx As Int32)
        moEffect.BeginPass(lIdx)
    End Sub
    Public Sub StopPass()
        moEffect.EndPass()
    End Sub
    Public Sub StopEffect()
        moEffect.End()
        GFXEngine.moDevice.RenderState.FogEnable = mbFogEnable
        GFXEngine.moDevice.RenderState.Lighting = True
    End Sub

    Private Sub ValidateTextures()
        If bRefreshTexture = True OrElse moNormalMap Is Nothing OrElse moNormalMap.Disposed = True Then
            Device.IsUsingEventHandlers = False
            'moNormalMap = TextureLoader.FromFile(moDevice, "C:\Program Files\NVIDIA Corporation\SDK 9.5\MEDIA\textures\2D\waves2.dds")
            moNormalMap = goResMgr.GetTexture("waves2.dds", GFXResourceManager.eGetTextureType.WaterTexture, "pring.pak")
            If mhdlNormalTexture Is Nothing = False Then moEffect.SetValue(mhdlNormalTexture, moNormalMap)
            Device.IsUsingEventHandlers = True
        End If
    End Sub

    Public Sub New()
        Dim oAssembly As System.Reflection.Assembly
        oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.Ocean.fx")
        Device.IsUsingEventHandlers = False
        moEffect = Effect.FromStream(GFXEngine.moDevice, oStream, Nothing, ShaderFlags.None, Nothing)
        'moEffect = Effect.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\ShaderLibrary\Ocean\HLSL\Ocean.fx", Nothing, ShaderFlags.None, Nothing, "")
        Device.IsUsingEventHandlers = True
        GetEffectHandles()

        moEffect.Technique = moEffect.GetTechnique("Main")
    End Sub

    Private Sub GetEffectHandles()
        Device.IsUsingEventHandlers = False
        mhdlWorldViewProj = moEffect.GetParameter(Nothing, "WvpXf")
        mhdlWorld = moEffect.GetParameter(Nothing, "WorldXf")
        'mhdlInverseView = moEffect.GetParameter(Nothing, "ViewIXf")
        mhdlTimer = moEffect.GetParameter(Nothing, "Timer")
        mhdlLightDir = moEffect.GetParameter(Nothing, "lightDir")
        mhdlNormalTexture = moEffect.GetParameter(Nothing, "NormalTexture")
        mhdlBumpHeight = moEffect.GetParameter(Nothing, "BumpScale")
        mhdlShininess = moEffect.GetParameter(Nothing, "Shininess")
        mhdlTextReptX = moEffect.GetParameter(Nothing, "TexReptX")
        mhdlTextReptY = moEffect.GetParameter(Nothing, "TexReptY")
        mhdlBumpSpeedX = moEffect.GetParameter(Nothing, "BumpSpeedX")
        mhdlBumpSpeedY = moEffect.GetParameter(Nothing, "BumpSpeedY")
        mhdlSpecularColor = moEffect.GetParameter(Nothing, "SpecularColor")
        mhdlDeepColor = moEffect.GetParameter(Nothing, "DeepColor")
        mhdlShallowColor = moEffect.GetParameter(Nothing, "ShallowColor")
        mhdlWaterClrStr = moEffect.GetParameter(Nothing, "KWater")
        mhdlWaterAlpha = moEffect.GetParameter(Nothing, "WaterAlpha")

        mhdlFogEnable = moEffect.GetParameter(Nothing, "gbRenderFog")
        mhdlFogStart = moEffect.GetParameter(Nothing, "gFogStart")
        mhdlFogRange = moEffect.GetParameter(Nothing, "gFogRange")
        mhdlFogColor = moEffect.GetParameter(Nothing, "gFogColor")

        mhdlEyePos = moEffect.GetParameter(Nothing, "gEyePosW")

        mhdlCycle = moEffect.GetParameter(Nothing, "cycle")

        Device.IsUsingEventHandlers = True
    End Sub

    Private Sub ResetParamDefaults()
        'Now, we set the defaults for all of them, just to be safe...
        mfSpecR = 1.0F : mfSpecG = 1.0F : mfSpecB = 1.0F
        moEffect.SetValue(mhdlLightDir, New Vector4(0.58F, -0.58F, 0.58F, 0.0F))
        moEffect.SetValue(mhdlBumpHeight, 5.0F)
        moEffect.SetValue(mhdlShininess, 10.0F)
        moEffect.SetValue(mhdlTextReptX, 1.0F)
        moEffect.SetValue(mhdlTextReptY, 1.0F)
        moEffect.SetValue(mhdlBumpSpeedX, -0.001F)
        moEffect.SetValue(mhdlBumpSpeedY, 0.0F)
        moEffect.SetValue(mhdlSpecularColor, New ColorValue(1.0F, 1.0F, 1.0F))
        moEffect.SetValue(mhdlDeepColor, New ColorValue(0.0F, 0.0F, 0.1F))
        moEffect.SetValue(mhdlShallowColor, New ColorValue(0.0F, 0.5F, 0.5F))
        moEffect.SetValue(mhdlWaterClrStr, 1.0F)
        moEffect.SetValue(mhdlWaterAlpha, 0.7F)
        mfBaseWaterAlpha = 0.7F
        mfCurrentWaterAlpha = mfBaseWaterAlpha
    End Sub

    Private Shared mvbQuad As VertexBuffer = Nothing
    Private Shared Sub RenderQuad(ByRef oDevice As Device)
        If mvbQuad Is Nothing OrElse mvbQuad.Disposed = True Then
            If InitVBQuad(oDevice) = False Then Return
        End If
        oDevice.VertexFormat = CustomVertex.PositionTextured.Format
        oDevice.SetStreamSource(0, mvbQuad, 0, CustomVertex.PositionTextured.StrideSize)
        oDevice.DrawPrimitives(PrimitiveType.TriangleStrip, 0, 2)
    End Sub
    Private Shared Function InitVBQuad(ByRef oDevice As Device) As Boolean
        Dim uVert As CustomVertex.PositionTextured
        If mvbQuad Is Nothing OrElse mvbQuad.Disposed = True Then mvbQuad = New VertexBuffer(uVert.GetType, 4, oDevice, Usage.WriteOnly, CustomVertex.PositionTextured.Format, Pool.Managed)
        Dim bResult As Boolean = False

        Try
            Dim gfxStream As GraphicsStream = mvbQuad.Lock(0, 0, 0)
            Dim uVerts(3) As CustomVertex.PositionTextured
            uVerts(0) = New CustomVertex.PositionTextured(New Vector3(-6000.0F, 0, -6000), 0, 1)
            uVerts(1) = New CustomVertex.PositionTextured(New Vector3(-6000.0F, 0, 6000), 0, 0)
            uVerts(2) = New CustomVertex.PositionTextured(New Vector3(6000, 0, -6000), 1, 1)
            uVerts(3) = New CustomVertex.PositionTextured(New Vector3(6000, 0, 6000.0F), 1, 0)
            gfxStream.Write(uVerts)
            mvbQuad.Unlock()
            bResult = True
        Catch ex As Exception
            bResult = False
        End Try
        Return bResult
    End Function

    Public Sub DisposeMe()
        mhdlWorldViewProj = Nothing
        mhdlWorld = Nothing
        'mhdlInverseView = Nothing
        mhdlTimer = Nothing
        mhdlLightDir = Nothing
        mhdlNormalTexture = Nothing
        mhdlBumpHeight = Nothing
        mhdlShininess = Nothing
        mhdlTextReptX = Nothing
        mhdlTextReptY = Nothing
        mhdlBumpSpeedX = Nothing
        mhdlBumpSpeedY = Nothing
        mhdlSpecularColor = Nothing
        mhdlDeepColor = Nothing
        mhdlShallowColor = Nothing
        mhdlWaterClrStr = Nothing
        mhdlWaterAlpha = Nothing
        mhdlEyePos = Nothing
        mhdlCycle = Nothing
        Try
            If moNormalMap Is Nothing = False Then moNormalMap.Dispose()
        Catch
        End Try
        moNormalMap = Nothing

        Try
            If moEffect Is Nothing = False Then moEffect.Dispose()
        Catch
        End Try
        moEffect = Nothing
    End Sub

    Public Sub SetAsPlanetType(ByVal yPlanetType As Byte, ByVal lCellSpacing As Int32)
        ResetParamDefaults()

        Me.yCurrentPlanetType = yPlanetType
        Me.lCurrentCellSpacing = lCellSpacing

        Select Case yPlanetType
            Case PlanetType.eAcidic
                mfSpecR = (32 / 255.0F)
                mfSpecG = (64 / 255.0F)
                mfSpecB = (32 / 255.0F)
                moEffect.SetValue(mhdlSpecularColor, New ColorValue(mfSpecR, mfSpecG, mfSpecB))
                moEffect.SetValue(mhdlDeepColor, New ColorValue(32, 128, 32))
                moEffect.SetValue(mhdlShallowColor, New ColorValue(16, 72, 16))
                moEffect.SetValue(mhdlBumpHeight, 0.75F)
            Case PlanetType.eAdaptable
                moEffect.SetValue(mhdlDeepColor, New ColorValue(32, 32, 92))
                moEffect.SetValue(mhdlShallowColor, New ColorValue(64, 64, 92))
                mfBaseWaterAlpha = 0.4F
                moEffect.SetValue(mhdlWaterAlpha, 0.4F)
            Case PlanetType.eDesert
                moEffect.SetValue(mhdlDeepColor, New ColorValue(0, 32, 32))
                moEffect.SetValue(mhdlShallowColor, New ColorValue(24, 48, 64))
                moEffect.SetValue(mhdlWaterAlpha, 0.9F)
                mfBaseWaterAlpha = 0.9F
            Case PlanetType.eGeoPlastic
                mfSpecR = 0 : mfSpecG = 0 : mfSpecB = 0
                moEffect.SetValue(mhdlSpecularColor, New ColorValue(0, 0, 0))
                moEffect.SetValue(mhdlDeepColor, New ColorValue(255, 0, 0))
                moEffect.SetValue(mhdlShallowColor, New ColorValue(32, 16, 0))
                moEffect.SetValue(mhdlWaterAlpha, 1.0F)
                mfBaseWaterAlpha = 1.0F
            Case PlanetType.eTerran
                moEffect.SetValue(mhdlDeepColor, New ColorValue(0, 0, 25))
                moEffect.SetValue(mhdlShallowColor, New ColorValue(0, 72, 92))
            Case PlanetType.eTundra
                moEffect.SetValue(mhdlDeepColor, New ColorValue(0, 12, 25))
                moEffect.SetValue(mhdlShallowColor, New ColorValue(0, 192, 192))
            Case PlanetType.eWaterWorld
                moEffect.SetValue(mhdlDeepColor, New ColorValue(0, 32, 72))
                moEffect.SetValue(mhdlShallowColor, New ColorValue(32, 64, 92))
                moEffect.SetValue(mhdlWaterAlpha, 0.6F)
                mfBaseWaterAlpha = 0.6F
        End Select

        Select Case lCellSpacing
            Case gl_LARGE_PLANET_CELL_SPACING
                moEffect.SetValue(mhdlTextReptX, 2.0F)
                moEffect.SetValue(mhdlTextReptY, 1.0F)
                moEffect.SetValue(mhdlBumpSpeedX, -0.005F)
            Case gl_HUGE_PLANET_CELL_SPACING
                moEffect.SetValue(mhdlTextReptX, 2.0F)
                moEffect.SetValue(mhdlTextReptY, 2.0F)
                moEffect.SetValue(mhdlBumpSpeedX, -0.01F)
        End Select
        mfCurrentWaterAlpha = mfBaseWaterAlpha
    End Sub
    Public Function SetFromLightVec(ByVal vecDir As Vector3) As Single
        Dim fResult As Single = 0.0F
        moEffect.SetValue(mhdlLightDir, New Vector4(vecDir.X, vecDir.Y, vecDir.Z, 0))
        'Now, affect our light...
        If vecDir.Y > 0 Then
            Dim vecTmp As Vector3 = vecDir
            vecTmp.Normalize()
            fResult = Math.Min(1, Math.Abs(vecTmp.Y) * 3)
        Else : fResult = 0
        End If
        Return fResult
    End Function
	Public Sub SetWaterClrStr(ByVal fValue As Single)
		moEffect.SetValue(mhdlWaterClrStr, fValue)
	End Sub
	Public Sub AttenuateSpecular(ByVal fMult As Single)
		Dim newVal As ColorValue = New ColorValue(mfSpecR * fMult, mfSpecG * fMult, mfSpecB * fMult)
		moEffect.SetValue(mhdlSpecularColor, newVal)
	End Sub

    Public Sub SetForMapWrap(ByVal bXGrtr As Boolean, ByVal lFullSize As Int32)
        'Now, set our parameters... 
        With moEffect
            Dim moDevice As Device = GFXEngine.moDevice
            Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
            .SetValue(mhdlWorldViewProj, matWVP)
            .SetValue(mhdlWorld, moDevice.Transform.World)
            '.SetValue(mhdlInverseView, Matrix.Invert(moDevice.Transform.View))

            Dim vecEyePos As Vector4
            If bXGrtr = True Then
                'on left edge
                vecEyePos = New Vector4(goCamera.mlCameraX + lFullSize, goCamera.mlCameraY, goCamera.mlCameraZ, 0.0F)
            Else
                'on right edge
                vecEyePos = New Vector4(goCamera.mlCameraX - lFullSize, goCamera.mlCameraY, goCamera.mlCameraZ, 0.0F)
            End If
            moEffect.SetValue(mhdlEyePos, vecEyePos)



            .CommitChanges()
        End With
    End Sub
End Class
