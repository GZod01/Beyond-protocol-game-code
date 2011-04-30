Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class PostShader
    Private moEffect As Effect = Nothing

    Private moSceneMap As Texture
    Private moDownsampleMap As Texture
    Private moBlurMap1 As Texture
    Private moBlurMap2 As Texture

    Private mhdlWindowSize As EffectHandle
    Private mhdlSceneMap As EffectHandle
    Private mhdlDownsample As EffectHandle
    Private mhdlBlurMap1 As EffectHandle
    Private mhdlBlurMap2 As EffectHandle
    Private mhdlGlowIntensity As EffectHandle

    Private mfPreviousGlow As Single = -1.0F
    Private mfPreviousfWS0 As Single = 0
    Private mfPreviousfWS1 As Single = 0

    Private mfWS_Values() As Single

    Public Sub ExecutePostProcess()

        Dim moDevice As Device = GFXEngine.moDevice

        'Validate our parameters
        If mhdlSceneMap Is Nothing = False Then moEffect.SetValue(mhdlSceneMap, moSceneMap)
        If mfPreviousGlow <> muSettings.PostGlowAmt Then
            mfPreviousGlow = muSettings.PostGlowAmt
            If mhdlGlowIntensity Is Nothing = False Then moEffect.SetValue(mhdlGlowIntensity, muSettings.PostGlowAmt)
        End If
        If ValidateTextures() = False Then Return

        Dim oOriginal As Surface = moDevice.GetBackBuffer(0, 0, BackBufferType.Mono)
        Dim oScene As Surface = moSceneMap.GetSurfaceLevel(0)

        moDevice.StretchRectangle(oOriginal, New Rectangle(0, 0, moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight), _
          oScene, New System.Drawing.Rectangle(0, 0, moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight), TextureFilter.None)

        If oScene Is Nothing = False Then oScene.Dispose()
        oScene = Nothing

        'Turn off the Z buffer
        moDevice.RenderState.ZBufferEnable = False
		'moDevice.RenderState.ZBufferWriteEnable = False
        'And alpha blending/light
        moDevice.RenderState.AlphaBlendEnable = False
        moDevice.RenderState.Lighting = False

        moDevice.Indices = Nothing
        moDevice.SetTexture(0, Nothing)
        moDevice.SetTexture(1, Nothing)
        moDevice.RenderState.CullMode = Cull.None
        moDevice.RenderState.FogEnable = False

        Dim lPasses As Int32 = moEffect.Begin(FX.None)
        For X As Int32 = 0 To lPasses - 1
            If oScene Is Nothing = False Then oScene.Dispose()
            If X = 0 Then
                'downsample
                oScene = moDownsampleMap.GetSurfaceLevel(0)
                moDevice.SetRenderTarget(0, oScene)
            ElseIf X = 1 Then
                'blurmap1
                oScene = moBlurMap1.GetSurfaceLevel(0)
                moDevice.SetRenderTarget(0, oScene)
            ElseIf X = 2 Then
                'blurmap2
                oScene = moBlurMap2.GetSurfaceLevel(0)
                moDevice.SetRenderTarget(0, oScene)
            Else
                'reset render target
                moDevice.SetRenderTarget(0, oOriginal)
            End If

            Try
                moEffect.BeginPass(X)
            Catch
            End Try
            RenderQuad(moDevice)
            Try
                moEffect.EndPass()
            Catch
            End Try
        Next X
        moEffect.End() 
        If oScene Is Nothing = False Then oScene.Dispose()
        oScene = Nothing

        'Turn on the Z buffer
        moDevice.RenderState.ZBufferEnable = True
		'moDevice.RenderState.ZBufferWriteEnable = True
        'And alpha blending/light
        moDevice.RenderState.AlphaBlendEnable = True
        moDevice.RenderState.Lighting = True
        oOriginal.Dispose()
        oOriginal = Nothing
    End Sub

    Public Sub ReleaseTextures()
        If moSceneMap Is Nothing = False Then moSceneMap.Dispose()
        moSceneMap = Nothing
        If moDownsampleMap Is Nothing = False Then moDownsampleMap.Dispose()
        moDownsampleMap = Nothing
        If moBlurMap1 Is Nothing = False Then moBlurMap1.Dispose()
        moBlurMap1 = Nothing
        If moBlurMap2 Is Nothing = False Then moBlurMap2.Dispose()
        moBlurMap2 = Nothing

        If mvbQuad Is Nothing = False Then mvbQuad.Dispose()
        mvbQuad = Nothing

        Try
            If moEffect Is Nothing = False Then moEffect.Dispose()
            moEffect = Nothing
        Catch
        End Try

    End Sub

    Private Function ValidateTextures() As Boolean

        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then
            If moSceneMap Is Nothing = False Then moSceneMap.Dispose()
            moSceneMap = Nothing
            If moDownsampleMap Is Nothing = False Then moDownsampleMap.Dispose()
            moDownsampleMap = Nothing
            If moBlurMap1 Is Nothing = False Then moBlurMap1.Dispose()
            moBlurMap1 = Nothing
            If moBlurMap2 Is Nothing = False Then moBlurMap2.Dispose()
            moBlurMap2 = Nothing
            Return False
        End If

        Dim moDevice As Device = GFXEngine.moDevice

        If moSceneMap Is Nothing OrElse moSceneMap.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moSceneMap = New Texture(moDevice, moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight, 1, Usage.RenderTarget, moDevice.PresentationParameters.BackBufferFormat, Pool.Default)
            Device.IsUsingEventHandlers = True
        End If
        If moDownsampleMap Is Nothing OrElse moDownsampleMap.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moDownsampleMap = New Texture(moDevice, moDevice.PresentationParameters.BackBufferWidth \ 4, moDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, moDevice.PresentationParameters.BackBufferFormat, Pool.Default)
            If mhdlDownsample Is Nothing = False Then moEffect.SetValue(mhdlDownsample, moDownsampleMap)
            Device.IsUsingEventHandlers = True
        End If
        If moBlurMap1 Is Nothing OrElse moBlurMap1.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moBlurMap1 = New Texture(moDevice, moDevice.PresentationParameters.BackBufferWidth \ 4, moDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, moDevice.PresentationParameters.BackBufferFormat, Pool.Default)
            If mhdlBlurMap1 Is Nothing = False Then moEffect.SetValue(mhdlBlurMap1, moBlurMap1)
            Device.IsUsingEventHandlers = True
        End If
        If moBlurMap2 Is Nothing OrElse moBlurMap2.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moBlurMap2 = New Texture(moDevice, moDevice.PresentationParameters.BackBufferWidth \ 4, moDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, moDevice.PresentationParameters.BackBufferFormat, Pool.Default)
            If mhdlBlurMap2 Is Nothing = False Then moEffect.SetValue(mhdlBlurMap2, moBlurMap2)
            Device.IsUsingEventHandlers = True
        End If


        mfWS_Values(0) = moDevice.PresentationParameters.BackBufferWidth
        mfWS_Values(1) = moDevice.PresentationParameters.BackBufferHeight
        If mfPreviousfWS0 <> mfWS_Values(0) OrElse mfPreviousfWS1 <> mfWS_Values(1) Then
            mfPreviousfWS0 = mfWS_Values(0) : mfPreviousfWS1 = mfWS_Values(1)
            If mhdlWindowSize Is Nothing = False Then moEffect.SetValue(mhdlWindowSize, mfWS_Values)
        End If

        Return True
    End Function

    Public Sub New()
        Dim moDevice As Device = GFXEngine.moDevice

        ReDim mfWS_Values(1)
        Try
            Dim oAssembly As System.Reflection.Assembly
            oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
            Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.PostScreenGlow.fx")
            Device.IsUsingEventHandlers = False
            moEffect = Effect.FromStream(moDevice, oStream, Nothing, ShaderFlags.None, Nothing)
            Device.IsUsingEventHandlers = True

            mhdlWindowSize = moEffect.GetParameter(Nothing, "windowSize")
            mhdlSceneMap = moEffect.GetParameter(Nothing, "sceneMap")
            mhdlDownsample = moEffect.GetParameter(Nothing, "downsampleMap")
            mhdlBlurMap1 = moEffect.GetParameter(Nothing, "blurMap1")
            mhdlBlurMap2 = moEffect.GetParameter(Nothing, "blurMap2")
            mhdlGlowIntensity = moEffect.GetParameter(Nothing, "GlowIntensity")

            moEffect.Technique = "ScreenGlow"
        Catch
            muSettings.PostGlowAmt = 0
        End Try
    End Sub

    Private Shared mvbQuad As VertexBuffer = Nothing
    Private Shared Sub RenderQuad(ByRef oDevice As Device)
        If mvbQuad Is Nothing OrElse mvbQuad.Disposed = True Then
            If InitVBQuad(oDevice) = False Then Return
        End If
        Try
            oDevice.VertexFormat = CustomVertex.PositionTextured.Format
            oDevice.SetStreamSource(0, mvbQuad, 0, CustomVertex.PositionTextured.StrideSize)
            oDevice.DrawPrimitives(PrimitiveType.TriangleStrip, 0, 2)
        Catch
        End Try
    End Sub
    Private Shared Function InitVBQuad(ByRef oDevice As Device) As Boolean
        Dim uVert As CustomVertex.PositionTextured
        If mvbQuad Is Nothing OrElse mvbQuad.Disposed = True Then mvbQuad = New VertexBuffer(uVert.GetType, 4, oDevice, Usage.WriteOnly, CustomVertex.PositionTextured.Format, Pool.Managed)
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

    Public Sub DisposeMe()
        mhdlWindowSize = Nothing
        mhdlSceneMap = Nothing
        mhdlDownsample = Nothing
        mhdlBlurMap1 = Nothing
        mhdlBlurMap2 = Nothing
        mhdlGlowIntensity = Nothing

        Try
            If mvbQuad Is Nothing = False Then mvbQuad.Dispose()
            mvbQuad = Nothing
        Catch
        End Try

        Try
            If moSceneMap Is Nothing = False Then moSceneMap.Dispose()
        Catch
        End Try
        moSceneMap = Nothing
        Try
            If moDownsampleMap Is Nothing = False Then moDownsampleMap.Dispose()
        Catch
        End Try
        moDownsampleMap = Nothing
        Try
            If moBlurMap1 Is Nothing = False Then moBlurMap1.Dispose()
        Catch
        End Try
        moBlurMap1 = Nothing
        Try
            If moBlurMap2 Is Nothing = False Then moBlurMap2.Dispose()
        Catch
        End Try
        moBlurMap2 = Nothing
        Try
            If moEffect Is Nothing = False Then moEffect.Dispose()
        Catch
        End Try
        moEffect = Nothing
    End Sub
End Class
 