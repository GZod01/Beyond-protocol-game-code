Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class CosmoShader

    Private moEffect As Effect = Nothing

    'Texture Samplers
    Private mhndlDiffuseMap As EffectHandle
    Private mhndlLightMap As EffectHandle
    Private mhndlWVP As EffectHandle            'WorldViewProjection
    Private mhndlGamma As EffectHandle

    Public Sub BeginRender(ByRef texDiffuse As Texture, ByRef texLight As Texture, ByVal fGamma As Single)
        'mlRenderFlags = 0

        Dim moDevice As Device = GFXEngine.moDevice
        
        'Now, determine what technique to use...
        moEffect.Technique = "textured"

        moEffect.SetValue(mhndlDiffuseMap, texDiffuse)
        moEffect.SetValue(mhndlLightMap, texLight)
        moEffect.SetValue(mhndlGamma, fGamma)
        Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
        moEffect.SetValue(mhndlWVP, matWVP)
        moEffect.CommitChanges()

        Dim lPasses As Int32 = moEffect.Begin(FX.None)
        moEffect.BeginPass(0)

    End Sub

    Public Sub EndRender()
        moEffect.EndPass()
        moEffect.End()
        GFXEngine.moDevice.RenderState.CullMode = Cull.CounterClockwise
    End Sub

    Public Sub New()
        Device.IsUsingEventHandlers = False

        Dim oAssembly As System.Reflection.Assembly
        oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
        Dim sBase As String = AppDomain.CurrentDomain.FriendlyName
        Dim lIdx As Int32 = sBase.IndexOf("."c)
        sBase = sBase.Substring(0, lIdx)
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.Cosmos.fx")
        Device.IsUsingEventHandlers = False
        moEffect = Effect.FromStream(GFXEngine.moDevice, oStream, Nothing, ShaderFlags.None, Nothing)

        mhndlDiffuseMap = moEffect.GetParameter(Nothing, "diffuseTexture")
        mhndlLightMap = moEffect.GetParameter(Nothing, "LightTexture")
        mhndlWVP = moEffect.GetParameter(Nothing, "worldViewProj")

        mhndlGamma = moEffect.GetParameter(Nothing, "GammaBright")

        Device.IsUsingEventHandlers = True
    End Sub

    Protected Overrides Sub Finalize()
        mhndlDiffuseMap = Nothing
        mhndlLightMap = Nothing
        mhndlWVP = Nothing
        mhndlGamma = Nothing
        
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
