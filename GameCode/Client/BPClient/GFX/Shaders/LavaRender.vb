Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class LavaRender

    Private moEffect As Effect = Nothing

    Private moTex1 As Texture
    Private moTex2 As Texture
    Private moTex3 As Texture

    Private msTex1 As String
    Private msTex2 As String
    Private msTex3 As String
    Private mlCurrentSeed As Int32 = -1

    Private mhdlDiff1 As EffectHandle = Nothing
    Private mhdlDiff2 As EffectHandle = Nothing
    Private mhdlDiff3 As EffectHandle = Nothing

    Private mhdlWorldViewProj As EffectHandle = Nothing

    'Other values...
    Private mhdlBumpSpeedX As EffectHandle = Nothing
    Private mhdlBumpSpeedY As EffectHandle = Nothing
    Private mhdlTextReptX As EffectHandle = Nothing
    Private mhdlTextReptY As EffectHandle = Nothing
    Private mhdlCycle As EffectHandle = Nothing

    Public yCurrentPlanetType As Byte = 255
    Public lCurrentCellSpacing As Int32 = 0
    Public fWaterHeight As Single

    Private mbFogEnable As Boolean = False

    Private Shared moSW As Stopwatch

    Public Function PrepareEffect(ByVal bLavaPlanet As Boolean) As Int32
        If moSW Is Nothing Then moSW = Stopwatch.StartNew

        Dim moDevice As Device = GFXEngine.moDevice
        'set up our world to be at our camera
        Dim matWorld As Matrix = Matrix.Identity
        matWorld.Multiply(Matrix.Translation(0, fWaterHeight, 0))
        moDevice.Transform.World = matWorld

        'Now, set our parameters...
        With moEffect
            Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
            .SetValue(mhdlWorldViewProj, matWVP)
            Dim fCycle As Single = CSng(Microsoft.VisualBasic.Timer Mod 100.0F) ' CSng(((Microsoft.VisualBasic.Timer Mod 20) - 10))
            .SetValue(mhdlCycle, fCycle)
        End With

        moDevice.VertexFormat = CustomVertex.PositionTextured.Format
        moDevice.Indices = Nothing
        moDevice.SetTexture(0, Nothing)
        moDevice.RenderState.Lighting = False
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
        Dim moDevice As Device = GFXEngine.moDevice
        moDevice.RenderState.FogEnable = mbFogEnable
        moDevice.RenderState.Lighting = True
    End Sub

    Public Sub InitializeFromPlanet(ByVal lSeed As Int32)
        If mlCurrentSeed = lSeed Then Return

        mlCurrentSeed = lSeed
        Dim bUsed(4) As Boolean
        For X As Int32 = 0 To 4
            bUsed(X) = False
        Next X

        'Rnd(-1)
        'Randomize(mlCurrentSeed)
        Dim oRandom As New Random(mlCurrentSeed)

        Dim lVals(2) As Int32

        For X As Int32 = 0 To 2
            Do
                lVals(X) = CInt(oRandom.NextDouble() * 4)
            Loop Until bUsed(lVals(X)) = False
            bUsed(lVals(X)) = True
        Next X

        Dim sTexs() As String = {"L1.DDS", "L2.DDS", "L3.DDS", "L4.DDS", "L5.DDS"}
        msTex1 = sTexs(lVals(0))
        msTex2 = sTexs(lVals(1))
        msTex3 = sTexs(lVals(2))

        moTex1 = goResMgr.GetTexture(msTex1, GFXResourceManager.eGetTextureType.NoSpecifics, "pliq.pak")
        moTex2 = goResMgr.GetTexture(msTex2, GFXResourceManager.eGetTextureType.NoSpecifics, "pliq.pak")
        moTex3 = goResMgr.GetTexture(msTex3, GFXResourceManager.eGetTextureType.NoSpecifics, "pliq.pak")

        moEffect.SetValue(mhdlDiff1, moTex1)
        moEffect.SetValue(mhdlDiff2, moTex2)
        moEffect.SetValue(mhdlDiff3, moTex3)
    End Sub

    Public Sub New()
        Dim oAssembly As System.Reflection.Assembly
        oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.Lava.fx")
        Device.IsUsingEventHandlers = False
        moEffect = Effect.FromStream(GFXEngine.moDevice, oStream, Nothing, ShaderFlags.None, Nothing)
        Device.IsUsingEventHandlers = True
        GetEffectHandles()

        moEffect.Technique = moEffect.GetTechnique("textured")
    End Sub

    Private Sub GetEffectHandles()
        Device.IsUsingEventHandlers = False
        mhdlWorldViewProj = moEffect.GetParameter(Nothing, "worldViewProj")
        
        mhdlTextReptX = moEffect.GetParameter(Nothing, "TexReptX")
        mhdlTextReptY = moEffect.GetParameter(Nothing, "TexReptY")
        mhdlBumpSpeedX = moEffect.GetParameter(Nothing, "BumpSpeedX")
        mhdlBumpSpeedY = moEffect.GetParameter(Nothing, "BumpSpeedY")
        mhdlCycle = moEffect.GetParameter(Nothing, "cycle")

        mhdlDiff1 = moEffect.GetParameter(Nothing, "diffuseTexture")
        mhdlDiff2 = moEffect.GetParameter(Nothing, "diffuseTexture2")
        mhdlDiff3 = moEffect.GetParameter(Nothing, "diffuseTexture3")

        Device.IsUsingEventHandlers = True
    End Sub

    Private Sub ResetParamDefaults()
        'Now, we set the defaults for all of them, just to be safe...
        moEffect.SetValue(mhdlTextReptX, 1.0F)
        moEffect.SetValue(mhdlTextReptY, 1.0F)
        moEffect.SetValue(mhdlBumpSpeedX, -0.01F)
        moEffect.SetValue(mhdlBumpSpeedY, 0.0F)
    End Sub
 
    Public Sub DisposeMe()
        mhdlWorldViewProj = Nothing
        mhdlTextReptX = Nothing
        mhdlTextReptY = Nothing
        mhdlBumpSpeedX = Nothing
        mhdlBumpSpeedY = Nothing
        mhdlCycle = Nothing
        mhdlDiff1 = Nothing
        mhdlDiff2 = Nothing
        mhdlDiff3 = Nothing

        Try
            If moTex1 Is Nothing = False Then moTex1.Dispose()
        Catch
        End Try
        moTex1 = Nothing

        Try
            If moTex2 Is Nothing = False Then moTex2.Dispose()
        Catch
        End Try
        moTex2 = Nothing

        Try
            If moTex3 Is Nothing = False Then moTex3.Dispose()
        Catch
        End Try
        moTex3 = Nothing

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

        'Select Case lCellSpacing
        '    Case gl_LARGE_PLANET_CELL_SPACING
        '        moEffect.SetValue(mhdlTextReptX, 3.0F)
        '        moEffect.SetValue(mhdlTextReptY, 2.0F)
        '        moEffect.SetValue(mhdlBumpSpeedX, -0.015F)
        '    Case gl_HUGE_PLANET_CELL_SPACING
        '        moEffect.SetValue(mhdlTextReptX, 4.0F)
        '        moEffect.SetValue(mhdlTextReptY, 3.0F)
        '        moEffect.SetValue(mhdlBumpSpeedX, -0.02F)
        'End Select
    End Sub

    Public Sub SetForMapWrap()
        'Now, set our parameters... 
        With moEffect
            Dim moDevice As Device = GFXEngine.moDevice
            Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
            .SetValue(mhdlWorldViewProj, matWVP)
            .CommitChanges()
        End With
    End Sub
End Class
