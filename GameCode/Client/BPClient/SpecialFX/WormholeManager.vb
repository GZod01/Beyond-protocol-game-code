Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class WormholeManager
    Protected Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32
    Private Structure Particle
        Public vecLoc As Vector3
        Public vecAcc As Vector3
        Public vecSpeed As Vector3

        Public mfA As Single
        Public mfR As Single
        Public mfG As Single
        Public mfB As Single

        Public fAChg As Single
        Public fRChg As Single
        Public fGChg As Single
        Public fBChg As Single

        Public fAngle As Single
        Public fAngleChg As Single

        Public ParticleColor As System.Drawing.Color

        Public ParticleActive As Boolean

        Public Sub Update(ByVal fElapsedTime As Single)
            vecLoc.Add(Vector3.Multiply(vecSpeed, fElapsedTime))
            vecSpeed.Add(Vector3.Multiply(vecAcc, fElapsedTime))

            mfA += (fAChg * fElapsedTime)
            mfR += (fRChg * fElapsedTime)
            mfG += (fGChg * fElapsedTime)
            mfB += (fBChg * fElapsedTime)

            If mfA < 0 Then mfA = 0
            If mfA > 255 Then mfA = 255
            If mfR < 0 Then mfR = 0
            If mfR > 255 Then mfR = 255
            If mfG < 0 Then mfG = 0
            If mfG > 255 Then mfG = 255
            If mfB < 0 Then mfB = 0
            If mfB > 255 Then mfB = 255

            Me.fAngle += (Me.fAngleChg * fElapsedTime)
            If Me.fAngle > 360 Then fAngle = 0
            If Me.fAngle < 0 Then fAngle = 360

            ParticleColor = System.Drawing.Color.FromArgb(CInt(mfA), CInt(mfR), CInt(mfG), CInt(mfB))
        End Sub

        Public Sub Reset(ByVal fX As Single, ByVal fY As Single, ByVal fZ As Single, ByVal fXSpeed As Single, ByVal fYSpeed As Single, ByVal fZSpeed As Single, ByVal fXAcc As Single, ByVal fYAcc As Single, ByVal fZAcc As Single, ByVal fR As Single, ByVal fG As Single, ByVal fB As Single, ByVal fA As Single)
            vecLoc.X = fX : vecLoc.Y = fY : vecLoc.Z = fZ
            vecAcc.X = fXAcc : vecAcc.Y = fYAcc : vecAcc.Z = fZAcc
            vecSpeed.X = fXSpeed : vecSpeed.Y = fYSpeed : vecSpeed.Z = fZSpeed
            mfA = fA : mfR = fR : mfG = fG : mfB = fB

            ParticleActive = True
        End Sub
    End Structure

    Private Class WormholeEffect
        Public oParentWormhole As Wormhole
        Public vecEmitter As Vector3    'location of this wormhole effect (for an inverted effect, the wormhole is -200
        Public oCullBox As CullBox
        Public bCullObj As Boolean
        Private moParticles() As Particle
        Public mlParticleUB As Int32
        Private mbEmitterStopping As Boolean = False
        Private EmitterStopped As Boolean = False

        Public moPoints() As CustomVertex.PositionColoredTextured

        Private Const ml_WORMHOLE_RADIUS As Int32 = 200
        Private mlCutoff As Int32
        Private mbInverted As Boolean = False

        Public Sub New(ByVal oVecLoc As Vector3, ByVal lParticleCnt As Int32, ByVal bInverted As Boolean)
            Dim X As Int32

            vecEmitter = oVecLoc
            mlParticleUB = lParticleCnt - 1
            mlCutoff = mlParticleUB \ 5
            mbInverted = bInverted

            ReDim moParticles(mlParticleUB)
            ReDim moPoints(mlParticleUB)

            For X = 0 To mlParticleUB
                moParticles(X) = New Particle()
                ResetParticle(X)
            Next X

            oCullBox = CullBox.GetCullBox(vecEmitter.X, vecEmitter.Y, vecEmitter.Z, vecEmitter.X - ml_WORMHOLE_RADIUS, vecEmitter.Y - ml_WORMHOLE_RADIUS, vecEmitter.Z - ml_WORMHOLE_RADIUS, vecEmitter.X + ml_WORMHOLE_RADIUS, vecEmitter.Y + ml_WORMHOLE_RADIUS, vecEmitter.Z + ml_WORMHOLE_RADIUS)
        End Sub

        Public Sub StopEmitter()
            mbEmitterStopping = True
        End Sub

        Public Sub StartEmitter()
            Dim X As Int32
            mbEmitterStopping = False
            For X = 0 To mlParticleUB
                moParticles(X).ParticleActive = True
                ResetParticle(X)
            Next X
        End Sub

        Private Sub ResetParticle(ByVal lIndex As Int32)
            Dim lX As Int32

            If mbEmitterStopping = True Then
                moParticles(lIndex).mfA = 0
                moParticles(lIndex).ParticleActive = False

                EmitterStopped = True
                For lX = 0 To mlParticleUB
                    If moParticles(lIndex).ParticleActive = True Then
                        EmitterStopped = False
                        Exit For
                    End If
                Next lX
            Else

                If lIndex < mlCutoff Then
                    Dim fY As Single = -(ml_WORMHOLE_RADIUS / 2)
                    Dim fYS As Single
                    Dim fYA As Single
                    Dim fX As Single = Rnd() * ml_WORMHOLE_RADIUS / 10.0F
                    Dim fZ As Single = Rnd() * ml_WORMHOLE_RADIUS / 10.0F
                    Dim fXS As Single = Rnd() * 0.01F
                    Dim fZS As Single = Rnd() * 0.01F
                    fYS = (Rnd() * 0.01F) '+ 0.005F
                    fYA = (Rnd() * 0.1F) '+ 0.05F

                    If mbInverted = True Then
                        fYS = -fYS
                        fY = -fY
                        fYA = -fYA
                    End If

                    'fX += vecEmitter.X
                    'fY += vecEmitter.Y
                    'fZ += vecEmitter.Z

                    Dim fR As Single = 255
                    Dim fG As Single = 255
                    Dim fB As Single = 255
                    'fR = 220 + (GetNxtRnd() * 30) : fG = GetNxtRnd() * 64 + 32 : fB = GetNxtRnd() * 64 + 32
                    Select Case lIndex Mod 5
                        Case 0
                            fR = 220 + (Rnd() * 30) : fG = Rnd() * 64 + 32 : fB = Rnd() * 64 + 32
                        Case 1
                            fR = Rnd() * 64 + 32 : fG = 220 + (Rnd() * 30) : fB = Rnd() * 64 + 32
                        Case 2
                            fR = Rnd() * 64 + 32 : fG = Rnd() * 64 + 32 : fB = 220 + (Rnd() * 30)
                        Case 3
                            fR = Rnd() * 64 + 32 : fG = 220 + (Rnd() * 30) : fB = 220 + (Rnd() * 30)
                        Case 4
                            fR = 220 + (Rnd() * 30) : fG = Rnd() * 64 + 32 : fB = 220 + (Rnd() * 30)
                        Case 5
                            fR = 220 + (Rnd() * 30) : fG = 220 + (Rnd() * 30) : fB = 220 + (Rnd() * 30)
                    End Select
                    With moParticles(lIndex)
                        .Reset(fX, fY, fZ, fXS, fYS, fZS, 0, fYA, 0, fR, fG, fB, 0)
                        .mfA = 255
                        .fAChg = Math.Min(((2 * Rnd()) + 2) * -1, -1)
                        .ParticleColor = System.Drawing.Color.FromArgb(0, 0, 0, 0)
                        moParticles(lIndex).fAngle = 0
                        moParticles(lIndex).fAngleChg = (Rnd() * 3.0F) - 1.5F
                    End With
                Else
                    Dim fAngle As Single = Rnd() * 360.0F

                    Dim fX As Single = ((ml_WORMHOLE_RADIUS \ 10) * 9) + (Rnd() * (ml_WORMHOLE_RADIUS \ 10))
                    Dim fY As Single = 0
                    Dim fZ As Single = 0

                    Dim fXA As Single = 0 '(GetNxtRnd() * 2) - 1
                    Dim fYA As Single = Rnd() * -0.01F
                    Dim fZA As Single = 0 '(GetNxtRnd() * 2) - 1

                    Dim fXS As Single
                    Dim fYS As Single = -1.0F
                    Dim fZS As Single

                    RotatePoint(0, 0, fX, fZ, fAngle)

                    fXS = Rnd() + 2
                    fZS = 0
                    RotatePoint(0, 0, fXS, fZS, fAngle - 180.0F)

                    If mbInverted = True Then
                        fYS = -fYS
                        fY = -fY
                        fYA = -fYA
                    End If

                    'fX += vecEmitter.X
                    'fY += vecEmitter.Y
                    'fZ += vecEmitter.Z

                    Dim fR As Single = 255
                    Dim fG As Single = 255
                    Dim fB As Single = 255
                    Select Case lIndex Mod 5
                        Case 0
                            fR = 220 + (Rnd() * 30) : fG = Rnd() * 64 + 32 : fB = Rnd() * 64 + 32
                        Case 1
                            fR = Rnd() * 64 + 32 : fG = 220 + (Rnd() * 30) : fB = Rnd() * 64 + 32
                        Case 2
                            fR = Rnd() * 64 + 32 : fG = Rnd() * 64 + 32 : fB = 220 + (Rnd() * 30)
                        Case 3
                            fR = Rnd() * 64 + 32 : fG = 220 + (Rnd() * 30) : fB = 220 + (Rnd() * 30)
                        Case 4
                            fR = 220 + (Rnd() * 30) : fG = Rnd() * 64 + 32 : fB = 220 + (Rnd() * 30)
                        Case 5
                            fR = 220 + (Rnd() * 30) : fG = 220 + (Rnd() * 30) : fB = 220 + (Rnd() * 30)
                    End Select

                    'moParticles(lIndex).Reset(fX, fY, fZ, fXS, -1, fZS, fXA, fYA, fZA, 128 + (GetNxtRnd() * 128), 128 + (GetNxtRnd() * 128), 128 + (GetNxtRnd() * 128), 0)
                    moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, fR, fG, fB, 0)
                    moParticles(lIndex).mfA = 0
                    moParticles(lIndex).fAChg = (2 * Rnd()) + 1
                    moParticles(lIndex).ParticleColor = System.Drawing.Color.FromArgb(0, 0, 0, 0)

                    moParticles(lIndex).fAngle = 0
                    moParticles(lIndex).fAngleChg = (Rnd() * 3.0F) - 1.5F


                    'moParticles(lIndex).Reset(vecEmitter.X + ((GetNxtRnd() * 10) - 5), vecEmitter.Y, vecEmitter.Z + ((GetNxtRnd() * 10) - 5), -0.4 + (GetNxtRnd() * 0.8), Rnd, -0.4 + (GetNxtRnd() * 0.8), 0, (GetNxtRnd() * 0.03), 0, 255, 128, 50, (0.6 + (0.2 * GetNxtRnd())) * 255)
                    'moParticles(lIndex).fAChg = -((0.01 + GetNxtRnd() * 0.05) * 255)
                End If


            End If


        End Sub

        Public Sub Update(ByVal fElapsed As Single)
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            Dim fX As Single
            Dim fZ As Single

            If EmitterStopped = True Then Exit Sub

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    .Update(fElapsed)
                    'If X = 399 Then Stop
                    If (X > mlCutoff AndAlso .mfA >= 255) OrElse (X <= mlCutoff AndAlso .mfA <= 0) Then
                        .mfA = 0
                        ResetParticle(X)
                    Else

                        If (.vecSpeed.X < 0 AndAlso .vecLoc.X < 0) OrElse (.vecSpeed.X > 0 AndAlso .vecLoc.X > 0) Then
                            .vecSpeed.X = 0
                            '.vecAcc.Y *= 200
                            '.fAChg *= 2
                        End If

                        If (.vecSpeed.Z < 0 AndAlso .vecLoc.Z < 0) OrElse (.vecSpeed.Z > 0 AndAlso .vecLoc.Z > 0) Then
                            .vecSpeed.Z = 0
                            '.vecAcc.Y *= 200
                            '.fAChg *= 2
                        End If
                        'If X >= mlCutoff Then
                        If mbInverted = False Then
                            If .vecLoc.Y < -(ml_WORMHOLE_RADIUS / 2) Then
                                .vecAcc.Y = 0
                                .vecSpeed.Y = 0
                            End If
                        Else
                            If .vecLoc.Y > (ml_WORMHOLE_RADIUS / 2) Then
                                .vecAcc.Y = 0
                                .vecSpeed.Y = 0
                            End If
                        End If
                        'End If
                    End If

                    moPoints(X).Color = .ParticleColor.ToArgb

                    fX = .vecLoc.X
                    Dim fY As Single = .vecLoc.Y
                    fZ = .vecLoc.Z

                    RotatePoint(0, 0, fX, fZ, .fAngle)
                    fX += vecEmitter.X
                    fY += vecEmitter.Y
                    fZ += vecEmitter.Z
                    moPoints(X).Position = New Vector3(fX, fY, fZ) 

                End With
            Next X

        End Sub
    End Class

    Private moWormholes() As WormholeEffect
    Private mlWormholeUB As Int32 = -1
    Private mlPrevFrame As Int32
    Private moParticle As Texture

    Private muWHVerts() As CustomVertex.PositionColoredTextured
    Private moWormholeCloudTex As Texture

    Public lCurrentSystemID As Int32 = -1

    Public Sub ClearAllWormholes()
        ReDim moWormholes(-1)
        mlWormholeUB = -1
    End Sub
    Public Sub AddWormhole(ByRef oWormhole As Wormhole, ByVal lSystemID As Int32)
        For X As Int32 = 0 To mlWormholeUB
            If moWormholes(X) Is Nothing = False AndAlso moWormholes(X).oParentWormhole.ObjectID = oWormhole.ObjectID Then Return
        Next X

        Dim vecLoc As Vector3
        If oWormhole.System1 Is Nothing = False AndAlso oWormhole.System1.ObjectID = lSystemID Then
            vecLoc = New Vector3(oWormhole.LocX1, -100, oWormhole.LocY1)
        ElseIf oWormhole.System2 Is Nothing = False Then
            vecLoc = New Vector3(oWormhole.LocX2, -100, oWormhole.LocY2)
        Else
            vecLoc = New Vector3(oWormhole.LocX2, 0, oWormhole.LocY2)
        End If

        mlWormholeUB += 2
        ReDim Preserve moWormholes(mlWormholeUB)
        moWormholes(mlWormholeUB - 1) = New WormholeEffect(vecLoc, 2000, False)
        moWormholes(mlWormholeUB - 1).oParentWormhole = oWormhole



        vecLoc.Y -= 200
        moWormholes(mlWormholeUB) = New WormholeEffect(vecLoc, 2000, True)
        moWormholes(mlWormholeUB).oParentWormhole = oWormhole
    End Sub
    Public Sub RenderFX(ByVal bUpdateNoRender As Boolean)
        Dim fElapsed As Single
        If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 10.0F
        mlPrevFrame = timeGetTime

        If mlWormholeUB = -1 Then Return

        For X As Int32 = 0 To mlWormholeUB
            With moWormholes(X)
                .oCullBox = CullBox.GetCullBox(.vecEmitter.X, .vecEmitter.Y, .vecEmitter.Z, -1000, -1000, -1000, 1000, 1000, 1000)
                .bCullObj = goCamera.CullObject(.oCullBox)
                If .bCullObj = False Then .Update(fElapsed)
            End With
        Next X

        If bUpdateNoRender = True Then Return

        If moParticle Is Nothing OrElse moParticle.Disposed = True Then moParticle = goResMgr.GetTexture("Flare2.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "xpl.pak")

        Dim moDevice As Device = GFXEngine.moDevice

        If glCurrentEnvirView = CurrentView.eSystemView AndAlso muSettings.WormholeAura = True Then
            Dim omat As Material
            With omat
                .Diffuse = System.Drawing.Color.FromArgb(128, 128, 0, 128)
                .Emissive = System.Drawing.Color.FromArgb(0, 0, 0, 0)
                .Ambient = System.Drawing.Color.FromArgb(128, 32, 0, 32)
                .Specular = System.Drawing.Color.FromArgb(128, 32, 0, 32)
                .SpecularSharpness = 2.0F
            End With
            With moDevice
                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

                .RenderState.SourceBlend = Blend.SourceAlpha
                .RenderState.DestinationBlend = Blend.One
                .RenderState.AlphaBlendEnable = True
                .RenderState.ZBufferEnable = False
                .RenderState.Lighting = True
            End With
            moDevice.Material = omat
            moDevice.SetTexture(0, moWormholeCloudTex)
            moDevice.RenderState.CullMode = Cull.None
            moDevice.RenderState.AlphaBlendEnable = True
            moDevice.Transform.World = Matrix.Identity
            moDevice.RenderState.Lighting = False
            moDevice.VertexFormat = CustomVertex.PositionColoredTextured.Format

            For X As Int32 = 0 To mlWormholeUB
                If moWormholes(X) Is Nothing = False Then

                    If moWormholes(X).bCullObj = False Then
                        If muWHVerts Is Nothing Then
                            ReDim muWHVerts(23)
                            moWormholeCloudTex = goResMgr.GetTexture("Cloud1.dds", GFXResourceManager.eGetTextureType.ModelTexture, "")
                        End If

                        Static xfTex1 As Single = 0.0F
                        Static xfTex2 As Single = 0.3F

                        xfTex1 += 0.00005F + (Rnd() * 0.00001F)
                        If xfTex1 > 1 Then xfTex1 = 0
                        xfTex2 -= 0.00005F
                        If xfTex2 < 0 Then xfTex2 = 1

                        Dim lEdge As Int32 = System.Drawing.Color.FromArgb(0, 128, 0, 128).ToArgb
                        Dim lCenter As Int32 = System.Drawing.Color.FromArgb(128, 128, 0, 128).ToArgb
                        PrepareWHVerts(-98, lEdge, lCenter, xfTex1, moWormholes(X).vecEmitter)

                        moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, 8, muWHVerts)

                        lEdge = System.Drawing.Color.FromArgb(0, 32, 64, 128).ToArgb
                        lCenter = System.Drawing.Color.FromArgb(255, 32, 64, 128).ToArgb
                        PrepareWHVerts(-102, lEdge, lCenter, xfTex2, moWormholes(X).vecEmitter)

                        moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, 8, muWHVerts)
                    End If
                End If
            Next X


            'moBPSprite.EndRender()
            With moDevice
                'reset state data
                .RenderState.SourceBlend = Blend.SourceAlpha        'Blend.One
                .RenderState.DestinationBlend = Blend.InvSourceAlpha    'Blend.Zero
                .RenderState.ZBufferEnable = True

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
            End With
        End If


        With moDevice
            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

            .RenderState.PointSpriteEnable = True
            .RenderState.PointScaleEnable = True
            .RenderState.PointSize = 32
            .RenderState.PointScaleA = 0
            .RenderState.PointScaleB = 0
            .RenderState.PointScaleC = 1
            .RenderState.SourceBlend = Blend.SourceAlpha
            .RenderState.DestinationBlend = Blend.One
            .RenderState.AlphaBlendEnable = True

            .RenderState.ZBufferWriteEnable = False

            .RenderState.Lighting = False

            .VertexFormat = CustomVertex.PositionColoredTextured.Format
            .SetTexture(0, moParticle)
        End With

        For X As Int32 = 0 To mlWormholeUB
            If moWormholes(X) Is Nothing = False Then
                If moWormholes(X).bCullObj = False Then
                    With moWormholes(X)
                        moDevice.DrawUserPrimitives(PrimitiveType.PointList, .mlParticleUB + 1, .moPoints)
                    End With
                End If
            End If
        Next X

        With moDevice
            .RenderState.ZBufferWriteEnable = True
            .RenderState.Lighting = True

            .RenderState.SourceBlend = Blend.One
            .RenderState.DestinationBlend = Blend.Zero
            .RenderState.AlphaBlendEnable = True

            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
        End With
    End Sub
    Private mfRotation As Single
    Private Sub PrepareWHVerts(ByVal lY As Int32, ByVal lEdge As Int32, ByVal lCenter As Int32, ByVal fTex As Single, ByVal vecPos As Vector3)
        Dim fMinTC As Single = fTex '* 5
        Dim fHalfTC As Single = (0.125F + fTex) '* 5
        Dim fFullTC As Single = (0.25F + fTex) '* 5

        Dim zAxis As Vector3 = New Vector3(goCamera.mlCameraX - goCamera.mlCameraAtX, goCamera.mlCameraY - goCamera.mlCameraAtY, goCamera.mlCameraZ - goCamera.mlCameraAtZ)
        zAxis.Normalize()
        Dim vecWorldUp As Vector3 = New Vector3(0, 1, 0)
        Dim xAxis As Vector3 = Vector3.Cross(zAxis, vecWorldUp)
        xAxis.Normalize()
        Dim yAxis As Vector3 = Vector3.Cross(xAxis, zAxis)
        yAxis.Normalize()

        Dim vecCol1 As Vector3 = New Vector3(xAxis.X, yAxis.X, zAxis.X)
        Dim vecCol2 As Vector3 = New Vector3(xAxis.Y, yAxis.Y, zAxis.Y)
        Dim vecCol3 As Vector3 = New Vector3(xAxis.Z, yAxis.Z, zAxis.Z)

        Dim lSecEdge As Int32 = System.Drawing.Color.FromArgb(64, System.Drawing.Color.FromArgb(lEdge).R, System.Drawing.Color.FromArgb(lEdge).G, System.Drawing.Color.FromArgb(lEdge).B).ToArgb

        'transform is DOT(vecPtX, New Vector3(M11,M21,M31))
        'DOT is (X1 * X2) + (Y1 * Y2) + (Z1 * Z2)
        'this is faster when dealing with large quantities of sprites... faster than DX's Vector3
        'Dim vecPosition As Vector3 = New Vector3(0, lY, 0)
        Dim fXPt As Single = 10000
        Dim fZPt As Single = 10000
        'mfRotation += 1.0F
        'RotatePoint(0, 0, fXPt, fZPt, mfRotation)


        '0 3 1
        muWHVerts(0) = New CustomVertex.PositionColoredTextured(-fXPt, lY, -fZPt, 0, fMinTC, fMinTC)
        muWHVerts(1) = New CustomVertex.PositionColoredTextured(-fXPt, lY, 0, 0, fMinTC, fHalfTC)
        muWHVerts(2) = New CustomVertex.PositionColoredTextured(0, lY, -fZPt, 0, fHalfTC, fMinTC)

        '3 4 1
        muWHVerts(3) = New CustomVertex.PositionColoredTextured(-fXPt, lY, 0, 0, fMinTC, fHalfTC)
        muWHVerts(4) = New CustomVertex.PositionColoredTextured(0, lY, 0, lCenter, fHalfTC, fHalfTC)
        muWHVerts(5) = New CustomVertex.PositionColoredTextured(0, lY, -fZPt, 0, fHalfTC, fMinTC)

        '1 4 2
        muWHVerts(6) = New CustomVertex.PositionColoredTextured(0, lY, -fZPt, 0, fHalfTC, fMinTC)
        muWHVerts(7) = New CustomVertex.PositionColoredTextured(0, lY, 0, lCenter, fHalfTC, fHalfTC)
        muWHVerts(8) = New CustomVertex.PositionColoredTextured(fXPt, lY, -fZPt, 0, fFullTC, fMinTC)

        '4 5 2
        muWHVerts(9) = New CustomVertex.PositionColoredTextured(0, lY, 0, lCenter, fHalfTC, fHalfTC)
        muWHVerts(10) = New CustomVertex.PositionColoredTextured(fXPt, lY, 0, 0, fFullTC, fHalfTC)
        muWHVerts(11) = New CustomVertex.PositionColoredTextured(fXPt, lY, -fZPt, 0, fFullTC, fMinTC)

        '3 6 4
        muWHVerts(12) = New CustomVertex.PositionColoredTextured(-fXPt, lY, 0, 0, fMinTC, fHalfTC)
        muWHVerts(13) = New CustomVertex.PositionColoredTextured(-fXPt, lY, fZPt, 0, fMinTC, fFullTC)
        muWHVerts(14) = New CustomVertex.PositionColoredTextured(0, lY, 0, lCenter, fHalfTC, fHalfTC)

        '6 7 4
        muWHVerts(15) = New CustomVertex.PositionColoredTextured(-fXPt, lY, fZPt, 0, fMinTC, fFullTC)
        muWHVerts(16) = New CustomVertex.PositionColoredTextured(0, lY, fZPt, 0, fHalfTC, fFullTC)
        muWHVerts(17) = New CustomVertex.PositionColoredTextured(0, lY, 0, lCenter, fHalfTC, fHalfTC)

        '4 7 5
        muWHVerts(18) = New CustomVertex.PositionColoredTextured(0, lY, 0, lCenter, fHalfTC, fHalfTC)
        muWHVerts(19) = New CustomVertex.PositionColoredTextured(0, lY, fZPt, 0, fHalfTC, fFullTC)
        muWHVerts(20) = New CustomVertex.PositionColoredTextured(fXPt, lY, 0, 0, fFullTC, fHalfTC)

        '7 8 5 
        muWHVerts(21) = New CustomVertex.PositionColoredTextured(0, lY, fZPt, 0, fHalfTC, fFullTC)
        muWHVerts(22) = New CustomVertex.PositionColoredTextured(fXPt, lY, fZPt, 0, fFullTC, fFullTC)
        muWHVerts(23) = New CustomVertex.PositionColoredTextured(fXPt, lY, 0, 0, fFullTC, fHalfTC)

        For X As Int32 = 0 To muWHVerts.GetUpperBound(0)
            With muWHVerts(X)
                Dim fX As Single = (.X * vecCol1.X) + (.Z * vecCol1.Y) + (0 * vecCol1.Z) + vecPos.X
                Dim fY As Single = (.X * vecCol2.X) + (.Z * vecCol2.Y) + (0 * vecCol2.Z) + vecPos.Y + lY
                Dim fZ As Single = (.X * vecCol3.X) + (.Z * vecCol3.Y) + (0 * vecCol3.Z) + vecPos.Z
                .X = fX : .Y = fY : .Z = fZ
            End With
        Next X

    End Sub

    Public Sub ClearTextures()
        moParticle = Nothing
        moWormholeCloudTex = Nothing
        muWHVerts = Nothing
    End Sub
End Class