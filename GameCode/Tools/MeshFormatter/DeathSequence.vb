Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class DeathSequence
    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Private Const ml_MINOR_EXP_PCNT As Int32 = 200
    Private Const ml_FINALE_EXP_PCNT As Int32 = 500

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

    Private Structure MinorExplosion
        'Where this minor explosion is located
        Public mvecEmitter As Vector3

        Private moParticles() As Particle
        Private mlPrevFrame As Int32

        Public mbEmitterStopping As Boolean
        Public EmitterStopped As Boolean

        'REFERENCE To the Actual array!!!
        Private moPoints() As CustomVertex.PositionColoredTextured
        Private mlPointStartIndex As Int32      'Tells us where to put the points

        Public vecFlashPoint As Vector3
        Public fFlashAlpha As Single

        Public mlParticleUB As Int32

        Public bInitialized As Boolean

        Private mlSmallestSize As Int32
        Private mlHalfSmallestSize As Int32

        Public Sub Initialize(ByVal vecLoc As Vector3, ByVal lStartIndex As Int32, ByRef poPoints() As CustomVertex.PositionColoredTextured, ByVal lSmallestSizeVal As Int32)
            mlSmallestSize = lSmallestSizeVal \ 4
            If mlSmallestSize > 70 Then mlSmallestSize = 70
            mlHalfSmallestSize = mlSmallestSize \ 2

            mbEmitterStopping = False
            EmitterStopped = False
            mlPointStartIndex = lStartIndex
            moPoints = poPoints

            Dim X As Int32

            mvecEmitter = vecLoc

            mlParticleUB = ml_MINOR_EXP_PCNT - 1
            ReDim moParticles(mlParticleUB)

            For X = 0 To mlParticleUB
                ResetParticle(X)
            Next X

            vecFlashPoint = New Vector3(vecLoc.X, vecLoc.Y, vecLoc.Z)
            fFlashAlpha = 255

            bInitialized = True
        End Sub

        Private Sub ResetParticle(ByVal lIndex As Int32)
            Dim fX As Single
            Dim fY As Single
            Dim fZ As Single
            Dim fXS As Single
            Dim fYS As Single
            Dim fZS As Single
            Dim fXA As Single
            Dim fYA As Single
            Dim fZA As Single

            Dim lX As Int32

            If mbEmitterStopping = True Then
                moParticles(lIndex).mfA = 0
                moParticles(lIndex).ParticleActive = False
                moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

                EmitterStopped = True
                For lX = 0 To mlParticleUB
                    If moParticles(lX).ParticleActive = True Then
                        EmitterStopped = False
                        Exit For
                    End If
                Next lX
            Else
                'fX = mvecEmitter.X + (Rnd() * 50 - 25)
                'fY = mvecEmitter.Y + (Rnd() * 50 - 25)
                'fZ = mvecEmitter.Z + (Rnd() * 50 - 25)
                'fXS = (Rnd() * 40) - 20
                'fYS = (Rnd() * 40) - 20
                'fZS = (Rnd() * 40) - 20

                
                fX = mvecEmitter.X + (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fY = mvecEmitter.Y + (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fZ = mvecEmitter.Z + (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fXS = (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fYS = (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fZS = (Rnd() * mlSmallestSize - mlHalfSmallestSize)

                fXA = Rnd() * 0.2F : fYA = Rnd() * 0.2F : fZA = Rnd() * 0.2F
                If fXS > 0 Then fXA = -fXA
                If fYS > 0 Then fYA = -fYA
                If fZS > 0 Then fZA = -fZA

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, 200 + (Rnd() * 50), 90 + (Rnd() * 30), 32 + (Rnd() * 32), 255)
                moParticles(lIndex).fAChg = -((Rnd() * 16) + 8)
            End If


        End Sub

        Public Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            If EmitterStopped = True Then Exit Sub

            If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            mlPrevFrame = timeGetTime

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    If .ParticleActive = True Then
                        .Update(fElapsed)
                        If .mfA <= 0 Then
                            ResetParticle(X)
                        End If

                        moPoints(X + mlPointStartIndex).Color = .ParticleColor.ToArgb
                        moPoints(X + mlPointStartIndex).Position = .vecLoc
                    End If
                End With
            Next X

            fFlashAlpha -= 15 * fElapsed
            If fFlashAlpha < 0 Then fFlashAlpha = 0

            If fFlashAlpha = 0 AndAlso mbEmitterStopping = False Then
                vecFlashPoint = New Vector3(mvecEmitter.X + (Rnd() * 50 - 25), mvecEmitter.Y + (Rnd() * 50 - 25), mvecEmitter.Z + (Rnd() * 50 - 25))
                fFlashAlpha = 255
            End If
        End Sub
    End Structure

    Private Structure FinaleExplosion
        Public mvecEmitter As Vector3

        Private moParticles() As Particle
        Private mlPrevFrame As Int32
        Public mbEmitterStopping As Boolean

        Public moPoints() As CustomVertex.PositionColoredTextured
        Public mlParticleUB As Int32

        Public EmitterStopped As Boolean

        Public bInitialized As Boolean

        Private mfPraxisAlpha As Single
        Private mfPraxisRed As Single
        Private mfPraxisGreen As Single
        Private mfPraxisBlue As Single
        Public mfPraxisScale As Single
        Public moPraxisColor As System.Drawing.Color

        Private mfMaxPraxisSize As Single

        Public Event FinaleComplete()

        Public Sub Initialize(ByVal vecLoc As Vector3, ByVal lMaxSize As Int32)
            mfMaxPraxisSize = lMaxSize / 30.0F

            mbEmitterStopping = False
            EmitterStopped = False

            Dim X As Int32

            mvecEmitter = vecLoc

            mlParticleUB = ml_FINALE_EXP_PCNT - 1
            ReDim moParticles(mlParticleUB)
            ReDim moPoints(mlParticleUB + 1)

            For X = 0 To mlParticleUB
                ResetParticle(X)
            Next X

            mfPraxisAlpha = 255.0F
            mfPraxisRed = 255.0F
            mfPraxisGreen = 255.0F
            mfPraxisBlue = 255.0F
            mfPraxisScale = 0.0F

            bInitialized = True
            mbEmitterStopping = True
        End Sub

        Private Sub ResetParticle(ByVal lIndex As Int32)
            Dim fX As Single
            Dim fY As Single
            Dim fZ As Single
            Dim fXS As Single
            Dim fYS As Single
            Dim fZS As Single
            Dim fXA As Single
            Dim fYA As Single
            Dim fZA As Single

            Dim lX As Int32

            If mbEmitterStopping = True Then
                moParticles(lIndex).mfA = 0
                moParticles(lIndex).ParticleActive = False
                moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

                EmitterStopped = True
                For lX = 0 To mlParticleUB
                    If moParticles(lX).ParticleActive = True Then
                        EmitterStopped = False
                        Exit For
                    End If
                Next lX

                If EmitterStopped = True Then RaiseEvent FinaleComplete()
            Else
                fX = mvecEmitter.X + (Rnd() * 10 - 5)  '(Rnd() - 0.5)
                fY = mvecEmitter.Y + (Rnd() * 10 - 5) '(Rnd() - 0.5)
                fZ = mvecEmitter.Z + (Rnd() * 10 - 5) ' (Rnd() - 0.5)
                fXS = (Rnd() * 10) - 5
                fYS = (Rnd() * 10) - 5
                fZS = (Rnd() * 10) - 5

                fXA = Rnd() * 6 - 3 : fYA = Rnd() * 6 - 3 : fZA = Rnd() * 6 - 3
                If fXS > 0 Then fXA = -fXA
                If fYS > 0 Then fYA = -fYA
                If fZS > 0 Then fZA = -fZA

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, 200 + (Rnd() * 50), 90 + (Rnd() * 30), 32 + (Rnd() * 32), 255)
                moParticles(lIndex).fAChg = -(Rnd() + 3.0F)
            End If


        End Sub

        Public Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            If EmitterStopped = True Then Return

            If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            mlPrevFrame = timeGetTime

            'Update the praxis
            'TODO: Make the Praxis effect dependant on the size of the entity exploding
            mfPraxisAlpha -= (10.0F * fElapsed)
            mfPraxisRed -= (0.001F * fElapsed)
            mfPraxisGreen -= (0.5F * fElapsed)
            mfPraxisBlue -= (1.0F * fElapsed)
            mfPraxisScale = (1.0F - (mfPraxisAlpha / 255.0F)) * mfMaxPraxisSize
            'mfPraxisScale += (0.7F * fElapsed)
            If mfPraxisAlpha < 0 Then mfPraxisAlpha = 0
            If mfPraxisRed < 0 Then mfPraxisRed = 0
            If mfPraxisGreen < 0 Then mfPraxisGreen = 0
            If mfPraxisBlue < 0 Then mfPraxisBlue = 0
            moPraxisColor = Color.FromArgb(CInt(mfPraxisAlpha), CInt(mfPraxisRed), CInt(mfPraxisGreen), CInt(mfPraxisBlue))

            If mfPraxisAlpha < 0 Then mfPraxisAlpha = 0
            If mfPraxisRed < 200 Then mfPraxisRed = 200
            If mfPraxisGreen < 120 Then mfPraxisGreen = 120
            If mfPraxisBlue < 0 Then mfPraxisBlue = 0

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    If .ParticleActive = True Then
                        .Update(fElapsed)
                        If .mfA <= 0 Then
                            ResetParticle(X)
                        End If

                        moPoints(X).Color = .ParticleColor.ToArgb
                        moPoints(X).Position = .vecLoc
                    End If
                End With
            Next X

        End Sub

    End Structure

    'Particle texture...
    Public Shared moTex As Texture
    Public Shared moPraxisTex As Texture
    Private moDevice As Device  'reference only
    Private moLoc As Vector3
    Private moSize As Vector3

    'Helper for half of moSize
    Private moHalfSize As Vector3

    'Based on parameter in New routine
    Private mbDoFinalExplosion As Boolean = False

    'The individual explosions
    Private moExp() As MinorExplosion
    'That lead up to the finale
    Private moFinale As FinaleExplosion

    'Last time an explosion was added
    Private mlLastExpAdd As Int32 = 0

    'Indicates that this sequence has ended in the event that any subsequent calls are made to this sequence after termination
    Public bSequenceEnded As Boolean = False

    'Indicates that the model is to be rendered still or not... at a point in the sequence, the model will be invisible
    Public bRenderMesh As Boolean = True

    'The final explosion causes a bright light, this manages that bright light
    Private mlFlashPointAlpha As Int32 = 0

    'The actual point list used for all children
    Private moPoints() As CustomVertex.PositionColoredTextured
    Private mlPointUB As Int32 = -1

    Private moFlashPoints() As CustomVertex.PositionColoredTextured
    Private mlCurrentFlashPointUB As Int32 = -1

    Private mlCurrentMaxPointUB As Int32 = -1

    Private mbRenderExplosions As Boolean = True
    Private mlEntityIndex As Int32 = -1

    Public Sub New(ByRef oDevice As Device, ByVal oLoc As Vector3, ByVal lExpCnt As Int32, ByVal oSize As Vector3, ByVal bDoFinalExp As Boolean, ByVal lEntityIndex As Int32)
        moDevice = oDevice
        mlEntityIndex = lEntityIndex

        If moTex Is Nothing Then
            Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sFile.EndsWith("\") = False Then sFile &= "\"
            moTex = TextureLoader.FromFile(moDevice, sFile & "Particle.dds")
            moPraxisTex = TextureLoader.FromFile(moDevice, sFile & "Praxis.dds")
        End If

        moLoc = oLoc
        moSize = oSize
        moHalfSize = New Vector3(oSize.X / 2.0F, oSize.Y / 2.0F, oSize.Z / 2.0F)
        mbDoFinalExplosion = bDoFinalExp

        'Ok, our points
        ReDim moPoints(lExpCnt * ml_MINOR_EXP_PCNT - 1)
        ReDim moFlashPoints(lExpCnt - 1)
        ReDim moExp(lExpCnt - 1)
        InitializeExplosion(0)

    End Sub

    Private Sub InitializeExplosion(ByVal lIdx As Int32)
        Dim fX As Single = (Rnd() * moSize.X) - moHalfSize.X
        Dim fY As Single = (Rnd() * moSize.Y) - moHalfSize.Y
        Dim fZ As Single = (Rnd() * moSize.Z) - moHalfSize.Z

        RotateVals(fX, fY, fZ, CShort(Form1.hscrRotate.Value) - 900S, CShort(Form1.hscrYaw.Value))

        Dim lSmallest As Int32 = CInt(Math.Min(Math.Min(moSize.X, moSize.Y), moSize.Z))

        moExp(lIdx).Initialize(New Vector3(moLoc.X + fX, moLoc.Y + fY, moLoc.Z + fZ), lIdx * ml_MINOR_EXP_PCNT, moPoints, lSmallest)
        mlCurrentMaxPointUB = ((lIdx + 1) * ml_MINOR_EXP_PCNT) - 1
        mlCurrentFlashPointUB = lIdx
    End Sub

    Public Sub Render(ByVal bUpdateNoRender As Boolean)
        If bSequenceEnded = True Then Return

        If moExp Is Nothing = False AndAlso mbRenderExplosions = True Then

            mbRenderExplosions = False
            For X As Int32 = 0 To moExp.GetUpperBound(0)
                If moExp(X).bInitialized = True Then
                    moExp(X).Update()
                    If moExp(X).EmitterStopped = False Then mbRenderExplosions = True
                ElseIf timeGetTime - mlLastExpAdd > 75 Then
                    InitializeExplosion(X)
                    mlLastExpAdd = timeGetTime
                    Exit For
                End If
            Next X

            If bRenderMesh = True Then
                bRenderMesh = moExp(moExp.GetUpperBound(0)).bInitialized = False
                If bRenderMesh = False Then

                    For X As Int32 = 0 To moExp.GetUpperBound(0)
                        moExp(X).mbEmitterStopping = True
                        mlFlashPointAlpha = 255
                    Next

                    If mbDoFinalExplosion = True Then
                        Dim lMax As Int32 = CInt(Math.Max(Math.Max(moSize.X, moSize.Y), moSize.Z))
                        moFinale.Initialize(moLoc, lMax)
                        AddHandler moFinale.FinaleComplete, AddressOf BigFinale_Ended
                    End If
                End If
            End If

        End If

        If mbRenderExplosions = False AndAlso mbDoFinalExplosion = False Then
            BigFinale_Ended()
            Return
        End If

        If bUpdateNoRender = False Then

            'Now, do our rendering of everything...
            With moDevice
                .Transform.World = Matrix.Identity

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, Form1.l_DS_Op)

                .RenderState.PointSpriteEnable = True
                .RenderState.PointScaleEnable = True
                .RenderState.PointSize = 64
                .RenderState.PointScaleA = 0
                .RenderState.PointScaleB = 0
                .RenderState.PointScaleC = 1
                .RenderState.SourceBlend = Blend.SourceAlpha
                .RenderState.DestinationBlend = Blend.One
                .RenderState.AlphaBlendEnable = True

                .RenderState.ZBufferWriteEnable = False

                .RenderState.Lighting = False

                .VertexFormat = CustomVertex.PositionColoredTextured.Format
                .SetTexture(0, moTex)

                If mbRenderExplosions = True Then
                    'Render the Minor Explosion Points in one call...
                    .DrawUserPrimitives(PrimitiveType.PointList, mlCurrentMaxPointUB + 1, moPoints)

                    'Prepare to render flashpoints
                    .RenderState.PointSize = 512

                    'set up our flashpoints
                    For X As Int32 = 0 To mlCurrentFlashPointUB
                        moFlashPoints(X).Color = System.Drawing.Color.FromArgb(CInt(moExp(X).fFlashAlpha), 255, 255, 255).ToArgb
                        moFlashPoints(X).Position = moExp(X).vecFlashPoint
                    Next X

                    'Render the MinorExplosion's FlashPoints
                    .DrawUserPrimitives(PrimitiveType.PointList, mlCurrentFlashPointUB + 1, moFlashPoints)
                End If

                'Render the Finale Explosion Points in one call...
                If mbDoFinalExplosion = True AndAlso moFinale.bInitialized = True Then
                    .RenderState.PointSize = 128
                    moFinale.Update()
                    .DrawUserPrimitives(PrimitiveType.PointList, moFinale.moPoints.Length, moFinale.moPoints)

                    'Then, render the Praxis effect
                    If moFinale.moPraxisColor.A > 0 Then
                        moDevice.Transform.World = Matrix.Multiply(Matrix.Identity, Matrix.Scaling(moFinale.mfPraxisScale, moFinale.mfPraxisScale, moFinale.mfPraxisScale))
                        Using moPraxis As Sprite = New Sprite(moDevice)
                            moPraxis.SetWorldViewLH(moDevice.Transform.World, moDevice.Transform.View)
                            moPraxis.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.ObjectSpace)
                            moPraxis.Draw(moPraxisTex, System.Drawing.Rectangle.Empty, New Vector3(64, 64, 0), Vector3.Multiply(Me.moLoc, 1 / moFinale.mfPraxisScale), moFinale.moPraxisColor)
                            moPraxis.End()
                        End Using
                    End If
                End If

                'Then, reset our device...
                .RenderState.ZBufferWriteEnable = True
                .RenderState.Lighting = True

                .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
                .RenderState.AlphaBlendEnable = True

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
            End With
        End If

        If mlFlashPointAlpha <> 0 Then
            If bUpdateNoRender = False Then
                Using oSpr As New Sprite(moDevice)
                    Dim oTmpMat As Matrix = Matrix.Identity
                    Dim oTmpVec As Vector3 = New Vector3(moLoc.X / moHalfSize.X, moLoc.Y / moHalfSize.Y, moLoc.Z / moHalfSize.Z)
                    oTmpMat.Multiply(Matrix.Translation(oTmpVec))
                    oTmpMat.Multiply(Matrix.Scaling(moHalfSize))
                    moDevice.Transform.World = oTmpMat
                    oSpr.SetWorldViewLH(oTmpMat, moDevice.Transform.View)
                    oSpr.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace)
                    oSpr.Draw(moTex, System.Drawing.Rectangle.Empty, New Vector3(16, 16, 0), New Vector3(0, 0, 0), Color.FromArgb(mlFlashPointAlpha, 255, 255, 255))
                    oSpr.End()
                    oSpr.Dispose()
                    moDevice.Transform.World = Matrix.Identity
                End Using
            End If

            mlFlashPointAlpha -= 10
            If mlFlashPointAlpha < 0 Then mlFlashPointAlpha = 0
        End If
    End Sub

    Private Sub BigFinale_Ended()
        bSequenceEnded = True
        Me.bRenderMesh = True
    End Sub


    Private Sub RotateVals(ByRef fXLoc As Single, ByRef fYLoc As Single, ByRef fZLoc As Single, ByVal iLocAngle As Int16, ByVal iLocYaw As Int16)
        Dim fDX As Single
        Dim fDZ As Single

        Const gdRadPerDegree As Single = Math.PI / 180.0F

        Dim fRads As Single = -((iLocYaw / 10.0F)) * gdRadPerDegree
        Dim fCosR As Single = CSng(Math.Cos(fRads))
        Dim fSinR As Single = CSng(Math.Sin(fRads))

        'Yaw...
        fDX = fXLoc
        fDZ = fYLoc
        fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
        fYLoc = -((fDX * fSinR) - (fDZ * fCosR))

        'Now set up for standard rotation...
        fRads = ((iLocAngle / 10.0F) - 90) * gdRadPerDegree
        fCosR = CSng(Math.Cos(fRads))
        fSinR = CSng(Math.Sin(fRads))

        'Loc
        fDX = fXLoc
        fDZ = fZLoc
        fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
        fZLoc = -((fDX * fSinR) - (fDZ * fCosR))

    End Sub

End Class