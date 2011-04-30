Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class ExplosionManager

    Private Structure Explosion
        Public vecLoc As Vector3
        Public fRotation As Single
        Public fRotationChange As Single
        Public lExplosionType As Int32
        Public lState As Int32
        Public fSize As Single
        Public fSizeChange As Single
        Public clr As System.Drawing.Color

        Public fBrightLightSize As Single
        Public clrBrightLight As System.Drawing.Color
        Public yBrightLightShownFull As Byte

        Public fFadeDelay As Single

        Public sSoundFX As String
        Public bPlaySound As Boolean

        Private fState As Single
        Public Sub Update(ByVal fElapsed As Single)
            fRotation += fRotationChange * fElapsed
            fSize += fSizeChange * fElapsed
            fState += fElapsed * fFadeDelay
            lState = CInt(fState)
            If yBrightLightShownFull = 0 Then
                yBrightLightShownFull = 1
            Else
                fBrightLightSize -= (100.0F * fElapsed)
            End If
        End Sub

        Public Sub SetAccurateState(ByVal fVal As Single)
            fState = fVal
        End Sub
    End Structure

    Public Shared CurrentAddCount As Int32 = 0

    Private moExplosions() As Explosion
    Private myExplosionUsed() As Byte
    Private mlExplosionUB As Int32 = -1

    Private moNewExp1() As Explosion
    Private mlNewExp1UB As Int32 = -1
    Private moNewExp2() As Explosion
    Private mlNewExp2UB As Int32 = -1
    Private myNewExpAddTo As Byte = 0

    Private mlAddToQueue1 As Int32 = 0
    Private mlAddToQueue2 As Int32 = 0

    Private mlSpriteCnts() As Int32
    Private moExplTex() As Texture
    Private msExplTexName() As String

    Private moSprite() As BPSprite
    Private mswTimer As Stopwatch

    Private msExplSnd() As String

    Private mrctSrc() As Rectangle

    Public Sub New()
        ReDim moNewExp1(-1)
        ReDim moNewExp2(-1)

        ReDim mlSpriteCnts(4)
        mlSpriteCnts(0) = 0
        mlSpriteCnts(1) = 0
        mlSpriteCnts(2) = 0
        mlSpriteCnts(3) = 0
        mlSpriteCnts(4) = 0

        ReDim moExplTex(4)
        ReDim msExplTexName(4)

        msExplTexName(0) = "ani2.dds"
        msExplTexName(1) = "ani3.dds"
        msExplTexName(2) = "ani5.dds"
        msExplTexName(3) = "ani8.dds"
        msExplTexName(4) = "Flare2.dds"

        ReDim msExplSnd(5)
        msExplSnd(0) = "Explosions\HullHit1.wav"
        msExplSnd(1) = "Explosions\HullHit2.wav"
        msExplSnd(2) = "Explosions\HullHit3.wav"
        msExplSnd(3) = "Explosions\HullHit4.wav"
        msExplSnd(4) = "Explosions\SmallSpaceDeath1.wav"
        msExplSnd(5) = "Explosions\SmallSpaceDeath2.wav"


        ReDim mrctSrc(15)
        For lIdx As Int32 = 0 To 15
            Dim lY As Int32 = lIdx \ 4
            Dim lX As Int32 = lIdx - (lY * 4)
            mrctSrc(lIdx) = New Rectangle(lX * 64, lY * 64, 64, 64)
        Next lIdx
        mswTimer = Stopwatch.StartNew()

        ReDim moSprite(4)
        For X As Int32 = 0 To 4
            moSprite(X) = New BPSprite()
        Next X
    End Sub

    Public Sub Add(ByVal pvecLoc As Vector3, ByVal pfRotation As Single, ByVal pfRotationChange As Single, ByVal plExplosionType As Int32, ByVal pfSize As Single, ByVal pfSizeChange As Single, ByVal clr As System.Drawing.Color, ByVal clrBright As System.Drawing.Color, ByVal lFadeDelay As Int32, ByVal fFlashSize As Single, ByVal bPlaySound As Boolean)
        If muSettings.RenderExplosionFX = False Then Return
        If CurrentAddCount > 500 Then Return
        CurrentAddCount += 1

        plExplosionType = Math.Max(0, Math.Min(plExplosionType, 3))

        Dim lSndUB As Int32 = msExplSnd.GetUpperBound(0)
        Dim lSnd As Int32 = CInt(Rnd() * lSndUB)
        If lSnd > lSndUB Then lSnd = lSndUB
        If lSnd < 0 Then lSnd = 0

        Dim uExpl As Explosion
        With uExpl
            .fRotation = pfRotation
            .fRotationChange = pfRotationChange
            .fSize = pfSize
            .fSizeChange = pfSizeChange
            .lExplosionType = plExplosionType
            .lState = 0
            .vecLoc = pvecLoc
            .clr = clr
            .fBrightLightSize = fFlashSize
            .clrBrightLight = clrBright
            .fFadeDelay = 1.0F / (lFadeDelay / 30.0F)
            .sSoundFX = msExplSnd(lSnd)
            .SetAccurateState(0)
            .yBrightLightShownFull = 0
            .bPlaySound = bPlaySound
        End With

        If myNewExpAddTo = 0 Then
            mlAddToQueue1 += 1
            SyncLock moNewExp1
                mlNewExp1UB += 1
                If mlNewExp1UB > moNewExp1.GetUpperBound(0) Then
                    ReDim Preserve moNewExp1(mlNewExp1UB + 100)
                End If

                moNewExp1(mlNewExp1UB) = uExpl
            End SyncLock
            mlAddToQueue1 -= 1
        Else
            mlAddToQueue2 += 1
            SyncLock moNewExp2
                mlNewExp2UB += 1
                If mlNewExp2UB > moNewExp2.GetUpperBound(0) Then
                    ReDim Preserve moNewExp2(mlNewExp2UB + 100)
                End If

                moNewExp2(mlNewExp2UB) = uExpl
            End SyncLock
            mlAddToQueue2 -= 1
        End If

    End Sub

    Public Shared Function GetWeaponTypeBrightColor(ByVal yWeaponType As Byte) As System.Drawing.Color
        Select Case CType(yWeaponType, WeaponType)
            'TODO: add others in here...
            Case WeaponType.eCapitalSolidGreenBeam, WeaponType.eFlickerGreenBeam, WeaponType.eShortGreenPulse, WeaponType.eSolidGreenBeam
                Return System.Drawing.Color.FromArgb(255, 64, 255, 64)
            Case WeaponType.eCapitalSolidPurpleBeam, WeaponType.eFlickerPurpleBeam, WeaponType.eShortPurplePulse, WeaponType.eSolidPurpleBeam
                Return System.Drawing.Color.FromArgb(255, 255, 64, 255)
            Case WeaponType.eCapitalSolidRedBeam, WeaponType.eFlickerRedBeam, WeaponType.eShortRedPulse, WeaponType.eSolidRedBeam
                Return System.Drawing.Color.FromArgb(255, 255, 64, 64)
            Case WeaponType.eCapitalSolidTealBeam, WeaponType.eFlickerTealBeam, WeaponType.eShortTealPulse, WeaponType.eSolidTealBeam
                Return System.Drawing.Color.FromArgb(255, 64, 255, 255)
            Case WeaponType.eCapitalSolidYellowBeam, WeaponType.eFlickerYellowBeam, WeaponType.eShortYellowPulse, WeaponType.eSolidYellowBeam
                Return System.Drawing.Color.FromArgb(255, 255, 192, 128)
            Case WeaponType.eMetallicProjectile_Bronze, WeaponType.eMetallicProjectile_Copper, WeaponType.eMetallicProjectile_Gold, WeaponType.eMetallicProjectile_Lead, WeaponType.eMetallicProjectile_Silver
                Return System.Drawing.Color.FromArgb(255, 255, 128, 64)
            Case Else
                Return Color.White
        End Select
    End Function

    Public Sub AddBeamExplosionToEntity(ByRef oTarget As RenderObject, ByVal yWeaponTypeID As Byte, ByVal vecOnModelTo As Vector3)
        Dim vecFinal As Vector3
        If oTarget Is Nothing = False Then
            'vecFinal = Vector3.TransformCoordinate(vecOnModelTo, oTarget.GetWorldMatrix)
            With oTarget.GetWorldMatrix()
                Dim fX As Single = vecOnModelTo.X
                Dim fY As Single = vecOnModelTo.Y
                Dim fZ As Single = vecOnModelTo.Z
                vecFinal.X = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                vecFinal.Y = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                vecFinal.Z = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
            End With
        Else : vecFinal = vecOnModelTo
        End If
        Dim clrBright As System.Drawing.Color = ExplosionManager.GetWeaponTypeBrightColor(yWeaponTypeID)
        Dim fSize As Single = Rnd() * 20 + 100
        Add(vecFinal, Rnd() * 360, Rnd(), CInt(Rnd() * 4), fSize, 0, Color.White, clrBright, 30, fSize * 2, True)
    End Sub

    Private Sub LoadExplTex(ByVal lIdx As Int32)
        moExplTex(lIdx) = Nothing
        moExplTex(lIdx) = goResMgr.GetTexture(msExplTexName(lIdx), GFXResourceManager.eGetTextureType.NoSpecifics, "xpl.pak")
    End Sub

    Public Sub Render(ByVal bUpdateNoRender As Boolean)
        If muSettings.RenderExplosionFX = False Then Return
        If myNewExpAddTo = 0 Then myNewExpAddTo = 1 Else myNewExpAddTo = 0

        glExplosionRendered = 0
        Dim moDevice As Device = GFXEngine.moDevice
        Try
            Dim bNoNeed As Boolean = True
            For X As Int32 = 0 To 4
                If mlSpriteCnts(X) > 0 Then
                    If bUpdateNoRender = False Then
                        If moExplTex(X) Is Nothing OrElse moExplTex(X).Disposed = True Then LoadExplTex(X)
                        moSprite(X).BeginRender(mlSpriteCnts(X), moExplTex(X), moDevice)
                    End If
                    bNoNeed = False
                ElseIf mlSpriteCnts(X) < 0 Then
                    mlSpriteCnts(X) = 0
                End If
            Next X

            Dim fElapsed As Single = mswTimer.ElapsedMilliseconds / 30.0F
            mswTimer.Reset()
            mswTimer.Start()
            If bNoNeed = True Then
                ProcessQueue()
                Return
            End If
            For X As Int32 = 0 To mlExplosionUB
                If myExplosionUsed(X) <> 0 Then
                    With moExplosions(X)
                        .Update(fElapsed)

                        If .fBrightLightSize <> Single.MinValue Then
                            If .fBrightLightSize < 0 Then
                                .fBrightLightSize = Single.MinValue
                                mlSpriteCnts(4) -= 1
                            Else
                                glExplosionRendered += 1
                                If bUpdateNoRender = False Then moSprite(4).Draw(.vecLoc, .fBrightLightSize, .clrBrightLight, 0) ' Color.FromArgb(255, 64, 128, 255), 0)
                            End If
                        End If

                        If .lState > 15 Then
                            mlSpriteCnts(.lExplosionType) -= 1
                            myExplosionUsed(X) = 0
                        Else
                            glExplosionRendered += 1
                            If bUpdateNoRender = False Then moSprite(.lExplosionType).Draw(.vecLoc, .fSize, .clr, .fRotation, mrctSrc(.lState))
                        End If
                    End With
                End If
            Next X

            If bUpdateNoRender = False Then
                With moDevice
                    .RenderState.SourceBlend = Blend.SourceAlpha
                    .RenderState.DestinationBlend = Blend.One

                    moSprite(4).EndRender()
                    For X As Int32 = 0 To 3
                        If mlSpriteCnts(X) > 0 Then moSprite(X).EndRender()
                    Next X

                    .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                    .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
                End With
            End If

        Catch
            Try
                moDevice.RenderState.SourceBlend = Blend.SourceAlpha
                moDevice.RenderState.DestinationBlend = Blend.InvSourceAlpha
            Catch
            End Try
        End Try

        ProcessQueue()
    End Sub

    Private Sub ProcessQueue()
        If myNewExpAddTo = 0 Then
            'ok, we are checking queue2
            Dim lWaits As Int32 = 0
            While mlAddToQueue2 > 0
                Threading.Thread.Sleep(1)
                lWaits += 1
                If lWaits > 10 Then Exit While
            End While

            Dim lLastAvailableIdx As Int32 = 0
            For X As Int32 = 0 To mlNewExp2UB
                Dim bFound As Boolean = False
                For Y As Int32 = lLastAvailableIdx To mlExplosionUB
                    If myExplosionUsed(Y) = 0 Then
                        lLastAvailableIdx = Y
                        bFound = True
                        moExplosions(Y) = moNewExp2(X)
                        myExplosionUsed(Y) = 255
                        Exit For
                    End If
                Next Y
                If bFound = False Then
                    mlExplosionUB += 1
                    ReDim Preserve moExplosions(mlExplosionUB)
                    ReDim Preserve myExplosionUsed(mlExplosionUB)
                    moExplosions(mlExplosionUB) = moNewExp2(X)
                    myExplosionUsed(mlExplosionUB) = 255
                    lLastAvailableIdx = mlExplosionUB
                End If
                mlSpriteCnts(moNewExp2(X).lExplosionType) += 1
                mlSpriteCnts(4) += 1
                If goSound Is Nothing = False AndAlso moNewExp2(X).bPlaySound = True Then goSound.StartSound(moNewExp2(X).sSoundFX, False, SoundMgr.SoundUsage.eNonDeathExplosions, moNewExp2(X).vecLoc, New Vector3(0, 0, 0))
            Next X

            ReDim moNewExp2(100)
            mlNewExp2UB = -1
        Else
            'we are checking queue1
            Dim lWaits As Int32 = 0
            While mlAddToQueue1 > 0
                Threading.Thread.Sleep(1)
                lWaits += 1
                If lWaits > 10 Then Exit While
            End While

            Dim lLastAvailableIdx As Int32 = 0
            For X As Int32 = 0 To mlNewExp1UB
                Dim bFound As Boolean = False
                For Y As Int32 = lLastAvailableIdx To mlExplosionUB
                    If myExplosionUsed(Y) = 0 Then
                        lLastAvailableIdx = Y
                        bFound = True
                        moExplosions(Y) = moNewExp1(X)
                        myExplosionUsed(Y) = 255
                        Exit For
                    End If
                Next Y
                If bFound = False Then
                    mlExplosionUB += 1
                    ReDim Preserve moExplosions(mlExplosionUB)
                    ReDim Preserve myExplosionUsed(mlExplosionUB)
                    moExplosions(mlExplosionUB) = moNewExp1(X)
                    myExplosionUsed(mlExplosionUB) = 255
                    lLastAvailableIdx = mlExplosionUB
                End If
                mlSpriteCnts(moNewExp1(X).lExplosionType) += 1
                mlSpriteCnts(4) += 1
                If goSound Is Nothing = False AndAlso moNewExp1(X).bPlaySound = True Then goSound.StartSound(moNewExp1(X).sSoundFX, False, SoundMgr.SoundUsage.eNonDeathExplosions, moNewExp1(X).vecLoc, New Vector3(0, 0, 0))
            Next X

            ReDim moNewExp1(100)
            mlNewExp1UB = -1
        End If
    End Sub

    Public Sub RemoveAll()
        mlExplosionUB = -1
        ReDim moExplosions(-1)
    End Sub

    Public Sub ClearResources()
        If moExplTex Is Nothing = False Then
            For X As Int32 = 0 To moExplTex.GetUpperBound(0)
                If moExplTex(X) Is Nothing = False Then moExplTex(X).Dispose()
                moExplTex(X) = Nothing
            Next X
        End If
    End Sub

    Protected Overrides Sub Finalize()
        RemoveAll()
        If moSprite Is Nothing = False Then
            For X As Int32 = 0 To moSprite.GetUpperBound(0)
                If moSprite(X) Is Nothing = False Then
                    moSprite(X).DisposeMe()
                    moSprite(X) = Nothing
                End If
            Next X
            Erase moSprite
        End If
        MyBase.Finalize()
    End Sub
End Class
