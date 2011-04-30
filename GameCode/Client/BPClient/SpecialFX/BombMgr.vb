Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class BombMgr
    Private Shared moRandom As New Random
    Public Shared goBombMgr As New BombMgr()

    Private Class BombShot
        Public vecLoc As Vector3
        Public vecVel As Vector3

        Public lFloor As Int32
        Private mlLastUpdate As Int32

        Public clrBase As System.Drawing.Color
        Public yAOE As Byte = 0

        Private mvecTrailVec() As Vector3
        Private mclrTrailClr() As System.Drawing.Color
        Private mfTrailSize() As Single

        Public Function Update(ByRef oSpr As BPSprite, ByVal bUpdateNoRender As Boolean) As Boolean

            Dim bResult As Boolean = False

            'CHANGE THIS OUT WITH CYCLE-BASED CALCS
            Dim lElapsed As Int32 = glCurrentCycle - lElapsed
            mlLastUpdate = lElapsed

            vecLoc += vecVel
            If vecLoc.Y < lFloor Then
                bResult = True
            End If
            Dim bHadUpdate As Boolean = False
            For X As Int32 = mvecTrailVec.GetUpperBound(0) To 0 Step -1  '0 To mvecTrailVec.GetUpperBound(0)
                If bResult = True Then
                    Dim clrTemp As System.Drawing.Color = mclrTrailClr(X)
                    Dim lA As Int32 = clrTemp.A - 15
                    If lA < 0 Then lA = 0
                    mclrTrailClr(X) = System.Drawing.Color.FromArgb(lA, clrTemp.R, clrTemp.G, clrTemp.B)
                End If
                If bUpdateNoRender = False Then
                    Dim vecPos As Vector3 = vecLoc
                    vecPos.Add(mvecTrailVec(X))
                    'If X = 1 Then
                    '    Dim lLowVal As Int32 = 0
                    '    If mclrTrailClr(X) = System.Drawing.Color.FromArgb(255, lLowVal, lLowVal, lLowVal) Then
                    '        mclrTrailClr(X) = System.Drawing.Color.FromArgb(255, 255, 255, 255)
                    '    Else
                    '        mclrTrailClr(X) = System.Drawing.Color.FromArgb(255, lLowVal, lLowVal, lLowVal)
                    '    End If
                    'End If
                    oSpr.Draw(vecPos, mfTrailSize(X), mclrTrailClr(X), moRandom.Next(0, 360))
                End If
            Next X

            Return bResult
        End Function

        Public Sub New(ByVal vecOrigin As Vector3, ByVal lMinY As Int32, ByVal clrBomb As System.Drawing.Color, ByVal vel As Vector3)
            vecLoc = vecOrigin
            lFloor = lMinY
            clrBase = clrBomb
            mlLastUpdate = glCurrentCycle
            vecVel = vel

            Dim lTrailUB As Int32 = 35
            ReDim mvecTrailVec(lTrailUB)
            ReDim mclrTrailClr(lTrailUB)
            ReDim mfTrailSize(lTrailUB)

            'Now, configure...
            Dim clrTemp As System.Drawing.Color = clrBomb
            For X As Int32 = 0 To lTrailUB
                If X < 5 Then
                    If X < 3 Then
                        If X = 0 Then
                            mclrTrailClr(X) = System.Drawing.Color.FromArgb(255, 255, 255, 255)
                        Else
                            mclrTrailClr(X) = System.Drawing.Color.FromArgb(255, Math.Min(clrBase.R * 2, 255), Math.Min(clrBase.G * 2, 255), Math.Min(clrBase.B * 2, 255))

                        End If

                        mfTrailSize(X) = 40 + (3 * X)
                        mvecTrailVec(X) = New Vector3(0, 3 * X, 0)
                    Else
                        If X = 3 Then
                            mvecTrailVec(X) = New Vector3(0, 0, 0)
                        Else
                            mvecTrailVec(X) = New Vector3(0, 0, 0)
                        End If
                        mclrTrailClr(X) = clrTemp
                        mfTrailSize(X) = 40 + (5 * X)     '30 + 2 * x
                    End If


                    'Dim lR As Int32 = clrTemp.R - 25
                    'Dim lG As Int32 = clrTemp.G - 25
                    'Dim lB As Int32 = clrTemp.B - 25
                    'Dim lA As Int32 = clrTemp.A
                    'If lR < 0 Then lR = 0
                    'If lG < 0 Then lG = 0
                    'If lB < 0 Then lB = 0
                    'If lR = 0 AndAlso lG = 0 AndAlso lB = 0 Then
                    '    lA -= 12
                    '    If lA < 0 Then
                    '        lA = 0
                    '    End If
                    'End If
                    'clrTemp = System.Drawing.Color.FromArgb(lA, lR, lG, lB)

                Else
                    Dim lMult As Int32 = X - 4
                    mvecTrailVec(X) = New Vector3(-vel.X * lMult, 32 * lMult, -vel.Z * lMult)        '16
                    mclrTrailClr(X) = clrTemp
                    mfTrailSize(X) = 50 + (10 * (X - 4))     '30 + 2 * x
                    Dim lR As Int32 = clrTemp.R - 25
                    Dim lG As Int32 = clrTemp.G - 25
                    Dim lB As Int32 = clrTemp.B - 25
                    Dim lA As Int32 = clrTemp.A
                    If lR < 0 Then lR = 0
                    If lG < 0 Then lG = 0
                    If lB < 0 Then lB = 0
                    If lR = 0 AndAlso lG = 0 AndAlso lB = 0 Then
                        lA -= 12
                        If lA < 0 Then
                            lA = 0
                        End If
                    End If
                    clrTemp = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
                End If
            Next X

        End Sub
    End Class

    Private moShots() As BombShot

    Public Sub Update(ByVal bUpdateNoRender As Boolean)

        Dim oSpr As BPSprite = Nothing
        glBombFXRendered = 0
        If moShots Is Nothing = False Then
            For X As Int32 = 0 To moShots.GetUpperBound(0)
                If moShots(X) Is Nothing = False Then
                    If bUpdateNoRender = False AndAlso oSpr Is Nothing Then
                        oSpr = New BPSprite()
                        oSpr.BeginRender(0, goResMgr.GetTexture("p3.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "xpl.pak"), GFXEngine.moDevice)
                    Else
                        glBombFXRendered += 1
                    End If

                    If moShots(X).Update(oSpr, bUpdateNoRender) = True Then
                        'weapon end - do AOE, etc...
                        Dim lMinSndDist As Int32
                        Dim lMaxSndDist As Int32
                        If goSound Is Nothing = False Then
                            lMinSndDist = SoundMgr.lMinDistance ' gosound.lMinDistance = 1
                            lMaxSndDist = SoundMgr.lMaxDistance
                            SoundMgr.lMinDistance = 2000
                            SoundMgr.lMaxDistance = 60000
                        End If
                        
                        If moShots(X).yAOE > 0 Then
                            AOEExplosionMgr.goMgr.AddNew(moShots(X).vecLoc, Math.Min(255, CInt(moShots(X).yAOE) * 20I))      'verify this   was x20
                        ElseIf goSound Is Nothing = False Then
                            If Rnd() * 100 < 50 Then
                                goSound.StartSound("Explosions\MediumGroundDeath1.wav", False, SoundMgr.SoundUsage.eNonDeathExplosions, moShots(X).vecLoc, New Vector3(0, 0, 0))
                            Else
                                goSound.StartSound("Explosions\MediumGroundDeath2.wav", False, SoundMgr.SoundUsage.eNonDeathExplosions, moShots(X).vecLoc, New Vector3(0, 0, 0))
                            End If
                        End If
                        If goSound Is Nothing = False Then
                            With moShots(X).vecLoc
                                'If Math.Abs(.X - goCamera.mlCameraX) > 45000 OrElse Math.Abs(.Y - goCamera.mlCameraY) > 45000 OrElse Math.Abs(.Z - goCamera.mlCameraZ) > 45000 Then
                                SoundMgr.lMinDistance = 20000
                                SoundMgr.lMaxDistance = 120000
                                goSound.StartSound("DistantBomb.wav", False, SoundMgr.SoundUsage.eNonDeathExplosions, moShots(X).vecLoc, New Vector3(0, 0, 0))
                                'End If
                            End With
                            SoundMgr.lMinDistance = lMinSndDist
                            SoundMgr.lMaxDistance = lMaxSndDist
                        End If
                        moShots(X) = Nothing
                    End If
                End If
            Next X
            If oSpr Is Nothing = False Then
                GFXEngine.moDevice.Transform.World = Matrix.Identity
                oSpr.EndRender()
                oSpr.DisposeMe()
            End If
        End If

    End Sub

    Public Sub AddNewBomb(ByVal lFromX As Int32, ByVal lFromY As Int32, ByVal lFromZ As Int32, ByVal lToX As Int32, ByVal lToY As Int32, ByVal lToZ As Int32, ByVal yWpnType As Byte, ByVal yAOE As Byte)
        Dim clrBomb As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        Select Case yWpnType
            Case WeaponType.eBomb_Green
                clrBomb = System.Drawing.Color.FromArgb(255, 64, 255, 64)
            Case WeaponType.eBomb_Purple
                clrBomb = System.Drawing.Color.FromArgb(255, 255, 64, 255)
            Case WeaponType.eBomb_Red
                clrBomb = System.Drawing.Color.FromArgb(255, 255, 64, 64)
            Case WeaponType.eBomb_Teal
                clrBomb = System.Drawing.Color.FromArgb(255, 64, 255, 255)
            Case WeaponType.eBomb_Yellow
                clrBomb = System.Drawing.Color.FromArgb(255, 255, 255, 64)
                'Case Else 'WeaponType.eBomb_Gray
        End Select

        'NOTE: A neat effect might be to have the velocity be based on the gravity of the planet...

        Dim vecFrom As Vector3 = New Vector3(lFromX, lFromY, lFromZ)
        Dim vecTo As Vector3 = New Vector3(lToX, lToY, lToZ)
        Dim fDescent As Single = -54

        Dim fSteps As Single = Math.Abs((lFromY - lToY) / fDescent)
        Dim fDX As Single = (lToX - lFromX) / fSteps
        Dim fDZ As Single = (lToZ - lFromZ) / fSteps

        Dim vecVel As Vector3 = New Vector3(fDX, fDescent, fDZ)

        Dim oNew As New BombShot(vecFrom, lToY, clrBomb, vecVel)
        oNew.yAOE = yAOE

        If goSound Is Nothing = False Then
            If goSound.MusicOn = True Then
                goSound.IncrementExcitementLevel(5) 'we are attacking + 15
            End If
        End If

        If goWpnMgr.GenerateSoundFX = True AndAlso goSound Is Nothing = False Then
            'goSound.StartSound(sSoundFile, False, ySFXType, New Vector3(lToX, lToY, lToZ), New Vector3(0, 0, 0))
            Dim lMinSndDist As Int32 = SoundMgr.lMinDistance ' gosound.lMinDistance = 1
            Dim lMaxSndDist As Int32 = SoundMgr.lMaxDistance
            SoundMgr.lMinDistance = 2000
            SoundMgr.lMaxDistance = 55000
            goSound.StartSound("Bomb.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(lFromX, lFromY, lFromZ), vecVel)
            SoundMgr.lMinDistance = lMinSndDist
            SoundMgr.lMaxDistance = lMaxSndDist
        End If

        If moShots Is Nothing Then ReDim moShots(-1)

        For X As Int32 = 0 To moShots.GetUpperBound(0)
            If moShots(X) Is Nothing Then
                moShots(X) = oNew
                Return
            End If
        Next X
        If moShots Is Nothing = False AndAlso moShots.GetUpperBound(0) > 100 Then Return
        ReDim Preserve moShots(moShots.GetUpperBound(0) + 1)
        moShots(moShots.GetUpperBound(0)) = oNew

    End Sub

End Class
