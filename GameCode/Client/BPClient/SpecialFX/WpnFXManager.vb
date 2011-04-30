Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class WpnFXManager
    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32
    Private MustInherit Class BaseWpnEffect
        Protected ml_PULSE_WPN_SPEED As Int32 = 120
        Protected ml_PROJ_WPN_SPEED As Int32 = 120

        Public lIndex As Int32      'index in the Manager array
        Public fDestX As Single     'where the effect goes to X
        Public fDestY As Single     'where the effect goes to Y
        Public fDestZ As Single     'where the effect goes to Z
        Public fCurrentX As Single  'where the effect comes from or is currently located at X
        Public fCurrentY As Single  'where the effect comes from or is currently located at Y
        Public fCurrentZ As Single  'where the effect comes from or is currently located at Z

        Public fPrevY As Single     'for single shot

        Public oTarget As RenderObject      'the target being shot at
        Public oTargetMissile As MissileMgr.Missile
        Public oAttacker As RenderObject    'the attacker firing the shot

        Public colWpnColor As System.Drawing.Color

        Private myWeaponType As Byte        'from the WeaponType enum 
        Public WpnState As Single = 0
        Protected mlWpnStateCnt As Int32
        Public WeaponHit As Boolean     'if true, the weapon hits, else, it misses...

        Public WpnStateRect() As Rectangle
        Public vecCenter As Vector3

        Public WpnMaxSpeed As Int32

        Public vecOriginalFireFrom As Vector3

        Protected mlPreviousFrame As Int32
        Protected mbEffectActive As Boolean

        Public bUsesSprite As Boolean       'for optimization purposes

        Public vecOnModelTo As Vector3
        Public ySideHit As Byte

        Public MustOverride Function Update() As Boolean

        Public fSizeMult As Single = 1.0F

        Public yAOE As Byte = 0

        'Public Shared Function QuickTransformCoordinate(ByVal vecCoord As Vector3, ByVal mat As Matrix) As Vector3
        '    'With matTemp
        '    '    .M11 = xAxis.X : .M12 = xAxis.Y : .M13 = xAxis.Z : .M14 = 0
        '    '    .M21 = yAxis.X : .M22 = yAxis.Y : .M23 = yAxis.Z : .M24 = 0
        '    '    .M31 = zAxis.X : .M32 = zAxis.Y : .M33 = zAxis.Z : .M34 = 0
        '    '    '.M41 = vecPosition.X : .M42 = vecPosition.Y : .M43 = vecPosition.Z : .M44 = 1
        '    'End With
        '    'With mat
        '    '    vecCol1 = New Vector3(.M11, .M21, .M31) '  New Vector3(xAxis.X, yAxis.X, zAxis.X)
        '    '    vecCol2 = New Vector3(.M12, .M22, .M32) '  New Vector3(xAxis.Y, yAxis.Y, zAxis.Y)
        '    '    vecCol3 = New Vector3(.M13, .M23, .M33) ' New Vector3(xAxis.Z, yAxis.Z, zAxis.Z)
        '    'End With

        '    With mat
        '        Return New Vector3((vecCoord.X * .M11) + (vecCoord.Y * .M21) + (vecCoord.Z * .M31) + .M41, (vecCoord.X * .M12) + (vecCoord.Y * .M22) + (vecCoord.Z * .M32) + .M42, (vecCoord.X * .M13) + (vecCoord.Y * .M23) + (vecCoord.Z * .M33) + .M43)
        '        'Dim fX As Single = (vecCoord.X * .M11) + (vecCoord.Y * .M21) + (vecCoord.Z * .M31) + .M41
        '        'Dim fY As Single = (vecCoord.X * .M12) + (vecCoord.Y * .M22) + (vecCoord.Z * .M32) + .M42
        '        'Dim fZ As Single = (vecCoord.X * .M13) + (vecCoord.Y * .M23) + (vecCoord.Z * .M33) + .M43
        '        'Return New Vector3(fX, fY, fZ)
        '    End With
        'End Function

        Public Property EffectActive() As Boolean
            Get
                Return mbEffectActive
            End Get
            Set(ByVal Value As Boolean)
                If mbEffectActive = False AndAlso Value = True Then
                    'turning on, set our last update
                    mlPreviousFrame = timeGetTime
                End If
                mbEffectActive = Value
            End Set
        End Property

        Public Property WeaponTypeID() As Byte
            Get
                Return myWeaponType
            End Get
            Set(ByVal Value As Byte)
                myWeaponType = Value

                Select Case myWeaponType
                    Case WeaponType.eFlickerGreenBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        WpnStateRect(1) = New Rectangle(0, 32, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 64, 255, 64)
                        vecCenter = New Vector3(32, 16, 0)
                    Case WeaponType.eFlickerPurpleBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        WpnStateRect(1) = New Rectangle(0, 32, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 64, 255)
                        vecCenter = New Vector3(32, 16, 0)
                    Case WeaponType.eFlickerRedBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        WpnStateRect(1) = New Rectangle(0, 32, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 64, 64)
                        vecCenter = New Vector3(32, 16, 0)
                    Case WeaponType.eFlickerTealBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        WpnStateRect(1) = New Rectangle(0, 32, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 64, 255, 255)
                        vecCenter = New Vector3(32, 16, 0)
                    Case WeaponType.eFlickerYellowBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        WpnStateRect(1) = New Rectangle(0, 32, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 255, 64)
                        vecCenter = New Vector3(32, 16, 0)
                    Case WeaponType.eMetallicProjectile_Bronze
                        If glCurrentEnvirView = CurrentView.eWeaponBuilder Then WpnMaxSpeed = 50 Else WpnMaxSpeed = ml_PROJ_WPN_SPEED
                        WpnStateCnt = 1
                        'WpnStateRect(0) = New Rectangle(64, 32, 32, 32)
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 242, 164, 64)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eMetallicProjectile_Copper
                        If glCurrentEnvirView = CurrentView.eWeaponBuilder Then WpnMaxSpeed = 50 Else WpnMaxSpeed = ml_PROJ_WPN_SPEED
                        WpnStateCnt = 1
                        'WpnStateRect(0) = New Rectangle(64, 32, 32, 32)
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 64, 128)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eMetallicProjectile_Gold
                        If glCurrentEnvirView = CurrentView.eWeaponBuilder Then WpnMaxSpeed = 50 Else WpnMaxSpeed = ml_PROJ_WPN_SPEED
                        WpnStateCnt = 1
                        'WpnStateRect(0) = New Rectangle(64, 32, 32, 32)
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 230, 64)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eMetallicProjectile_Lead
                        If glCurrentEnvirView = CurrentView.eWeaponBuilder Then WpnMaxSpeed = 50 Else WpnMaxSpeed = ml_PROJ_WPN_SPEED
                        WpnStateCnt = 1
                        'WpnStateRect(0) = New Rectangle(64, 32, 32, 32)
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eMetallicProjectile_Silver
                        If glCurrentEnvirView = CurrentView.eWeaponBuilder Then WpnMaxSpeed = 50 Else WpnMaxSpeed = ml_PROJ_WPN_SPEED
                        WpnStateCnt = 1
                        'WpnStateRect(0) = New Rectangle(64, 32, 32, 32)
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eShortGreenPulse
                        WpnMaxSpeed = ml_PULSE_WPN_SPEED
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        WpnStateRect(1) = New Rectangle(96, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 64, 255, 64)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eShortPurplePulse
                        WpnMaxSpeed = ml_PULSE_WPN_SPEED
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        WpnStateRect(1) = New Rectangle(96, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 64, 255)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eShortRedPulse
                        WpnMaxSpeed = ml_PULSE_WPN_SPEED
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        WpnStateRect(1) = New Rectangle(96, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 64, 64)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eShortTealPulse
                        WpnMaxSpeed = ml_PULSE_WPN_SPEED
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        WpnStateRect(1) = New Rectangle(96, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 64, 255, 255)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eShortYellowPulse
                        WpnMaxSpeed = ml_PULSE_WPN_SPEED
                        WpnStateCnt = 2
                        WpnStateRect(0) = New Rectangle(64, 0, 32, 32)
                        WpnStateRect(1) = New Rectangle(96, 0, 32, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 255, 64)
                        vecCenter = New Vector3(16, 16, 0)
                    Case WeaponType.eSolidGreenBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 1
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 64, 255, 64)
                        vecCenter = New Vector3(32, 16, 0)
                    Case WeaponType.eSolidPurpleBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 1
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 64, 255)
                        vecCenter = New Vector3(32, 16, 0)
                    Case WeaponType.eSolidRedBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 1
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 64, 64)
                        vecCenter = New Vector3(32, 16, 0)
                    Case WeaponType.eSolidTealBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 1
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 64, 255, 255)
                        vecCenter = New Vector3(32, 16, 0)
                    Case WeaponType.eSolidYellowBeam
                        WpnMaxSpeed = Int32.MaxValue
                        WpnStateCnt = 1
                        WpnStateRect(0) = New Rectangle(0, 0, 64, 32)
                        colWpnColor = System.Drawing.Color.FromArgb(255, 255, 255, 64)
                        vecCenter = New Vector3(32, 16, 0)
                End Select

            End Set
        End Property

        Public Property WpnStateCnt() As Int32
            Get
                Return mlWpnStateCnt
            End Get
            Set(ByVal Value As Int32)
                mlWpnStateCnt = Value

                ReDim WpnStateRect(mlWpnStateCnt - 1)
            End Set
        End Property

        'Private myPrevMapWrapSituation As Byte
        'Public Sub UpdateMapWrapSituation(ByVal yMapWrapSituation As Byte, ByVal lLocXMapWrapCheck As Integer)
        '    If yMapWrapSituation <> myPrevMapWrapSituation Then

        '        'Reset our loc
        '        If fCurrentX < goCurrentEnvir.lMinXPos Then
        '            fCurrentX += goCurrentEnvir.lMapWrapAdjustX
        '        ElseIf fCurrentX > goCurrentEnvir.lMaxXPos Then
        '            fCurrentX -= goCurrentEnvir.lMapWrapAdjustX
        '        End If
        '        If fDestX < goCurrentEnvir.lMinXPos Then
        '            fDestX += goCurrentEnvir.lMapWrapAdjustX
        '        ElseIf fDestX > goCurrentEnvir.lMaxXPos Then
        '            fDestX -= goCurrentEnvir.lMapWrapAdjustX
        '        End If

        '        If yMapWrapSituation = 1 Then
        '            'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
        '            If fCurrentX > lLocXMapWrapCheck Then
        '                fCurrentX -= goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '            If fDestX > lLocXMapWrapCheck Then
        '                fDestX -= goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '        Else
        '            'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
        '            If fCurrentX < lLocXMapWrapCheck Then
        '                fCurrentX += goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '            If fDestX < lLocXMapWrapCheck Then
        '                fDestX += goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '        End If
        '    End If
        '    myPrevMapWrapSituation = yMapWrapSituation
        'End Sub
    End Class

    Private Class SingleShotWpn     'pulse lasers, projectile gun fire, mortars, shells, etc...
        Inherits BaseWpnEffect

        'Modifiers to the movement pattern of the shot
        Public lModX As Int32
        Public lModY As Int32
        Public lModZ As Int32

        Private mvecPrevTarget As Vector3
        Private mlLastTargetUpdate As Int32 = 0


        Public Overrides Function Update() As Boolean
            Dim fElapsed As Single
            Dim fTotDist As Single
            Dim fTotVel As Single
            Dim fVelX As Single
            Dim fVelY As Single
            Dim fVelZ As Single

            Dim bResult As Boolean = False

            fElapsed = (timeGetTime - mlPreviousFrame) / 30.0F

            fPrevY = fCurrentY

            If EffectActive = True AndAlso fElapsed <> 0.0F Then
                If WeaponHit = True AndAlso ((oTarget Is Nothing = False AndAlso oTarget.oMesh Is Nothing = False) OrElse oTargetMissile Is Nothing = False) Then

                    Dim fTmpDX As Single
                    If oTarget Is Nothing = False Then
                        
                        If glCurrentCycle - mlLastTargetUpdate > 30 Then
                            mlLastTargetUpdate = glCurrentCycle

                            'fDestY = oTarget.LocY + oTarget.oMesh.YMidPoint
                            'fDestZ = oTarget.LocZ

                            'Dim vecTemp As Vector3 = vecOnModelTo
                            'Dim vecRes As Vector3 = Vector3.TransformCoordinate(vecTemp, oTarget.GetWorldMatrix)
                            'Dim vecRes As Vector3 '= QuickTransformCoordinate(vecTemp, oTarget.GetWorldMatrix) 'New Vector3(oTarget.LocX, oTarget.LocY, oTarget.LocZ)
                            With oTarget.GetWorldMatrix
                                Dim fX As Single = vecOnModelTo.X
                                Dim fY As Single = vecOnModelTo.Y
                                Dim fZ As Single = vecOnModelTo.Z
                                mvecPrevTarget.X = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                                mvecPrevTarget.Y = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                                mvecPrevTarget.Z = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
                            End With
                            'fTmpDX = vecRes.X
                            'fDestY = vecRes.Y
                            'fDestZ = vecRes.Z
                        End If
                        fTmpDX = mvecPrevTarget.X
                        fDestY = mvecPrevTarget.Y
                        fDestZ = mvecPrevTarget.Z

                    Else
                        fTmpDX = oTargetMissile.vecMissile.X
                        fDestY = oTargetMissile.vecMissile.Y
                        fDestZ = oTargetMissile.vecMissile.Z
                    End If

                    'If goCurrentEnvir Is Nothing = False Then
                    '    If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                    '        If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then

                    '            'Now, determine if map wrapping is the better choice
                    '            Dim oPlanet As Planet = CType(goCurrentEnvir.oGeoObject, Planet)
                    '            Dim lTotalWidth As Int32 = TerrainClass.Width * oPlanet.CellSpacing

                    '            'Check for a map wrap crossover...
                    '            If goCamera.mlCameraX < 0 Then
                    '                'Ok, if our loc is > map extent
                    '                If fCurrentX > goCurrentEnvir.lMaxXPos OrElse fDestX > goCurrentEnvir.lMaxXPos Then
                    '                    fCurrentX -= lTotalWidth
                    '                    fTmpDX -= lTotalWidth
                    '                End If
                    '            Else
                    '                'Ok, if our loc is < map extent
                    '                If fCurrentX < goCurrentEnvir.lMinXPos OrElse fDestX < goCurrentEnvir.lMinXPos Then
                    '                    fCurrentX += lTotalWidth
                    '                    fTmpDX += lTotalWidth
                    '                End If
                    '            End If

                    '            If Math.Abs(goCamera.mlCameraX - fTmpDX) > Math.Abs(goCamera.mlCameraX - fCurrentX) Then
                    '                If fTmpDX < fCurrentX Then
                    '                    'ok, normally, we would go left...
                    '                    Dim fTmp As Single = fTmpDX + lTotalWidth
                    '                    If Math.Abs(fTmp - fCurrentX) < Math.Abs(fCurrentX - fTmpDX) Then fTmpDX = fTmp
                    '                Else
                    '                    'ok, normally, we would go right..
                    '                    Dim fTmp As Single = fTmpDX - lTotalWidth
                    '                    If Math.Abs(fCurrentX - fTmp) < Math.Abs(fCurrentX - fTmpDX) Then fTmpDX = fTmp
                    '                End If
                    '            End If
                    '        End If
                    '    End If
                    'End If

                    fDestX = fTmpDX
                End If

                'Special Effects Weapon Fire
                fTotDist = System.Math.Abs(fDestX - fCurrentX) + System.Math.Abs(fDestZ - fCurrentZ) + System.Math.Abs(fDestY - fCurrentY)
                fTotVel = fElapsed * WpnMaxSpeed
                If fTotVel = 0.0F OrElse fTotDist = 0 Then
                    bResult = True
                Else
                    mlPreviousFrame = timeGetTime

                    If fTotDist > 0.0000001F Then
                        fTotDist = 1.0F / fTotDist
                        fVelX = ((fDestX - fCurrentX) * fTotDist) * fTotVel
                        fVelY = ((fDestY - fCurrentY) * fTotDist) '* fTotVel
                        fPrevY = (fVelY * -40)
                        fVelY *= fTotVel
                        fVelZ = ((fDestZ - fCurrentZ) * fTotDist) * fTotVel
                    Else
                        fVelX = ((fDestX - fCurrentX) / fTotDist) * fTotVel
                        fVelY = ((fDestY - fCurrentY) / fTotDist) '* fTotVel
                        fPrevY = (fVelY * -40)
                        fVelY *= fTotVel
                        fVelZ = ((fDestZ - fCurrentZ) / fTotDist) * fTotVel
                    End If

                    fCurrentX += fVelX
                    fCurrentY += fVelY
                    fCurrentZ += fVelZ
                    fPrevY += fCurrentY

                    If Math.Abs(fCurrentX - fDestX) < fVelX Then fCurrentX = fDestX
                    If Math.Abs(fCurrentY - fDestY) < fVelY Then fCurrentY = fDestY
                    If Math.Abs(fCurrentZ - fDestZ) < fVelZ Then fCurrentZ = fDestZ

                    'Increment our wpn state
                    WpnState += fElapsed
                    If WpnState >= WpnStateCnt Then WpnState = 0.0F

                    'return whether this wpn event is over or not
                    bResult = fDestX = fCurrentX AndAlso fDestY = fCurrentY AndAlso fDestZ = fCurrentZ
                End If

                If bResult = True Then
                    If oTarget Is Nothing = False Then
                        Dim vecBurn As Vector3 = vecOnModelTo
                        If vecOnModelTo.X > 0 Then
                            vecBurn.X += 5
                        Else
                            vecBurn.X -= 5
                        End If
                        oTarget.AddBurnMark(vecBurn, 1, Rnd() * 64 + 24)
                    End If
                End If

            End If
            Return bResult
        End Function

    End Class

    Private Class BeamWpn           'Solid beam lasers
        Inherits BaseWpnEffect

        Private Const ml_BEAM_SIZE As Int32 = 13        'TODO: we could make capital beams larger...
        Private Const ml_TEXTURE_WIDTH As Int32 = 128
        Private Const ml_TEXTURE_HEIGHT As Int32 = 128
        Private Const ml_DURATION As Int32 = 2000

        Public muVerts1(3) As CustomVertex.PositionTextured
        Public muVerts2(3) As CustomVertex.PositionTextured

        Private mfTuLow As Single
        Private mfTvLow As Single
        Private mfTuHi As Single
        Private mfTvHi As Single

        Private mlStart As Int32 = Int32.MinValue

        Private mlLastState As Int32
        Private mfSizeMult As Single = 0.3

        Private mfExplosions As Single = 0.0F
        Private mvecDestLoc As Vector3 = Vector3.Empty
        Private mvecLastModelTo As Vector3 '= vecOnModelTo 

        Public Overrides Function Update() As Boolean
            Dim fElapsed As Single

            If EffectActive = True Then
                If mvecDestLoc = Vector3.Empty Then
                    mvecLastModelTo = Me.vecOnModelTo
                    mvecDestLoc = Me.vecOnModelTo
                End If

                Dim bAddBurnToTarget As Boolean = False

                If mlPreviousFrame = 0 Then fElapsed = 1.0F Else fElapsed = (timeGetTime - mlPreviousFrame) / 30.0F
                mlPreviousFrame = timeGetTime

                'Increase our Size
                If mlStart = Int32.MinValue Then
                    mfSizeMult += (0.03F * fElapsed)
                    If mfSizeMult > 1 Then
                        mfSizeMult = 1
                        mlStart = timeGetTime
                    End If
                ElseIf WpnStateCnt = 1 AndAlso (timeGetTime - mlStart) > 2000 Then
                    mfSizeMult -= (0.02F * fElapsed)
                    If mfSizeMult < 0 Then mfSizeMult = 0
                ElseIf WpnStateCnt = 2 AndAlso (timeGetTime - mlStart) > 800 Then
                    mfSizeMult -= (0.08F * fElapsed)
                    If mfSizeMult < 0 Then mfSizeMult = 0
                End If

                If oAttacker Is Nothing = False Then ' AndAlso oAttacker.oMesh Is Nothing = False Then
                    With oAttacker.GetWorldMatrix
                        Dim fX As Single = vecOriginalFireFrom.X
                        Dim fY As Single = vecOriginalFireFrom.Y
                        Dim fZ As Single = vecOriginalFireFrom.Z
                        fCurrentX = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                        fCurrentY = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                        fCurrentZ = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
                    End With

                    'Dim vecFrom As Vector3 = Me.vecOriginalFireFrom
                    'vecFrom = QuickTransformCoordinate(vecFrom, oAttacker.GetWorldMatrix)
                    'fCurrentX = vecFrom.X
                    'fCurrentY = vecFrom.Y
                    'fCurrentZ = vecFrom.Z
                End If

                Dim vecActualEnd As Vector3 = vecOnModelTo

                If WeaponHit = True Then
                    If (oTarget Is Nothing = False AndAlso oTarget.oMesh Is Nothing = False) OrElse oTargetMissile Is Nothing = False Then  'If oTarget Is Nothing = False AndAlso oTarget.oMesh Is Nothing = False Then

                        If oTarget Is Nothing = False Then
                            If vecOnModelTo = mvecDestLoc Then
                                mvecDestLoc = oTarget.oMesh.GetRandomMeshPoint(ySideHit)
                            End If
                            Dim vecDiff As Vector3 = mvecDestLoc - vecOnModelTo
                            vecDiff.Normalize()
                            vecDiff.Multiply(fElapsed * 2)
                            vecOnModelTo += vecDiff

                            'ok, determine where on the unit the beam is presently hitting
                            'Dim oInt As IntersectInformation
                            'Dim vecDir As Vector3 = vecOnModelTo
                            'vecDir.Normalize()
                            'Dim bTruHit As Boolean = oTarget.oMesh.oMesh.Intersect(New Vector3(0, 0, 0), vecDir, oInt)
                            'vecOnModelTo = Vector3.Multiply(vecDir, oInt.Dist + 50)
                            Dim vecFrom As Vector3 ' = New Vector3(fCurrentX, fCurrentY, fCurrentZ)
                            Dim matTarget As Matrix = oTarget.GetWorldMatrix
                            Dim matInvTarget As Matrix = Matrix.Invert(matTarget)
                            'matTemp.Invert()
                            'vecFrom = Vector3.TransformCoordinate(vecFrom, matTemp)
                            With matInvTarget
                                Dim fX As Single = fCurrentX
                                Dim fY As Single = fCurrentY
                                Dim fZ As Single = fCurrentZ
                                vecFrom.X = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                                vecFrom.Y = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                                vecFrom.Z = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
                            End With
                            'vecFrom = QuickTransformCoordinate(vecFrom, matTemp)

                            'vecfrom is in model space now
                            'vecFrom = vecFrom - vecOnModelTo
                            Dim oInt As IntersectInformation
                            Dim vecDir As Vector3 = vecFrom - vecOnModelTo
                            vecDir.Multiply(-1)
                            vecDir.Normalize()
                            Dim bTruHit As Boolean = oTarget.oMesh.oMesh.Intersect(vecFrom, vecDir, oInt)
                            If bTruHit = False Then
                                vecActualEnd = vecOnModelTo
                            Else
                                vecActualEnd = vecFrom + (vecDir * oInt.Dist)
                            End If

                            If Vector3.Subtract(mvecLastModelTo, vecActualEnd).Length > 30 AndAlso bTruHit = True Then
                                mvecLastModelTo = vecActualEnd
                                Dim vecBurn As Vector3 = vecActualEnd

                                Select Case ySideHit
                                    Case 0
                                        vecBurn.Z -= 5
                                    Case 1
                                        vecBurn.X -= 5
                                    Case 2
                                        vecBurn.Z += 5
                                    Case 3
                                        vecBurn.X += 5
                                End Select
                                oTarget.AddBurnMark(vecBurn, ySideHit, Rnd() * 64 + 24)
                            End If


                            With matTarget
                                Dim fX As Single = vecActualEnd.X
                                Dim fY As Single = vecActualEnd.Y
                                Dim fZ As Single = vecActualEnd.Z
                                fDestX = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                                fDestY = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                                fDestZ = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
                            End With
                            'Dim vecTemp As Vector3 = vecActualEnd
                            'Dim vecRes As Vector3 = QuickTransformCoordinate(vecTemp, oTarget.GetWorldMatrix) 'New Vector3(oTarget.LocX, oTarget.LocY, oTarget.LocZ) '
                            'fDestX = vecRes.X
                            'fDestY = vecRes.Y
                            'fDestZ = vecRes.Z
                        End If

                        Dim fTmpDX As Single '= oTarget.LocX
                        If oTarget Is Nothing = False Then fTmpDX = fDestX Else fTmpDX = oTargetMissile.vecMissile.X
                        'If goCurrentEnvir Is Nothing = False Then
                        '    If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                        '        If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                        '            'Now, determine if map wrapping is the better choice
                        '            Dim oPlanet As Planet = CType(goCurrentEnvir.oGeoObject, Planet)
                        '            Dim lTotalWidth As Int32 = TerrainClass.Width * oPlanet.CellSpacing

                        '            'Check for a map wrap crossover...
                        '            If goCamera.mlCameraX < 0 Then
                        '                'Ok, if our loc is > map extent
                        '                If fCurrentX > goCurrentEnvir.lMaxXPos OrElse fDestX > goCurrentEnvir.lMaxXPos Then
                        '                    fCurrentX -= lTotalWidth
                        '                    fTmpDX -= lTotalWidth
                        '                End If
                        '            Else
                        '                'Ok, if our loc is < map extent
                        '                If fCurrentX < goCurrentEnvir.lMinXPos OrElse fDestX < goCurrentEnvir.lMinXPos Then
                        '                    fCurrentX += lTotalWidth
                        '                    fTmpDX += lTotalWidth
                        '                End If
                        '            End If

                        '            If Math.Abs(goCamera.mlCameraX - fTmpDX) > Math.Abs(goCamera.mlCameraX - fCurrentX) Then
                        '                If fTmpDX < fCurrentX Then
                        '                    'ok, normally, we would go left...
                        '                    Dim fTmp As Single = fTmpDX + lTotalWidth
                        '                    If Math.Abs(fTmp - fCurrentX) < Math.Abs(fCurrentX - fTmpDX) Then fTmpDX = fTmp
                        '                Else
                        '                    'ok, normally, we would go right..
                        '                    Dim fTmp As Single = fTmpDX - lTotalWidth
                        '                    If Math.Abs(fCurrentX - fTmp) < Math.Abs(fCurrentX - fTmpDX) Then fTmpDX = fTmp
                        '                End If
                        '            End If
                        '        End If
                        '    End If
                        'End If

                        fDestX = fTmpDX
                        If oTarget Is Nothing Then
                            fDestY = oTargetMissile.vecMissile.Y
                            fDestZ = oTargetMissile.vecMissile.Z
                        End If
                    End If
                Else
                    fDestX += Rnd() * fElapsed
                    fDestY += Rnd() * fElapsed
                    fDestZ += Rnd() * fElapsed
                    vecActualEnd.X = fDestX : vecActualEnd.Y = fDestY : vecActualEnd.Z = fDestZ
                End If

                mfExplosions += fElapsed
                If mfExplosions > 2 AndAlso WeaponHit = True Then
                    mfExplosions = 0
                    If goExplMgr Is Nothing = False Then goExplMgr.AddBeamExplosionToEntity(oTarget, WeaponTypeID, vecActualEnd)
                End If


                'Increment our wpn state
                WpnState += fElapsed
                If WpnState > WpnStateCnt - 1 Then WpnState = 0.0F


                Dim fWpnStateSizeMult As Single = 1.0F
                If mlLastState <> CInt(WpnState) OrElse (mfTuLow = 0 AndAlso mfTvLow = 0 AndAlso mfTvHi = 0 AndAlso mfTvLow = 0) Then
                    mlLastState = CInt(WpnState)

                    With WpnStateRect(mlLastState)
                        mfTuLow = CSng(.Left / ml_TEXTURE_WIDTH)
                        mfTuHi = CSng(.Right / ml_TEXTURE_WIDTH)
                        mfTvLow = CSng(.Top / ml_TEXTURE_HEIGHT)
                        mfTvHi = CSng(.Bottom / ml_TEXTURE_HEIGHT)
                    End With

                End If

                If WpnStateCnt = 2 Then fWpnStateSizeMult = Rnd()

                Dim fSize As Single = ml_BEAM_SIZE * (mfSizeMult * MyBase.fSizeMult) * fWpnStateSizeMult

                muVerts1(0) = New CustomVertex.PositionTextured(fCurrentX - fSize, fCurrentY, fCurrentZ + (fSize / 2), mfTuLow, mfTvLow)
                muVerts1(1) = New CustomVertex.PositionTextured(fCurrentX + fSize, fCurrentY, fCurrentZ - (fSize / 2), mfTuLow, mfTvHi)
                muVerts1(2) = New CustomVertex.PositionTextured(fDestX - fSize, fDestY, fDestZ + (fSize / 2), mfTuHi, mfTvLow)
                muVerts1(3) = New CustomVertex.PositionTextured(fDestX + fSize, fDestY, fDestZ - (fSize / 2), mfTuHi, mfTvHi)

                muVerts2(0) = New CustomVertex.PositionTextured(fCurrentX + (fSize / 2), fCurrentY + fSize, fCurrentZ, mfTuLow, mfTvLow)
                muVerts2(1) = New CustomVertex.PositionTextured(fCurrentX - (fSize / 2), fCurrentY - fSize, fCurrentZ, mfTuLow, mfTvHi)
                muVerts2(2) = New CustomVertex.PositionTextured(fDestX + (fSize / 2), fDestY + fSize, fDestZ, mfTuHi, mfTvLow)
                muVerts2(3) = New CustomVertex.PositionTextured(fDestX - (fSize / 2), fDestY - fSize, fDestZ, mfTuHi, mfTvHi)

                If mfSizeMult < 0.001 Then
                    mlStart = Int32.MinValue
                    mfSizeMult = 0.3
                    Return True
                End If
            End If
            Return False
        End Function
    End Class

    Private moFX() As BaseWpnEffect
    Private myFXUsed() As Byte
    Private mlFXUB As Int32 = -1

    Public Shared moTexture As Texture

    Public Event WeaponEnd(ByVal oTarget As BaseEntity, ByVal yWeaponType As Byte, ByVal bHit As Boolean, ByVal vecOnModelTo As Vector3)

    Public Event WpnEnd_WpnBldrOnly(ByVal yWeaponType As Byte)

    Public GenerateSoundFX As Boolean = True

    Private moRand As Random

    Private moBPSprite As BPSprite

    Public Sub New()
        moRand = New Random()
    End Sub

    'Private Sub SpriteDispose(ByVal sender As Object, ByVal e As EventArgs)
    '    moSprite = Nothing
    'End Sub

    'Private Sub SpriteLost(ByVal sender As Object, ByVal e As EventArgs)
    '    If moSprite Is Nothing = False Then moSprite.Dispose()
    '    moSprite = Nothing
    'End Sub

    Private Function HandleMapWrapShift(ByVal fTargetX As Single, ByVal fFromX As Single) As Single
        'Handle Map Wrap...
        If goCurrentEnvir Is Nothing = False Then
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                    'Now, determine if map wrapping is the better choice
                    Dim oPlanet As Planet = CType(goCurrentEnvir.oGeoObject, Planet)
                    Dim lTotalWidth As Int32 = TerrainClass.Width * oPlanet.CellSpacing
                    If fTargetX < fFromX Then
                        'ok, normally, we would go left...
                        Dim fTmp As Single = fTargetX + lTotalWidth
                        If Math.Abs(fTmp - fFromX) < Math.Abs(fFromX - fTargetX) Then
                            If Math.Abs(goCamera.mlCameraX - fTargetX) < Math.Abs(goCamera.mlCameraX - fFromX) Then
                                fFromX -= lTotalWidth
                            Else : fTargetX = fTmp
                            End If
                        End If
                    Else
                        'ok, normally, we would go right..
                        Dim fTmp As Single = fTargetX - lTotalWidth
                        If Math.Abs(fFromX - fTmp) < Math.Abs(fFromX - fTargetX) Then
                            If Math.Abs(goCamera.mlCameraX - fTargetX) < Math.Abs(goCamera.mlCameraX - fFromX) Then
                                fFromX += lTotalWidth
                            Else : fTargetX = fTmp
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Function

    Public Sub AddNewEffect(ByVal oTarget As BaseEntity, ByVal oAttacker As BaseEntity, ByVal yWpnType As Byte, ByVal bHit As Boolean, ByVal bPointDefense As Boolean, ByVal yAOE As Byte)
        'weapon type indicates what weapon type it is from the WeaponType enum 
        Dim X As Int32
        Dim lIdx As Int32
        Dim sSoundFile As String
        Dim lTemp As Int32

        'Should not happen, but we gotta make sure
        If oTarget Is Nothing Then Return
        If oAttacker Is Nothing Then Return

        If oAttacker.OwnerID = glPlayerID AndAlso oTarget.OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID Then
            If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
                If oTarget.ObjTypeID = ObjectType.eFacility Then
                    'If goTutorial.EventTriggered(TutorialManager.TutorialTriggerType.AssaultPirateBase) = False Then
                    '    goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.AssaultPirateBase)
                    'End If
                ElseIf goTutorial.EventTriggered(TutorialManager.TutorialTriggerType.FirstEngagement) = False Then
                    goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.FirstEngagement)
                End If
            End If
        End If

        If goSound Is Nothing = False Then
            If goSound.MusicOn = True Then
                If oTarget.OwnerID = glPlayerID Then goSound.IncrementExcitementLevel(20) 'we are being attacked +20
                If oAttacker.OwnerID = glPlayerID Then goSound.IncrementExcitementLevel(15) 'we are attacking + 15
            End If
        End If


        lIdx = -1
        For X = 0 To mlFXUB
            If myFXUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlFXUB += 1
            ReDim Preserve moFX(mlFXUB)
            ReDim Preserve myFXUsed(mlFXUB)
            lIdx = mlFXUB
        End If

        'TODO: we will want to put in better rotations for turrets... but for now, just set it
        If oAttacker.oMesh Is Nothing = False Then
            If oAttacker.oMesh.bTurretMesh = True Then
                oAttacker.TurretRot = CShort(LineAngleDegrees(CInt(oAttacker.LocX), CInt(oAttacker.LocZ), CInt(oTarget.LocX), CInt(oTarget.LocZ)) * 10)
            End If
        End If

        'TODO: in some cases, the weapon types have only one sfx file, we will add more later
        'TODO: BOMB SFX when they are done
        'TODO: Capital Solid Beam SFX when they are done
        'TODO: Charged Missile SFX when they are done
        'TODO: Land Based / Space Based mine... isn't that an explosion?
        'TODO: Missile SFX when they are done

        Dim ySFXType As SoundMgr.SoundUsage

        Select Case yWpnType
            Case WeaponType.eSolidGreenBeam, WeaponType.eSolidPurpleBeam, WeaponType.eSolidRedBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                sSoundFile = "Beam2a.wav"
                'sSoundFile = "Energy Weapons\Beam2a.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eSolidTealBeam, WeaponType.eSolidYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                sSoundFile = "Beam2b.wav"
                'sSoundFile = "Energy Weapons\Beam2b.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eFlickerGreenBeam, WeaponType.eFlickerPurpleBeam, WeaponType.eFlickerRedBeam, _
              WeaponType.eFlickerTealBeam, WeaponType.eFlickerYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                sSoundFile = "Beam1" & Chr(lsub) & ".wav"
                'sSoundFile = "Energy Weapons\Beam1" & Chr(lsub) & ".wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eMetallicProjectile_Bronze, WeaponType.eMetallicProjectile_Copper, _
              WeaponType.eMetallicProjectile_Gold, WeaponType.eMetallicProjectile_Lead, WeaponType.eMetallicProjectile_Silver
                lTemp = moRand.Next(1, 5)
                sSoundFile = "LargeProj" & lTemp & ".wav"
                'sSoundFile = "Projectile Weapons\LargeProj" & lTemp & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireProjectile
            Case Else
                lTemp = CInt(yWpnType) + 1
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                sSoundFile = "Laser" & lTemp & Chr(lsub) & ".wav"
                'sSoundFile = "Energy Weapons\Laser" & lTemp & Chr(lsub) & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
        End Select

        'UnitArcs
        Dim fBaseAngle As Single = LineAngleDegrees(CInt(oAttacker.LocX), CInt(oAttacker.LocZ), CInt(oTarget.LocX), CInt(oTarget.LocZ))
        Dim fAngle As Single = fBaseAngle

        Dim fMyAngle As Single = (oAttacker.LocAngle / 10.0F) '- 90.0F
        'If glCurrentEnvirView = CurrentView.eStartupLogin Then fMyAngle -= 90.0F
        fAngle -= fMyAngle
        If fAngle > 360 Then fAngle -= 360
        If fAngle < 0 Then fAngle += 360

        Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
        Dim vecFrom As Vector3 = oAttacker.GetFireFromLoc(ySide)

        Dim ySideHit As Byte = 0
        Dim vecOnModelTarget As Vector3
        If bHit = True Then
            'Now, determine what side I hit...
            fAngle = fBaseAngle - 180
            'If fAngle > 360 Then fAngle -= 360
            'If fAngle < 0 Then fAngle += 360
            fMyAngle = (oTarget.LocAngle / 10.0F)
            fAngle -= fMyAngle
            If fAngle > 360 Then fAngle -= 360
            If fAngle < 0 Then fAngle += 360
            ySideHit = AngleToQuadrant(CInt(fAngle))
            vecOnModelTarget = oTarget.oMesh.GetRandomMeshPoint(ySideHit)
        Else
            vecOnModelTarget = New Vector3(0, 0, 0)
        End If

        With moFX(lIdx)
            .vecOriginalFireFrom = vecFrom

            'If .bUsesSprite = True Then
            '    'vecFrom = BaseWpnEffect.QuickTransformCoordinate(vecFrom, oAttacker.GetWorldMatrix)
            With oAttacker.GetWorldMatrix()
                Dim fX As Single = vecFrom.X
                Dim fY As Single = vecFrom.Y
                Dim fZ As Single = vecFrom.Z
                moFX(lIdx).fCurrentX = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                moFX(lIdx).fCurrentY = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                moFX(lIdx).fCurrentZ = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
            End With
            'Else
            '    .fCurrentX = vecFrom.X
            '    .fCurrentY = vecFrom.Y
            '    .fCurrentZ = vecFrom.Z
            'End If

            'Ok, I'll put the hook in here for sound FX
            If GenerateSoundFX = True AndAlso goSound Is Nothing = False AndAlso sSoundFile <> "" Then
                'once and done sound, we don't care
                If (oTarget Is Nothing = False AndAlso oTarget.yVisibility = eVisibilityType.Visible) OrElse (oAttacker Is Nothing = False AndAlso oAttacker.yVisibility = eVisibilityType.Visible) Then
                    goSound.StartSound(sSoundFile, False, ySFXType, New Vector3(.fCurrentX, .fCurrentY, .fCurrentZ), New Vector3(0, 0, 0))
                End If
            End If

            .ySideHit = ySideHit
            .vecOnModelTo = vecOnModelTarget

            If bHit = True Then
                .fDestX = oTarget.LocX
                .fDestY = oTarget.LocY
                .fDestZ = oTarget.LocZ                          'Dest Z
            Else
                Dim vecTmp As Vector3 = oTarget.oMesh.vecDeathSeqSize
                .fDestX = oTarget.LocX + ((Rnd() * vecTmp.X * 2) - vecTmp.X)
                If Rnd() * 100 < 50 Then
                    .fDestY = oTarget.LocY + ((Rnd() * vecTmp.Y * 0.5F) + vecTmp.Y)
                Else
                    .fDestY = oTarget.LocY - ((Rnd() * vecTmp.Y * 0.5F) + vecTmp.Y)
                End If
                .fDestZ = oTarget.LocZ + ((Rnd() * vecTmp.Z * 2) - vecTmp.Z)
            End If

            .lIndex = lIdx
            .WeaponHit = bHit
            .WeaponTypeID = yWpnType
            .EffectActive = True
            .oTarget = oTarget
            .oAttacker = oAttacker
            .yAOE = yAOE
            If bPointDefense = False Then
                .fSizeMult = 1.0F
            Else
                .fSizeMult = 0.3F
                If .WpnMaxSpeed < Int32.MaxValue Then .WpnMaxSpeed *= 2
            End If
        End With
        myFXUsed(lIdx) = 255

    End Sub

    Public Sub AddNewEffect(ByVal oTarget As MissileMgr.Missile, ByVal oAttacker As BaseEntity, ByVal yWpnType As Byte, ByVal bHit As Boolean, ByVal bPointDefense As Boolean)
        'weapon type indicates what weapon type it is from the WeaponType enum 
        Dim X As Int32
        Dim lIdx As Int32
        Dim sSoundFile As String
        Dim lTemp As Int32

        'Should not happen, but we gotta make sure
        If oTarget Is Nothing Then Return
        If oAttacker Is Nothing Then Return

        If goSound Is Nothing = False Then
            If goSound.MusicOn = True Then
                If oAttacker.OwnerID = glPlayerID Then goSound.IncrementExcitementLevel(15) 'we are attacking + 15
            End If
        End If

        lIdx = -1
        For X = 0 To mlFXUB
            If myFXUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlFXUB += 1
            ReDim Preserve moFX(mlFXUB)
            ReDim Preserve myFXUsed(mlFXUB)
            lIdx = mlFXUB
        End If

        'TODO: we will want to put in better rotations for turrets... but for now, just set it
        If oAttacker.oMesh Is Nothing = False Then
            If oAttacker.oMesh.bTurretMesh = True Then
                oAttacker.TurretRot = CShort(LineAngleDegrees(CInt(oAttacker.LocX), CInt(oAttacker.LocZ), CInt(oTarget.vecMissile.X), CInt(oTarget.vecMissile.Z)) * 10)
            End If
        End If

        'TODO: in some cases, the weapon types have only one sfx file, we will add more later
        'TODO: BOMB SFX when they are done
        'TODO: Capital Solid Beam SFX when they are done
        'TODO: Charged Missile SFX when they are done
        'TODO: Land Based / Space Based mine... isn't that an explosion?
        'TODO: Missile SFX when they are done

        Dim ySFXType As SoundMgr.SoundUsage

        Select Case yWpnType
            Case WeaponType.eSolidGreenBeam, WeaponType.eSolidPurpleBeam, WeaponType.eSolidRedBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                'sSoundFile = "Energy Weapons\Beam2a.wav"
                sSoundFile = "Beam2a.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eSolidTealBeam, WeaponType.eSolidYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                'sSoundFile = "Energy Weapons\Beam2b.wav"
                sSoundFile = "Beam2b.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eFlickerGreenBeam, WeaponType.eFlickerPurpleBeam, WeaponType.eFlickerRedBeam, _
              WeaponType.eFlickerTealBeam, WeaponType.eFlickerYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                'sSoundFile = "Energy Weapons\Beam1" & Chr(lsub) & ".wav"
                sSoundFile = "Beam1" & Chr(lsub) & ".wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eMetallicProjectile_Bronze, WeaponType.eMetallicProjectile_Copper, _
              WeaponType.eMetallicProjectile_Gold, WeaponType.eMetallicProjectile_Lead, WeaponType.eMetallicProjectile_Silver
                lTemp = moRand.Next(1, 5)
                'sSoundFile = "Projectile Weapons\LargeProj" & lTemp & ".wav"
                sSoundFile = "LargeProj" & lTemp & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireProjectile
            Case Else
                lTemp = CInt(yWpnType) + 1
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                'sSoundFile = "Energy Weapons\Laser" & lTemp & Chr(lsub) & ".wav"
                sSoundFile = "Laser" & lTemp & Chr(lsub) & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
        End Select

        'UnitArcs
        Dim fAngle As Single = LineAngleDegrees(CInt(oAttacker.LocX), CInt(oAttacker.LocZ), CInt(oTarget.vecMissile.X), CInt(oTarget.vecMissile.Z))

        Dim fMyAngle As Single = (oAttacker.LocAngle / 10.0F) ' - 90.0F
        'If glCurrentEnvirView = CurrentView.eStartupLogin Then fMyAngle -= 90.0F
        fAngle -= fMyAngle
        If fAngle > 360 Then fAngle -= 360
        If fAngle < 0 Then fAngle += 360

        Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
        Dim vecFrom As Vector3 = oAttacker.GetFireFromLoc(ySide)
        'Dim vecOrigFrom As Vector3 = vecFrom

        Dim fTmpDX As Single = oTarget.vecMissile.X
        'If goCurrentEnvir Is Nothing = False Then
        '    If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
        '        If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
        '            'Now, determine if map wrapping is the better choice
        '            Dim oPlanet As Planet = CType(goCurrentEnvir.oGeoObject, Planet)
        '            Dim lTotalWidth As Int32 = TerrainClass.Width * oPlanet.CellSpacing
        '            If fTmpDX < vecFrom.X Then
        '                'ok, normally, we would go left...
        '                Dim fTmp As Single = fTmpDX + lTotalWidth
        '                If Math.Abs(fTmp - vecFrom.X) < Math.Abs(vecFrom.X - fTmpDX) Then
        '                    If Math.Abs(goCamera.mlCameraX - fTmpDX) < Math.Abs(goCamera.mlCameraX - vecFrom.X) Then
        '                        vecFrom.X -= lTotalWidth
        '                    Else : fTmpDX = fTmp
        '                    End If
        '                End If
        '            Else
        '                'ok, normally, we would go right..
        '                Dim fTmp As Single = fTmpDX - lTotalWidth
        '                If Math.Abs(vecFrom.X - fTmp) < Math.Abs(vecFrom.X - fTmpDX) Then
        '                    If Math.Abs(goCamera.mlCameraX - fTmpDX) < Math.Abs(goCamera.mlCameraX - vecFrom.X) Then
        '                        vecFrom.X += lTotalWidth
        '                    Else : fTmpDX = fTmp
        '                    End If
        '                End If
        '            End If
        '        End If
        '    End If
        'End If

        'Ok, I'll put the hook in here for sound FX
        If GenerateSoundFX = True AndAlso goSound Is Nothing = False AndAlso sSoundFile <> "" Then
            'once and done sound, we don't care
            'If (oAttacker Is Nothing = False AndAlso oAttacker.yVisibility = eVisibilityType.Visible) Then
            'goSound.StartSound(sSoundFile, False, ySFXType, vecFrom, New Vector3(0, 0, 0))
            goSound.StartSound(sSoundFile, False, ySFXType, New Vector3(oAttacker.LocX, oAttacker.LocY, oAttacker.LocZ), New Vector3(0, 0, 0))
            'End If
        End If

        With moFX(lIdx)
            .ySideHit = 0
            .vecOnModelTo = New Vector3(0, 0, 0)
            .vecOriginalFireFrom = vecFrom

            .fCurrentX = vecFrom.X
            .fCurrentY = vecFrom.Y
            .fCurrentZ = vecFrom.Z

            With oAttacker.GetWorldMatrix()
                Dim fX As Single = vecFrom.X
                Dim fY As Single = vecFrom.Y
                Dim fZ As Single = vecFrom.Z
                moFX(lIdx).fCurrentX = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                moFX(lIdx).fCurrentY = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                moFX(lIdx).fCurrentZ = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
            End With

            If bHit = True Then
                .fDestX = fTmpDX
                .fDestY = oTarget.vecMissile.Y
                .fDestZ = oTarget.vecMissile.Z
            Else
                .fDestX = fTmpDX + ((Rnd() * 150) - 75)
                .fDestY = oTarget.vecMissile.Y + (Rnd() * 50 - 25)
                .fDestZ = oTarget.vecMissile.Z + ((Rnd() * 150) - 75)
            End If

            .lIndex = lIdx
            .WeaponHit = bHit
            .WeaponTypeID = yWpnType
            .EffectActive = True
            .oTargetMissile = oTarget
            .oAttacker = oAttacker
            .yAOE = 0
            If bPointDefense = False Then
                .fSizeMult = 1.0F
            Else
                .fSizeMult = 0.3F
                If .WpnMaxSpeed <> Int32.MaxValue Then .WpnMaxSpeed *= 2
            End If
        End With
        myFXUsed(lIdx) = 255

    End Sub

    Public Sub AddNewEffect_PointToEntity(ByVal lFromX As Int32, ByVal lFromY As Int32, ByVal lFromZ As Int32, ByVal oTarget As BaseEntity, ByVal yWpnType As Byte, ByVal bHit As Boolean, ByVal bPointDefense As Boolean)
        'weapon type indicates what weapon type it is from the WeaponType enum 
        Dim X As Int32
        Dim lIdx As Int32
        Dim sSoundFile As String
        Dim lTemp As Int32

        'Should not happen, but we gotta make sure
        If oTarget Is Nothing Then Return

        lIdx = -1
        For X = 0 To mlFXUB
            If myFXUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlFXUB += 1
            ReDim Preserve moFX(mlFXUB)
            ReDim Preserve myFXUsed(mlFXUB)
            lIdx = mlFXUB
        End If

        'TODO: in some cases, the weapon types have only one sfx file, we will add more later
        'TODO: BOMB SFX when they are done
        'TODO: Capital Solid Beam SFX when they are done
        'TODO: Charged Missile SFX when they are done
        'TODO: Land Based / Space Based mine... isn't that an explosion?
        'TODO: Missile SFX when they are done

        Dim ySFXType As SoundMgr.SoundUsage

        Select Case yWpnType
            Case WeaponType.eSolidGreenBeam, WeaponType.eSolidPurpleBeam, WeaponType.eSolidRedBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                'sSoundFile = "Energy Weapons\Beam2a.wav"
                sSoundFile = "Beam2a.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eSolidTealBeam, WeaponType.eSolidYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                'sSoundFile = "Energy Weapons\Beam2b.wav"
                sSoundFile = "Beam2b.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eFlickerGreenBeam, WeaponType.eFlickerPurpleBeam, WeaponType.eFlickerRedBeam, _
              WeaponType.eFlickerTealBeam, WeaponType.eFlickerYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                'sSoundFile = "Energy Weapons\Beam1" & Chr(lsub) & ".wav"
                sSoundFile = "Beam1" & Chr(lsub) & ".wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eMetallicProjectile_Bronze, WeaponType.eMetallicProjectile_Copper, _
              WeaponType.eMetallicProjectile_Gold, WeaponType.eMetallicProjectile_Lead, WeaponType.eMetallicProjectile_Silver
                lTemp = moRand.Next(1, 5)
                'sSoundFile = "Projectile Weapons\LargeProj" & lTemp & ".wav"
                sSoundFile = "LargeProj" & lTemp & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireProjectile
            Case Else
                lTemp = CInt(yWpnType) + 1
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                'sSoundFile = "Energy Weapons\Laser" & lTemp & Chr(lsub) & ".wav"
                sSoundFile = "Laser" & lTemp & Chr(lsub) & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
        End Select

        'UnitArcs
        Dim vecFrom As Vector3 = New Vector3(lFromX, lFromY, lFromZ)

        Dim fTmpDX As Single = oTarget.LocX
        Dim ySideHit As Byte = 0
        Dim vecOnModelTarget As Vector3
        If bHit = True Then
            'Now, determine what side I hit...
            Dim fAngle As Single = LineAngleDegrees(lFromX, lFromZ, CInt(oTarget.LocX), CInt(oTarget.LocZ)) - 180
            Dim fMyAngle As Single = (oTarget.LocAngle / 10.0F)
            fAngle -= fMyAngle
            If fAngle > 360 Then fAngle -= 360
            If fAngle < 0 Then fAngle += 360
            ySideHit = AngleToQuadrant(CInt(fAngle))
            vecOnModelTarget = oTarget.oMesh.GetRandomMeshPoint(ySideHit)
        Else
            vecOnModelTarget = New Vector3(0, 0, 0)
        End If

        With moFX(lIdx)
            .ySideHit = ySideHit
            .vecOnModelTo = vecOnModelTarget
            .vecOriginalFireFrom = vecFrom
            .fCurrentX = vecFrom.X
            .fCurrentY = vecFrom.Y
            .fCurrentZ = vecFrom.Z
            If bHit = True Then
                .fDestX = oTarget.LocX
                .fDestY = oTarget.LocY
                .fDestZ = oTarget.LocZ                          'Dest Z
            Else
                .fDestX = oTarget.LocX + ((Rnd() * 200 * 2) - 200)
                If Rnd() * 100 < 50 Then
                    .fDestY = oTarget.LocY + ((Rnd() * 600 * 0.5F) + 600)
                Else
                    .fDestY = oTarget.LocY - ((Rnd() * 600 * 0.5F) + 600)
                End If
                .fDestZ = oTarget.LocZ + ((Rnd() * 1500 * 2) - 1500)
            End If

            .lIndex = lIdx
            .WeaponHit = bHit
            .WeaponTypeID = yWpnType
            .EffectActive = True
            .oTargetMissile = Nothing
            .oTarget = oTarget
            .oAttacker = Nothing
            .yAOE = 0
            If bPointDefense = False Then
                .fSizeMult = 1.0F
            Else
                .fSizeMult = 0.3F
                .WpnMaxSpeed *= 2
            End If
        End With
        myFXUsed(lIdx) = 255

    End Sub

    Public Sub AddNewEffect(ByVal lFromX As Int32, ByVal lFromY As Int32, ByVal lFromZ As Int32, ByVal lToX As Int32, ByVal lToY As Int32, ByVal lToZ As Int32, ByVal yWpnType As Byte, ByVal bBombard As Boolean, ByVal yAOE As Byte)
        'ok... this is used for 'indirect' fire... the problem with this is that we can't make effects based off
        '  of this effect because we don't know what is affected by this wpn shot

        'weapon type indicates what weapon type it is from the WeaponType enum 
        Dim X As Int32
        Dim lIdx As Int32
        Dim sSoundFile As String
        Dim lTemp As Int32

        lIdx = -1
        For X = 0 To mlFXUB
            If myFXUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlFXUB += 1
            ReDim Preserve moFX(mlFXUB)
            ReDim Preserve myFXUsed(mlFXUB)
            lIdx = mlFXUB
        End If

        'TODO: in some cases, the weapon types have only one sfx file, we will add more later
        'TODO: BOMB SFX when they are done
        'TODO: Capital Solid Beam SFX when they are done
        'TODO: Charged Missile SFX when they are done
        'TODO: Land Based / Space Based mine... isn't that an explosion?
        'TODO: Missile SFX when they are done

        Dim ySFXType As SoundMgr.SoundUsage

        Select Case yWpnType
            Case WeaponType.eSolidGreenBeam, WeaponType.eSolidPurpleBeam, WeaponType.eSolidRedBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                'sSoundFile = "Energy Weapons\Beam2a.wav"
                sSoundFile = "Beam2a.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eSolidTealBeam, WeaponType.eSolidYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                'sSoundFile = "Energy Weapons\Beam2b.wav"
                sSoundFile = "Beam2b.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eFlickerGreenBeam, WeaponType.eFlickerPurpleBeam, WeaponType.eFlickerRedBeam, _
              WeaponType.eFlickerTealBeam, WeaponType.eFlickerYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                'sSoundFile = "Energy Weapons\Beam1" & Chr(lsub) & ".wav"
                sSoundFile = "Beam1" & Chr(lsub) & ".wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eMetallicProjectile_Bronze, WeaponType.eMetallicProjectile_Copper, _
              WeaponType.eMetallicProjectile_Gold, WeaponType.eMetallicProjectile_Lead, WeaponType.eMetallicProjectile_Silver
                lTemp = moRand.Next(1, 5)
                'sSoundFile = "Projectile Weapons\LargeProj" & lTemp & ".wav"
                sSoundFile = "LargeProj" & lTemp & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireProjectile
            Case Else
                lTemp = CInt(yWpnType) + 1
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                'sSoundFile = "Energy Weapons\Laser" & lTemp & Chr(lsub) & ".wav"
                sSoundFile = "Laser" & lTemp & Chr(lsub) & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
        End Select

        'Ok, I'll put the hook in here for sound FX
        If GenerateSoundFX = True AndAlso goSound Is Nothing = False AndAlso sSoundFile <> "" Then
            'once and done sound, we don't care
            'TODO: we need to determine if the point indicated is 'visible' to the user
            goSound.StartSound(sSoundFile, False, ySFXType, New Vector3(lToX, lToY, lToZ), New Vector3(0, 0, 0))
        End If

        With moFX(lIdx)
            .vecOriginalFireFrom.X = lFromX : .vecOriginalFireFrom.Y = lFromY : .vecOriginalFireFrom.Z = lFromZ
            .vecOnModelTo = New Vector3(0, 0, 0)
            .ySideHit = 0
            .fCurrentX = lFromX
            .fCurrentY = lFromY
            .fCurrentZ = lFromZ
            .fDestX = lToX
            .fDestY = lToY
            .fDestZ = lToZ
            .lIndex = lIdx
            .WeaponHit = False  '???
            .WeaponTypeID = yWpnType
            .EffectActive = True
            .oTarget = Nothing
            .oAttacker = Nothing
            .yAOE = yAOE

            If bBombard = True Then
                If .bUsesSprite = True Then .fSizeMult = 10 Else .fSizeMult = 3
            End If
        End With
        myFXUsed(lIdx) = 255

    End Sub

    Public Sub AddNewEffect(ByVal oAttacker As BaseEntity, ByVal lToX As Int32, ByVal lToY As Int32, ByVal lToZ As Int32, ByVal yWpnType As Byte, ByVal bBombard As Boolean, ByVal yAOE As Byte)
        'ok... this is used for 'indirect' fire... the problem with this is that we can't make effects based off
        '  of this effect because we don't know what is affected by this wpn shot

        'weapon type indicates what weapon type it is from the WeaponType enum 
        Dim X As Int32
        Dim lIdx As Int32
        Dim sSoundFile As String
        Dim lTemp As Int32

        lIdx = -1
        For X = 0 To mlFXUB
            If myFXUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlFXUB += 1
            ReDim Preserve moFX(mlFXUB)
            ReDim Preserve myFXUsed(mlFXUB)
            lIdx = mlFXUB
        End If

        If goSound Is Nothing = False Then
            If goSound.MusicOn = True Then
                If oAttacker.OwnerID = glPlayerID Then goSound.IncrementExcitementLevel(15) 'we are attacking + 15
            End If
        End If

        'TODO: in some cases, the weapon types have only one sfx file, we will add more later
        'TODO: BOMB SFX when they are done
        'TODO: Capital Solid Beam SFX when they are done
        'TODO: Charged Missile SFX when they are done
        'TODO: Land Based / Space Based mine... isn't that an explosion?
        'TODO: Missile SFX when they are done

        Dim ySFXType As SoundMgr.SoundUsage

        Select Case yWpnType
            Case WeaponType.eSolidGreenBeam, WeaponType.eSolidPurpleBeam, WeaponType.eSolidRedBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                'sSoundFile = "Energy Weapons\Beam2a.wav"
                sSoundFile = "Beam2a.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eSolidTealBeam, WeaponType.eSolidYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                'sSoundFile = "Energy Weapons\Beam2b.wav"
                sSoundFile = "Beam2b.wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eFlickerGreenBeam, WeaponType.eFlickerPurpleBeam, WeaponType.eFlickerRedBeam, _
              WeaponType.eFlickerTealBeam, WeaponType.eFlickerYellowBeam
                moFX(lIdx) = New BeamWpn()
                moFX(lIdx).bUsesSprite = False
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                'sSoundFile = "Energy Weapons\Beam1" & Chr(lsub) & ".wav"
                sSoundFile = "Beam1" & Chr(lsub) & ".wav"
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
            Case WeaponType.eMetallicProjectile_Bronze, WeaponType.eMetallicProjectile_Copper, _
              WeaponType.eMetallicProjectile_Gold, WeaponType.eMetallicProjectile_Lead, WeaponType.eMetallicProjectile_Silver
                lTemp = moRand.Next(1, 5)
                'sSoundFile = "Projectile Weapons\LargeProj" & lTemp & ".wav"
                sSoundFile = "LargeProj" & lTemp & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireProjectile
            Case Else
                lTemp = CInt(yWpnType) + 1
                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
                'sSoundFile = "Energy Weapons\Laser" & lTemp & Chr(lsub) & ".wav"
                sSoundFile = "Laser" & lTemp & Chr(lsub) & ".wav"
                moFX(lIdx) = New SingleShotWpn()
                moFX(lIdx).bUsesSprite = True
                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
        End Select

        'Ok, I'll put the hook in here for sound FX
        If GenerateSoundFX = True AndAlso goSound Is Nothing = False AndAlso sSoundFile <> "" Then
            'once and done sound, we don't care
            'TODO: we need to determine if the point indicated is 'visible' to the user
            goSound.StartSound(sSoundFile, False, ySFXType, New Vector3(lToX, lToY, lToZ), New Vector3(0, 0, 0))
        End If

        'UnitArcs
        Dim fAngle As Single = LineAngleDegrees(CInt(oAttacker.LocX), CInt(oAttacker.LocZ), lToX, lToZ)

        Dim fMyAngle As Single = (oAttacker.LocAngle / 10.0F) '- 90.0F
        'If glCurrentEnvirView = CurrentView.eStartupLogin Then fMyAngle -= 90.0F
        fAngle -= fMyAngle
        If fAngle > 360 Then fAngle -= 360
        If fAngle < 0 Then fAngle += 360

        Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
        Dim vecFrom As Vector3 = oAttacker.GetFireFromLoc(ySide)
        'Dim vecOrigFrom As Vector3 = vecFrom

        With moFX(lIdx)
            .vecOriginalFireFrom = vecFrom
            If .bUsesSprite = True Then
                'vecFrom = BaseWpnEffect.QuickTransformCoordinate(vecFrom, oAttacker.GetWorldMatrix)
                With oAttacker.GetWorldMatrix()
                    Dim fX As Single = vecFrom.X
                    Dim fY As Single = vecFrom.Y
                    Dim fZ As Single = vecFrom.Z
                    moFX(lIdx).fCurrentX = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                    moFX(lIdx).fCurrentY = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                    moFX(lIdx).fCurrentZ = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
                End With
            Else
                .fCurrentX = vecFrom.X
                .fCurrentY = vecFrom.Y
                .fCurrentZ = vecFrom.Z
            End If

            .fDestX = lToX
            .fDestY = lToY
            .fDestZ = lToZ
            .lIndex = lIdx
            .WeaponHit = False  '???
            .WeaponTypeID = yWpnType
            .EffectActive = True
            .oTarget = Nothing
            .oAttacker = oAttacker
            .yAOE = yAOE

            If bBombard = True Then
                .fSizeMult = 3      '3 times the size it would normally be
            End If
        End With
        myFXUsed(lIdx) = 255

    End Sub

    Private Sub WpnEffectEnd(ByVal Index As Int32)
        Dim yType As Byte
        Dim bHit As Boolean
        Dim oObj As RenderObject
        Dim vecOnModelTo As Vector3
        Dim yAOE As Byte = 0

        yType = moFX(Index).WeaponTypeID
        bHit = moFX(Index).WeaponHit
        oObj = moFX(Index).oTarget
        yAOE = moFX(Index).yAOE

        vecOnModelTo = moFX(Index).vecOnModelTo
        If bHit = False Then
            With moFX(Index)
                vecOnModelTo = New Vector3(.fCurrentX, .fCurrentY, .fCurrentZ)
            End With
        End If

        If yAOE > 0 Then 
            With (moFX(Index))
                AOEExplosionMgr.goMgr.AddNew(New Vector3(.fCurrentX, .fCurrentY, .fCurrentZ), yAOE)
            End With
        End If

        myFXUsed(Index) = 0
        moFX(Index) = Nothing
        If oObj Is Nothing = False Then RaiseEvent WeaponEnd(CType(oObj, BaseEntity), yType, bHit, vecOnModelTo)
        RaiseEvent WpnEnd_WpnBldrOnly(yType)
    End Sub

    Public Sub RenderFX(ByVal bUpdateNoRender As Boolean)
        Dim X As Int32
        Dim bOver As Boolean
        Dim oMat As Material

        Try
            Dim moDevice As Device = GFXEngine.moDevice
            If bUpdateNoRender Then
                'Ok, update the effect but do not render it
                For X = 0 To mlFXUB
                    If myFXUsed(X) > 0 AndAlso moFX(X) Is Nothing = False Then ' AndAlso moFX(X) Is Nothing = False Then
                        With moFX(X)
                            If .Update() = True Then WpnEffectEnd(X)
                        End With
                    End If
                Next X
            Else
                If moTexture Is Nothing OrElse moTexture.Disposed = True Then moTexture = goResMgr.GetTexture("WpnFire.dds", GFXResourceManager.eGetTextureType.NoSpecifics)

                'Dim yMapWrapSituation As Byte = 0           '0 = no map wrap, 1 = left edge, 2 = right edge
                'Dim lLocXMapWrapCheck As Int32 = 0
                'If goCurrentEnvir Is Nothing = False Then
                '    Dim lTmpMapWrapVal As Int32 = Math.Min((goCurrentEnvir.lMaxXPos * 2) \ 3, muSettings.EntityClipPlane)
                '    If goCamera.mlCameraX < goCurrentEnvir.lMinXPos + lTmpMapWrapVal Then
                '        yMapWrapSituation = 1
                '        lLocXMapWrapCheck = goCurrentEnvir.lMaxXPos - lTmpMapWrapVal
                '    ElseIf goCamera.mlCameraX > goCurrentEnvir.lMaxXPos - lTmpMapWrapVal Then
                '        yMapWrapSituation = 2
                '        lLocXMapWrapCheck = goCurrentEnvir.lMinXPos + lTmpMapWrapVal
                '    End If
                'End If

                'Go thru and get our shot count
                Dim lShotCnt As Int32 = 0
                Dim lCurrIdx As Int32 = 0
                Dim lHeadClr As Int32
                Dim lNormalTailClr As Int32 = System.Drawing.Color.FromArgb(128, 0, 0, 0).ToArgb
                Dim lTailClr As Int32
                Dim lInvisHeadClr As Int32 = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb
                Dim lInvisTailClr As Int32 = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb
                Dim fX As Single
                Dim fZ As Single

                Dim kf_SHOT_SIZE As Single = 3.0F

                For X = 0 To mlFXUB
                    If myFXUsed(X) <> 0 AndAlso moFX(X) Is Nothing = False AndAlso moFX(X).bUsesSprite = True Then
                        lShotCnt += 1
                    End If
                Next
                Dim uVerts(lShotCnt * 12) As CustomVertex.PositionColored
                Dim fTmpAngle As Single

                glWpnFXRendered = 0

                If muSettings.RenderPulseBolts = True Then
                    If moBPSprite Is Nothing Then moBPSprite = New BPSprite()
                    moBPSprite.BeginRender(lShotCnt, goResMgr.GetTexture("Particle.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "textures.pak"), GFXEngine.moDevice)
                End If

                For X = 0 To mlFXUB
                    If myFXUsed(X) <> 0 AndAlso moFX(X) Is Nothing = False AndAlso moFX(X).bUsesSprite = True Then
                        With moFX(X)
                            'If yMapWrapSituation <> myPrevMapWrapSituation Then .UpdateMapWrapSituation(yMapWrapSituation, lLocXMapWrapCheck)

                            bOver = .Update()

                            'See if this fixes the bug... I fear it may be a race condition...
                            If lCurrIdx + 11 > uVerts.GetUpperBound(0) Then Exit For

                            'an attacker or target not set indicates an orbital bombardment
                            '  otherwise, the weapon has to be visible somewhere
                            If (.oAttacker Is Nothing OrElse .oTarget Is Nothing) OrElse (.oAttacker.yVisibility = eVisibilityType.Visible OrElse .oTarget.yVisibility = eVisibilityType.Visible) Then
                                lHeadClr = .colWpnColor.ToArgb
                                lTailClr = lNormalTailClr
                                glWpnFXRendered += 1

                                If muSettings.RenderPulseBolts = True AndAlso .WeaponTypeID <= WeaponType.eShortPurplePulse Then
                                    moBPSprite.Draw(New Vector3(.fCurrentX, .fCurrentY, .fCurrentZ), Rnd() * 25 + 8, .colWpnColor, 0)
                                End If
                            Else
                                lHeadClr = lInvisHeadClr
                                lTailClr = lInvisTailClr
                            End If

                            Dim fThisShotSize As Single = kf_SHOT_SIZE * .fSizeMult
                            fX = 0 : fZ = -fThisShotSize

                            fTmpAngle = LineAngleDegrees(CInt(.fCurrentX), CInt(.fCurrentZ), CInt(.fDestX), CInt(.fDestZ))

                            RotatePoint(0, 0, fX, fZ, fTmpAngle)
                            uVerts(lCurrIdx) = New CustomVertex.PositionColored(.fCurrentX + fX, .fCurrentY - fThisShotSize, .fCurrentZ + fZ, lHeadClr)

                            uVerts(lCurrIdx + 1) = New CustomVertex.PositionColored(.fCurrentX, .fCurrentY + fThisShotSize, .fCurrentZ, lHeadClr)

                            fX = 0 : fZ = fThisShotSize
                            RotatePoint(0, 0, fX, fZ, fTmpAngle)
                            uVerts(lCurrIdx + 2) = New CustomVertex.PositionColored(.fCurrentX + fX, .fCurrentY - fThisShotSize, .fCurrentZ + fZ, lHeadClr)

                            fX = -40 : fZ = 0
                            RotatePoint(0, 0, fX, fZ, fTmpAngle)
                            uVerts(lCurrIdx + 3) = uVerts(lCurrIdx)
                            uVerts(lCurrIdx + 4) = New CustomVertex.PositionColored(.fCurrentX + fX, .fPrevY, .fCurrentZ + fZ, lTailClr)
                            uVerts(lCurrIdx + 5) = uVerts(lCurrIdx + 1)

                            uVerts(lCurrIdx + 6) = uVerts(lCurrIdx + 1)
                            uVerts(lCurrIdx + 7) = uVerts(lCurrIdx + 4)
                            uVerts(lCurrIdx + 8) = uVerts(lCurrIdx + 2)

                            uVerts(lCurrIdx + 9) = uVerts(lCurrIdx + 2)
                            uVerts(lCurrIdx + 10) = uVerts(lCurrIdx + 4)
                            uVerts(lCurrIdx + 11) = uVerts(lCurrIdx)

                            lCurrIdx += 12

                            If bOver = True Then WpnEffectEnd(X)
                        End With
                    End If
                Next X

                'reset identity to nothing...
                moDevice.Transform.World = Matrix.Identity
                moDevice.RenderState.CullMode = Cull.None
                moDevice.RenderState.ZBufferWriteEnable = False
                moDevice.RenderState.SourceBlend = Blend.SourceAlpha
                moDevice.RenderState.DestinationBlend = Blend.One
                moDevice.RenderState.AlphaBlendEnable = True

                'Ok, render our shots...
                If lShotCnt > 0 Then
                    moDevice.RenderState.Lighting = False
                    moDevice.SetTexture(0, Nothing)
                    moDevice.VertexFormat = CustomVertex.PositionColored.Format

                    moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, lShotCnt * 4, uVerts)

                    moDevice.RenderState.Lighting = True
                End If

                'Now, go back thru and do our non-sprite based rendering
                moDevice.VertexFormat = CustomVertex.PositionTextured.Format
                moDevice.SetTexture(0, moTexture)

                For X = 0 To mlFXUB
                    If myFXUsed(X) > 0 AndAlso moFX(X) Is Nothing = False AndAlso moFX(X).bUsesSprite = False Then
                        With CType(moFX(X), BeamWpn)
                            bOver = .Update()

                            oMat.Ambient = .colWpnColor
                            oMat.Diffuse = .colWpnColor
                            oMat.Emissive = .colWpnColor
                            moDevice.Material = oMat

                            'TODO: We can streamline this... lots of beams would cause LOTS of lag
                            'Instead, we could use a list of verts and use it as a Triangle List instead of strip...
                            '  then, we just do one draw call...
                            If .oAttacker Is Nothing = False AndAlso .oTarget Is Nothing = False Then
                                If .oAttacker.yVisibility = eVisibilityType.Visible OrElse .oTarget.yVisibility = eVisibilityType.Visible Then
                                    moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, .muVerts1)
                                    moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, .muVerts2)

                                    glWpnFXRendered += 1
                                End If
                            Else
                                'Always render orbital bombardment
                                moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, .muVerts1)
                                moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, .muVerts2)

                                glWpnFXRendered += 1
                            End If

                            If bOver = True Then WpnEffectEnd(X)
                        End With
                    End If
                Next X

                moDevice.RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                moDevice.RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero 
                moDevice.RenderState.ZBufferWriteEnable = True

                If muSettings.RenderPulseBolts = True AndAlso moBPSprite Is Nothing = False Then moBPSprite.EndRender()

                'myPrevMapWrapSituation = yMapWrapSituation
            End If


        Catch
            'do nothing
        End Try
    End Sub

    Public Sub CleanAll()
        mlFXUB = -1
        ReDim moFX(-1)
        ReDim myFXUsed(-1)
        GenerateSoundFX = True
    End Sub

End Class
