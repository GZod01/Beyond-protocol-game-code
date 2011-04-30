Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class RenderObject
    Inherits Base_GUID

	Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32
	Public Shared Function GetCurrTime() As Int32
		Return timeGetTime
	End Function

    Public oMesh As BaseMesh

    'for a renderable unit we track only that which we need
    Public LocX As Single   'int32
	Public LocY As Single ' Int32
    Public LocZ As Single   'int32  
    Public LocAngle As Int16
    Public LocYaw As Int16

    'Public fMapWrapLocX As Single = Single.MinValue

    Private matWorld As Matrix
    'By setting these to the min values, they ensure that oure matrix gets set the first time...
    Private mfLastX As Single = Single.MinValue
	Private mfLastY As Single = Single.MinValue
	Private mfLastZ As Single = Single.MinValue
	Private miLastA As Int16 = Int16.MinValue
	Private miLastYaw As Int16 = Int16.MinValue
    Private mbForceTurrentUpdate As Boolean = False
    Private mlLastWorldMatrixSet As Int32

	'These get set in the add object and in the movement events
	Public CellLocX As Int32
	Public CellLocZ As Int32

	'For new Z-Level and Y Movement
	Private miLocPitch As Int16 = 0
	Public mfTrueLocPitch As Single = 0.0F
	Private mlLastYChgTime As Int32 = Int32.MinValue
	Public mlTargetY As Int32 = 0
	Private miLastLocPitch As Int16 = Int16.MinValue
	Public bDoNotRunSetYTarget As Boolean = False
	Public bForceSetY As Boolean = False

	Public yVisibility As Byte

	'For turrets
	Private matTurrWorld As Matrix
	Private miTurretRot As Int16 'rotation of the turret in 10s degrees... set from Fire Weapon
	Private miLastTurretRot As Int16
	Private mlLastTurretSetTime As Int32
    Private mlLastMoveYCycle As Int32

	Private mlFireFromLocID(3) As Int32

	'Burn FX
	Protected mlBurnFXActive() As Int32 = Nothing

	'For EngineFX color
	Public clrEngineFX As System.Drawing.Color = Color.FromArgb(255, 64, 192, 255)
	Public clrShieldFX As System.Drawing.Color = Color.FromArgb(255, 0, 255, 255)

    Public bCulled As Boolean       'indicates whether this object was culled this frame (if so, don't render its burn fx and other related fx)

#Region "  Burn Mark Management  "
    Public muBurnMarks() As BurnEffectPoint
    Public Sub AddBurnMark(ByVal vecLoc As Vector3, ByVal ySide As Byte, ByVal fSize As Single)
        Dim fMinBurnAlpha As Single = 255.0F
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To muBurnMarks.GetUpperBound(0)
            If muBurnMarks(X).bActive = False Then
                With muBurnMarks(X)
                    .fBurnAlpha = 255.0F
                    .fRotation = Rnd() * 360
                    .vecLoc = vecLoc
                    .ySide = ySide
                    .bActive = True
                    .fSize = fSize
                End With
                Return
            ElseIf muBurnMarks(X).fBurnAlpha < fMinBurnAlpha Then
                fMinBurnAlpha = muBurnMarks(X).fBurnAlpha
                lIdx = -1
            End If
        Next X
        If lIdx <> -1 Then
            With muBurnMarks(lIdx)
                .fBurnAlpha = 255.0F
                .fRotation = Rnd() * 360
                .vecLoc = vecLoc
                .ySide = ySide
                .bActive = True
                .fSize = fSize
            End With
        End If
    End Sub
#End Region

    Public Function GetBurnFXLoc(ByVal lSide As Int32) As Int32
        If mlBurnFXActive Is Nothing Then
            If oMesh Is Nothing OrElse oMesh.muBurnLocs Is Nothing Then Return -1

            ReDim mlBurnFXActive(oMesh.muBurnLocs.GetUpperBound(0))

            For X As Int32 = 0 To oMesh.muBurnLocs.GetUpperBound(0)
                mlBurnFXActive(X) = -1
            Next X
        End If

        For X As Int32 = 0 To mlBurnFXActive.GetUpperBound(0)
            If (oMesh.muBurnLocs(X).lSide = lSide OrElse oMesh.muBurnLocs(X).lSide = -1) AndAlso mlBurnFXActive(X) = -1 Then
                Return X
            End If
        Next X

        Return -1
    End Function

    Public Function GetFireFromLoc(ByVal ySide As Byte) As Vector3
        Dim lVal As Int32 = mlFireFromLocID(ySide)
        If oMesh Is Nothing Then Return New Vector3(0, 0, 0)
        If lVal > oMesh.muFireLocs(ySide).GetUpperBound(0) Then lVal = 0
        mlFireFromLocID(ySide) = lVal + 1

        If oMesh.muFireLocs(ySide).GetUpperBound(0) >= lVal Then
            'Dim vecResult As Vector3 = oMesh.muFireLocs(ySide)(lVal).GetFireFromLoc(Me)
            'Dim matWorld As Matrix = GetWorldMatrix()
            'With matWorld
            '    Dim fX As Single = (vecResult.X * .M11) + (vecResult.Y * .M21) + (vecResult.Z * .M31) '+ .M41
            '    Dim fY As Single = (vecResult.X * .M12) + (vecResult.Y * .M22) + (vecResult.Z * .M32) '+ .M42
            '    Dim fZ As Single = (vecResult.X * .M13) + (vecResult.Y * .M23) + (vecResult.Z * .M33) '+ .M43
            '    vecResult.X = fX
            '    vecResult.Y = fY
            '    vecResult.Z = fZ
            'End With
            Dim vecTemp As Vector3
            With oMesh.muFireLocs(ySide)(lVal)
                vecTemp.X = .lOffsetX : vecTemp.Y = .lOffsetY : vecTemp.Z = .lOffsetZ
            End With
            Return vecTemp 'oMesh.muFireLocs(ySide)(lVal).GetFireFromLoc(Me)
        Else : Return New Vector3(0, 0, 0) 'Return New Vector3(LocX, LocY, LocZ)
        End If
    End Function

    Public Property TurretRot() As Int16
        Get
            Return miTurretRot
        End Get
        Set(ByVal value As Int16)
            miTurretRot = value - 900S
            If miTurretRot < 0 Then miTurretRot += 3600S
            mlLastTurretSetTime = timeGetTime
        End Set
    End Property

    Private matPerpWorldMatrix As Matrix
    Public Function GetPerpWorldMatrix() As Matrix
        Return matPerpWorldMatrix
    End Function

    Public bForceGetWorldMatrix As Boolean = False
    Public Function GetWorldMatrix() As Matrix

        'If glCurrentCycle = mlLastWorldMatrixSet Then Return matWorld
        'mlLastWorldMatrixSet = glCurrentCycle

        'If fMapWrapLocX = Single.MinValue Then fMapWrapLocX = LocX

        If bForceGetWorldMatrix = True OrElse LocX <> mfLastX OrElse LocY <> mfLastY OrElse LocZ <> mfLastZ OrElse LocAngle <> miLastA OrElse LocYaw <> miLastYaw OrElse miLocPitch <> miLastLocPitch OrElse miLocPitch <> 0 Then

            'If Me.ObjTypeID = 8191 Then Stop

            Dim matTemp As Matrix
            Dim fPitch As Single = 0
            Dim fRoll As Single = DegreeToRadian(LocYaw / 10.0F)
            Dim fYaw As Single = LocAngle - 900

            Dim bProcessAsLandUnit As Boolean

            'mfLastX = fMapWrapLocX
            mfLastX = LocX
            mfLastZ = LocZ
            miLastA = LocAngle
            miLastYaw = LocYaw
            mbForceTurrentUpdate = True
            miLastLocPitch = miLocPitch

            fYaw = DegreeToRadian(fYaw * 0.1F)

            'Ok, if we are on a planet we need to check something...
            Dim lTempY As Int32 = CInt(LocY)

            bProcessAsLandUnit = False

            If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                If goCurrentEnvir.oGeoObject Is Nothing = False Then

                    If ObjTypeID = ObjectType.eFacility Then
                        'ok, its a facility... check the four corners...
                        Dim lMinHt As Int32 = Int32.MaxValue
                        Dim lMaxHt As Int32 = Int32.MinValue
                        Dim lTotal As Int32 = 0
                        Dim lHalfVal As Int32
                        Dim lCenterPt As Int32
                        If oMesh Is Nothing = False Then
                            lHalfVal = CInt(Math.Ceiling(oMesh.ShieldXZRadius / 2.0F))
                        Else : lHalfVal = 100
                        End If
                        With CType(goCurrentEnvir.oGeoObject, Planet)
                            Dim lTemp As Int32 = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ - lHalfVal, False))
                            lTotal += lTemp
                            If lTemp > lMaxHt Then lMaxHt = lTemp
                            If lTemp < lMinHt Then lMinHt = lTemp
                            lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ - lHalfVal, False))
                            lTotal += lTemp
                            If lTemp > lMaxHt Then lMaxHt = lTemp
                            If lTemp < lMinHt Then lMinHt = lTemp
                            lTemp = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ + lHalfVal, False))
                            lTotal += lTemp
                            If lTemp > lMaxHt Then lMaxHt = lTemp
                            If lTemp < lMinHt Then lMinHt = lTemp
                            lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ + lHalfVal, False))
                            lTotal += lTemp
                            If lTemp > lMaxHt Then lMaxHt = lTemp
                            If lTemp < lMinHt Then lMinHt = lTemp
                            lTemp = CInt(.GetHeightAtPoint(LocX, LocZ, False))
                            lTotal += lTemp
                            lCenterPt = lTemp
                            If lTemp > lMaxHt Then lMaxHt = lTemp
                            If lTemp < lMinHt Then lMinHt = lTemp
                        End With
                        'placed this here to fix facilities from being in mountains should no longer be required going forward but we'll leave
                        If lMaxHt - lMinHt > 200 Then lMinHt = ((lTotal \ 5) + lCenterPt) \ 2
                        If lMinHt = Int32.MaxValue Then
                            LocY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(LocX, LocZ, False)) '+ oMesh.PlanetYAdjust
                        Else : LocY = lMinHt '+ oMesh.PlanetYAdjust
                        End If
                        If oMesh Is Nothing = False Then LocY += oMesh.PlanetYAdjust
                        bProcessAsLandUnit = False
                        Me.bForceSetY = False
                    Else
                        bProcessAsLandUnit = True
                    End If
                End If
            End If

            If bProcessAsLandUnit = True Then
                lTempY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(LocX, LocZ, True))
                If oMesh Is Nothing = False AndAlso oMesh.bLandBased = True Then
                    lTempY += oMesh.PlanetYAdjust

                    Dim lAdjAngle As Int32 = LocAngle
                    Dim fX1 As Single = 1.0F
                    Dim fZ1 As Single = 0.0F
                    Call RotatePoint(0, 0, fX1, fZ1, lAdjAngle / 10.0F)

                    'Now, get our angle...
                    Dim vN As Vector3 = CType(goCurrentEnvir.oGeoObject, Planet).GetTriangleNormal(LocX, LocZ)
                    Dim vU As Vector3 = New Vector3(0, 1, 0)
                    Dim vF As Vector3 = New Vector3(fX1, 0, fZ1)
                    Dim vcpNU As Vector3 = Vector3.Cross(vU, vN)
                    Dim fdpNU As Single = Vector3.Dot(vU, vN)
                    fYaw = RadianToDegree(CSng(-Math.Acos(Vector3.Dot(vF, vU))))

                    If fYaw = 90 Then fYaw = 0

                    If lAdjAngle < 0 Then lAdjAngle += 3600
                    fYaw += (lAdjAngle / 10.0F)
                    If fYaw < 0 Then fYaw += 360
                    If fYaw > 360 Then fYaw -= 360

                    'Now, assign our values...
                    If fYaw > 315 Or fYaw < 45 Then
                        fPitch = vcpNU.X
                        fRoll = vcpNU.Z
                    ElseIf fYaw < 135 Then
                        fPitch = -vcpNU.Z
                        fRoll = vcpNU.X
                    ElseIf fYaw < 225 Then
                        fPitch = -vcpNU.X
                        fRoll = -vcpNU.Z
                    Else
                        fPitch = vcpNU.Z
                        fRoll = -vcpNU.X
                    End If
                    fYaw = DegreeToRadian(fYaw)
                    LocY = lTempY
                Else
                    lTempY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetMaxHeight()) + 400
                    If LocY < lTempY Then LocY = lTempY
                    'lTempY += 1000
                    'lTempY = CInt(Math.Ceiling(lTempY / 100) * 100)

                    SetYTarget(lTempY)
                    If mlTargetY < lTempY Then mlTargetY = lTempY

                    If CInt(LocY) <> mlTargetY AndAlso Me.ObjTypeID = ObjectType.eUnit Then
                        If (CType(Me, BaseEntity).CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                            If CInt(LocY) < mlTargetY Then
                                LocY += 0.5F
                            Else : LocY -= 0.5F
                            End If
                        End If
                        'If Me.mlLastMoveYCycle = Int32.MinValue Then Me.mlLastMoveYCycle = glCurrentCycle - 30
                        'HandleMoveY(1, 0.01F, Me.mlLastMoveYCycle + 30)
                        'CType(Me, BaseEntity).CurrentStatus = CType(Me, BaseEntity).CurrentStatus Or elUnitStatus.eUnitMoving
                    End If

                    If glCurrentCycle - mlLastMoveYCycle > 30 Then
                        If miLocPitch < 0 Then
                            'mlLocPitch += TurnRate
                            mfTrueLocPitch += 2.0F
                            miLocPitch = CShort(mfTrueLocPitch)
                            If miLocPitch > 0 Then
                                miLocPitch = 0
                                mfTrueLocPitch = 0.0F
                            End If
                        ElseIf miLocPitch > 0 Then
                            'mlLocPitch -= TurnRate
                            mfTrueLocPitch -= 2.0F
                            miLocPitch = CShort(mfTrueLocPitch)
                            If miLocPitch < 0 Then
                                miLocPitch = 0
                                mfTrueLocPitch = 0.0F
                            End If
                        End If
                    End If

                    'lastly, we need to set our pitch
                    fPitch = CSng(miLocPitch / 10.0F) * gdRadPerDegree
                End If
            Else
                SetYTarget(0)

                If glCurrentCycle - mlLastMoveYCycle > 30 Then
                    If miLocPitch < 0 Then
                        'mlLocPitch += TurnRate
                        mfTrueLocPitch += 2.0F
                        miLocPitch = CShort(mfTrueLocPitch)
                        If miLocPitch > 0 Then
                            miLocPitch = 0
                            mfTrueLocPitch = 0.0F
                        End If
                    ElseIf miLocPitch > 0 Then
                        'mlLocPitch -= TurnRate
                        mfTrueLocPitch -= 2.0F
                        miLocPitch = CShort(mfTrueLocPitch)
                        If miLocPitch < 0 Then
                            miLocPitch = 0
                            mfTrueLocPitch = 0.0F
                        End If
                    End If
                End If

                'lastly, we need to set our pitch
                fPitch = CSng(miLocPitch / 10.0F) * gdRadPerDegree
            End If

            mfLastY = LocY

            matWorld = Matrix.Identity
            matWorld.RotateYawPitchRoll(fYaw, fPitch, fRoll)
            matTemp = Matrix.Identity
            matTemp.Translate(LocX, LocY, LocZ)
            matWorld.Multiply(matTemp)

            matPerpWorldMatrix = Matrix.Identity
            Dim fPerpYaw As Single = fYaw - 90
            If fPerpYaw < 0 Then fPerpYaw += 360
            matPerpWorldMatrix.RotateYawPitchRoll(fPerpYaw, fPitch, fRoll)
            matPerpWorldMatrix.Multiply(matTemp)

            matTemp = Nothing
        End If

        bForceGetWorldMatrix = False

        Return matWorld
    End Function

    Private mlLastSetYTarget As Int32
    Private mlLastBaseY As Int32 = Int32.MinValue
    Private Sub SetYTarget(ByVal lYOffset As Int32)
        Dim rcMe As Rectangle

        If bDoNotRunSetYTarget = True Then Return
        If oMesh Is Nothing Then Return
        If mlLastSetYTarget = glCurrentCycle Then Return
        mlLastSetYTarget = glCurrentCycle

        'ok, now, get our size..
        rcMe.Width = oMesh.YLevelRectSize * 2
        rcMe.Height = rcMe.Width
        rcMe.X = CInt(LocX)
        rcMe.Y = CInt(LocZ)

        rcMe.X -= oMesh.YLevelRectSize
        rcMe.Y -= oMesh.YLevelRectSize

        Dim lTargetY As Int32 = lYOffset
        Dim b2s As Boolean = ObjectID Mod 2 = 0
        Dim lMaxID As Int32 = Int32.MinValue

        Dim lMyAdjust As Int32 = Me.oMesh.YLevelEntityHeight * 5

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then

            If oEnvir.ObjTypeID = ObjectType.ePlanet Then b2s = False
            If oEnvir.lEntityIdx Is Nothing OrElse oEnvir.oEntity Is Nothing Then Return
            Dim lCurUB As Int32 = Math.Min(Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0)), oEnvir.oEntity.GetUpperBound(0))

            Try

                Dim oSystem As SolarSystem = Nothing
                If oEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso oEnvir.oGeoObject Is Nothing = False AndAlso CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.eSolarSystem Then oSystem = CType(oEnvir.oGeoObject, SolarSystem)

                If oSystem Is Nothing = False Then
                    Dim lBaseY As Int32 = 0 'oSystem.GetBaseY(LocX, LocZ)
                    If mlLastBaseY = Int32.MinValue Then
                        mlLastBaseY = lBaseY
                        Me.LocY = lBaseY
                    End If
                    If lBaseY < mlLastBaseY Then
                        Me.LocY -= (mlLastBaseY - lBaseY)
                        mlLastBaseY = lBaseY
                    ElseIf lBaseY > mlLastBaseY Then
                        Me.LocY += (lBaseY - mlLastBaseY)
                        mlLastBaseY = lBaseY
                    End If
                    'If Me.LocY < lBaseY Then Me.LocY = lBaseY
                    lTargetY = Math.Abs(lTargetY) + lBaseY
                    If Me.ObjTypeID = ObjectType.eFacility Then
                        For X As Int32 = 0 To oSystem.WormholeUB
                            If oSystem.moWormholes(X) Is Nothing = False Then
                                With oSystem.moWormholes(X)
                                    Dim rcTemp As Rectangle
                                    If .System1 Is Nothing = False AndAlso .System1.ObjectID = oSystem.ObjectID Then
                                        'system 1 coords
                                        rcTemp = New Rectangle(.LocX1 - 100, .LocY1 - 100, 200, 200) '.LocX1 + 100, .LocY1 + 100)
                                    Else
                                        'system 2 coords
                                        rcTemp = New Rectangle(.LocX2 - 100, .LocY2 - 100, 200, 200) ' .LocX2 + 100, .LocY2 + 100)
                                    End If
                                    If rcMe.IntersectsWith(rcTemp) = True Then
                                        lTargetY = -CInt(Me.oMesh.vecDeathSeqSize.Y + 300)
                                    End If
                                End With
                            End If
                        Next X
                    End If
                    

                End If


                For X As Int32 = 0 To lCurUB
                    'If oEnvir.lEntityIdx(X) <> -1 AndAlso (oEnvir.ObjTypeID = ObjectType.ePlanet OrElse (oEnvir.lEntityIdx(X) Mod 2 = 0) = b2s) AndAlso oEnvir.lEntityIdx(X) <> Me.ObjectID Then
                    If oEnvir.lEntityIdx(X) <> -1 AndAlso (oEnvir.lEntityIdx(X) Mod 2 = 0) = b2s AndAlso oEnvir.lEntityIdx(X) <> Me.ObjectID Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False AndAlso (oEnvir.ObjTypeID = ObjectType.eSolarSystem OrElse oEntity.ObjTypeID = ObjectType.eUnit) Then

                            If oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEntity.oMesh.bLandBased = True Then Continue For

                            Dim rcTemp As Rectangle

                            Dim lShieldXZ As Int32 = oEntity.oMesh.YLevelRectSize
                            Dim lMaxWH As Int32 = oEntity.oMesh.YLevelRectSize * 2

                            rcTemp.Width = lMaxWH
                            rcTemp.Height = lMaxWH

                            rcTemp.X = CInt(oEntity.LocX - lShieldXZ)
                            rcTemp.Y = CInt(oEntity.LocZ - lShieldXZ)
                            If rcMe.IntersectsWith(rcTemp) = True Then
                                'ok, is this unit bigger than me?
                                If (oEntity.oMesh.ShieldXZRadius > Me.oMesh.ShieldXZRadius) OrElse (oEntity.ObjectID > Me.ObjectID AndAlso CInt(oEntity.oMesh.ShieldXZRadius) = CInt(Me.oMesh.ShieldXZRadius)) Then
                                    If b2s = True Then 'AndAlso oEntity.ObjTypeID <> ObjectType.eFacility Then
                                        'Dim lNewVal As Int32 = CInt(oEntity.LocY - oEntity.oMesh.YLevelEntityHeight) ' - lMyAdjust)
                                        'lTargetY = Math.Min(lTargetY, lNewVal)
                                        'If lTargetY = lNewVal Then lMaxID = X
                                        lTargetY -= oEntity.oMesh.YLevelEntityHeight '- lMyAdjust
                                    Else
                                        'Dim lNewVal As Int32 = CInt(oEntity.LocY + oEntity.oMesh.YLevelEntityHeight) ' + lMyAdjust)
                                        'lTargetY = Math.Max(lTargetY, lNewVal)
                                        'If lTargetY = lNewVal Then lMaxID = X
                                        lTargetY += oEntity.oMesh.YLevelEntityHeight '+ lMyAdjust
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next X
            Catch
                'ignore, we'll hit it next time
                'TODO: This try catch may cause massive performance issues
            End Try
        End If

        mlTargetY = lTargetY
        If bForceSetY = True Then
            LocY = mlTargetY
            bForceSetY = False
        End If

    End Sub

    Public Function GetTurretMatrix() As Matrix
        If mbForceTurrentUpdate = True OrElse miLastTurretRot <> miTurretRot OrElse LocX <> mfLastX OrElse LocY <> mfLastY OrElse LocZ <> mfLastZ OrElse LocAngle <> miLastA OrElse LocYaw <> miLastYaw Then
            'matTurrWorld = Matrix.Identity
            mbForceTurrentUpdate = False
            mfLastX = LocX
            mfLastY = LocY
            mfLastZ = LocZ
            miLastA = LocAngle
            miLastYaw = LocYaw

            If timeGetTime - mlLastTurretSetTime > 10000 Then
                miTurretRot = LocAngle - 900S
                If miTurretRot < 0 Then miTurretRot += 3600S
                'mlLastTurretSetTime = timeGetTime
            End If

            Dim matTemp As Matrix
            Dim fYaw As Single

            Dim fX As Single
            Dim fZ As Single

            'Ok, if we are on a planet we need to check something...
            Dim lTempY As Int32
            If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then lTempY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(LocX, LocZ, True))
            If oMesh Is Nothing = False AndAlso oMesh.bLandBased = True Then
                lTempY += oMesh.PlanetYAdjust
            Else
                lTempY += 1000
            End If
            'TODO: Remarked this out, if turrets appear oddly now, this may be the cause
            'LocY = lTempY

            matTurrWorld = Matrix.Identity
            fYaw = DegreeToRadian(miTurretRot / 10.0F)
            matTurrWorld.RotateY(fYaw)
            matTemp = Matrix.Identity

            fX = -oMesh.lTurretZOffset
            fZ = 0
            RotatePoint(0, 0, fX, fZ, LocAngle / 10.0F)

            matTemp.Translate(LocX + fX, LocY, LocZ + fZ)
            matTurrWorld.Multiply(matTemp)

            matTemp = Nothing
        End If

        miLastTurretRot = miTurretRot
        Return matTurrWorld
    End Function

    Public Function GetScreenPos(ByRef moDevice As Device) As System.Drawing.Point
        'If Me.fMapWrapLocX = Single.MinValue Then Return Point.Empty
        If moDevice Is Nothing Then Return Point.Empty

        Dim ScreenPos As Vector3
        Dim vpMain As Viewport = moDevice.Viewport()
        Dim matProj As Matrix = moDevice.Transform.Projection
        Dim matView As Matrix = moDevice.Transform.View
        Dim matWorld As Matrix = Matrix.Identity

        ScreenPos = Vector3.Project(New Vector3(Me.LocX, Me.LocY, Me.LocZ), vpMain, matProj, matView, matWorld)

        If ScreenPos.X > Int32.MaxValue OrElse ScreenPos.X < Int32.MinValue Then
            Return Point.Empty
        End If
        If ScreenPos.Y > Int32.MaxValue OrElse ScreenPos.Y < Int32.MinValue Then
            Return Point.Empty
        End If
        Dim pt As Point = Point.Empty
        Try
            pt = New System.Drawing.Point(CInt(ScreenPos.X), CInt(ScreenPos.Y))
        Catch
        End Try

        Return pt
    End Function

    Public Sub SetWorldMatrix(ByVal mat As Matrix)
        matWorld = mat
    End Sub

    Public Sub SetWorldMatrixCurrent()
        'miLocPitch <> miLastLocPitch OrElse miLocPitch <> 0 Then
        bForceGetWorldMatrix = False
        'fMapWrapLocX = LocX
        mfLastX = LocX
        mfLastY = LocY
        mfLastZ = LocZ
        miLastA = LocAngle
        miLastYaw = LocYaw
        miLocPitch = 0
        miLastLocPitch = miLocPitch
    End Sub

    Public ReadOnly Property CurrentWorldMatrix() As Matrix
        Get
            Return matWorld
        End Get
    End Property
    Public ReadOnly Property CurrentTurretMatrix() As Matrix
        Get
            Return matTurrWorld
        End Get
    End Property

    Public Sub HandleMoveY(ByVal iTurnRate As Short, ByVal lCurrTime As Int32, ByVal fAcceleration As Single)
        Dim fElapsed As Single
        If mlLastYChgTime = Int32.MinValue Then fElapsed = 1 Else fElapsed = (lCurrTime - mlLastYChgTime) / 30.0F
        mlLastYChgTime = lCurrTime

        If fElapsed > 30 Then fElapsed = 30

        mlLastMoveYCycle = glCurrentCycle

        Dim yTilt As Byte = 0
        Dim lTempLocY As Int32 = CInt(LocY)

        Dim lMaxVal As Int16 = 250
        If Me.bDoNotRunSetYTarget = True Then lMaxVal = 500

        If mlTargetY < lTempLocY Then

            mfTrueLocPitch -= (iTurnRate * fElapsed)
            'mlLocPitch -= TurnRate
            miLocPitch = CShort(mfTrueLocPitch)
            If miLocPitch < -lMaxVal Then
                miLocPitch = -lMaxVal
                mfTrueLocPitch = miLocPitch
            End If

            'LocY -= Me.Acceleration * 15
            'If LocY < lTargetY Then LocY = lTargetY
            'yTilt = 1
        ElseIf mlTargetY > lTempLocY Then
            mfTrueLocPitch += (iTurnRate * fElapsed)
            'mlLocPitch += TurnRate
            miLocPitch = CShort(mfTrueLocPitch)
            If miLocPitch > lMaxVal Then
                miLocPitch = lMaxVal
                mfTrueLocPitch = miLocPitch
            End If
            'LocY += Me.Acceleration * 15
            'If LocY > lTargetY Then LocY = lTargetY
            'yTilt = 2
        Else
            If miLocPitch < 0 Then
                'mlLocPitch += TurnRate
                mfTrueLocPitch += (iTurnRate * fElapsed)
                miLocPitch = CShort(mfTrueLocPitch)
                If miLocPitch > 0 Then
                    miLocPitch = 0
                    mfTrueLocPitch = 0.0F
                End If
            ElseIf miLocPitch > 0 Then
                'mlLocPitch -= TurnRate
                mfTrueLocPitch -= (iTurnRate * fElapsed)
                miLocPitch = CShort(mfTrueLocPitch)
                If miLocPitch < 0 Then
                    miLocPitch = 0
                    mfTrueLocPitch = 0.0F
                End If
            End If
        End If

        'Now, change...
        Dim fTempYChg As Single = (CSng(miLocPitch / lMaxVal) * (fAcceleration * 30)) * fElapsed 'fMaxSpeed * 0.3F) * fElapsed 
        If bDoNotRunSetYTarget = True Then fTempYChg *= 20
        If mlTargetY > lTempLocY Then
            LocY += fTempYChg
            If mlTargetY < LocY Then LocY = mlTargetY
        ElseIf mlTargetY < LocY Then
            LocY += fTempYChg
            If mlTargetY > LocY Then LocY = mlTargetY
        End If

        'Now, for minimum
        'If glCurrentEnvirView = CurrentView.ePlanetView Then
        '	Dim oEnvir As BaseEnvironment = goCurrentEnvir
        '	If oEnvir Is Nothing = False Then
        '		If oEnvir.ObjTypeID = ObjectType.ePlanet Then
        '			If oEnvir.oGeoObject Is Nothing = False AndAlso CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
        '				Dim fTempYVal As Single = CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(LocX, LocZ, True) + 1000
        '				If LocY < fTempYVal Then LocY = fTempYVal
        '				If mlTargetY < fTempYVal Then mlTargetY = CInt(fTempYVal)
        '			End If
        '		End If
        '	End If
        'End If

    End Sub

    Public Sub New()
        ReDim muBurnMarks(49)
        For X As Int32 = 0 To muBurnMarks.GetUpperBound(0)
            muBurnMarks(X).bActive = False
        Next X
    End Sub
End Class

