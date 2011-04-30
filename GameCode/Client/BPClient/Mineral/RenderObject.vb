Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class RenderObject
    Inherits Epica_GUID

    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Public oMesh As EpicaMesh

    'for a renderable unit we track only that which we need
    Public LocX As Single   'int32
    Public LocY As Int32
    Public LocZ As Single   'int32  
    Public LocAngle As Int16
    Public LocYaw As Int16

    Public fMapWrapLocX As Single = Single.MinValue

    Private matWorld As Matrix
    'By setting these to the min values, they ensure that oure matrix gets set the first time...
    Private mfLastX As Single = Single.MinValue
    Private mlLastY As Int32 = Int32.MinValue
    Private mfLastZ As Single = Single.MinValue
    Private miLastA As Int16 = Int16.MinValue
    Private miLastY As Int16 = Int16.MinValue
    Private mbForceTurrentUpdate As Boolean = False

    'These get set in the add object and in the movement events
    Public CellLocX As Int32
    Public CellLocZ As Int32

    Public yVisibility As Byte

    'For turrets
    Private matTurrWorld As Matrix
    Private miTurretRot As Int16 'rotation of the turret in 10s degrees... set from Fire Weapon
    Private miLastTurretRot As Int16
    Private mlLastTurretSetTime As Int32

    Private mlFireFromLocID(3) As Int32

    'Burn FX
    Protected mlBurnFXActive() As Int32 = Nothing

    'For EngineFX color
    Public clrEngineFX As System.Drawing.Color = Color.FromArgb(255, 64, 192, 255)
    Public clrShieldFX As System.Drawing.Color = Color.FromArgb(255, 0, 255, 255)

    Public bCulled As Boolean       'indicates whether this object was culled this frame (if so, don't render its burn fx and other related fx)

    Public Function GetBurnFXLoc(ByVal lSide As Int32) As Int32
        If mlBurnFXActive Is Nothing Then
            If oMesh.muBurnLocs Is Nothing Then Return -1

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
        If lVal > oMesh.muFireLocs(ySide).GetUpperBound(0) Then lVal = 0
        mlFireFromLocID(ySide) = lVal + 1
        If oMesh.muFireLocs(ySide).GetUpperBound(0) >= lVal Then
            Return oMesh.muFireLocs(ySide)(lVal).GetFireFromLoc(Me)
        Else : Return New Vector3(LocX, LocY, LocZ)
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

    Public Function GetWorldMatrix() As Matrix
        If fMapWrapLocX = Single.MinValue Then fMapWrapLocX = LocX

        If fMapWrapLocX <> mfLastX OrElse LocY <> mlLastY OrElse LocZ <> mfLastZ OrElse LocAngle <> miLastA OrElse LocYaw <> miLastY Then
            Dim matTemp As Matrix
            Dim fPitch As Single = 0
            Dim fRoll As Single = DegreeToRadian(LocYaw / 10.0F)
            Dim fYaw As Single = LocAngle - 900

            Dim bProcessAsLandUnit As Boolean

            mfLastX = fMapWrapLocX
            mfLastZ = LocZ
            miLastA = LocAngle
            miLastY = LocYaw
            mbForceTurrentUpdate = True

            fYaw = DegreeToRadian(fYaw / 10.0F)

            'Ok, if we are on a planet we need to check something...
            Dim lTempY As Int32 = LocY

            bProcessAsLandUnit = False

            If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                If goCurrentEnvir.oGeoObject Is Nothing = False Then

                    If ObjTypeID = ObjectType.eFacility Then
                        LocY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(LocX, LocZ)) + oMesh.PlanetYAdjust
                        bProcessAsLandUnit = False
                    Else
                        bProcessAsLandUnit = True
                    End If
                End If
            End If

            If bProcessAsLandUnit = True Then
                lTempY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(LocX, LocZ))
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
                Else
                    lTempY += 1000        '???
                    'Yaw is not affected
                End If
                LocY = lTempY
            End If

            If Me.ObjectID = 467 Then LocY += CInt(Rnd() * 3000)
            mlLastY = LocY

            matWorld = Matrix.Identity
            matWorld.RotateYawPitchRoll(fYaw, fPitch, fRoll)
            matTemp = Matrix.Identity
            'matTemp.Translate(LocX, LocY, LocZ)
            matTemp.Translate(fMapWrapLocX, LocY, LocZ)
            matWorld.Multiply(matTemp)
            matTemp = Nothing
        End If

        Return matWorld
    End Function

    Public Function GetTurretMatrix() As Matrix
        If mbForceTurrentUpdate = True OrElse miLastTurretRot <> miTurretRot OrElse fMapWrapLocX <> mfLastX OrElse LocY <> mlLastY OrElse LocZ <> mfLastZ OrElse LocAngle <> miLastA OrElse LocYaw <> miLastY Then
            'matTurrWorld = Matrix.Identity
            mbForceTurrentUpdate = False
            mfLastX = fMapWrapLocX
            mlLastY = LocY
            mfLastZ = LocZ
            miLastA = LocAngle
            miLastY = LocYaw



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
            Dim lTempY As Int32 = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(LocX, LocZ))
            If oMesh Is Nothing = False AndAlso oMesh.bLandBased = True Then
                lTempY += oMesh.PlanetYAdjust
            Else
                lTempY += 1000
            End If
            LocY = lTempY

            matTurrWorld = Matrix.Identity
            fYaw = DegreeToRadian(miTurretRot / 10.0F)
            matTurrWorld.RotateY(fYaw)
            matTemp = Matrix.Identity

            fX = -oMesh.lTurretZOffset
            fZ = 0
            RotatePoint(0, 0, fX, fZ, LocAngle / 10.0F)

            matTemp.Translate(fMapWrapLocX + fX, LocY, LocZ + fZ)
            matTurrWorld.Multiply(matTemp)

            matTemp = Nothing
        End If

        miLastTurretRot = miTurretRot
        Return matTurrWorld
    End Function

    Public Function GetScreenPos(ByRef moDevice As Device) As System.Drawing.Point
        If Me.fMapWrapLocX = Single.MinValue Then Return Point.Empty

        Dim ScreenPos As Vector3
        Dim vpMain As Viewport = moDevice.Viewport()
        Dim matProj As Matrix = moDevice.Transform.Projection
        Dim matView As Matrix = moDevice.Transform.View
        Dim matWorld As Matrix = Matrix.Identity

        ScreenPos = Vector3.Project(New Vector3(Me.fMapWrapLocX, Me.LocY, Me.LocZ), vpMain, matProj, matView, matWorld)

        Return New System.Drawing.Point(CInt(ScreenPos.X), CInt(ScreenPos.Y))
    End Function

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
End Class

