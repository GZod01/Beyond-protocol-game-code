Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class MineralCache
    Inherits Base_GUID

    Public Event DetailsArrived(ByVal oCache As MineralCache)

    Public LocX As Int32
    Public LocY As Int32 = Int32.MinValue
    Public LocZ As Int32

    Public CacheTypeID As Byte

    Public lMapWrapLocX As Int32

    Public MineralID As Int32

    Public Concentration As Int32 = Int32.MinValue
    Public Quantity As Int32 = Int32.MinValue

    Public ReadOnly Property oMineral() As Mineral
        Get
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = MineralID Then
                    Return goMinerals(X)
                End If
            Next X
            Return Nothing
        End Get
	End Property
	Private mlDebrisTexIdx As Int32 = -1
	Public ReadOnly Property lDebrisTexIdx() As Int32
		Get
			If mlDebrisTexIdx = -1 Then
				If (CacheTypeID And MineralCacheType.eFlying) <> 0 Then
					mlDebrisTexIdx = 0
				ElseIf (CacheTypeID And MineralCacheType.eNaval) <> 0 Then
					mlDebrisTexIdx = 0
				Else : mlDebrisTexIdx = 0		  'ground and all others
				End If
			End If
			Return mlDebrisTexIdx
		End Get
	End Property


	Public Shared Function GetCacheMaterial() As Material
		Dim oEnvir As BaseEnvironment = goCurrentEnvir
		Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)
		If oEnvir Is Nothing = False Then
			If oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEnvir.oGeoObject Is Nothing = False Then
				Dim yType As PlanetType = CType(CType(oEnvir.oGeoObject, Planet).MapTypeID, PlanetType)

				Select Case yType
					Case PlanetType.eAcidic
						clrVal = muSettings.AcidMineralCache
					Case PlanetType.eAdaptable
						clrVal = muSettings.AdaptableMineralCache
					Case PlanetType.eBarren
						clrVal = muSettings.BarrenMineralCache
					Case PlanetType.eDesert
						clrVal = muSettings.DesertMineralCache
					Case PlanetType.eGeoPlastic
						clrVal = muSettings.LavaMineralCache
					Case PlanetType.eTerran
						clrVal = muSettings.TerranMineralCache
					Case PlanetType.eTundra
						clrVal = muSettings.IceMineralCache
					Case PlanetType.eWaterWorld
						clrVal = muSettings.WaterworldMineralCache
				End Select
			End If
		End If

		Dim oMat As Material
		With oMat
			.Ambient = clrVal
			.Diffuse = .Ambient
			.Specular = .Ambient
			.Emissive = System.Drawing.Color.FromArgb(255, 0, 0, 0)
			.SpecularSharpness = 17
		End With

		Return oMat
	End Function

    Public Shared CacheMesh() As Mesh
	Public Shared texCache As Texture

	Public Shared DebrisMesh As Mesh
	Public Shared DebrisTex() As Texture

    Public Sub DetailsRefreshed()
        RaiseEvent DetailsArrived(Me)
    End Sub

    Private matWorld As Matrix
    Private mlPrevX As Int32
    Private mlPrevY As Int32
    Private mlPrevZ As Int32
    Public Function GetWorldMatrix() As Matrix
        Try
            If lMapWrapLocX <> mlPrevX OrElse LocY <> mlPrevY OrElse LocZ <> mlPrevZ OrElse Me.CacheTypeID <> MineralCacheType.eMineable Then

                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                If oEnvir Is Nothing Then Return matWorld

                mlPrevX = lMapWrapLocX
                mlPrevY = LocY
                mlPrevZ = LocZ

                Dim fPitch As Single
                Dim fRoll As Single
                Dim fYaw As Single
                Dim matTemp As Matrix

                If oEnvir Is Nothing = False AndAlso oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEnvir.oGeoObject Is Nothing = False Then
                    fPitch = 0
                    fRoll = DegreeToRadian(0)
                    fYaw = -900

                    Dim oObj As Object = oEnvir.oGeoObject
                    If oObj Is Nothing Then
                        LocY = mlPrevY - 1
                        Return matWorld
                    End If
                    If CType(oObj, Base_GUID).ObjTypeID <> ObjectType.ePlanet Then
                        LocY = mlPrevY - 1
                        Return matWorld
                    End If

                    Dim oPlanet As Planet = CType(oObj, Planet)

                    fYaw = DegreeToRadian(fYaw / 10.0F)

                    If oPlanet Is Nothing = False Then
                        Dim fTmpY As Single = oPlanet.GetHeightAtPoint(lMapWrapLocX - 73, LocZ - 95, False)
                        fTmpY += oPlanet.GetHeightAtPoint(lMapWrapLocX + 73, LocZ - 95, False)
                        fTmpY += oPlanet.GetHeightAtPoint(lMapWrapLocX - 73, LocZ + 95, False)
                        fTmpY += oPlanet.GetHeightAtPoint(lMapWrapLocX + 73, LocZ + 95, False)
                        fTmpY += 16
                        fTmpY /= 4.0F

                        fTmpY += oPlanet.GetHeightAtPoint(lMapWrapLocX, LocZ, False) + 4
                        fTmpY /= 2.0F
                        LocY = CInt(fTmpY) 'oPlanet.GetHeightAtPoint(vecLoc.X, vecLoc.Z) + 4

                        Dim lAdjAngle As Int32 = 0
                        Dim fX1 As Single = 1.0F
                        Dim fZ1 As Single = 0.0F
                        Call RotatePoint(0, 0, fX1, fZ1, lAdjAngle / 10.0F)

                        'Now, get our angle...
                        Dim vN As Vector3 = oPlanet.GetTriangleNormal(lMapWrapLocX, LocZ)
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
                        LocY = mlPrevY - 1
                    End If

                End If

                If Me.CacheTypeID <> MineralCacheType.eMineable Then
                    If mfBaseYaw = Single.MinValue OrElse mfBasePitch = Single.MinValue OrElse mfBaseRoll = Single.MinValue Then
                        'Rnd(-1)
                        'Randomize(Me.ObjectID)
                        mfBaseYaw = Rnd() * gdPi
                        mfBasePitch = Rnd() * gdPi
                        mfBaseRoll = Rnd() * gdPi
                        If oEnvir Is Nothing = False AndAlso oEnvir.ObjTypeID <> ObjectType.ePlanet Then
                            mfYawRot = ((Rnd() * 0.005F) - 0.0025F)
                            mfPitchRot = ((Rnd() * 0.005F) - 0.0025F)
                            mfRollRot = ((Rnd() * 0.005F) - 0.0025F)
                        End If
                        Randomize(Val(Now.ToString("MMddHHmm")))
                    End If

                    If oEnvir Is Nothing = False AndAlso oEnvir.ObjTypeID <> ObjectType.ePlanet Then
                        'ok, rotation time
                        mfBaseYaw += mfYawRot
                        mfBasePitch += mfPitchRot
                        mfBaseRoll += mfRollRot
                    End If

                    fYaw = mfBaseYaw
                    fPitch = mfBasePitch
                    fRoll = mfBaseRoll
                End If

                matWorld = Matrix.Identity
                matWorld.RotateYawPitchRoll(fYaw, fPitch, fRoll)
                matTemp = Matrix.Identity
                matTemp.Translate(lMapWrapLocX, LocY, LocZ)
                matWorld.Multiply(matTemp)
                matTemp = Nothing
            End If
        Catch
            LocY = mlPrevY - 1
        End Try

        Return matWorld
    End Function
	Private mfBaseYaw As Single = Single.MinValue
	Private mfBasePitch As Single = Single.MinValue
	Private mfBaseRoll As Single = Single.MinValue
	Private mfYawRot As Single
	Private mfPitchRot As Single
	Private mfRollRot As Single

    Private mlModelIdx As Int32 = -1
    Public ReadOnly Property ModelIndex() As Int32
        Get
            If mlModelIdx = -1 Then
                mlModelIdx = Me.ObjectID Mod 3
            End If
            Return mlModelIdx
        End Get
    End Property

    Public Shared Sub Initialize3DResources()
		Device.IsUsingEventHandlers = False

		If MineralCache.CacheMesh Is Nothing = False Then
			For X As Int32 = 0 To MineralCache.CacheMesh.GetUpperBound(0)
				If MineralCache.CacheMesh(X) Is Nothing = False AndAlso MineralCache.CacheMesh(X).Disposed = False Then
					MineralCache.CacheMesh(X).Dispose()
				End If
				MineralCache.CacheMesh(X) = Nothing
			Next X
		End If
		If MineralCache.texCache Is Nothing = False Then
			If MineralCache.texCache.Disposed = False Then MineralCache.texCache.Dispose()
			MineralCache.texCache = Nothing
		End If
		If MineralCache.DebrisMesh Is Nothing = False Then
			If MineralCache.DebrisMesh.Disposed = False Then MineralCache.DebrisMesh.Dispose()
			MineralCache.DebrisMesh = Nothing
		End If
		If MineralCache.DebrisTex Is Nothing = False Then
			For X As Int32 = 0 To MineralCache.DebrisTex.GetUpperBound(0)
				If MineralCache.DebrisTex(X) Is Nothing = False AndAlso MineralCache.DebrisTex(X).Disposed = False Then
					MineralCache.DebrisTex(X).Dispose()
				End If
				MineralCache.DebrisTex(X) = Nothing
			Next X
		End If

		ReDim MineralCache.CacheMesh(2)
		MineralCache.CacheMesh(0) = goResMgr.LoadScratchMeshNoTextures("MinCache1.x", "MinCache.pak")
		MineralCache.CacheMesh(1) = goResMgr.LoadScratchMeshNoTextures("MinCache2.x", "MinCache.pak")
		MineralCache.CacheMesh(2) = goResMgr.LoadScratchMeshNoTextures("MinCache3.x", "MinCache.pak")
		MineralCache.texCache = goResMgr.GetTexture("MinCache1.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "MinCacheTex.pak")
		MineralCache.DebrisMesh = goResMgr.LoadScratchMeshNoTextures("debris.x", "MinCache.pak")

		ReDim MineralCache.DebrisTex(0)
		MineralCache.DebrisTex(0) = goResMgr.GetTexture("ff06.bmp", GFXResourceManager.eGetTextureType.ModelTexture, "textures.pak")
		'put in others here...

		Device.IsUsingEventHandlers = True
	End Sub
	Public Shared Function GetComponentTypeID(ByVal lValue As Int32) As Int16
		If (lValue And MineralCacheType.eArmorDebris) <> 0 Then Return ObjectType.eArmorTech
		If (lValue And MineralCacheType.eEngineDebris) <> 0 Then Return ObjectType.eEngineTech
		If (lValue And MineralCacheType.eRadarDebris) <> 0 Then Return ObjectType.eRadarTech
		If (lValue And MineralCacheType.eShieldDebris) <> 0 Then Return ObjectType.eShieldTech
		If (lValue And MineralCacheType.eWeaponDebris) <> 0 Then Return ObjectType.eWeaponTech
		Return 0
    End Function
    Public Shared Sub Clear3DResources()
        If MineralCache.CacheMesh Is Nothing = False Then
            For X As Int32 = 0 To MineralCache.CacheMesh.GetUpperBound(0)
                If MineralCache.CacheMesh(X) Is Nothing = False Then MineralCache.CacheMesh(X).Dispose()
                MineralCache.CacheMesh(X) = Nothing
            Next X
        End If
        If MineralCache.texCache Is Nothing = False Then MineralCache.texCache.Dispose()
        MineralCache.texCache = Nothing
        If MineralCache.DebrisMesh Is Nothing = False Then MineralCache.DebrisMesh.Dispose()
        MineralCache.DebrisMesh = Nothing
        If MineralCache.DebrisTex Is Nothing = False Then
            For X As Int32 = 0 To MineralCache.DebrisTex.GetUpperBound(0)
                If MineralCache.DebrisTex(X) Is Nothing = False Then
                    MineralCache.DebrisTex(X).Dispose()
                End If
                MineralCache.DebrisTex(X) = Nothing
            Next X
        End If
    End Sub
End Class
