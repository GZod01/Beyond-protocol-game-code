Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class Camera
    Private Const ml_ROTATE_CAMERA_THRESHOLD As Int32 = 5

    Public mlCameraX As Int32
    Public mlCameraY As Int32
    Public mlCameraZ As Int32
    Public mlCameraAtX As Int32
    Public mlCameraAtY As Int32
    Public mlCameraAtZ As Int32

    Public bScrollLeft As Boolean = False
    Public bScrollRight As Boolean = False
    Public bScrollUp As Boolean = False
    Public bScrollDown As Boolean = False

    Private mlScrollRate() As Int32
    Private mlWheelScrollRate As Int32

    Private muGrid() As CustomVertex.PositionTextured
    Private moGridTex As Texture
    Private matGrid As Material
    Private mbGridReady As Boolean = False

    Private mfSS_Intensity As Single = 0.0F
    Private mbSS As Boolean = False

    Public SelectionBoxStart As Point
    Public SelectionBoxEnd As Point
    Public ShowSelectionBox As Boolean = False

    Public ScrollEdgeWidth As Int32

    Private msw_SS As Stopwatch

	Private mlTrackingIndex As Int32 = -1

	Private mbZoomToLoc As Boolean = False
	Private mlZoomToLocX As Int32
	Private mlZoomToLocZ As Int32
    Private miZoomToAngle As Int16

	Public Property TrackingIndex() As Int32
		Get
			Return mlTrackingIndex
		End Get
		Set(ByVal value As Int32)
			mlTrackingIndex = value
			If mlTrackingIndex <> -1 Then
				mfPrecCX = mlCameraX
				mfPrecCZ = mlCameraZ
				mlCameraAtX = CInt(goCurrentEnvir.oEntity(mlTrackingIndex).LocX)
				mlCameraAtZ = CInt(goCurrentEnvir.oEntity(mlTrackingIndex).LocZ)
			End If
		End Set
	End Property

	Public Sub New()
		Dim oINI As InitFile = New InitFile()
		Dim X As Int32

		ReDim mlScrollRate(CurrentView.eFullScreenInterface - 1)		'fullscreen interface is the one after the last one

		For X = 0 To CurrentView.eFullScreenInterface - 1
            Dim lDef As Int32 = 400
            If X = CurrentView.eSystemMapView1 Then lDef = 40
            mlScrollRate(X) = CInt(Val(oINI.GetString("SETTINGS", "ScrollRate" & X, lDef.ToString)))
		Next X

		mlWheelScrollRate = CInt(Val(oINI.GetString("SETTINGS", "WheelScrollRate", "50")))

		ScrollEdgeWidth = CInt(Val(oINI.GetString("SETTINGS", "ScrollArea", "10")))

		oINI = Nothing
	End Sub

	Public Property ZoomRate() As Int32
		Get
			Return mlWheelScrollRate
		End Get
		Set(ByVal Value As Int32)
			Dim oINI As New InitFile()

			mlWheelScrollRate = Value
			oINI.WriteString("SETTINGS", "WheelScrollRate", Value.ToString)
			oINI = Nothing
		End Set
	End Property

	Public Property ScrollRate(ByVal ForEnvirView As Int32) As Int32
		Get
			If ForEnvirView < mlScrollRate.Length Then
				Return mlScrollRate(ForEnvirView)
			Else
				Return 15
			End If
		End Get
		Set(ByVal Value As Int32)
			Dim oINI As New InitFile()
			If ForEnvirView < mlScrollRate.Length Then
				mlScrollRate(ForEnvirView) = Value
				oINI.WriteString("SETTINGS", "ScrollRate" & ForEnvirView, Value.ToString())
			End If
			oINI = Nothing
		End Set
    End Property

    Public Sub SimplyPlaceCamera(ByVal lX As Int32, ByVal lY As Int32, ByVal lZ As Int32)
        Dim lDiffX As Int32 = mlCameraX - mlCameraAtX
        Dim lDiffY As Int32 = mlCameraY - mlCameraAtY
        Dim lDiffZ As Int32 = mlCameraZ - mlCameraAtZ

        mlCameraAtX = lX : mlCameraAtY = lY : mlCameraAtZ = lZ
        mlCameraX = lX + lDiffX : mlCameraY = lY + lDiffY : mlCameraZ = lZ + lDiffZ
    End Sub

	'This sub is called every frame
	Private mfCVelX As Single
	Private mfCVelZ As Single
	Private mfPrecCX As Single
	Private mfPrecCZ As Single

	Private mlRotateCameraThresholdX As Int32 = 0
	Private mlRotateCameraThresholdY As Int32 = 0

	Public Sub ResetRotateCameraThresholds()
		mlRotateCameraThresholdX = 0
		mlRotateCameraThresholdY = 0
	End Sub

	Public Sub SetupMatrices(ByRef oDevice As Device, ByVal CurrentEnvirView As Int32)
		Dim fX As Single = mlCameraX
		Dim fY As Single = mlCameraY
		Dim fZ As Single = mlCameraZ

		Dim fAtX As Single = mlCameraAtX
		Dim fAtY As Single = mlCameraAtY
		Dim fAtZ As Single = mlCameraAtZ

		If mbZoomToLoc = True Then

			If mlZoomToLocX = mlCameraAtX AndAlso mlZoomToLocZ = mlCameraAtZ Then mbZoomToLoc = False

			Dim fTLX As Single = mlZoomToLocX
			Dim fTLZ As Single = mlZoomToLocZ
			Dim fDestAtX As Single = fTLX
			Dim fDestAtZ As Single = fTLZ
			Dim fDestX As Single = fTLX - 500
			Dim fDestZ As Single = fTLZ
			RotatePoint(CInt(fTLX), CInt(fTLZ), fDestX, fDestZ, miZoomToAngle / 10.0F)

			Dim vecResult As Vector3
			vecResult.X = mfPrecCX + ((fDestX - mfPrecCX) / 16.0F)
			vecResult.Z = mfPrecCZ + ((fDestZ - mfPrecCZ) / 16.0F)

			mfPrecCX = vecResult.X : mfPrecCZ = vecResult.Z

			fX = mfPrecCX
			fZ = mfPrecCZ

			fAtX = fX - 500
			fAtZ = fZ

			Dim fAngle As Single = LineAngleDegrees(CInt(fTLX), CInt(fTLZ), CInt(fX), CInt(fZ))
			RotatePoint(CInt(fX), CInt(fZ), fAtX, fAtZ, fAngle)

			If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
				If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
					Dim fHt As Single = CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(CInt(fX), CInt(fZ), True)
					fY = fHt + muSettings.lPlanetViewCameraY
				End If
			End If
		ElseIf TrackingIndex <> -1 AndAlso goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.lEntityUB >= TrackingIndex AndAlso goCurrentEnvir.lEntityIdx(TrackingIndex) <> -1 AndAlso (glCurrentEnvirView = CurrentView.ePlanetView OrElse glCurrentEnvirView = CurrentView.eSystemView) Then
			With goCurrentEnvir.oEntity(TrackingIndex)
				Dim fTLX As Single = .LocX
				Dim fTLZ As Single = .LocZ
				Dim fDestAtX As Single = fTLX
				Dim fDestAtZ As Single = fTLZ
				Dim fDestX As Single = fTLX - 3000
				Dim fDestZ As Single = fTLZ
                RotatePoint(fTLX, fTLZ, fDestX, fDestZ, .LocAngle / 10.0F)

				Dim vecResult As Vector3
				vecResult.X = mfPrecCX + ((fDestX - mfPrecCX) / 16.0F)
				vecResult.Z = mfPrecCZ + ((fDestZ - mfPrecCZ) / 16.0F)

				mfPrecCX = vecResult.X : mfPrecCZ = vecResult.Z

				fX = mfPrecCX
				fZ = mfPrecCZ

				fAtX = fX - 3000
				fAtZ = fZ

                Dim fAngle As Single = LineAngleDegreesF(fTLX, fTLZ, fX, fZ)
                'RotatePoint(CInt(fX), CInt(fZ), fAtX, fAtZ, fAngle)
                RotatePoint(fX, fZ, fAtX, fAtZ, fAngle)
			End With
		End If

		mlCameraX = CInt(fX) : mlCameraY = CInt(fY) : mlCameraZ = CInt(fZ)
		mlCameraAtX = CInt(fAtX) : mlCameraAtY = CInt(fAtY) : mlCameraAtZ = CInt(fAtZ)

        ''Handle our screenshake...
        'If muSettings.ScreenShakeEnabled = True AndAlso mbSS = True Then
        '	If msw_SS Is Nothing Then msw_SS = New Stopwatch()
        '	If msw_SS.IsRunning = False Then msw_SS.Start()
        '	Dim fElapsed As Single = msw_SS.ElapsedMilliseconds / 30.0F
        '	msw_SS.Reset()
        '	msw_SS.Start()

        '	If fElapsed <> 0 Then

        '		'Dim fTemp As Single = 0.05F * fElapsed

        '		'now, multiple intensity by ftemp
        '		'mfSS_Intensity -= (mfSS_Intensity / (30 * fElapsed))
        '		'mfSS_Intensity *= fTemp

        '		If mfSS_Intensity <= 1 Then
        '			mfSS_Intensity = 0
        '		Else
        '			Dim fHalfIntensity As Single = mfSS_Intensity / 2
        '                  fX += ((Rnd() * mfSS_Intensity) - fHalfIntensity)
        '                  fY += ((Rnd() * mfSS_Intensity) - fHalfIntensity)
        '                  fZ += ((Rnd() * mfSS_Intensity) - fHalfIntensity)

        '			mfSS_Intensity -= (fHalfIntensity / 2.0F)
        '		End If

        '	End If
        'End If
        ''End of screenshake
        Try
            oDevice.Transform.View = Matrix.LookAtLH(New Vector3(fX, fY, fZ), _
              New Vector3(fAtX, fAtY, fAtZ), New Vector3(0.0F, 1.0F, 0.0F))    'up is always this

            If CurrentEnvirView = CurrentView.ePlanetView Then
                oDevice.Transform.Projection = Matrix.PerspectiveFovLH(0.7853982F, CSng(oDevice.PresentationParameters.BackBufferWidth / oDevice.PresentationParameters.BackBufferHeight), muSettings.TerrainNearClippingPlane, muSettings.FarClippingPlane)
            Else
                oDevice.Transform.Projection = Matrix.PerspectiveFovLH(0.7853982F, CSng(oDevice.PresentationParameters.BackBufferWidth / oDevice.PresentationParameters.BackBufferHeight), muSettings.NearClippingPlane, muSettings.FarClippingPlane)
            End If
        Catch
        End Try
    End Sub

    'This sub is called every frame
    Private mbScrolledLastTime As Boolean = False
    Private mlStartCameraX As Int32
    Private mlStartCameraZ As Int32

    Private mlPrevCX As Int32
    Private mlPrevCY As Int32
    Private mlPrevCZ As Int32

    Public Sub CheckUpdateListenerLoc()
        'Update our sound object's listener...
        If mlPrevCX <> mlCameraX OrElse mlPrevCY <> mlCameraY OrElse mlPrevCZ <> mlCameraZ Then
            mlPrevCX = mlCameraX : mlPrevCY = mlCameraY : mlPrevCZ = mlCameraZ
            If goSound Is Nothing = False Then
                'Ensure that no values equal each other
                If (mlCameraX <> mlCameraAtX OrElse mlCameraZ <> mlCameraAtZ OrElse mlCameraY <> mlCameraAtY) Then 'AndAlso glCurrentCycle Mod 10 = 0 Then
                    goSound.UpdateListenerLoc(New Vector3(mlCameraX, mlCameraY, mlCameraZ), Vector3.Normalize(Vector3.Subtract(New Vector3(mlCameraAtX, mlCameraAtY, mlCameraAtZ), New Vector3(mlCameraX, mlCameraY, mlCameraZ))), New Vector3(0, 1, 0))
                End If
            End If
        End If
    End Sub

    Public Sub ScrollCamera(ByVal CurrentEnvirView As Int32)
		If glCurrentEnvirView > CurrentView.eFullScreenInterface Then Return
		'If bScrollLeft = False AndAlso bScrollUp = False AndAlso bScrollDown = False AndAlso bScrollRight = False Then Return

		If bScrollLeft = True OrElse bScrollUp = True OrElse bScrollDown = True OrElse bScrollRight = True Then
			If NewTutorialManager.TutorialOn = True Then
				If goUILib Is Nothing = False AndAlso goUILib.CommandAllowed(True, "ScrollCamera") = False Then Return
				NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eScrollCamera, -1, -1, -1, "")
			End If
			mbZoomToLoc = False
		End If

		'handle view scrolling... however, this comes with a twist... we need to determine our angle to the target
		' deltaXx is the delta when scrolling camera's X along the X world axis
		' deltaXz is the delta when scrolling camera's X along the Z world axis
		' deltaZx is the delta when scrolling camera's Z along the X world axis
		' deltaZz is the delta when scrolling camera's Z along the Z world axis
		' which is to say, when changing X (scrolling horizontally), deltaXx is change to X and deltaXz is change to Z
		' and when changing Z (scrolling vertically), deltaZx is change to X and deltaZz is change to Z
		Dim vecCameraLoc As Vector3 = New Vector3(mlCameraX, mlCameraY, mlCameraZ)
		Dim vecCameraAt As Vector3 = New Vector3(mlCameraAtX, mlCameraAtY, mlCameraAtZ)
		Dim vecTemp As Vector3 = Vector3.Subtract(vecCameraLoc, vecCameraAt)

		Dim deltaXx As Int32
		Dim deltaXz As Int32
		Dim deltaZx As Int32
		Dim deltaZz As Int32

		Dim lScrollRate As Int32

		vecTemp.Normalize()
		Dim vecDot As Vector3 = Vector3.Cross(vecTemp, New Vector3(0, 1, 0))


		If CurrentEnvirView < mlScrollRate.Length Then
			lScrollRate = mlScrollRate(CurrentEnvirView)
		Else : lScrollRate = 15
        End If

        If bScrollDown = True OrElse bScrollLeft = True OrElse bScrollRight = True OrElse bScrollUp = True Then
            If mbScrolledLastTime = False Then
                mbScrolledLastTime = True
                mlStartCameraX = mlCameraAtX
                mlStartCameraZ = mlCameraAtZ
            End If
        ElseIf mbScrolledLastTime = True Then
            mbScrolledLastTime = False
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eScrollCamera, (mlStartCameraX - mlCameraAtX), (mlStartCameraZ - mlCameraAtZ))
        End If


		Try
			deltaXx = CInt(vecDot.X * lScrollRate)
			deltaXz = CInt(vecDot.Z * lScrollRate)
			deltaZx = CInt(vecTemp.X * lScrollRate)
			deltaZz = CInt(vecTemp.Z * lScrollRate)

			If bScrollLeft Then
				mlCameraX -= deltaXx
				mlCameraAtX -= deltaXx
				mlCameraZ -= deltaXz
				mlCameraAtZ -= deltaXz
				'Clear our tracking index
				TrackingIndex = -1
			ElseIf bScrollRight Then
				mlCameraX += deltaXx
				mlCameraAtX += deltaXx
				mlCameraZ += deltaXz
				mlCameraAtZ += deltaXz
				'Clear our tracking index
				TrackingIndex = -1
			End If
			If bScrollUp Then
				mlCameraZ -= deltaZz
				mlCameraAtZ -= deltaZz
				mlCameraX -= deltaZx
				mlCameraAtX -= deltaZx
				'Clear our tracking index
				TrackingIndex = -1
			ElseIf bScrollDown Then
				mlCameraZ += deltaZz
				mlCameraAtZ += deltaZz
				mlCameraX += deltaZx
				mlCameraAtX += deltaZx
				'Clear our tracking index
				TrackingIndex = -1
			End If
		Catch
        End Try



		Try
            TestForBelowGround()
            CheckUpdateListenerLoc()
        Catch
        End Try
		Try
			If CurrentEnvirView = CurrentView.eSystemView Then
				SolarSystem.lSysMapView2CameraX = mlCameraX \ 10000
				SolarSystem.lSysMapView2CameraZ = mlCameraZ \ 10000
			ElseIf CurrentEnvirView = CurrentView.eSystemMapView2 Then
				SolarSystem.lSysMapView2CameraX = (mlCameraX * 30) \ 10000
                SolarSystem.lSysMapView2CameraZ = (mlCameraZ * 30) \ 10000
            ElseIf CurrentEnvirView = CurrentView.eSystemMapView1 Then

                Dim lBase As Int32 = mlCameraAtX
                Dim lTmp As Int32 = mlCameraAtX * 10000

                If lTmp > 5000000 Then mlCameraAtX = 500
                If lTmp < -5000000 Then mlCameraAtX = -500

                If lBase <> mlCameraAtX Then
                    Dim lDiff As Int32 = mlCameraAtX - lBase
                    mlCameraX += lDiff
                End If

                lBase = mlCameraAtZ
                lTmp = mlCameraAtZ * 10000
                If lTmp > 5000000 Then mlCameraAtZ = 500
                If lTmp < -5000000 Then mlCameraAtZ = -500

                If lBase <> mlCameraAtZ Then
                    Dim lDiff As Int32 = mlCameraAtZ - lBase
                    mlCameraZ += lDiff
                End If

            ElseIf CurrentEnvirView = CurrentView.ePlanetView Then
                If goCurrentEnvir Is Nothing = False Then
                    If goCurrentEnvir.oGeoObject Is Nothing = False Then
                        If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                            Dim lHlfExt As Int32 = (CType(goCurrentEnvir.oGeoObject, Planet).CellSpacing * TerrainClass.Width) \ 2
                            'Dim lAdjust As Int32 = (CType(goCurrentEnvir.oGeoObject, Planet).CellSpacing * (TerrainClass.Width - 1))

                            Dim lBase As Int32 = mlCameraX

                            If mlCameraX < -lHlfExt Then
                                'Dim lTemp As Int32 = mlCameraZ - mlCameraAtZ
                                'Dim lTemp As Int32 = Math.Abs(mlCameraX) - Math.Abs(lHlfExt)
                                'Dim lTemp2 As Int32 = mlCameraX
                                'mlCameraX = -lHlfExt + lTemp
                                'lTemp2 -= mlCameraX
                                'mlCameraAtX += lTemp
                                mlCameraX = -lHlfExt
                            ElseIf mlCameraX > lHlfExt Then
                                'Dim lTemp As Int32 = Math.Abs(mlCameraX) - Math.Abs(lHlfExt)
                                'Dim lTemp2 As Int32 = mlCameraX
                                'mlCameraX = lHlfExt - lTemp
                                'lTemp2 -= mlCameraX
                                'mlCameraAtX -= lTemp
                                mlCameraX = lHlfExt
                            End If
                            If lBase <> mlCameraX Then
                                Dim lDiff As Int32 = mlCameraX - lBase
                                mlCameraAtX += lDiff
                            End If
                            'If mlCameraAtX < -lHlfExt Then
                            '    mlCameraAtX = -lHlfExt
                            'ElseIf mlCameraAtX > lHlfExt Then
                            '    mlCameraAtX = lHlfExt
                            'End If

                            'Ok, check for Z extremities
                            Dim lMaxZ As Int32 = (CType(goCurrentEnvir.oGeoObject, Planet).CellSpacing * TerrainClass.Width) \ 2
                            lBase = mlCameraZ
                            If mlCameraZ < -lMaxZ Then
                                'Dim lTemp As Int32 = mlCameraZ - mlCameraAtZ
                                'mlCameraAtZ = -lMaxZ
                                'mlCameraZ = mlCameraAtZ + lTemp
                                mlCameraZ = -lMaxZ
                            ElseIf mlCameraZ > lMaxZ Then
                                'Dim lTemp As Int32 = mlCameraZ - mlCameraAtZ
                                'mlCameraAtZ = lMaxZ
                                'mlCameraZ = mlCameraAtZ + lTemp
                                mlCameraZ = lMaxZ
                            End If
                            If lBase <> mlCameraZ Then
                                Dim lDiff As Int32 = mlCameraZ - lBase
                                mlCameraAtZ += lDiff
                            End If
                        End If
                    End If
                End If
			End If
		Catch
		End Try



		vecTemp = Nothing
		vecCameraAt = Nothing
		vecCameraLoc = Nothing
        vecDot = Nothing

        If lAutoRotateDir <> 0 Then
            RotateCamera(lAutoRotateDir, 0, False)
        End If
        If lAutoZoom <> 0 Then
            ModifyZoom(lAutoZoom)
        End If
    End Sub

    Public Sub ModifyZoom(ByVal lAmt As Int32)
        If glCurrentEnvirView > CurrentView.eFullScreenInterface Then
            If glCurrentEnvirView = CurrentView.eHullResearch Then
                Dim frmHull As frmHullBuilder = CType(goUILib.GetWindow("frmHullBuilder"), frmHullBuilder)
                If frmHull Is Nothing = False Then
                    frmHull.AdjustZoom(lAmt)
                End If
                frmHull = Nothing
            End If
            Return
        End If
        Dim oVec3 As Vector3 = New Vector3(mlCameraX - mlCameraAtX, mlCameraY - mlCameraAtY, mlCameraZ - mlCameraAtZ)
        oVec3.Normalize()

        'TODO: Bug with implemtation.  Zoom 1 away from the planet surface.  Zoom one more.  You are prevented from the Y axis but the camera does slighty rotate clockwise.  Not sure why this is, again probably a vector issue.
        If glCurrentEnvirView = CurrentView.ePlanetView Then
            Dim lTemp As Int32 = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraX, goCamera.mlCameraZ, True))
            If mlCameraY > 0 AndAlso mlCameraY + CInt(lAmt * oVec3.Y) < lTemp Then Return
        ElseIf glCurrentEnvirView = CurrentView.eGalaxyMapView Then
            If mlCameraY > 0 AndAlso mlCameraY + CInt(lAmt * oVec3.Y) < 100 Then Return
            If mlCameraY < 0 AndAlso mlCameraY + CInt(lAmt * oVec3.Y) > 0 Then Return
        Else
            If mlCameraY > 0 AndAlso mlCameraY + CInt(lAmt * oVec3.Y) < 100 Then Return
            If mlCameraY < 0 AndAlso mlCameraY + CInt(lAmt * oVec3.Y) > 0 Then Return
        End If


        mlCameraX += CInt(lAmt * oVec3.X)
        mlCameraY += CInt(lAmt * oVec3.Y)
        mlCameraZ += CInt(lAmt * oVec3.Z)

        If glCurrentEnvirView = CurrentView.ePlanetView Then
            Dim lMaxScroll As Int32 = CInt(muSettings.EntityClipPlane * 0.5F)
            If mlCameraY - mlCameraAtY > lMaxScroll Then
                Dim lTemp As Int32 = (mlCameraY - mlCameraAtY) - lMaxScroll
                mlCameraY -= lTemp
                If muSettings.ZoomChangesView = True Then
                    If lAmt > 0 Then frmMain.GoBackAView()
                    frmMain.lBackedOutToMapViewCycle = glCurrentCycle
                End If
            End If
        ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
            Dim lDiff As Int32 = mlCameraY - mlCameraAtY
            If lDiff > 40000 AndAlso muSettings.ZoomChangesView = True Then
                frmMain.GoBackAView()
            End If
        ElseIf glCurrentEnvirView = CurrentView.eSystemMapView1 Then
            Dim lDiff As Int32 = mlCameraY - mlCameraAtY
            If lDiff > 4500 AndAlso muSettings.ZoomChangesView = True Then
                frmMain.GoBackAView()
            End If
        End If

        oVec3 = Nothing
    End Sub

    Public lAutoRotateDir As Int32 = 0
    Public lAutoZoom As Int32 = 0

    Public Function RotateCamera(ByVal deltaX As Int32, ByVal deltaY As Int32, ByVal bRightDrag As Boolean) As Boolean
        Dim fXYaw As Single
		Dim fZYaw As Single

        TrackingIndex = -1

		If bRightDrag = False Then
			mlRotateCameraThresholdX += Math.Abs(deltaX)
			mlRotateCameraThresholdY += Math.Abs(deltaY)

			bRightDrag = mlRotateCameraThresholdX > ml_ROTATE_CAMERA_THRESHOLD OrElse mlRotateCameraThresholdY > ml_ROTATE_CAMERA_THRESHOLD
		End If
		'If we are still false, then return false
		If bRightDrag = False Then Return False

		If NewTutorialManager.TutorialOn = True Then
			If goUILib Is Nothing = False AndAlso goUILib.CommandAllowed(True, "RotateCamera") = False Then Return False
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eRotateCamera, -1, -1, -1, "")
		End If

		'X is easy
		RotatePoint(mlCameraAtX, mlCameraAtZ, mlCameraX, mlCameraZ, deltaX)

		'Ok, Y is the real pain...
		Dim vecAngle As Vector3 = New Vector3(mlCameraAtX - mlCameraX, mlCameraAtY - mlCameraY, mlCameraAtZ - mlCameraZ)
		vecAngle.Normalize()

		fXYaw = vecAngle.X * deltaY
		fZYaw = vecAngle.Z * deltaY

		RotatePoint(mlCameraAtX, mlCameraAtY, mlCameraX, mlCameraY, fXYaw)
		RotatePoint(mlCameraAtZ, mlCameraAtY, mlCameraZ, mlCameraY, fZYaw)

		TestForBelowGround()

		If glCurrentEnvirView = CurrentView.ePlanetView Then
			muSettings.lPlanetViewCameraX = mlCameraX - mlCameraAtX
			muSettings.lPlanetViewCameraZ = mlCameraZ - mlCameraAtZ

			If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
				Dim lTemp As Int32 = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraX, goCamera.mlCameraZ, True))
				muSettings.lPlanetViewCameraY = Math.Abs(goCamera.mlCameraY - lTemp)
            End If
        ElseIf glCurrentEnvirView = CurrentView.eGalaxyMapView Then
            muSettings.lGalaxyViewCameraX = mlCameraX - mlCameraAtX
            muSettings.lGalaxyViewCameraY = mlCameraY - mlCameraAtY
            muSettings.lGalaxyViewCameraZ = mlCameraZ - mlCameraAtZ
        End If

		If mlCameraAtX = mlCameraX AndAlso mlCameraAtZ = mlCameraZ Then
			mlCameraX += 2
			mlCameraZ += 2
		End If

		Return True
    End Function

    Private Sub TestForBelowGround()
        If glCurrentEnvirView <> CurrentView.ePlanetView AndAlso glCurrentEnvirView <> CurrentView.eSystemView Then Return

        If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
            Dim lHt As Int32 = 0
            Dim lHtAt As Int32 = 0
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet AndAlso glCurrentEnvirView = CurrentView.ePlanetView Then
                    'Ok... everything is good
                    lHt = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(mlCameraX, mlCameraZ, True)) + 500

                    'Next do the ht at camera point
                    lHtAt = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(mlCameraAtX, mlCameraAtZ, True))
                    If lHtAt < CType(goCurrentEnvir.oGeoObject, Planet).WaterHeight Then lHtAt = CType(goCurrentEnvir.oGeoObject, Planet).WaterHeight
                End If
                'ElseIf goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                '    If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.eSolarSystem AndAlso glCurrentEnvirView = CurrentView.eSystemView Then
                '        With CType(goCurrentEnvir.oGeoObject, SolarSystem)
                '            lHt = .GetBaseY(mlCameraX, mlCameraZ)
                '            lHtAt = .GetBaseY(mlCameraAtX, mlCameraAtZ)

                '            'If lHtAt <> mlCameraAtY Then
                '            '    Dim lAtDiff As Int32 = lHtAt - mlCameraAtY 'mlCameraAtY - lHtAt
                '            '    mlCameraY += lAtDiff
                '            'End If
                '        End With
                '    End If
            End If

            'Now... what's our current y?
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso mlCameraY < lHt Then mlCameraY = lHt
            Dim lDiff As Int32 = lHtAt - mlCameraAtY
            mlCameraAtY = lHtAt
            mlCameraY += lDiff

            If Math.Abs(mlCameraY - mlCameraAtY) > 25000 Then
                If mlCameraY < mlCameraAtY Then
                    mlCameraY = mlCameraAtY - 25000
                Else
                    mlCameraY = mlCameraAtY + 25000
                    If glCurrentEnvirView = CurrentView.eSystemView AndAlso muSettings.ZoomChangesView = True Then
                        frmMain.GoBackAView()
                    End If
                End If
            End If
        End If
    End Sub

    Public Function GetCameraAngleDegrees() As Single
        Return LineAngleDegrees(mlCameraX, mlCameraZ, mlCameraAtX, mlCameraAtZ)
    End Function

    Private Sub InitializeGrid()
        'set up our grid
        ReDim muGrid(3)
        Dim fMin As Single = -1000
        Dim fExt As Single = 1000
        Dim fTexMax As Single = (fExt - fMin) / 256
        muGrid(0) = New CustomVertex.PositionTextured(fMin, -100, fMin, 0, 0)
        muGrid(1) = New CustomVertex.PositionTextured(fMin, -100, fExt, 0, fTexMax)
        muGrid(2) = New CustomVertex.PositionTextured(fExt, -100, fMin, fTexMax, 0)
        muGrid(3) = New CustomVertex.PositionTextured(fExt, -100, fExt, fTexMax, fTexMax)
        matGrid.Diffuse = System.Drawing.Color.White
        matGrid.Ambient = matGrid.Diffuse
        matGrid.Emissive = System.Drawing.Color.Black
        mbGridReady = True
    End Sub

    Public Sub DrawGrid(ByRef oDevice As Device)
        If mbGridReady = False Then InitializeGrid()

        If moGridTex Is Nothing OrElse moGridTex.Disposed = True Then moGridTex = goResMgr.GetTexture("Grid.dds", GFXResourceManager.eGetTextureType.NoSpecifics)

        'set up our grid
        With muGrid(0)
            .X = mlCameraAtX - 5000
            .Y = -100
            .Z = mlCameraAtZ - 5000
        End With
        With muGrid(1)
            .X = mlCameraAtX - 5000
            .Y = -100
            .Z = mlCameraAtZ + 5000
        End With
        With muGrid(2)
            .X = mlCameraAtX + 5000
            .Y = -100
            .Z = mlCameraAtZ - 5000
        End With
        With muGrid(3)
            .X = mlCameraAtX + 5000
            .Y = -100
            .Z = mlCameraAtZ + 5000
        End With

        oDevice.SetRenderState(RenderStates.Lighting, 0)
        oDevice.RenderState.AlphaBlendEnable = True

        With oDevice
            .RenderState.SourceBlend = Blend.SourceAlpha
            .RenderState.DestinationBlend = Blend.One
            .RenderState.AlphaBlendEnable = True
            .RenderState.ZBufferEnable = False
        End With
        Try
            oDevice.Transform.World = Matrix.Identity()
            oDevice.VertexFormat = CustomVertex.PositionTextured.Format
            oDevice.Material = matGrid
            oDevice.SetTexture(0, moGridTex)
            oDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, muGrid)
        Catch
        End Try
        With oDevice
            .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
            .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
            .RenderState.AlphaBlendEnable = True
            .RenderState.ZBufferEnable = True
        End With
    End Sub

    Public Sub ScreenShake(ByVal fIntensity As Single)
        'If msw_SS Is Nothing Then msw_SS = Stopwatch.StartNew()
        'mbSS = True
        'mfSS_Intensity += fIntensity
    End Sub

    Public Sub SetupSystemMinimapMatrices(ByRef oDevice As Device)
        oDevice.Transform.View = Matrix.LookAtLH(New Vector3(mlCameraX, 10000, mlCameraZ), _
          New Vector3(mlCameraX, 0, mlCameraZ), New Vector3(0.0#, 0.0#, 1.0#))    'up is always this
        oDevice.Transform.Projection = Matrix.OrthoLH(10000, 10000, 0.1, 20000)
	End Sub

	Public Sub ZoomToPosition(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal iLocA As Int16)
		TrackingIndex = -1

		mlZoomToLocX = lLocX
		mlZoomToLocZ = lLocZ
		miZoomToAngle = iLocA
		mbZoomToLoc = True
	End Sub

#Region " Frustrum Culling "
    Private mp_Frustrum() As Plane

    Public Sub CalculateFrustrum(ByRef oDevice As Device)

        Dim lClipPlane As Int32 = muSettings.EntityClipPlane

        If glCurrentEnvirView = CurrentView.eSystemView Then
            lClipPlane *= 10
        End If

        Dim matProj As Matrix = Matrix.PerspectiveFovLH(0.7853982F, CSng(oDevice.PresentationParameters.BackBufferWidth / oDevice.PresentationParameters.BackBufferHeight), 1, lClipPlane)
        Dim matComb As Matrix = Matrix.Multiply(oDevice.Transform.View, matProj)

        If mp_Frustrum Is Nothing Then ReDim mp_Frustrum(5)

        'Left clipping plane
        With mp_Frustrum(0)
            .A = matComb.M14 + matComb.M11
            .B = matComb.M24 + matComb.M21
            .C = matComb.M34 + matComb.M31
            .D = matComb.M44 + matComb.M41
        End With

        'Right clipping plane
        With mp_Frustrum(1)
            .A = matComb.M14 - matComb.M11
            .B = matComb.M24 - matComb.M21
            .C = matComb.M34 - matComb.M31
            .D = matComb.M44 - matComb.M41
        End With

        'Top Clipping Plane
        With mp_Frustrum(2)
            .A = matComb.M14 - matComb.M12
            .B = matComb.M24 - matComb.M22
            .C = matComb.M34 - matComb.M32
            .D = matComb.M44 - matComb.M42
        End With

        'Bottom clipping plane
        With mp_Frustrum(3)
            .A = matComb.M14 + matComb.M12
            .B = matComb.M24 + matComb.M22
            .C = matComb.M34 + matComb.M32
            .D = matComb.M44 + matComb.M42
        End With

        'Near clipping plane
        With mp_Frustrum(4)
            .A = matComb.M13
            .B = matComb.M23
            .C = matComb.M33
            .D = matComb.M43
        End With

        'Far clipping Plane
        With mp_Frustrum(5)
            .A = matComb.M14 - matComb.M13
            .B = matComb.M24 - matComb.M23
            .C = matComb.M34 - matComb.M33
            .D = matComb.M44 - matComb.M43
        End With

        For X As Int32 = 0 To 5
            mp_Frustrum(X).Normalize()
        Next X
    End Sub

    Public Function CullObject(ByRef bBox As CullBox) As Boolean
        'for each plane in the frustrum
        Dim c1 As Vector3
        Dim c2 As Vector3

        For X As Int32 = 0 To 5
            With mp_Frustrum(X)
                'find furthest and nears point to the plane
                If .A > 0.0F Then
                    c1.X = bBox.vecMax.X : c2.X = bBox.vecMin.X
                Else
                    c1.X = bBox.vecMin.X : c2.X = bBox.vecMax.X
                End If
                If .B > 0.0F Then
                    c1.Y = bBox.vecMax.Y : c2.Y = bBox.vecMin.Y
                Else
                    c1.Y = bBox.vecMin.Y : c2.Y = bBox.vecMax.Y
                End If
                If .C > 0.0F Then
                    c1.Z = bBox.vecMax.Z : c2.Z = bBox.vecMin.Z
                Else
                    c1.Z = bBox.vecMin.Z : c2.Z = bBox.vecMax.Z
                End If

                Dim fD1 As Single = .A * c1.X + .B * c1.Y + .C * c1.Z + .D
                If fD1 < 0.0F Then
                    Dim fD2 As Single = .A * c2.X + .B * c2.Y + .C * c2.Z + .D
                    If fD2 < 0.0F Then Return True
                End If
            End With
        Next X

        Return False
    End Function

    Public Function CullObject_NoMaxDistance(ByRef bBox As CullBox) As Boolean
        'for each plane in the frustrum
        Dim c1 As Vector3
        Dim c2 As Vector3

        For X As Int32 = 0 To 4
            With mp_Frustrum(X)
                'find furthest and nears point to the plane
                If .A > 0.0F Then
                    c1.X = bBox.vecMax.X : c2.X = bBox.vecMin.X
                Else
                    c1.X = bBox.vecMin.X : c2.X = bBox.vecMax.X
                End If
                If .B > 0.0F Then
                    c1.Y = bBox.vecMax.Y : c2.Y = bBox.vecMin.Y
                Else
                    c1.Y = bBox.vecMin.Y : c2.Y = bBox.vecMax.Y
                End If
                If .C > 0.0F Then
                    c1.Z = bBox.vecMax.Z : c2.Z = bBox.vecMin.Z
                Else
                    c1.Z = bBox.vecMin.Z : c2.Z = bBox.vecMax.Z
                End If

                Dim fD1 As Single = .A * c1.X + .B * c1.Y + .C * c1.Z + .D
                If fD1 < 0.0F Then
                    Dim fD2 As Single = .A * c2.X + .B * c2.Y + .C * c2.Z + .D
                    If fD2 < 0.0F Then Return True
                End If
            End With
        Next X

        Return False
    End Function
#End Region

End Class

Public Structure CullBox
    Public vecMin As Vector3
    Public vecMax As Vector3

    Public Shared Function GetCullBox(ByVal fLocX As Single, ByVal fLocY As Single, ByVal fLocZ As Single, ByVal fMinX As Single, ByVal fMinY As Single, ByVal fMinZ As Single, ByVal fMaxX As Single, ByVal fMaxY As Single, ByVal fMaxZ As Single) As CullBox
        Dim oBox As CullBox
        oBox.vecMax.X = fLocX + fMaxX
        oBox.vecMax.Y = fLocY + fMaxY
        oBox.vecMax.Z = fLocZ + fMaxZ
        oBox.vecMin.X = fLocX + fMinX
        oBox.vecMin.Y = fLocY + fMinY
        oBox.vecMin.Z = fLocZ + fMinZ
        Return oBox
    End Function
End Structure