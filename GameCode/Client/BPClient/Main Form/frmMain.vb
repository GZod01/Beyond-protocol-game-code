Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Structure Ray
    Public Origin As Microsoft.DirectX.Vector3
    Public Direction As Microsoft.DirectX.Vector3
End Structure

Public Class frmMain
    Inherits System.Windows.Forms.Form

    Public Enum eSelectNextType As Byte
        eAnyType = 0
        eEngineer = 1
        eFacility = 2
        eUnit = 3
        eUnpoweredResidence = 4
        eUnpoweredFacility = 5
        eIdleEngineer = 6
        ePreviousBitShift = 128
    End Enum

    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32
    Private WithEvents moEngine As GFXEngine

    Private mbRightDown As Boolean
    Private mbRightDrag As Boolean
    Public Shared mlMouseX As Int32
    Public Shared mlMouseY As Int32
    Private mbMouseWheelOverride As Boolean

    Private mbKeyScrollLeft As Boolean = False
    Private mbKeyScrollRight As Boolean = False
    Private mbKeyScrollUp As Boolean = False
    Private mbKeyScrollDown As Boolean = False

    Private WithEvents moMsgSys As MsgSystem

    Private mbFKeyDown As Boolean
    Public Shared mbShiftKeyDown As Boolean
    Public Shared mbCtrlKeyDown As Boolean
    Public Shared mbAltKeyDown As Boolean
    Private mbSetRallyPoint As Boolean = False
    Private mbLoginFailed As Boolean
    Private mbEnvirReady As Boolean = False
	Public Shared mbChangingEnvirs As Boolean = False

    Public Shared mbIgnoreNextMouseUp As Boolean = False
    Private mbIgnoreNextKeyUp As Boolean = False
    Public Shared mbIgnoreMouseMove As Boolean = False

    Private mbStartedUp As Boolean = False

    Public Shared mbfrmConfirmHandled As Boolean = False

	Private mlLastKeepAliveSend As Int32 = 0

	Private mbFormClosing As Boolean = False

	Private moClickPlane As Plane = Plane.FromPointNormal(New Vector3(0, 0, 0), New Vector3(0, 1, 0))

	Private mlLastUIEvent As Int32 = 0

	'Private moMainThread As Threading.Thread
	'Private mbMainThread As Boolean = True


    Private mbInFormClosing As Boolean = False
    Public Shared bRestartWithUpdater As Boolean = False

    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        CleanupOldFiles()

        Process.GetCurrentProcess.Kill()
    End Sub
    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If mbInFormClosing = True Then Return
        mbInFormClosing = True

        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 255 AndAlso mbfrmConfirmHandled = False Then
            Dim sResult As String = frmConfirmQuit.CheckLogoffCapability()
            If sResult <> "" Then
                Dim ofrm As frmConfirmQuit = CType(goUILib.GetWindow("frmConfirmQuit"), frmConfirmQuit)
                If ofrm Is Nothing = True Then
                    e.Cancel = True
                    ofrm = New frmConfirmQuit(goUILib, sResult)
                    ofrm = Nothing
                    mbInFormClosing = False
                    Return
                End If
            End If
        End If

        'On Error Resume Next
        GFXEngine.gbPaused = True
		mbFormClosing = True

        'disconnect from servers
        Dim lCnt As Int32 = 0
        Timer1.Enabled = False
        While mbInTimer1Tick = True
            Threading.Thread.Sleep(1)
            Timer1.Enabled = False
            lCnt += 1
            If lCnt > 100000 Then mbInTimer1Tick = False
        End While
        Timer1.Enabled = False
		'mbMainThread = False 

        'Me.Visible = False
		Application.DoEvents()
        Threading.Thread.Sleep(1)

        If moEngine Is Nothing = False AndAlso moEngine.RecreateEverything(Me, moMsgSys, False) = False Then moEngine.ForceDeviceLost()

        If moMsgSys Is Nothing = False Then
            moMsgSys.DisconnectAll()
            moMsgSys = Nothing
        End If

        If goControlGroups Is Nothing = False Then goControlGroups.SaveGroups()
        goControlGroups = Nothing
        muSettings.SaveSettings()

        goGalaxy = Nothing          'kill the galaxy
        goCurrentEnvir = Nothing    'kill the envir
        If goSound Is Nothing = False Then goSound.DisposeMe()
        goSound = Nothing           'kill the sound engine
        goResMgr = Nothing          'kill the res mgr
        moEngine = Nothing          'kill the engine
        goCamera = Nothing          'kill our camera



        'Try
        '    Application.Exit()
        'Catch
        'End Try
        'Try
        '    End
        'Catch
        'End Try
    End Sub

    Private Sub Login_ExitProgram()
        'this sub handles the exit program event
        Me.Close()
    End Sub

	Public Shared Function SelectAndGotoNextUnit(ByVal yType As eSelectNextType) As Int32
        Try
            If HasAliasedRights(AliasingRights.eViewUnitsAndFacilities) = False Then
                goUILib.AddNotification("You lack rights to view Units and Facilities.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                Return -1
            End If

            Dim bPrevious As Boolean = False
            If (yType And eSelectNextType.ePreviousBitShift) <> 0 Then
                bPrevious = True
                yType = yType Xor eSelectNextType.ePreviousBitShift
            End If

            Select Case yType
                Case eSelectNextType.eAnyType
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eSelectNextEntity, 0, 0)
                Case eSelectNextType.eEngineer
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eSelectNextEngineer, 0, 0)
                Case eSelectNextType.eFacility
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eSelectNextFacility, 0, 0)
                Case eSelectNextType.eIdleEngineer
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eSelectNextIdleEngineer, 0, 0)
                Case eSelectNextType.eUnit
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eSelectNextUnit, 0, 0)
                Case eSelectNextType.eUnpoweredFacility
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eSelectNextUnpoweredFacility, 0, 0)
                Case eSelectNextType.eUnpoweredResidence
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eSelectNextUnpoweredResidence, 0, 0)
            End Select

            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing Then Return -1

            Dim X As Int32
            Dim lIdx As Int32 = -1
            Dim lFirstMine As Int32 = -1

            Dim lResult As Int32 = -1

            Dim lEnd As Int32
            Dim lBegin As Int32
            Dim lStep As Int32


            If goCamera Is Nothing = False Then goCamera.TrackingIndex = -1
            If oEnvir.ObjTypeID = ObjectType.ePlanet Then
                If glCurrentEnvirView <> CurrentView.ePlanetView Then glCurrentEnvirView = CurrentView.ePlanetView
            ElseIf oEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                If glCurrentEnvirView <> CurrentView.eSystemView Then glCurrentEnvirView = CurrentView.eSystemView
            End If

            Dim lCurUB As Int32 = -1
            If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityIdx.GetUpperBound(0), oEnvir.lEntityUB)

            If bPrevious = True Then
                lEnd = 0
                lBegin = lCurUB
                lStep = -1
            Else
                lEnd = lCurUB
                lBegin = 0
                lStep = 1
            End If

            For X = lBegin To lEnd Step lStep '0 To lCurUB
                If oEnvir.lEntityIdx(X) <> -1 Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity.bSelected = True Then
                        lIdx = X
                        Exit For
                    ElseIf oEntity.OwnerID = glPlayerID AndAlso lFirstMine = -1 Then
                        Select Case yType
                            Case eSelectNextType.eEngineer
                                If oEntity.yProductionType = ProductionType.eFacility Then lFirstMine = X
                            Case eSelectNextType.eFacility
                                If oEntity.ObjTypeID = ObjectType.eFacility Then lFirstMine = X
                            Case eSelectNextType.eUnit
                                If oEntity.ObjTypeID = ObjectType.eUnit Then lFirstMine = X
                            Case eSelectNextType.eUnpoweredResidence
                                If oEntity.ObjTypeID = ObjectType.eFacility AndAlso _
                                   (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 AndAlso _
                                   oEntity.yProductionType = ProductionType.eColonists Then
                                    lFirstMine = X
                                End If
                            Case eSelectNextType.eUnpoweredFacility
                                If oEntity.ObjTypeID = ObjectType.eFacility AndAlso _
                                   (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 Then
                                    lFirstMine = X
                                End If
                            Case eSelectNextType.eIdleEngineer
                                If oEntity.yProductionType = ProductionType.eFacility Then
                                    If oEntity.bProducing = False AndAlso (oEntity.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                                        If NewTutorialManager.TutorialOn = True AndAlso NewTutorialManager.lRouteEngineerID = oEntity.ObjectID Then Continue For
                                        If glCurrentCycle - oEntity.LastUpdateCycle > 300 Then lFirstMine = X
                                    End If
                                End If
                            Case Else
                                lFirstMine = X
                        End Select
                    End If
                End If
            Next X

            'ensure all entitys are unselected
            oEnvir.DeselectAll()

            If lIdx = -1 Then
                If lFirstMine < 0 Then Return -1
                If lFirstMine > lCurUB Then Return -1
                Dim oEntity As BaseEntity = oEnvir.oEntity(lFirstMine)
                If oEntity Is Nothing Then Return -1

                oEntity.bSelected = True
                goCamera.mlCameraX = CInt(oEntity.LocX)
                goCamera.mlCameraAtX = goCamera.mlCameraX
                goCamera.mlCameraAtZ = CInt(oEntity.LocZ)
                goCamera.mlCameraZ = goCamera.mlCameraAtZ - 500

                If oEnvir.ObjTypeID = ObjectType.ePlanet Then
                    'goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ)) + 3000
                    goCamera.mlCameraY = CInt(CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

                    goCamera.mlCameraX += muSettings.lPlanetViewCameraX
                    goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
                End If
                lResult = lFirstMine
            Else
                lFirstMine = -1

                If bPrevious = True Then
                    lBegin = lIdx - 1
                    lEnd = 0
                    lStep = -1
                Else
                    lBegin = lIdx + 1
                    lEnd = lCurUB
                    lStep = 1
                End If

                For X = lBegin To lEnd Step lStep 'lIdx + 1 To lCurUB
                    If oEnvir.lEntityIdx(X) <> -1 Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing OrElse oEntity.OwnerID <> glPlayerID Then Continue For

                        Select Case yType
                            Case eSelectNextType.eEngineer
                                If oEntity.yProductionType = ProductionType.eFacility Then
                                    lFirstMine = X
                                    Exit For
                                End If
                            Case eSelectNextType.eFacility
                                If oEntity.ObjTypeID = ObjectType.eFacility Then
                                    lFirstMine = X
                                    Exit For
                                End If
                            Case eSelectNextType.eUnit
                                If oEntity.ObjTypeID = ObjectType.eUnit Then
                                    lFirstMine = X
                                    Exit For
                                End If
                            Case eSelectNextType.eUnpoweredResidence
                                If oEntity.ObjTypeID = ObjectType.eFacility AndAlso _
                                   (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 AndAlso _
                                   oEntity.yProductionType = ProductionType.eColonists Then
                                    lFirstMine = X
                                    Exit For
                                End If
                            Case eSelectNextType.eUnpoweredFacility
                                If oEntity.ObjTypeID = ObjectType.eFacility AndAlso _
                                   (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 Then
                                    lFirstMine = X
                                    Exit For
                                End If
                            Case eSelectNextType.eIdleEngineer
                                If oEntity.yProductionType = ProductionType.eFacility Then
                                    If oEntity.bProducing = False AndAlso (oEntity.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                                        If NewTutorialManager.TutorialOn = True AndAlso NewTutorialManager.lRouteEngineerID = oEntity.ObjectID Then Continue For
                                        If glCurrentCycle - oEntity.LastUpdateCycle > 300 Then lFirstMine = X
                                        Exit For
                                    End If
                                End If
                            Case Else
                                lFirstMine = X
                                Exit For
                        End Select
                    End If
                Next X

                If lFirstMine = -1 Then

                    If bPrevious = True Then
                        lBegin = lCurUB
                        lEnd = lIdx
                        lStep = -1
                    Else
                        lBegin = 0
                        lEnd = lIdx
                        lStep = 1
                    End If

                    For X = lBegin To lEnd Step lStep '0 To lIdx
                        If oEnvir.lEntityIdx(X) <> -1 Then
                            Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                            If oEntity Is Nothing OrElse oEntity.OwnerID <> glPlayerID Then Continue For

                            Select Case yType
                                Case eSelectNextType.eEngineer
                                    If oEntity.yProductionType = ProductionType.eFacility Then
                                        lFirstMine = X
                                        Exit For
                                    End If
                                Case eSelectNextType.eFacility
                                    If oEntity.ObjTypeID = ObjectType.eFacility Then
                                        lFirstMine = X
                                        Exit For
                                    End If
                                Case eSelectNextType.eUnit
                                    If oEntity.ObjTypeID = ObjectType.eUnit Then
                                        lFirstMine = X
                                        Exit For
                                    End If
                                Case eSelectNextType.eUnpoweredResidence
                                    If oEntity.ObjTypeID = ObjectType.eFacility AndAlso _
                                       (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 AndAlso _
                                       oEntity.yProductionType = ProductionType.eColonists Then
                                        lFirstMine = X
                                        Exit For
                                    End If
                                Case eSelectNextType.eUnpoweredFacility
                                    If oEntity.ObjTypeID = ObjectType.eFacility AndAlso _
                                      (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 Then
                                        lFirstMine = X
                                        Exit For
                                    End If
                                Case eSelectNextType.eIdleEngineer
                                    If oEntity.yProductionType = ProductionType.eFacility Then
                                        If oEntity.bProducing = False AndAlso (oEntity.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                                            If NewTutorialManager.TutorialOn = True AndAlso NewTutorialManager.lRouteEngineerID = oEntity.ObjectID Then Continue For
                                            If glCurrentCycle - oEntity.LastUpdateCycle > 300 Then lFirstMine = X
                                            Exit For
                                        End If
                                    End If
                                Case Else
                                    lFirstMine = X
                                    Exit For
                            End Select
                            'Exit For
                        End If
                    Next X
                End If

                If lFirstMine = -1 Then
                    Return -1
                Else
                    If lFirstMine < 0 OrElse lFirstMine > lCurUB Then Return -1
                    Dim oEntity As BaseEntity = oEnvir.oEntity(lFirstMine)
                    If oEntity Is Nothing Then Return -1
                    oEntity.bSelected = True
                    goCamera.mlCameraX = CInt(oEntity.LocX)
                    goCamera.mlCameraAtX = goCamera.mlCameraX
                    goCamera.mlCameraAtZ = CInt(oEntity.LocZ)
                    goCamera.mlCameraZ = goCamera.mlCameraAtZ - 500

                    If oEnvir.ObjTypeID = ObjectType.ePlanet Then
                        'goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ)) + 3000
                        goCamera.mlCameraY = CInt(CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

                        goCamera.mlCameraX += muSettings.lPlanetViewCameraX
                        goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
                    End If
                    lResult = lFirstMine
                End If
            End If

            Return lResult

        Catch
        End Try
	End Function

    Private Sub SelectAllSimilarUnits(ByVal bOnScreenOnly As Boolean)
        Try
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eSelectSimilar, 0, 0)
            If goCurrentEnvir Is Nothing = False Then
                Dim iModelIDs() As Int16 = Nothing
                Dim lUB As Int32 = -1
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                        Dim lIdx As Int32 = -1
                        Dim iModelID As Int16 = goCurrentEnvir.oEntity(X).oUnitDef.ModelID
                        For Y As Int32 = 0 To lUB
                            If iModelIDs(Y) = iModelID Then
                                lIdx = Y
                                Exit For
                            End If
                        Next Y
                        If lIdx = -1 Then
                            lUB += 1
                            ReDim Preserve iModelIDs(lUB)
                            iModelIDs(lUB) = iModelID
                        End If
                        goCurrentEnvir.oEntity(X).bSelected = False
                    End If
                Next X
                If lUB = -1 Then Return

                goUILib.ClearSelection()

                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso (goCurrentEnvir.oEntity(X).bCulled = False OrElse bOnScreenOnly = False) Then
                        Dim iModelID As Int16 = goCurrentEnvir.oEntity(X).oUnitDef.ModelID
                        For Y As Int32 = 0 To lUB
                            If iModelIDs(Y) = iModelID Then
                                goCurrentEnvir.oEntity(X).bSelected = True
                                Exit For
                            End If
                        Next Y
                    End If
                Next X
            End If
        Catch
        End Try
    End Sub

    Private Sub SelectAllSimilarNamedUnits(ByVal bOnScreenOnly As Boolean)
        Try
            If goCurrentEnvir Is Nothing = False Then
                Dim iNames() As String = Nothing
                Dim lUB As Int32 = -1
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                        Dim lIdx As Int32 = -1
                        Dim sName As String = GetCacheObjectValue(goCurrentEnvir.oEntity(X).ObjectID, goCurrentEnvir.oEntity(X).ObjTypeID)
                        If sName <> "Unknown" Then
                            For Y As Int32 = 0 To lUB
                                If iNames(Y) = sName Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Next Y
                            If lIdx = -1 Then
                                lUB += 1
                                ReDim Preserve iNames(lUB)
                                iNames(lUB) = sName
                            End If
                        End If
                        goCurrentEnvir.oEntity(X).bSelected = False
                    End If
                Next X
                If lUB = -1 Then Return

                goUILib.ClearSelection()

                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso (goCurrentEnvir.oEntity(X).bCulled = False OrElse bOnScreenOnly = False) Then
                        Dim sName As String = GetCacheObjectValue(goCurrentEnvir.oEntity(X).ObjectID, goCurrentEnvir.oEntity(X).ObjTypeID)
                        For Y As Int32 = 0 To lUB
                            If iNames(Y) = sName Then
                                goCurrentEnvir.oEntity(X).bSelected = True
                                Exit For
                            End If
                        Next Y
                    End If
                Next X
            End If
        Catch
        End Try
    End Sub

	Private Function GetEnvirClickedPoint(ByVal lMX As Int32, ByVal lMZ As Int32) As Vector3
		Dim bFindIntOnTerrain As Boolean = False
		Dim vecInt As Vector3
		'Get our height at that location
		If glCurrentEnvirView = CurrentView.ePlanetView Then
			If goCurrentEnvir Is Nothing = False Then
				If goCurrentEnvir.oGeoObject Is Nothing = False Then
					'ok, we must be displaying the planet...
					bFindIntOnTerrain = True
				End If
			End If
		End If

		Dim pickRay As Ray = CalcPickingRay(lMX, lMZ)

		If bFindIntOnTerrain = False Then
			If glCurrentEnvirView = CurrentView.ePlanetMapView Then
				Dim ptPos As Point = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).GetMapClickedCoords(lMX, lMZ)
				If ptPos.X <> Int32.MinValue And ptPos.Y <> Int32.MinValue Then
					vecInt = New Vector3(ptPos.X, 0, ptPos.Y)
				End If
			Else
				vecInt = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 20000)))

				If glCurrentEnvirView = CurrentView.eSystemMapView1 Then
					vecInt.X *= 10000 : vecInt.Y *= 10000 : vecInt.Z *= 10000
				ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
					vecInt.X *= 30 : vecInt.Y *= 30 : vecInt.Z *= 30
				End If
			End If
		ElseIf pickRay.Direction.Y < 0 Then
			'we only want to do this if the unit can move there, a unit can only move there if the click occurs on
			' the ground... a click can only occur on the ground if Y is negative
			Dim vecAddRay As Vector3 = pickRay.Direction

			vecAddRay.Scale(Math.Abs(1 / vecAddRay.Y) * 5)

			vecInt = pickRay.Origin

			'This is an attempt to reduce our test count... we get the maximum height and move our ray that much
			Dim lTmpVal As Int32 = CInt((CType(goCurrentEnvir.oGeoObject, Planet).GetMaxHeight() - vecInt.Y) / pickRay.Direction.Y)
			vecInt.Add(Vector3.Multiply(pickRay.Direction, lTmpVal))

			While vecInt.Y > 0
				vecInt.Add(vecAddRay)

				If CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(vecInt.X, vecInt.Z, True) > vecInt.Y Then
					'ok, that's it
					Exit While
				End If
			End While
		End If
		Return vecInt
	End Function

	'   Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)
	'       'I've left this here instead of hte camera class because I may want to capture these events for other things
	'       '  besides just zooming in and out... such as scrolling a listbox or something
	'       If e.Delta < 0 Then MouseWheelDown() Else MouseWheelUp()
    'End Sub
    Private mlPlanetMapZoomOuts As Int32 = 0
    Private mlPlanetMapZoomIns As Int32 = 0
	Private Sub frmMain_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
		mlLastUIEvent = glCurrentCycle
		If NewTutorialManager.TutorialOn = True Then
			If goUILib Is Nothing = False Then
				If goUILib.CommandAllowed(True, "ZoomCamera") = False Then Return
				NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eZoomCamera, -1, -1, -1, "")
			End If
		End If

        If mbFormClosing = True Then Return
        If Me.Visible = False Then Return

        If goUILib Is Nothing = False Then
            If goUILib.PostMouseWheelEvent(e) = True Then Return
        End If

        If glCurrentEnvirView = CurrentView.ePlanetMapView AndAlso muSettings.ZoomChangesView = True Then
            If e.Delta > 0 Then
                mlPlanetMapZoomIns += 1
                If mlPlanetMapZoomIns > 2 Then
                    'zoom in at the mouse cursor's location
                    HandleLeftMouseDown(mlMouseX, mlMouseY)
                    mlPlanetMapZoomIns = 0
                End If
                mlPlanetMapZoomOuts = 0
            ElseIf e.Delta < 0 Then
                If glCurrentCycle - lBackedOutToMapViewCycle > 10 Then
                    mlPlanetMapZoomOuts += 1
                    If mlPlanetMapZoomOuts > 2 Then
                        mlPlanetMapZoomOuts = 0
                        GoBackAView()
                    End If
                    mlPlanetMapZoomIns = 0
                End If
            End If
            Return
        End If

        If e.Delta < 0 Then MouseWheelDown() Else MouseWheelUp()
	End Sub

	Private Sub MouseWheelDown()
        If mbFormClosing = True Then Return

        If glCurrentEnvirView <> CurrentView.ePlanetMapView Then goCamera.ModifyZoom(goCamera.ZoomRate)

        Try
            If glCurrentEnvirView = CurrentView.ePlanetView AndAlso goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                Dim lTemp As Int32 = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraX, goCamera.mlCameraZ, True))
                muSettings.lPlanetViewCameraY = Math.Abs(goCamera.mlCameraY - lTemp)
            End If
        Catch
            'Do nothing
        End Try

        If glCurrentEnvirView = CurrentView.ePlanetView Then
            If goCamera.mlCameraAtY + 500 > goCamera.mlCameraY Then
                goCamera.mlCameraY = goCamera.mlCameraAtY + 500
                If goCamera.mlCameraAtX - goCamera.mlCameraX < 2 Then
                    goCamera.mlCameraX = goCamera.mlCameraAtX + 5
                End If
                If goCamera.mlCameraAtZ - goCamera.mlCameraZ < 2 Then
                    goCamera.mlCameraZ = goCamera.mlCameraAtZ + 5
                End If
            End If
        End If
    End Sub

	Private Sub MouseWheelUp()
		If mbFormClosing = True Then Return
		If glCurrentEnvirView <> CurrentView.ePlanetMapView Then goCamera.ModifyZoom(-goCamera.ZoomRate)

		Try
			If glCurrentEnvirView = CurrentView.ePlanetView AndAlso goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
				Dim lTemp As Int32 = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraX, goCamera.mlCameraZ, True))
				muSettings.lPlanetViewCameraY = Math.Abs(goCamera.mlCameraY - lTemp)
			End If
		Catch
			'Do nothing
        End Try

        If glCurrentEnvirView = CurrentView.ePlanetView Then
            If goCamera.mlCameraAtY + 500 > goCamera.mlCameraY Then
                goCamera.mlCameraY = goCamera.mlCameraAtY + 500
                If goCamera.mlCameraAtX - goCamera.mlCameraX < 2 Then
                    goCamera.mlCameraX = goCamera.mlCameraAtX + 5
                End If
                If goCamera.mlCameraAtZ - goCamera.mlCameraZ < 2 Then
                    goCamera.mlCameraZ = goCamera.mlCameraAtZ + 5
                End If
            End If
        ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
            Dim lDiff As Int32 = goCamera.mlCameraY - goCamera.mlCameraAtY
            If lDiff < 1400 AndAlso muSettings.ZoomChangesView = True Then
                HandleLeftMouseDown(mlMouseX, mlMouseY)
            End If
        End If
	End Sub

#Region "  MOUSE DOWN FUNCTIONS  "
    Private mlLastMouseDownEvent As Int32 = 0
    Private Sub frmMain_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        mlLastUIEvent = glCurrentCycle
        If mbFormClosing = True Then Return
        If Me.Visible = False Then Return
        If mbChangingEnvirs = True Then Return
        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return
        'Ok, first, check if we are in a place where mouse downs don't matter
        If glCurrentEnvirView = CurrentView.eStartupLogin Then
            Dim frmLogin As frmLoginDlg = CType(goUILib.GetWindow("frmLoginDlg"), frmLoginDlg)
            If frmLogin Is Nothing = False AndAlso frmLogin.Visible = False Then
                goUILib.RetainTooltip = False
                goUILib.SetToolTip(False)
                frmLogin.Visible = True
                mbIgnoreNextMouseUp = True
                Return
            End If
        End If
        If glCurrentEnvirView = CurrentView.eStartupDSELogo Then
            mbIgnoreNextMouseUp = True
            moEngine.mbBreakout = True 'If muSettings.bRanBefore = True Then
            Return
        End If

        'Ok, clear the tooltip
        goUILib.SetToolTip(False)
        'and then check our interfaces
        'first, check our interfaces...
        Dim bHaveGhost As Boolean = Not (goUILib.BuildGhost Is Nothing)
        Dim lIdx As Int32 = goUILib.BuildGhostID
        If goUILib.PostMessage(UILibMsgCode.eMouseDownMsgCode, e.X, e.Y, e.Button) = True Then
            If bHaveGhost = True And lIdx = goUILib.BuildGhostID Then
                goUILib.BuildGhost = Nothing
                goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
            End If
            Return
        End If

        'Ok, if we are here, our UI did not catch the event... now, check if we are left-clicking or right-clicking
        If e.Button = Windows.Forms.MouseButtons.Right Then
            HandleRightMouseDown(bHaveGhost, e.X, e.Y)
        ElseIf e.Button = Windows.Forms.MouseButtons.Left Then
            HandleLeftMouseDown(e.X, e.Y)
        End If

        mlLastMouseDownEvent = glCurrentCycle
    End Sub
    Private Sub HandleRightMouseDown(ByVal bHaveGhost As Boolean, ByVal lMouseX As Int32, ByVal lMouseY As Int32)
        mbRightDown = True
        goCamera.ShowSelectionBox = False

        'clear our gouilib.Buildghost
        If bHaveGhost = True Then
            goUILib.BuildGhost = Nothing
            goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
            mbIgnoreNextMouseUp = True
            mbRightDown = False
            mbRightDrag = False
        ElseIf goUILib.lUISelectState = UILib.eSelectState.eSetFleetDest Then
            Return
        ElseIf goUILib.lUISelectState = UILib.eSelectState.eSelectRouteLoc OrElse goUILib.lUISelectState = UILib.eSelectState.eSelectRouteTemplateLoc Then
            mbRightDown = False
            mbRightDrag = False
            mbIgnoreNextMouseUp = True
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False AndAlso goUILib.CommandAllowed(True, "SetRouteCancel") = False Then Return
            End If
            goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
            goUILib.AddNotification("Select route destination cancelled.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Dim oWin As UIWindow = Nothing
            If goUILib.lUISelectState = UILib.eSelectState.eSelectRouteLoc Then
                oWin = goUILib.GetWindow("frmRouteConfig")
            Else
                oWin = goUILib.GetWindow("frmRouteTemplate")
            End If
            If oWin Is Nothing = False Then oWin.Visible = True
        ElseIf goUILib.lUISelectState = UILib.eSelectState.eSelectRaceWaypoint Then
            mbRightDown = False
            mbRightDrag = False
            mbIgnoreNextMouseUp = True
            goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
            goUILib.yRenderUI = 255
            goUILib.AddNotification("Select route destination cancelled.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Dim oWin As UIWindow = goUILib.GetWindow("frmRaceConfig")
            If oWin Is Nothing = False Then oWin.Visible = True
        ElseIf goUILib.lUISelectState = UILib.eSelectState.eDismantleTarget OrElse goUILib.lUISelectState = UILib.eSelectState.eRepairTarget Then
            Dim pickRay As Ray = CalcPickingRay(lMouseX, lMouseY)
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing Then Return
            Dim lIdx As Int32 = oEnvir.PickObject(pickRay, muSettings.FarClippingPlane)
            If lIdx <> -1 Then
                Dim iMsgCode As Int16
                If goUILib.lUISelectState = UILib.eSelectState.eDismantleTarget Then
                    iMsgCode = GlobalMessageCode.eSetDismantleTarget
                Else : iMsgCode = GlobalMessageCode.eSetRepairTarget
                End If
                Dim lCurUB As Int32 = -1
                If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
                For X As Int32 = 0 To lCurUB
                    If oEnvir.lEntityIdx(X) <> -1 Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False AndAlso oEntity.bSelected = True Then
                            moMsgSys.SendSetMaintenanceTarget(iMsgCode, oEntity, oEnvir.oEntity(lIdx))
                            Exit For
                        End If
                    End If
                Next X

                mbIgnoreNextMouseUp = True
                mbRightDown = False
                mbRightDrag = False
            End If
            goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        ElseIf goUILib.lUISelectState = UILib.eSelectState.eSketchPad_SelectPoint Then
            goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        ElseIf mbCtrlKeyDown = True Then
            'ok, ctrl key is down, a facing move request... is a formation selected...

            Dim oWin As frmMultiDisplay = Nothing
            If goUILib Is Nothing = False Then
                oWin = CType(goUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
            End If
            If oWin Is Nothing = False AndAlso oWin.Visible = True Then
                Dim lFormationID As Int32 = oWin.GetFormationID
                If lFormationID > 0 Then
                    Dim vecInt As Vector3 = GetEnvirClickedPoint(lMouseX, lMouseY)
                    moEngine.lMouseDownX = CInt(vecInt.X)
                    moEngine.lMouseDownZ = CInt(vecInt.Z)
                    moEngine.lMouseDirX = moEngine.lMouseDownX
                    moEngine.lMouseDirZ = moEngine.lMouseDownZ + 1
                    moEngine.lFormationID = lFormationID
                    moEngine.bInCtrlMove = True
                End If
            End If
        End If
    End Sub
    Private Sub HandleLeftMouseDown(ByVal lMouseX As Int32, ByVal lMouseY As Int32)
        'If we're here, check for the buildghost
        If goUILib.BuildGhost Is Nothing = False Then
            HandleLeftMouseDownBuildGhost()
            Return
        ElseIf goUILib.lUISelectState <> UILib.eSelectState.eNoSelectState Then
            If HandleLeftMouseDownSelectState(lMouseX, lMouseY) = True Then Return
            'If mbIgnoreNextMouseUp = True Then Return
        End If

        'Check if we are changing environments....
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return
        Dim lCurUB As Int32 = -1
        If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))

        'First, we change environments if the player is in a SystemView or PlanetView and the
        '  view does not equal the current environment's type id
        If HandleLeftMouseDownChangeEnvir(oEnvir) = True Then Return

        'Set our pick ray, everything else will need it...
        Dim pickRay As Ray = CalcPickingRay(lMouseX, lMouseY)

        Select Case glCurrentEnvirView
            Case CurrentView.eGalaxyMapView
                HandleLeftMouseDownGalaxyView(pickRay, oEnvir)
            Case CurrentView.eSystemMapView1
                'Ok, what we do here is... we get our click loc
                Dim vecInt As Vector3 = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 200000)))

                With goCamera
                    Try
                        .mlCameraAtX = CInt((vecInt.X * 10000) / 30) : .mlCameraAtY = 0 : .mlCameraAtZ = CInt((vecInt.Z * 10000) / 30)
                    Catch
                        .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                    End Try
                    .mlCameraX = .mlCameraAtX : .mlCameraY = 10000 : .mlCameraZ = .mlCameraAtZ - 1000
                End With
                glCurrentEnvirView = CurrentView.eSystemMapView2
                If NewTutorialManager.TutorialOn = True Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eViewChanged, glCurrentEnvirView, -1, -1, "")

            Case CurrentView.eSystemMapView2
                Dim vecInt As Vector3 = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 200000)))

                With goCamera
                    Try
                        .mlCameraAtX = CInt(vecInt.X * 30) : .mlCameraAtY = 0 : .mlCameraAtZ = CInt(vecInt.Z * 30)
                    Catch
                        .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                    End Try
                    .mlCameraX = .mlCameraAtX : .mlCameraY = 3000 : .mlCameraZ = .mlCameraAtZ - 1000
                End With
                glCurrentEnvirView = CurrentView.eSystemView
                If NewTutorialManager.TutorialOn = True Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eViewChanged, glCurrentEnvirView, -1, -1, "")
            Case CurrentView.eSystemView
                If HandleLeftMouseDownSystemView(pickRay, oEnvir, lMouseX, lMouseY) = True Then Return
            Case CurrentView.ePlanetMapView
                'for planet view, we translate the coordinates straight to the map's coordinates
                Dim ptPos As Point
                Dim lTemp As Int32

                If goGalaxy Is Nothing = False AndAlso goGalaxy.CurrentSystemIdx <> -1 Then
                    lTemp = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx
                    If lTemp <> -1 Then
                        ptPos = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(lTemp).GetMapClickedCoords(lMouseX, lMouseY)
                        If ptPos.X <> Int32.MinValue And ptPos.Y <> Int32.MinValue Then
                            'Ok, we got a valid point on the map...
                            With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                                Dim bNeedBurst As Boolean = True
                                If oEnvir Is Nothing = False Then
                                    bNeedBurst = (oEnvir.ObjectID <> .moPlanets(.CurrentPlanetIdx).ObjectID OrElse _
                                     oEnvir.ObjTypeID <> .moPlanets(.CurrentPlanetIdx).ObjTypeID) AndAlso mbChangingEnvirs = False
                                End If

                                If bNeedBurst = True Then
                                    mbEnvirReady = False
                                    mbChangingEnvirs = True
                                    Dim lLoopCnt As Int32 = 0
                                    While gb_InHandleMovement
                                        Threading.Thread.Sleep(10)
                                        lLoopCnt += 1
                                        If lLoopCnt > 1000 Then Exit While
                                    End While
                                    moMsgSys.SendChangeEnvironment(.moPlanets(.CurrentPlanetIdx).ObjectID, .moPlanets(.CurrentPlanetIdx).ObjTypeID)
                                Else
                                    glCurrentEnvirView = CurrentView.ePlanetView
                                    If NewTutorialManager.TutorialOn = True Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eViewChanged, glCurrentEnvirView, -1, -1, "")
                                    'now set our location on the map, use cInt to round for us...
                                    lTemp = CInt(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(lTemp).GetHeightAtPoint(ptPos.X, ptPos.Y, True))

                                    With goCamera
                                        .mlCameraAtX = ptPos.X : .mlCameraAtY = 0 : .mlCameraAtZ = ptPos.Y
                                        .mlCameraX = ptPos.X + muSettings.lPlanetViewCameraX : .mlCameraY = lTemp + muSettings.lPlanetViewCameraY : .mlCameraZ = ptPos.Y + 200 + muSettings.lPlanetViewCameraZ
                                    End With
                                End If
                            End With

                        End If
                    End If
                ElseIf NewTutorialManager.TutorialOn = True Then
                    If goUILib Is Nothing = False OrElse goUILib.CommandAllowed(True, "UIEnable:ChangeView_" & CInt(CurrentView.ePlanetView).ToString) = True Then
                        glCurrentEnvirView = CurrentView.ePlanetView
                        'now set our location on the map, use cInt to round for us...
                        lTemp = CInt(Planet.GetTutorialPlanet().GetHeightAtPoint(ptPos.X, ptPos.Y, True))

                        With goCamera
                            .mlCameraAtX = ptPos.X : .mlCameraAtY = 0 : .mlCameraAtZ = ptPos.Y
                            .mlCameraX = ptPos.X + muSettings.lPlanetViewCameraX : .mlCameraY = lTemp + muSettings.lPlanetViewCameraY : .mlCameraZ = ptPos.Y + 200 + muSettings.lPlanetViewCameraZ
                        End With
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eViewChanged, glCurrentEnvirView, -1, -1, "")
                    End If
                End If
            Case CurrentView.ePlanetView
                HandleLeftMouseDownPlanetView(pickRay, oEnvir, lMouseX, lMouseY)
        End Select

        pickRay = Nothing
    End Sub
    Private Sub HandleLeftMouseDownBuildGhost()
        'Ok... player is selecting a location to build the object referred to in BuildGhost...
        mbIgnoreNextMouseUp = True
        mbIgnoreMouseMove = True

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return
        Dim lCurUB As Int32 = -1
        If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))

        If NewTutorialManager.TutorialOn = True Then
            For X As Int32 = 0 To lCurUB
                If oEnvir.lEntityIdx(X) <> -1 Then
                    ' AndAlso ((oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEnvir.oEntity(X).oMesh.bLandBased = True) OrElse oEnvir.ObjTypeID = ObjectType.eFacility) Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity Is Nothing = False AndAlso oEntity.bSelected = True AndAlso oEntity.ObjTypeID = ObjectType.eUnit Then
                        If oEntity.bProducing = True Then
                            If goUILib.CommandAllowed(True, "CancelFacilityProduction") = False Then Return
                        End If
                        Exit For
                    End If
                End If
            Next X

            Dim sParms() As String = {goUILib.BuildGhostID.ToString, goUILib.BuildGhostTypeID.ToString}
            If goUILib.CommandAllowedWithParms(True, "FacilityPlacement", sParms, False) = False Then Return
        End If

        'Verify that the facility's build location is not near any other facilities or in the case of station, near a planet
        Dim lTempVal As Int32 = CInt(goUILib.BuildGhost.ShieldXZRadius)
        Dim rcFac As Rectangle = Rectangle.FromLTRB(CInt(goUILib.vecBuildGhostLoc.X) - lTempVal, CInt(goUILib.vecBuildGhostLoc.Z) - lTempVal, CInt(goUILib.vecBuildGhostLoc.X) + lTempVal, CInt(goUILib.vecBuildGhostLoc.Z) + lTempVal)
        'Now, cycle thru to see if anyone is within that rect...
        For X As Int32 = 0 To lCurUB
            If oEnvir.lEntityIdx(X) <> -1 Then
                ' AndAlso ((oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEnvir.oEntity(X).oMesh.bLandBased = True) OrElse oEnvir.ObjTypeID = ObjectType.eFacility) Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.oMesh Is Nothing = False AndAlso ((oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEntity.oMesh.bLandBased = True) OrElse oEnvir.ObjTypeID = ObjectType.eFacility) AndAlso oEntity.bSelected = False Then
                    Dim lShieldRad As Int32 = CInt(oEntity.oMesh.ShieldXZRadius)
                    Dim rcTemp As Rectangle = Rectangle.FromLTRB(CInt(oEntity.LocX - lShieldRad), _
                    CInt(oEntity.LocZ - lShieldRad), CInt(oEntity.LocX + lShieldRad), _
                    CInt(oEntity.LocZ + lShieldRad))

                    If oEntity.yVisibility = eVisibilityType.Visible Then
                        If rcFac.IntersectsWith(rcTemp) = True Then
                            If goSound Is Nothing = False Then
                                goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                            End If
                            goUILib.AddNotification("Unable to build there, obstacle in the way.", System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next X

        If oEnvir.ObjTypeID = ObjectType.ePlanet Then

            Dim bIgnoreSlopeTest As Boolean = False
            Dim bNaval As Boolean = False
            Dim bMine As Boolean = False
            For X As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(X) = goUILib.BuildGhostID AndAlso goEntityDefs(X).ObjTypeID = goUILib.BuildGhostTypeID Then
                    If goEntityDefs(X).ProductionTypeID = ProductionType.eMining Then
                        bIgnoreSlopeTest = True
                        bMine = True
                    End If
                    bNaval = (goEntityDefs(X).yChassisType And ChassisType.eNavalBased) <> 0
                    Exit For
                End If
            Next X

            Dim lHalfVal As Int32 = CInt(Math.Ceiling(goUILib.BuildGhost.ShieldXZRadius / 2.0F))
            Dim lExtent As Int32 = (CType(goCurrentEnvir.oGeoObject, Planet).GetExtent \ 2)
            Dim lCellSpace As Int32 = CType(goCurrentEnvir.oGeoObject, Planet).CellSpacing
            Dim bValid As Boolean = True
            If goUILib.vecBuildGhostLoc.Z + lHalfVal * 4 > (lExtent - lCellSpace) OrElse goUILib.vecBuildGhostLoc.Z - (lHalfVal * 4) < -lExtent Then bValid = False
            If goUILib.vecBuildGhostLoc.X + lHalfVal * 4 > (lExtent - lCellSpace) OrElse goUILib.vecBuildGhostLoc.X - (lHalfVal * 4) < -lExtent Then bValid = False
            If bMine = True Then bValid = True
            If bValid = False Then
                If goSound Is Nothing = False Then
                    goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                End If
                goUILib.AddNotification("Unable to build that close to the edge of the map.", System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            If bNaval = False AndAlso goUILib.vecBuildGhostLoc.Y < CType(oEnvir.oGeoObject, Planet).WaterHeight Then
                If goSound Is Nothing = False Then
                    goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                End If
                goUILib.AddNotification("Unable to build that facility in liquid.", System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            ElseIf bNaval = True AndAlso (goUILib.vecBuildGhostLoc.Y > CType(oEnvir.oGeoObject, Planet).WaterHeight Or CType(oEnvir.oGeoObject, Planet).MapTypeID = PlanetType.eAcidic Or CType(oEnvir.oGeoObject, Planet).MapTypeID = PlanetType.eGeoPlastic) Then
                If goSound Is Nothing = False Then
                    goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                End If
                goUILib.AddNotification("Unable to build that facility out of liquid.", System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            Dim vecNormal As Vector3 = CType(oEnvir.oGeoObject, Planet).GetTriangleNormal(goUILib.vecBuildGhostLoc.X, goUILib.vecBuildGhostLoc.Z)
            If (bIgnoreSlopeTest = False AndAlso (Math.Abs(vecNormal.X) > 0.4 OrElse Math.Abs(vecNormal.Z) > 0.4)) OrElse (bValid = False AndAlso bIgnoreSlopeTest = False) Then
                If goSound Is Nothing = False Then
                    goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                End If
                'goUILib.AddNotification("Unable to build there, obstacle in the way.", System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                goUILib.AddNotification("Unable to build there, invalid terrain slope or height.", System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Exit Sub
            End If
        Else        'must be system
            'Now, determine intersection distance
            Dim lMinDist As Int32 = 15000   '15k by default with a min of 6k
            If goCurrentPlayer.StationPlacementCloserToPlanet = True Then lMinDist = 6000

            For X As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(X) = goUILib.BuildGhostID AndAlso goEntityDefs(X).ObjTypeID = goUILib.BuildGhostTypeID Then
                    If (goEntityDefs(X).ModelID And 255) = 148 Then
                        lMinDist = 0
                        If goCurrentEnvir.PositionOnPlanetRing(goUILib.vecBuildGhostLoc) = False Then
                            goUILib.AddNotification("Unable to build an Orbital Mining Platform except over a planet ring.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return
                        End If
                    End If
                    Exit For
                End If
            Next X

            'Now, find any planets/moons nearby
            With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                For X As Int32 = 0 To .PlanetUB
                    If Distance(CInt(goUILib.vecBuildGhostLoc.X), CInt(goUILib.vecBuildGhostLoc.Z), .moPlanets(X).LocX, .moPlanets(X).LocZ) < lMinDist Then
                        If goSound Is Nothing = False Then
                            goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                        End If
                        goUILib.AddNotification("Unable to build that close to a planet's gravity well!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                Next X

                For X As Int32 = 0 To .WormholeUB

                    Dim lTmpX As Int32
                    Dim lTmpZ As Int32
                    If .moWormholes(X).System1.ObjectID = .ObjectID Then
                        lTmpX = .moWormholes(X).LocX1 : lTmpZ = .moWormholes(X).LocY1
                    Else : lTmpX = .moWormholes(X).LocX2 : lTmpZ = .moWormholes(X).LocY2
                    End If

                    If Distance(CInt(goUILib.vecBuildGhostLoc.X), CInt(goUILib.vecBuildGhostLoc.Z), lTmpX, lTmpZ) < 1000 Then
                        If goSound Is Nothing = False Then
                            goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                        End If
                        goUILib.AddNotification("Unable to build that facility that close to a celestial body.", System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                Next X
            End With
        End If

        'If we're still good... then send build message to Region
        If goUILib.lUISelectState = UILib.eSelectState.ePlaceGuildFacility Then
            moMsgSys.SendSetGuildFacility(goUILib.BuildGhostID, goUILib.BuildGhostTypeID, CInt(goUILib.vecBuildGhostLoc.X), CInt(goUILib.vecBuildGhostLoc.Z), goUILib.BuildGhostAngle)
        Else
            For X As Int32 = 0 To lCurUB
                If oEnvir.lEntityIdx(X) <> -1 Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity Is Nothing = False AndAlso oEntity.bSelected = True Then
                        If goSound Is Nothing = False Then
                            goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUnitSpeech, Nothing, Nothing)
                        End If
                        If mbShiftKeyDown = True Then
                            moMsgSys.SendQueueSetProductionMsg(CType(oEntity, Base_GUID), goUILib.BuildGhostID, goUILib.BuildGhostTypeID, CInt(goUILib.vecBuildGhostLoc.X), CInt(goUILib.vecBuildGhostLoc.Z), goUILib.BuildGhostAngle)
                        Else
                            moMsgSys.SendSetProductionMsg(CType(oEntity, Base_GUID), goUILib.BuildGhostID, goUILib.BuildGhostTypeID, CInt(goUILib.vecBuildGhostLoc.X), CInt(goUILib.vecBuildGhostLoc.Z), goUILib.BuildGhostAngle)
                        End If


                        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
                            For Y As Int32 = 0 To glEntityDefUB
                                If glEntityDefIdx(Y) = goUILib.BuildGhostID AndAlso goEntityDefs(Y).ObjTypeID = goUILib.BuildGhostTypeID Then
                                    If goEntityDefs(Y).ProductionTypeID = ProductionType.eCommandCenterSpecial Then
                                        goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.CommandCenterPlaced)
                                    ElseIf goEntityDefs(Y).ProductionTypeID = ProductionType.eMining Then
                                        goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.EngineerOrderedToBuildMine)
                                    ElseIf goEntityDefs(Y).ProductionTypeID = ProductionType.eResearch Then
                                        goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.ResearchFacilityPlaced)
                                    End If
                                    Exit For
                                End If
                            Next Y
                        End If

                        Exit For
                    End If

                End If
            Next X
        End If

        If mbShiftKeyDown = False Then
            'Bring the build window back now that the placement has occured. Pass only
            Dim ofrm As UIWindow = goUILib.GetWindow("frmBuildWindow")
            If ofrm Is Nothing = False Then ofrm.Visible = True

            'and remove our buildghost
            goUILib.BuildGhost = Nothing
            goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        End If

    End Sub
    Private Function HandleLeftMouseDownSelectState(ByVal lMouseX As Int32, ByVal lMouseY As Int32) As Boolean
        'in a select state for the UI
        mbIgnoreNextMouseUp = True

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return False
        Dim lCurUB As Int32 = -1
        If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))

        Dim vecAddRay As Vector3
        Dim vecInt As Vector3

        'Ok, now, what select state am i in?
        Select Case goUILib.lUISelectState
            Case UILib.eSelectState.eSelectRaceWaypoint
                If glCurrentEnvirView = CurrentView.eSystemView OrElse glCurrentEnvirView = CurrentView.ePlanetView Then
                    'ok, check for if we are hitting an object...
                    Dim pickRay As Ray = CalcPickingRay(lMouseX, lMouseY)
                    'Ok, this can be either a planet or system...
                    Dim bFindIntOnTerrain As Boolean = False
                    If glCurrentEnvirView = CurrentView.ePlanetView Then
                        If oEnvir Is Nothing = False Then
                            If oEnvir.oGeoObject Is Nothing = False Then
                                'ok, we must be displaying the planet...
                                bFindIntOnTerrain = True
                            End If
                        End If
                    End If

                    'Nope, okay, then treat it as a move request
                    If bFindIntOnTerrain = False Then
                        If glCurrentEnvirView = CurrentView.ePlanetMapView Then
                            Dim ptPos As Point = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).GetMapClickedCoords(lMouseX, lMouseY)
                            If ptPos.X <> Int32.MinValue And ptPos.Y <> Int32.MinValue Then
                                vecInt = New Vector3(ptPos.X, 0, ptPos.Y)
                            End If
                        Else
                            vecInt = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 20000)))

                            If glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                                vecInt.X *= 10000 : vecInt.Y *= 10000 : vecInt.Z *= 10000
                            ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                                vecInt.X *= 30 : vecInt.Y *= 30 : vecInt.Z *= 30
                            End If
                        End If
                    ElseIf pickRay.Direction.Y < 0 Then
                        'we only want to do this if the unit can move there, a unit can only move there if the click occurs on
                        ' the ground... a click can only occur on the ground if Y is negative
                        vecAddRay = pickRay.Direction
                        vecAddRay.Scale(Math.Abs(1 / vecAddRay.Y) * 5)
                        vecInt = pickRay.Origin

                        'This is an attempt to reduce our test count... we get the maximum height and move our ray that much
                        Dim lTempVar As Int32 = CInt((CType(oEnvir.oGeoObject, Planet).GetMaxHeight() - vecInt.Y) / pickRay.Direction.Y)
                        vecInt.Add(Vector3.Multiply(pickRay.Direction, lTempVar))

                        While vecInt.Y > 0
                            vecInt.Add(vecAddRay)

                            If CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(vecInt.X, vecInt.Z, True) > vecInt.Y Then
                                'ok, that's it
                                Exit While
                            End If
                        End While
                    End If

                    Dim oTmpWin As frmRaceConfig = CType(goUILib.GetWindow("frmRaceConfig"), frmRaceConfig)
                    If oTmpWin Is Nothing = False Then
                        oTmpWin.SetLocResultVector(vecInt)
                    End If
                    oTmpWin = Nothing
                End If
                mbIgnoreNextMouseUp = False
                Return True
            Case UILib.eSelectState.eSelectRouteLoc, UILib.eSelectState.eSelectRouteTemplateLoc
                goCamera.SelectionBoxStart.X = lMouseX : goCamera.SelectionBoxStart.Y = lMouseY
                If glCurrentEnvirView = CurrentView.eSystemView OrElse glCurrentEnvirView = CurrentView.ePlanetView Then

                    'ok, check for if we are hitting an object...
                    Dim pickRay As Ray = CalcPickingRay(lMouseX, lMouseY)
                    Dim lIdx As Int32 = oEnvir.PickObject(pickRay, muSettings.FarClippingPlane)
                    If lIdx = -1 Then
                        'Ok, this can be either a planet or system...
                        Dim bFindIntOnTerrain As Boolean = False
                        If glCurrentEnvirView = CurrentView.ePlanetView Then
                            If oEnvir Is Nothing = False Then
                                If oEnvir.oGeoObject Is Nothing = False Then
                                    'ok, we must be displaying the planet...
                                    bFindIntOnTerrain = True
                                End If
                            End If
                        End If

                        'Nope, okay, then treat it as a move request
                        If bFindIntOnTerrain = False Then
                            If glCurrentEnvirView = CurrentView.ePlanetMapView Then
                                Dim ptPos As Point = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).GetMapClickedCoords(lMouseX, lMouseY)
                                If ptPos.X <> Int32.MinValue And ptPos.Y <> Int32.MinValue Then
                                    vecInt = New Vector3(ptPos.X, 0, ptPos.Y)
                                End If
                            Else
                                'are we in system view?
                                If glCurrentEnvirView = CurrentView.eSystemView Then
                                    'ok, did the player select a wormhole?
                                    If (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSuperSpecials) And 4) <> 0 Then
                                        Dim oTmpSphere As BoundingSphere
                                        For X As Int32 = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).WormholeUB
                                            If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes Is Nothing = False Then
                                                With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(X)
                                                    Dim bSystem1 As Boolean = False
                                                    If .System1 Is Nothing = False AndAlso goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID = .System1.ObjectID Then
                                                        oTmpSphere.SphereCenter.X = .LocX1
                                                        oTmpSphere.SphereCenter.Z = .LocY1
                                                        oTmpSphere.SphereCenter.Y = -200
                                                        bSystem1 = True
                                                    Else
                                                        oTmpSphere.SphereCenter.X = .LocX2
                                                        oTmpSphere.SphereCenter.Z = .LocY2
                                                        oTmpSphere.SphereCenter.Y = -200
                                                    End If
                                                    oTmpSphere.SphereRadius = 200

                                                    If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then

                                                        If goUILib.lUISelectState = UILib.eSelectState.eSelectRouteLoc Then
                                                            Dim oRouteCfg As frmRouteConfig = CType(goUILib.GetWindow("frmRouteConfig"), frmRouteConfig)
                                                            If oRouteCfg Is Nothing = False Then
                                                                If oEnvir Is Nothing = False Then
                                                                    'If oEnvir.lEntityIdx(lIdx) <> -1 Then
                                                                    oRouteCfg.SetLocResultWormhole(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(X), bSystem1)
                                                                    'End If
                                                                End If
                                                            End If
                                                            oRouteCfg = Nothing
                                                        Else
                                                            Dim oRouteCfg As frmRouteTemplate = CType(goUILib.GetWindow("frmRouteTemplate"), frmRouteTemplate)
                                                            If oRouteCfg Is Nothing = False Then
                                                                If oEnvir Is Nothing = False Then 
                                                                    oRouteCfg.SetLocResultWormhole(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(X), bSystem1)
                                                                End If
                                                            End If
                                                            oRouteCfg = Nothing
                                                        End If

                                                        
                                                        mbRightDrag = False
                                                        mbRightDown = False
                                                        goCamera.ShowSelectionBox = False
                                                        mbIgnoreNextMouseUp = False
                                                        Return True
                                                    End If
                                                End With
                                            End If
                                        Next X
                                    End If
                                End If

                                vecInt = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 20000)))

                                If glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                                    vecInt.X *= 10000 : vecInt.Y *= 10000 : vecInt.Z *= 10000
                                ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                                    vecInt.X *= 30 : vecInt.Y *= 30 : vecInt.Z *= 30
                                End If
                            End If
                        ElseIf pickRay.Direction.Y < 0 Then
                            'we only want to do this if the unit can move there, a unit can only move there if the click occurs on
                            ' the ground... a click can only occur on the ground if Y is negative
                            vecAddRay = pickRay.Direction
                            vecAddRay.Scale(Math.Abs(1 / vecAddRay.Y) * 5)
                            vecInt = pickRay.Origin

                            'This is an attempt to reduce our test count... we get the maximum height and move our ray that much
                            Dim lTempVar As Int32 = CInt((CType(oEnvir.oGeoObject, Planet).GetMaxHeight() - vecInt.Y) / pickRay.Direction.Y)
                            vecInt.Add(Vector3.Multiply(pickRay.Direction, lTempVar))

                            While vecInt.Y > 0
                                vecInt.Add(vecAddRay)

                                If CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(vecInt.X, vecInt.Z, True) > vecInt.Y Then
                                    'ok, that's it
                                    Exit While
                                End If
                            End While
                        End If

                        If goUILib.lUISelectState = UILib.eSelectState.eSelectRouteLoc Then
                            Dim oTmpWin As frmRouteConfig = CType(goUILib.GetWindow("frmRouteConfig"), frmRouteConfig)
                            If oTmpWin Is Nothing = False Then
                                oTmpWin.SetLocResultVector(vecInt)
                            End If
                            oTmpWin = Nothing
                        Else
                            Dim oRouteCfg As frmRouteTemplate = CType(goUILib.GetWindow("frmRouteTemplate"), frmRouteTemplate)
                            If oRouteCfg Is Nothing = False Then
                                If oEnvir Is Nothing = False Then
                                    oRouteCfg.SetLocResultVector(vecInt)
                                End If
                            End If
                            oRouteCfg = Nothing
                        End If
                        
                    Else
                        'ok, 
                        If goUILib.lUISelectState = UILib.eSelectState.eSelectRouteLoc Then
                            Dim oTmpWin As frmRouteConfig = CType(goUILib.GetWindow("frmRouteConfig"), frmRouteConfig)
                            If oTmpWin Is Nothing = False Then
                                If oEnvir Is Nothing = False Then
                                    If oEnvir.lEntityIdx(lIdx) <> -1 Then
                                        If oEnvir.oEntity(lIdx).OwnerID = glPlayerID Then
                                            oTmpWin.SetLocResultGUID(oEnvir.oEntity(lIdx))
                                        Else
                                            goUILib.AddNotification("Unable to set a route destination on a foreign Unit or Facility!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                                        End If
                                    End If
                                End If
                            End If
                            oTmpWin = Nothing
                        Else
                            Dim oRouteCfg As frmRouteTemplate = CType(goUILib.GetWindow("frmRouteTemplate"), frmRouteTemplate)
                            If oRouteCfg Is Nothing = False Then
                                If oEnvir Is Nothing = False Then
                                    If oEnvir.lEntityIdx(lIdx) <> -1 Then
                                        If oEnvir.oEntity(lIdx).OwnerID = glPlayerID Then
                                            oRouteCfg.SetLocResultGUID(oEnvir.oEntity(lIdx))
                                        Else
                                            goUILib.AddNotification("Unable to set a route destination on a foreign Unit or Facility!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                                        End If
                                    End If
                                End If
                            End If
                            oRouteCfg = Nothing
                        End If
                        
                    End If
                End If
                mbIgnoreNextMouseUp = False
                Return True
            Case UILib.eSelectState.eSetFleetDest
                'Ok, make sure we are in galaxy view
                If glCurrentEnvirView = CurrentView.eGalaxyMapView Then
                    Dim pickRay As Ray = CalcPickingRay(lMouseX, lMouseY)
                    'X = goGalaxy.PickSystem(pickRay, muSettings.FarClippingPlane)
                    'Now, do our check
                    Dim lSystemIdx As Int32 = -1
                    Dim oTmpSphere As BoundingSphere
                    Dim bHit As Boolean = False
                    For lIdx As Int32 = 0 To goGalaxy.mlSystemUB
                        With goGalaxy.moSystems(lIdx)
                            oTmpSphere.SphereCenter.X = .LocX
                            oTmpSphere.SphereCenter.Y = .LocY + 16.0F           'for the sprite offset
                            oTmpSphere.SphereCenter.Z = .LocZ
                            oTmpSphere.SphereRadius = 8
                            If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                                'Ok... set our tooltip
                                lSystemIdx = lIdx
                                bHit = True
                                Exit For
                            End If
                        End With
                    Next lIdx

                    'Now, if X <> -1 then
                    If bHit = True AndAlso lSystemIdx <> -1 Then
                        'ok, we have it...
                        If goGalaxy.moSystems(lSystemIdx).SystemType = SolarSystem.elSystemType.RespawnSystem Then
                            If goUILib Is Nothing = False Then
                                goUILib.AddNotification("Unable to travel to a Locked Respawn system!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            End If
                            mbIgnoreNextMouseUp = False
                            Return True
                        ElseIf goGalaxy.moSystems(lSystemIdx).SystemType = SolarSystem.elSystemType.SpawnSystem Then
                            If goUILib Is Nothing = False Then
                                goUILib.AddNotification("Unable to travel to a Locked Spawn system!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            End If
                            mbIgnoreNextMouseUp = False
                            Return True
                        End If
                        Dim ofrmFOrders As frmFleetOrders = CType(goUILib.GetWindow("frmFleetOrders"), frmFleetOrders)
                        If ofrmFOrders Is Nothing = False Then ofrmFOrders.SetDestNewSystem(lSystemIdx)
                        ofrmFOrders = Nothing
                    End If
                End If
                mbIgnoreNextMouseUp = False
                Return True
            Case UILib.eSelectState.eBombTargetSelection
                'ok, assuming we are in a planet...
                If oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso glCurrentEnvirView = CurrentView.ePlanetView AndAlso HasAliasedRights(AliasingRights.eChangeBehavior) = True Then
                    'Ok, selecting a location on the planet to bomb
                    Dim pickRay As Ray = CalcPickingRay(lMouseX, lMouseY)
                    If pickRay.Direction.Y > 0 Then
                        mbIgnoreNextMouseUp = False
                    Else
                        vecAddRay = pickRay.Direction
                        vecAddRay.Scale(Math.Abs(1 / vecAddRay.Y) * 5)
                        vecInt = pickRay.Origin

                        'This is an attempt to reduce our test count... we get the maximum height and move our ray that much
                        Dim lTempVar As Int32 = CInt((CType(oEnvir.oGeoObject, Planet).GetMaxHeight() - vecInt.Y) / pickRay.Direction.Y)
                        vecInt.Add(Vector3.Multiply(pickRay.Direction, lTempVar))

                        While vecInt.Y > 0
                            vecInt.Add(vecAddRay)

                            If CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(vecInt.X, vecInt.Z, True) > vecInt.Y Then
                                'ok, that's it
                                Exit While
                            End If
                        End While

                        'Ok, got our bombardment location
                        moMsgSys.SendBombardRequest(oEnvir.ObjectID, CInt(vecInt.X), CInt(vecInt.Z), goUILib.yBombardType)
                    End If
                Else : mbIgnoreNextMouseUp = False
                End If
            Case UILib.eSelectState.eDockWithObjectSelection
                Dim pickRay As Ray = CalcPickingRay(lMouseX, lMouseY)
                Dim lIdx As Int32 = oEnvir.PickObject(pickRay, muSettings.FarClippingPlane)
                If lIdx <> -1 Then
                    goUILib.SetBehaviorExtendedVals(oEnvir.oEntity(lIdx).ObjectID, oEnvir.oEntity(lIdx).ObjTypeID)
                Else : mbIgnoreNextMouseUp = False
                End If
            Case UILib.eSelectState.eEvadeToLocation
                'Ok, this can be either a planet or system...
                Dim bFindIntOnTerrain As Boolean = False
                If glCurrentEnvirView = CurrentView.ePlanetView Then
                    If oEnvir Is Nothing = False Then
                        If oEnvir.oGeoObject Is Nothing = False Then
                            'ok, we must be displaying the planet...
                            bFindIntOnTerrain = True
                        End If
                    End If
                End If

                Dim pickRay As Ray = CalcPickingRay(lMouseX, lMouseY)

                'Nope, okay, then treat it as a move request
                If bFindIntOnTerrain = False Then
                    If glCurrentEnvirView = CurrentView.ePlanetMapView Then
                        Dim ptPos As Point = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).GetMapClickedCoords(lMouseX, lMouseY)
                        If ptPos.X <> Int32.MinValue And ptPos.Y <> Int32.MinValue Then
                            vecInt = New Vector3(ptPos.X, 0, ptPos.Y)
                        End If
                    Else
                        vecInt = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 20000)))

                        If glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                            vecInt.X *= 10000 : vecInt.Y *= 10000 : vecInt.Z *= 10000
                        ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                            vecInt.X *= 30 : vecInt.Y *= 30 : vecInt.Z *= 30
                        End If
                    End If
                ElseIf pickRay.Direction.Y < 0 Then
                    'we only want to do this if the unit can move there, a unit can only move there if the click occurs on
                    ' the ground... a click can only occur on the ground if Y is negative
                    vecAddRay = pickRay.Direction
                    vecAddRay.Scale(Math.Abs(1 / vecAddRay.Y) * 5)
                    vecInt = pickRay.Origin

                    'This is an attempt to reduce our test count... we get the maximum height and move our ray that much
                    Dim lTempVar As Int32 = CInt((CType(oEnvir.oGeoObject, Planet).GetMaxHeight() - vecInt.Y) / pickRay.Direction.Y)
                    vecInt.Add(Vector3.Multiply(pickRay.Direction, lTempVar))

                    While vecInt.Y > 0
                        vecInt.Add(vecAddRay)

                        If CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(vecInt.X, vecInt.Z, True) > vecInt.Y Then
                            'ok, that's it
                            Exit While
                        End If
                    End While
                End If

                'VecInt has our location...
                goUILib.SetBehaviorExtendedVals(CInt(vecInt.X), CInt(vecInt.Z))
                goUILib.AddNotification("Evade location set.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Case UILib.eSelectState.eSketchPad_SelectPoint
                Dim oWin As frmSketchPad = CType(goUILib.GetWindow("frmSketchPad"), frmSketchPad)
                If oWin Is Nothing = False Then
                    oWin.PointSelected(lMouseX, lMouseY)
                End If
                mbIgnoreNextMouseUp = True
                mbRightDown = False
                mbRightDrag = False
                Return True
            Case UILib.eSelectState.eSketchPad_TextEntry
                goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                mbIgnoreNextMouseUp = True
                mbRightDown = False
                mbRightDrag = False
                Return True
        End Select

        'Now, clear our ui select state... and exit sub
        If mbIgnoreNextMouseUp = True Then
            goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
            Return True
        End If
        Return False
    End Function
    Private Function HandleLeftMouseDownChangeEnvir(ByRef oEnvir As BaseEnvironment) As Boolean
        If glCurrentEnvirView = CurrentView.ePlanetView Then
            If oEnvir Is Nothing = False Then
                Dim bNeedBurst As Boolean = oEnvir.ObjTypeID <> ObjectType.ePlanet AndAlso mbChangingEnvirs = False
                If bNeedBurst = True Then
                    mbEnvirReady = False
                    mbChangingEnvirs = True
                    If goUILib Is Nothing = False Then goUILib.ClearSelection()

                    Dim lLoopCnt As Int32 = 0
                    While gb_InHandleMovement
                        Threading.Thread.Sleep(10)
                        lLoopCnt += 1
                        If lLoopCnt > 1000 Then Exit While
                    End While

                    With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                        moMsgSys.SendChangeEnvironment(.moPlanets(.CurrentPlanetIdx).ObjectID, .moPlanets(.CurrentPlanetIdx).ObjTypeID)
                        mbIgnoreNextMouseUp = True
                    End With
                    Return True
                End If
            End If
        ElseIf glCurrentEnvirView = CurrentView.eSystemView Then
            If oEnvir Is Nothing = False Then
                Dim bNeedBurst As Boolean = oEnvir.ObjTypeID <> ObjectType.eSolarSystem AndAlso mbChangingEnvirs = False
                If bNeedBurst = True Then
                    mbEnvirReady = False
                    mbChangingEnvirs = True
                    If goUILib Is Nothing = False Then goUILib.ClearSelection()

                    Dim lLoopCnt As Int32 = 0
                    While gb_InHandleMovement
                        Threading.Thread.Sleep(10)
                        lLoopCnt += 1
                        If lLoopCnt > 1000 Then Exit While
                    End While

                    With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                        moMsgSys.SendChangeEnvironment(.ObjectID, .ObjTypeID)
                        .CurrentPlanetIdx = -1      '???
                        mbIgnoreNextMouseUp = True
                    End With

                    'If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
                    '    goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.ChangedToSolarSystemEnvir)
                    'End If
                    Return True
                End If
            End If
        End If
        Return False
    End Function
    Private Sub HandleLeftMouseDownGalaxyView(ByVal pickRay As Ray, ByRef oEnvir As BaseEnvironment)

        'a player cannot change systems while in Aurelium
        If goCurrentPlayer.yPlayerPhase <> 255 Then
            If goUILib Is Nothing = False Then
                goUILib.AddNotification("Unable to change systems while in Aurelium.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                goUILib.AddNotification("Use the Budget Window to go to your colony or press Home.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
            Return
        End If

        'Ok, select from our systems...
        'goGalaxy.CurrentSystemIdx = -1

        'X = goGalaxy.PickSystem(pickRay, muSettings.FarClippingPlane)
        Dim lSystemIdx As Int32 = -1
        Dim oTmpSphere As BoundingSphere
        Dim bHit As Boolean = False
        For lIdx As Int32 = 0 To goGalaxy.mlSystemUB
            With goGalaxy.moSystems(lIdx)
                oTmpSphere.SphereCenter.X = .LocX
                oTmpSphere.SphereCenter.Y = .LocY + 16.0F           'for the sprite offset
                oTmpSphere.SphereCenter.Z = .LocZ
                oTmpSphere.SphereRadius = 8
                If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                    'Ok... set our tooltip
                    lSystemIdx = lIdx
                    bHit = True
                    Exit For
                End If
            End With
        Next lIdx

        If lSystemIdx = -1 Then Exit Sub
        If goGalaxy.CurrentSystemIdx = -1 Then Exit Sub
        If goGalaxy.GalaxySelectionIdx <> lSystemIdx Then
            goGalaxy.GalaxySelectionIdx = lSystemIdx
            goGalaxy.GalaxySelectionChange = True
            Exit Sub
        End If

        goGalaxy.CurrentSystemIdx = lSystemIdx


        If goGalaxy.CurrentSystemIdx <> -1 Then
            'Ok, check if the system has planets defined
            If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB = -1 Then
                moMsgSys.RequestSystemDetails(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID)
            End If

            Dim bNeedBurst As Boolean = True
            If oEnvir Is Nothing = False Then
                bNeedBurst = oEnvir.ObjectID <> goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID OrElse _
                 oEnvir.ObjTypeID <> goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjTypeID
            End If

            If bNeedBurst = True Then
                mbEnvirReady = False
                mbChangingEnvirs = True
                Dim lLoopCnt As Int32 = 0
                While gb_InHandleMovement
                    Threading.Thread.Sleep(10)
                    lLoopCnt += 1
                    If lLoopCnt > 1000 Then Exit While
                End While
                moMsgSys.SendChangeEnvironment(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID, goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjTypeID)
            End If

            glCurrentEnvirView = CurrentView.eSystemMapView1

            With goCamera
                .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                .mlCameraX = 0 : .mlCameraY = 1000 : .mlCameraZ = -1000
            End With
        Else
            For lIdx As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                If goCurrentPlayer.mlUnitGroupIdx(lIdx) <> -1 Then
                    With goCurrentPlayer.moUnitGroups(lIdx)
                        If .lInterSystemTargetID <> -1 AndAlso .lInterSystemOriginID <> -1 Then
                            oTmpSphere = .GetBoundingSphere()
                            If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                                Dim ofrmFleetOrders As frmFleetOrders = CType(goUILib.GetWindow("frmFleetOrders"), frmFleetOrders)
                                If ofrmFleetOrders Is Nothing Then ofrmFleetOrders = New frmFleetOrders(goUILib)
                                ofrmFleetOrders.Visible = True
                                ofrmFleetOrders.SetFromFleet(.ObjectID)
                                ofrmFleetOrders = Nothing
                                Exit For
                            End If
                        End If
                    End With
                End If
            Next lIdx
        End If
    End Sub
    Private Function HandleLeftMouseDownSystemView(ByVal pickRay As Ray, ByRef oEnvir As BaseEnvironment, ByVal lMouseX As Int32, ByVal lMouseY As Int32) As Boolean
        Dim bHit As Boolean = False

        'test for hit
        Dim lIdx As Int32 = oEnvir.PickObject(pickRay, muSettings.FarClippingPlane)
        If lIdx <> -1 Then
            Dim oEntity As BaseEntity = oEnvir.oEntity(lIdx)
            If oEntity Is Nothing = False Then
                If NewTutorialManager.TutorialOn = True Then
                    If goUILib Is Nothing = False Then
                        Dim sParms() As String = {oEntity.ObjectID.ToString, oEntity.ObjTypeID.ToString, oEntity.yProductionType.ToString}
                        If goUILib.CommandAllowedWithParms(True, "ClickOnSelect", sParms, False) = False Then
                            Return True
                        End If
                    End If
                End If

                'in this case, we want to know if the object is being selected
                If oEntity.bSelected = False Then
                    If mbShiftKeyDown = False Then
                        oEnvir.DeselectAll()
                    End If
                End If

                If glCurrentCycle - mlLastMouseDownEvent < 15 Then
                    SelectAllSimilarUnits(True)
                    Return True
                End If

                If mbShiftKeyDown = False OrElse oEntity.OwnerID = glPlayerID Then
                    oEnvir.DeselectNonPlayerSelected()
                    oEntity.bSelected = Not oEntity.bSelected
                End If

                If oEntity.bSelected = True AndAlso oEntity.bHPChanged = True Then
                    moMsgSys.RequestEntityHPUpdate(oEntity.ObjectID, oEntity.ObjTypeID)
                End If
                bHit = True
            End If

        End If

        If bHit = False AndAlso glCurrentCycle - mlLastMouseDownEvent < 15 Then
            Dim oTmpSphere As BoundingSphere
            Dim bFound As Boolean = False
            For lIdx = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).WormholeUB
                With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(lIdx)
                    oTmpSphere.SphereRadius = 200.0F

                    Dim bIsSys1 As Boolean = False
                    If .System1 Is Nothing = False AndAlso .System1.ObjectID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID Then
                        oTmpSphere.SphereCenter.X = .LocX1
                        oTmpSphere.SphereCenter.Y = -100
                        oTmpSphere.SphereCenter.Z = .LocY1
                        bIsSys1 = True
                    ElseIf .System2 Is Nothing = False AndAlso .System2.ObjectID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID Then
                        oTmpSphere.SphereCenter.X = .LocX2
                        oTmpSphere.SphereCenter.Y = -100
                        oTmpSphere.SphereCenter.Z = .LocY2
                    End If
                    If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then

                        'ok, double-clicked a wormhole, jump to the other system
                        Dim lX As Int32
                        Dim lZ As Int32
                        Dim lSysID As Int32 = -1
                        If bIsSys1 = True AndAlso .System2 Is Nothing = False Then
                            lSysID = .System2.ObjectID
                            lX = .LocX2 : lZ = .LocY2
                        ElseIf bIsSys1 = False AndAlso .System1 Is Nothing = False Then
                            lSysID = .System1.ObjectID
                            lX = .LocX1 : lZ = .LocY1
                        End If
                        If lSysID = -1 Then Return True

                        Dim lSystemIdx As Int32 = -1
                        For X As Int32 = 0 To goGalaxy.mlSystemUB
                            If goGalaxy.moSystems(X) Is Nothing = False AndAlso goGalaxy.moSystems(X).ObjectID = lSysID Then
                                lSystemIdx = X
                                Exit For
                            End If
                        Next X
                        If lSystemIdx = -1 Then Return True

                        goGalaxy.CurrentSystemIdx = lSystemIdx

                        If goGalaxy.CurrentSystemIdx <> -1 Then
                            'Ok, check if the system has planets defined
                            If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB = -1 Then
                                moMsgSys.RequestSystemDetails(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID)
                            End If

                            Dim bNeedBurst As Boolean = True
                            If oEnvir Is Nothing = False Then
                                bNeedBurst = oEnvir.ObjectID <> goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID OrElse _
                                 oEnvir.ObjTypeID <> goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjTypeID
                            End If

                            If bNeedBurst = True Then
                                mbEnvirReady = False
                                mbChangingEnvirs = True
                                Dim lLoopCnt As Int32 = 0
                                While gb_InHandleMovement
                                    Threading.Thread.Sleep(10)
                                    lLoopCnt += 1
                                    If lLoopCnt > 1000 Then Exit While
                                End While
                                moMsgSys.SendChangeEnvironment(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID, goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjTypeID)
                            End If

                            glCurrentEnvirView = CurrentView.eSystemView

                            With goCamera
                                .SimplyPlaceCamera(lX, 0, lZ)
                            End With
                        End If
                        Return True
                    End If
                End With
            Next lIdx

            If bFound = False Then
                goUILib.RemoveWindow("frmPlanetDetails")
            End If
        End If

        If bHit = False Then
            'Ok, still haven't hit anything, try for celestial bodies...
            If goGalaxy.CurrentSystemIdx <> -1 Then

                'Static xlLastEventTime As Int32
                'If glCurrentCycle - xlLastEventTime < 15 Then
                Dim oTmpSphere As BoundingSphere
                For X As Int32 = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB
                    With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(X)
                        oTmpSphere.SphereCenter.X = .LocX
                        oTmpSphere.SphereCenter.Y = .LocY
                        oTmpSphere.SphereCenter.Z = .LocZ
                        oTmpSphere.SphereRadius = .PlanetRadius

                        If RaySphereIntersectionTest(pickRay, oTmpSphere) Then
                            'Ok, planet was selected, let's go there
                            bHit = True

                            'Set our index
                            goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx = X
                            'set our view
                            glCurrentEnvirView = CurrentView.ePlanetMapView
                            If NewTutorialManager.TutorialOn = True Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eViewChanged, glCurrentEnvirView, -1, -1, "")
                            Exit For
                        End If
                    End With
                Next X
                'End If
                'xlLastEventTime = glCurrentCycle
            End If
        End If

        If bHit = False AndAlso mbShiftKeyDown = False Then
            Dim bIgnoreDeselect As Boolean = False
            If NewTutorialManager.TutorialOn = True AndAlso goUILib Is Nothing = False Then
                bIgnoreDeselect = Not goUILib.CommandAllowed(True, "Deselect")
            End If
            If bIgnoreDeselect = False Then oEnvir.DeselectAll()
        End If

        'Mark for the Selection Box
        If goCamera Is Nothing = False Then
            goCamera.SelectionBoxStart = New Point(lMouseX, lMouseY)
        End If

        Return False
    End Function
    Private Sub HandleLeftMouseDownPlanetView(ByVal pickRay As Ray, ByRef oEnvir As BaseEnvironment, ByVal lMouseX As Int32, ByVal lMouseY As Int32)
        Dim bHit As Boolean = False

        'test the minimap if it is there
        If muSettings.ShowMiniMap = True OrElse (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255) Then
            If (lMouseX > muSettings.MiniMapLocX AndAlso lMouseX < muSettings.MiniMapLocX + muSettings.MiniMapWidthHeight) AndAlso _
               (lMouseY > muSettings.MiniMapLocY AndAlso lMouseY < muSettings.MiniMapLocY + muSettings.MiniMapWidthHeight) Then
                bHit = True

                If NewTutorialManager.TutorialOn = True Then
                    If goUILib Is Nothing = False Then
                        If goUILib.CommandAllowed(True, "MinimapLeftClick") = False Then Return
                    End If

                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eLeftClickOnMinimap, -1, -1, -1, "")
                End If

                'ok, move the camera to that point...
                Dim oPlanet As Planet = Nothing
                If NewTutorialManager.TutorialOn = True Then
                    If oEnvir Is Nothing = False AndAlso oEnvir.ObjectID >= 500000000 Then
                        oPlanet = Planet.GetTutorialPlanet()
                    End If
                End If
                If oPlanet Is Nothing AndAlso goGalaxy.CurrentSystemIdx <> -1 Then
                    If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx <> -1 Then
                        With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                            oPlanet = .moPlanets(.CurrentPlanetIdx)
                        End With
                    End If
                End If

                If oPlanet Is Nothing = False Then
                    Dim vecLoc As Vector3 = oPlanet.MiniMapClick(lMouseX - muSettings.MiniMapLocX, lMouseY - muSettings.MiniMapLocY)

                    Dim fCurHt As Single = oPlanet.GetHeightAtPoint(goCamera.mlCameraX, goCamera.mlCameraZ, True)
                    Dim lCXMod As Int32 = goCamera.mlCameraX - goCamera.mlCameraAtX
                    Dim lCZMod As Int32 = goCamera.mlCameraZ - goCamera.mlCameraAtZ

                    goCamera.mlCameraAtX = CInt(vecLoc.X) : goCamera.mlCameraAtZ = CInt(vecLoc.Z)
                    goCamera.mlCameraX = CInt(vecLoc.X) + lCXMod : goCamera.mlCameraZ = CInt(vecLoc.Z) + lCZMod

                    fCurHt = goCamera.mlCameraY - fCurHt
                    goCamera.mlCameraY = CInt(fCurHt + oPlanet.GetHeightAtPoint(goCamera.mlCameraX, goCamera.mlCameraZ, True))
                End If
            End If
        End If

        If bHit = False Then
            'test for hit
            Dim lIdx As Int32 = oEnvir.PickObject(pickRay, muSettings.TerrainFarClippingPlane)
            If lIdx <> -1 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(lIdx)
                If oEntity Is Nothing = False Then
                    If NewTutorialManager.TutorialOn = True Then
                        If goUILib Is Nothing = False Then
                            Dim sParms() As String = {oEntity.ObjectID.ToString, oEntity.ObjTypeID.ToString, oEntity.yProductionType.ToString}
                            If goUILib.CommandAllowedWithParms(True, "ClickOnSelect", sParms, False) = False Then
                                Return
                            End If
                        End If

                        If NewTutorialManager.GetTutorialStepID() > 3 AndAlso oEntity.bSelected = False Then
                            If mbShiftKeyDown = False Then
                                oEnvir.DeselectAll()
                            End If
                        End If
                    Else
                        If oEntity.bSelected = False Then
                            If mbShiftKeyDown = False Then
                                oEnvir.DeselectAll()
                            End If
                        End If
                    End If

                    If glCurrentCycle - mlLastMouseDownEvent < 15 Then
                        SelectAllSimilarUnits(True)
                        Return
                    End If

                    If mbShiftKeyDown = False OrElse oEntity.OwnerID = glPlayerID Then
                        oEnvir.DeselectNonPlayerSelected()
                        oEntity.bSelected = Not oEntity.bSelected
                    End If

                    If oEntity.bSelected = True AndAlso oEntity.bHPChanged = True Then
                        moMsgSys.RequestEntityHPUpdate(oEntity.ObjectID, oEntity.ObjTypeID)
                    End If
                    bHit = True
                End If

            End If
        End If

        If bHit = False AndAlso mbShiftKeyDown = False Then
            Dim bIgnoreDeselect As Boolean = False
            If NewTutorialManager.TutorialOn = True AndAlso goUILib Is Nothing = False Then
                bIgnoreDeselect = Not goUILib.CommandAllowed(True, "Deselect")
            End If
            If bIgnoreDeselect = False Then oEnvir.DeselectAll()
        End If

        'Mark for the Selection Box
        If goCamera Is Nothing = False Then
            goCamera.SelectionBoxStart = New Point(lMouseX, lMouseY)
        End If
    End Sub
#End Region

#Region "  MOUSE MOVE FUNCTIONS  "

    Private Sub frmMain_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        'If True = True Then Return
        If mbFormClosing = True Then Return
        If Me.Visible = False Then Return

        Dim deltaX As Int32
        Dim deltaY As Int32
        Dim pickRay As Ray
        Dim lIdx As Int32

        If mbChangingEnvirs = True Then Return
        If mbIgnoreMouseMove = True Then Return

        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return

        If mbMouseWheelOverride = True Then
            mlLastUIEvent = glCurrentCycle
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(True, "ZoomCamera") = False Then Return
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eZoomCamera, -1, -1, -1, "")
                End If
            End If

            deltaY = e.Y - mlMouseY
            If deltaY > 0 Then MouseWheelDown() Else MouseWheelUp()
        ElseIf mbRightDown = True Then
            If moEngine.bInCtrlMove = True Then
                'ok, in a ctrl move...
                Dim vecInt As Vector3 = GetEnvirClickedPoint(e.X, e.Y)
                moEngine.lMouseDirX = CInt(vecInt.X)
                moEngine.lMouseDirZ = CInt(vecInt.Z)
            Else
                deltaX = e.X - mlMouseX
                deltaY = e.Y - mlMouseY

                'Disable Y Axis changes in System Map View 1 and 2
                If glCurrentEnvirView = CurrentView.eSystemMapView1 OrElse glCurrentEnvirView = CurrentView.eSystemMapView2 Then deltaY = 0
                If glCurrentEnvirView <> CurrentView.eStartupLogin AndAlso glCurrentEnvirView <> CurrentView.eStartupDSELogo Then
                    mbRightDrag = goCamera.RotateCamera(deltaX, deltaY, mbRightDrag)
                End If
            End If
        ElseIf goUILib.BuildGhost Is Nothing = False Then
            Dim bFindIntOnTerrain As Boolean = False
            Dim vecInt As Vector3
            Dim fTemp As Single

            pickRay = CalcPickingRay(e.X, e.Y)
            If glCurrentEnvirView = CurrentView.ePlanetView Then
                If goCurrentEnvir Is Nothing = False Then
                    If goCurrentEnvir.oGeoObject Is Nothing = False Then
                        'ok, we must be displaying the planet...
                        bFindIntOnTerrain = True
                    End If
                End If
            End If

            If bFindIntOnTerrain = False Then
                If glCurrentEnvirView = CurrentView.ePlanetMapView Then
                    Dim ptPos As Point = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).GetMapClickedCoords(e.X, e.Y)
                    If ptPos.X <> Int32.MinValue And ptPos.Y <> Int32.MinValue Then
                        vecInt = New Vector3(ptPos.X, 0, ptPos.Y)
                    End If
                Else
                    vecInt = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 20000)))
                    'If vecInt.X > goCurrentEnvir.lMaxXPos Then
                    '    vecInt.X -= goCurrentEnvir.lMapWrapAdjustX
                    'ElseIf vecInt.X < goCurrentEnvir.lMinXPos Then
                    '    vecInt.X -= goCurrentEnvir.lMapWrapAdjustX
                    'End If
                    'Ok, now, if the facility being built (selected) is a mining facility, then we need to check for nearby mineral caches for snap to
                    For lTempIdx As Int32 = 0 To glEntityDefUB
                        If glEntityDefIdx(lTempIdx) = goUILib.BuildGhostID AndAlso goEntityDefs(lTempIdx).ObjTypeID = goUILib.BuildGhostTypeID Then
                            If (goEntityDefs(lTempIdx).ModelID And 255) = 148 Then
                                vecInt = goCurrentEnvir.PlaceAlongPlanetRing(vecInt)
                            End If
                            Exit For
                        End If
                    Next lTempIdx
                End If
            ElseIf pickRay.Direction.Y < 0 Then
                'we only want to do this if the unit can move there, a unit can only move there if the click occurs on
                ' the ground... a click can only occur on the ground if Y is negative
                Dim vecAddRay As Vector3 = pickRay.Direction

                vecAddRay.Scale(Math.Abs(1 / vecAddRay.Y) * 5)

                vecInt = pickRay.Origin

                'This is an attempt to reduce our test count... we get the maximum height and move our ray that much
                fTemp = (CType(goCurrentEnvir.oGeoObject, Planet).GetMaxHeight() - vecInt.Y) / pickRay.Direction.Y
                vecInt.Add(Vector3.Multiply(pickRay.Direction, fTemp))

                While vecInt.Y > 0
                    vecInt.Add(vecAddRay)

                    'If vecInt.X > goCurrentEnvir.lMaxXPos Then vecInt.X -= goCurrentEnvir.lMapWrapAdjustX
                    'If vecInt.X < goCurrentEnvir.lMinXPos Then vecInt.X += goCurrentEnvir.lMapWrapAdjustX

                    If CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(vecInt.X, vecInt.Z, True) > vecInt.Y Then
                        'ok, that's it
                        Exit While
                    End If
                End While

                'Ok, now, if the facility being built (selected) is a mining facility, then we need to check for nearby mineral caches for snap to
                For lTempIdx As Int32 = 0 To glEntityDefUB
                    If glEntityDefIdx(lTempIdx) = goUILib.BuildGhostID AndAlso goEntityDefs(lTempIdx).ObjTypeID = goUILib.BuildGhostTypeID Then
                        If goEntityDefs(lTempIdx).ProductionTypeID = ProductionType.eMining Then
                            vecInt = goCurrentEnvir.LocateClosestMineralCache(vecInt)
                        ElseIf (goEntityDefs(lTempIdx).ModelID And 255) = 148 Then
                            vecInt = goCurrentEnvir.PlaceAlongPlanetRing(vecInt)
                        End If
                        Exit For
                    End If
                Next lTempIdx

            End If

            'Now, vecInt is where we go
            goUILib.vecBuildGhostLoc = vecInt
        ElseIf goUILib.lUISelectState = UILib.eSelectState.eSketchPad_SelectPoint Then
            Dim oWin As frmSketchPad = CType(goUILib.GetWindow("frmSketchPad"), frmSketchPad)
            If oWin Is Nothing = False Then
                oWin.PeakPointSelected(e.X, e.Y)
            End If
        ElseIf goUILib.lUISelectState = UILib.eSelectState.eBombTargetSelection Then
            Dim bFindIntOnTerrain As Boolean = False
            Dim vecInt As Vector3
            Dim fTemp As Single

            pickRay = CalcPickingRay(e.X, e.Y)
            If glCurrentEnvirView = CurrentView.ePlanetView Then
                If goCurrentEnvir Is Nothing = False Then
                    If goCurrentEnvir.oGeoObject Is Nothing = False Then
                        'ok, we must be displaying the planet...
                        bFindIntOnTerrain = True
                    End If
                End If
            End If

            If bFindIntOnTerrain = True AndAlso pickRay.Direction.Y < 0 Then
                'we only want to do this if the unit can move there, a unit can only move there if the click occurs on
                ' the ground... a click can only occur on the ground if Y is negative
                Dim vecAddRay As Vector3 = pickRay.Direction

                vecAddRay.Scale(Math.Abs(1 / vecAddRay.Y) * 5)

                vecInt = pickRay.Origin

                'This is an attempt to reduce our test count... we get the maximum height and move our ray that much
                fTemp = (CType(goCurrentEnvir.oGeoObject, Planet).GetMaxHeight() - vecInt.Y) / pickRay.Direction.Y
                vecInt.Add(Vector3.Multiply(pickRay.Direction, fTemp))

                While vecInt.Y > 0
                    vecInt.Add(vecAddRay)

                    If vecInt.X > goCurrentEnvir.lMaxXPos Then vecInt.X -= goCurrentEnvir.lMapWrapAdjustX
                    If vecInt.X < goCurrentEnvir.lMinXPos Then vecInt.X += goCurrentEnvir.lMapWrapAdjustX

                    If CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(vecInt.X, vecInt.Z, True) > vecInt.Y Then
                        'ok, that's it
                        Exit While
                    End If
                End While
            End If
            goUILib.vecBombardLoc = vecInt
        ElseIf goCamera.ShowSelectionBox = True OrElse goUILib Is Nothing OrElse goUILib.PostMessage(UILibMsgCode.eMouseMoveMsgCode, e.X, e.Y, e.Button) = False Then

            If glCurrentEnvirView = CurrentView.eGalaxyMapView Then
                'pickRay = CalcPickingRay(e.X, e.Y)
                Dim oTmpSphere As BoundingSphere

                pickRay = CalcPickingRay(e.X, e.Y)

                'Clear our tooltip
                goUILib.SetToolTip(False)

                Dim bHit As Boolean = False

                'Now, do our check
                For lIdx = 0 To goGalaxy.mlSystemUB
                    With goGalaxy.moSystems(lIdx)
                        oTmpSphere.SphereCenter.X = .LocX
                        oTmpSphere.SphereCenter.Y = .LocY + 16.0F           'for the sprite offset
                        oTmpSphere.SphereCenter.Z = .LocZ
                        oTmpSphere.SphereRadius = 8
                        If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                            'Ok... set our tooltip

                            'TODO: GalaxyEnv, MouseOver Star, Extended Tooltip
                            'If lIdx = goGalaxy.GalaxySelectionIdx Then
                            '    goUILib.SetToolTip(.SystemName, e.X, e.Y)
                            'Else
                            Dim ofrmPDetails As frmPlanetDetails = CType(goUILib.GetWindow("frmPlanetDetails"), frmPlanetDetails)
                            If ofrmPDetails Is Nothing Then ofrmPDetails = New frmPlanetDetails(goUILib)
                            If ofrmPDetails.PlanetID <> .ObjectID Then
                                ofrmPDetails.SetFromStar(goGalaxy.moSystems(lIdx), e.X + 10, e.Y)
                                ofrmPDetails.Visible = True
                            Else : ofrmPDetails.UpdateLoc(e.X + 10, e.Y)
                            End If
                            ofrmPDetails = Nothing
                            'End If
                            bHit = True
                            Exit For
                        End If
                    End With
                Next lIdx

                If bHit = False Then
                    goUILib.SetToolTip(False)
                    goUILib.RemoveWindow("frmPlanetDetails")
                    For lIdx = 0 To goCurrentPlayer.mlUnitGroupUB
                        If goCurrentPlayer.mlUnitGroupIdx(lIdx) <> -1 Then
                            With goCurrentPlayer.moUnitGroups(lIdx)
                                If .lInterSystemTargetID <> -1 AndAlso .lInterSystemOriginID <> -1 Then
                                    oTmpSphere = .GetBoundingSphere()
                                    If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                                        'TODO: Eventually, we will want to provide more UI functionality, but for now, a tooltip will do
                                        goUILib.SetToolTip(.sName & " to " & goGalaxy.GetSystemName(.lInterSystemTargetID), e.X, e.Y)
                                        bHit = True
                                        Exit For
                                    End If
                                End If
                            End With
                        End If
                    Next lIdx
                End If
            ElseIf glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                If goGalaxy.CurrentSystemIdx <> -1 Then
                    pickRay = CalcPickingRay(e.X, e.Y)
                    Dim vecInt As Vector3 = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 20000)))

                    goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).SquareCenterX = CInt(vecInt.X)
                    goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).SquareCenterZ = CInt(vecInt.Z)

                    lIdx = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CheckSystemMapHover(CInt(vecInt.X), CInt(vecInt.Z))

                    If lIdx <> -1 Then
                        Dim lStarCnt As Int32
                        If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).StarType1Idx <> -1 Then lStarCnt += 1
                        If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).StarType2Idx <> -1 Then lStarCnt += 1
                        If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).StarType3Idx <> -1 Then lStarCnt += 1

                        If lIdx < lStarCnt Then
                            If lStarCnt > 1 Then
                                goUILib.SetToolTip(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).SystemName & " " & (lIdx + 1), e.X, e.Y)
                            Else
                                goUILib.SetToolTip(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).SystemName, e.X, e.Y)
                                'TODO: SolarSystemEnv, MouseOver Star, Extended Tooltip
                                'Dim ofrmPDetails As frmPlanetDetails = CType(goUILib.GetWindow("frmPlanetDetails"), frmPlanetDetails)
                                'If ofrmPDetails Is Nothing Then ofrmPDetails = New frmPlanetDetails(goUILib)
                                'If ofrmPDetails.PlanetID <> goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID Then
                                '    ofrmPDetails.SetFromStar(goGalaxy.moSystems(lIdx), e.X + 10, e.Y)
                                '    ofrmPDetails.Visible = True
                                'Else : ofrmPDetails.UpdateLoc(e.X + 10, e.Y)
                                'End If
                                'ofrmPDetails = Nothing
                            End If
                        Else
                            'goUILib.SetToolTip(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(lIdx - lStarCnt).PlanetName, e.X, e.Y)
                            Dim ofrmPDetails As frmPlanetDetails = CType(goUILib.GetWindow("frmPlanetDetails"), frmPlanetDetails)
                            If ofrmPDetails Is Nothing Then ofrmPDetails = New frmPlanetDetails(goUILib)
                            If ofrmPDetails.PlanetID <> goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(lIdx - lStarCnt).ObjectID Then
                                ofrmPDetails.SetFromPlanet(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(lIdx - lStarCnt), e.X + 10, e.Y)
                                ofrmPDetails.Visible = True
                            Else : ofrmPDetails.UpdateLoc(e.X + 10, e.Y)
                            End If
                            ofrmPDetails = Nothing
                        End If
                    Else
                        Dim oTmpSphere As BoundingSphere
                        Dim bFound As Boolean = False
                        For lIdx = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).WormholeUB
                            With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(lIdx)
                                oTmpSphere.SphereRadius = 5.0F
                                If .System1 Is Nothing = False AndAlso .System1.ObjectID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID Then
                                    oTmpSphere.SphereCenter.X = .LocX1 / 10000.0F
                                    oTmpSphere.SphereCenter.Y = 0
                                    oTmpSphere.SphereCenter.Z = .LocY1 / 10000.0F
                                ElseIf .System2 Is Nothing = False AndAlso .System2.ObjectID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID Then
                                    oTmpSphere.SphereCenter.X = .LocX2 / 10000.0F
                                    oTmpSphere.SphereCenter.Y = 0
                                    oTmpSphere.SphereCenter.Z = .LocY2 / 10000.0F
                                End If
                                If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                                    'Ok... set our tooltip
                                    'goUILib.SetToolTip(.PlanetName, e.X, e.Y) 
                                    Dim ofrmPDetails As frmPlanetDetails = CType(goUILib.GetWindow("frmPlanetDetails"), frmPlanetDetails)
                                    If ofrmPDetails Is Nothing Then ofrmPDetails = New frmPlanetDetails(goUILib)
                                    If ofrmPDetails.PlanetID <> .ObjectID Then
                                        ofrmPDetails.SetFromWormhole(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(lIdx), goGalaxy.moSystems(goGalaxy.CurrentSystemIdx), e.X + 10, e.Y)
                                        ofrmPDetails.Visible = True
                                    Else : ofrmPDetails.UpdateLoc(e.X + 10, e.Y)
                                    End If
                                    ofrmPDetails = Nothing
                                    bFound = True
                                    Exit For
                                End If
                            End With
                        Next lIdx

                        If bFound = False Then

                            goUILib.SetToolTip(False)
                            goUILib.RemoveWindow("frmPlanetDetails")
                        End If
                    End If
                End If
            ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                If goGalaxy.CurrentSystemIdx <> -1 Then
                    Dim oTmpSphere As BoundingSphere
                    Dim bFound As Boolean = False
                    pickRay = CalcPickingRay(e.X, e.Y)

                    'Clear our tooltip
                    goUILib.SetToolTip(False)

                    'Now, do our check
                    For lIdx = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB
                        With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(lIdx)
                            oTmpSphere.SphereCenter.X = .LocX / 30.0F
                            oTmpSphere.SphereCenter.Y = .LocY / 30.0F
                            oTmpSphere.SphereCenter.Z = .LocZ / 30.0F
                            oTmpSphere.SphereRadius = .PlanetRadius / 30.0F
                            If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                                'Ok... set our tooltip
                                'goUILib.SetToolTip(.PlanetName, e.X, e.Y) 
                                Dim ofrmPDetails As frmPlanetDetails = CType(goUILib.GetWindow("frmPlanetDetails"), frmPlanetDetails)
                                If ofrmPDetails Is Nothing Then ofrmPDetails = New frmPlanetDetails(goUILib)
                                If ofrmPDetails.PlanetID <> .ObjectID Then
                                    ofrmPDetails.SetFromPlanet(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(lIdx), e.X + 10, e.Y)
                                    ofrmPDetails.Visible = True
                                Else : ofrmPDetails.UpdateLoc(e.X + 10, e.Y)
                                End If
                                ofrmPDetails = Nothing
                                bFound = True

                                Exit For
                            End If
                        End With
                    Next lIdx

                    For lIdx = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).WormholeUB
                        With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(lIdx)
                            oTmpSphere.SphereRadius = 20.0F
                            If .System1 Is Nothing = False AndAlso .System1.ObjectID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID Then
                                oTmpSphere.SphereCenter.X = .LocX1 / 30.0F
                                oTmpSphere.SphereCenter.Y = 0
                                oTmpSphere.SphereCenter.Z = .LocY1 / 30.0F
                            ElseIf .System2 Is Nothing = False AndAlso .System2.ObjectID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID Then
                                oTmpSphere.SphereCenter.X = .LocX2 / 30.0F
                                oTmpSphere.SphereCenter.Y = 0
                                oTmpSphere.SphereCenter.Z = .LocY2 / 30.0F
                            End If
                            If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                                'Ok... set our tooltip
                                'goUILib.SetToolTip(.PlanetName, e.X, e.Y) 
                                Dim ofrmPDetails As frmPlanetDetails = CType(goUILib.GetWindow("frmPlanetDetails"), frmPlanetDetails)
                                If ofrmPDetails Is Nothing Then ofrmPDetails = New frmPlanetDetails(goUILib)
                                If ofrmPDetails.PlanetID <> .ObjectID Then
                                    ofrmPDetails.SetFromWormhole(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(lIdx), goGalaxy.moSystems(goGalaxy.CurrentSystemIdx), e.X + 10, e.Y)
                                    ofrmPDetails.Visible = True
                                Else : ofrmPDetails.UpdateLoc(e.X + 10, e.Y)
                                End If
                                ofrmPDetails = Nothing
                                bFound = True

                                Exit For
                            End If
                        End With
                    Next lIdx

                    If bFound = False Then
                        goUILib.RemoveWindow("frmPlanetDetails")
                    End If
                End If
            ElseIf glCurrentEnvirView = CurrentView.ePlanetMapView Then
                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                If oEnvir Is Nothing = False AndAlso oEnvir.oGeoObject Is Nothing = False Then
                    If CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                        Dim ptLoc As Point = CType(oEnvir.oGeoObject, Planet).GetMapClickedCoords(e.X, e.Y)
                        Dim lCurUB As Int32 = -1
                        If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
                        Dim sDisplay As String = ""
                        For X As Int32 = 0 To lCurUB
                            If oEnvir.lEntityIdx(X) > -1 Then
                                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                                If oEntity Is Nothing = False AndAlso oEntity.yVisibility <> eVisibilityType.NoVisibility Then

                                    If Math.Abs(oEntity.LocX - ptLoc.X) < 1000 Then
                                        If Math.Abs(oEntity.LocZ - ptLoc.Y) < 1000 Then
                                            If oEntity.yVisibility = eVisibilityType.Visible Then
                                                If oEntity.OwnerID = glPlayerID Then
                                                    sDisplay = GetCacheObjectValue(oEntity.ObjectID, oEntity.ObjTypeID) & vbCrLf & GetCacheObjectValue(oEntity.OwnerID, ObjectType.ePlayer)
                                                Else
                                                    If oEntity.sOverrideOwnerName Is Nothing OrElse oEntity.sOverrideOwnerName = "" Then
                                                        Dim bUnknown As Boolean = False
                                                        oEntity.sOverrideOwnerName = GetCacheObjectValueCheckUnknowns(oEntity.OwnerID, ObjectType.ePlayer, bUnknown)
                                                        If bUnknown = True Then oEntity.sOverrideOwnerName = ""
                                                    End If
                                                    sDisplay = "Contact: " & oEntity.sOverrideOwnerName 'GetCacheObjectValue(oEntity.OwnerID, ObjectType.ePlayer)
                                                End If

                                                Exit For
                                            ElseIf oEntity.yVisibility = eVisibilityType.FacilityIntel Then
                                                If oEntity.sOverrideOwnerName Is Nothing OrElse oEntity.sOverrideOwnerName = "" Then
                                                    Dim bUnknown As Boolean = False
                                                    oEntity.sOverrideOwnerName = GetCacheObjectValueCheckUnknowns(oEntity.OwnerID, ObjectType.ePlayer, bUnknown)
                                                    If bUnknown = True Then oEntity.sOverrideOwnerName = ""
                                                End If
                                                sDisplay = "Intelligence Data" & vbCrLf & oEntity.sOverrideOwnerName 'GetCacheObjectValue(oEntity.OwnerID, ObjectType.ePlayer)
                                            ElseIf oEntity.yVisibility = eVisibilityType.InMaxRange Then
                                                sDisplay = "Unknown Contact in Detection Range"
                                            End If

                                        End If
                                    End If

                                End If
                            End If
                        Next X
                        If sDisplay <> "" Then goUILib.SetToolTip(sDisplay, e.X + 10, e.Y)
                    End If
                End If
            ElseIf glCurrentEnvirView = CurrentView.eSystemView OrElse glCurrentEnvirView = CurrentView.ePlanetView Then
                If e.Button = Windows.Forms.MouseButtons.Left Then
                    If Math.Abs(goCamera.SelectionBoxStart.X - e.X) > 20 OrElse Math.Abs(goCamera.SelectionBoxStart.Y - e.Y) > 20 Then

                        Dim bIgnoreSelectBoxStart As Boolean = False
                        If NewTutorialManager.TutorialOn = True Then
                            bIgnoreSelectBoxStart = Not goUILib.CommandAllowed(True, "SelectionBoxStart")
                        End If
                        If bIgnoreSelectBoxStart = False AndAlso goUILib.oButtonDown Is Nothing = True AndAlso goUILib.oOptionDown Is Nothing = True Then
                            goCamera.ShowSelectionBox = True
                            goCamera.SelectionBoxEnd.X = e.X
                            goCamera.SelectionBoxEnd.Y = e.Y
                        End If
                    End If

                ElseIf goCurrentEnvir Is Nothing = False AndAlso (glCurrentEnvirView <> CurrentView.ePlanetView OrElse muSettings.RenderMineralCaches = True OrElse (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255)) Then
                    mlCheckMineralCacheHover = glCurrentCycle
                End If
            End If

            With goCamera

                .bScrollLeft = mbKeyScrollLeft : .bScrollRight = mbKeyScrollRight : .bScrollUp = mbKeyScrollUp : .bScrollDown = mbKeyScrollDown

                Dim lWidth As Int32 = Me.ClientSize.Width
                Dim lHeight As Int32 = Me.ClientSize.Height
                Dim oDevice As Device = GFXEngine.moDevice
                If oDevice Is Nothing = False AndAlso oDevice.PresentationParameters.Windowed = False Then
                    lWidth = oDevice.PresentationParameters.BackBufferWidth
                    lHeight = oDevice.PresentationParameters.BackBufferHeight
                End If
                If .ShowSelectionBox = False Then
                    If e.X < .ScrollEdgeWidth AndAlso mbKeyScrollLeft = False Then
                        .bScrollLeft = True
                    ElseIf e.X > lWidth - .ScrollEdgeWidth AndAlso mbKeyScrollRight = False Then
                        .bScrollRight = True
                    End If
                    If e.Y < .ScrollEdgeWidth AndAlso mbKeyScrollUp = False Then
                        .bScrollUp = True
                    ElseIf e.Y > lHeight - .ScrollEdgeWidth AndAlso mbKeyScrollDown = False Then
                        .bScrollDown = True
                    End If
                End If

            End With
        Else
            'Check for scroll stop
            Dim lWidth As Int32 = Me.ClientSize.Width
            Dim lHeight As Int32 = Me.ClientSize.Height
            Dim oDevice As Device = GFXEngine.moDevice
            If oDevice Is Nothing = False AndAlso oDevice.PresentationParameters.Windowed = False Then
                lWidth = oDevice.PresentationParameters.BackBufferWidth
                lHeight = oDevice.PresentationParameters.BackBufferHeight
            End If

            With goCamera
                If .bScrollLeft = True Then
                    If e.X > .ScrollEdgeWidth AndAlso mbKeyScrollLeft = False Then .bScrollLeft = False
                ElseIf .bScrollRight = True Then
                    If e.X < (lWidth - .ScrollEdgeWidth) AndAlso mbKeyScrollRight = False Then .bScrollRight = False
                End If
                If .bScrollUp = True Then
                    If e.Y > .ScrollEdgeWidth AndAlso mbKeyScrollUp = False Then .bScrollUp = False
                ElseIf .bScrollDown = True Then
                    If e.Y < (lHeight - .ScrollEdgeWidth) AndAlso mbKeyScrollDown = False Then .bScrollDown = False
                End If
            End With
        End If

        mlMouseX = e.X
        mlMouseY = e.Y
    End Sub

    Private mlCheckMineralCacheHover As Int32 = Int32.MaxValue
    Private Sub CheckMineralCacheHover()
        If mlCheckMineralCacheHover = Int32.MaxValue Then Return
        If glCurrentCycle - mlCheckMineralCacheHover > 4 Then
            mlCheckMineralCacheHover = Int32.MaxValue

            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing Then Return

            Dim pickRay As Ray = CalcPickingRay(mlMouseX, mlMouseY)
            Dim lIdx As Int32 = -1
            If glCurrentEnvirView = CurrentView.ePlanetView Then
                lIdx = oEnvir.PickCache(pickRay, muSettings.TerrainFarClippingPlane)
            Else : lIdx = oEnvir.PickCache(pickRay, muSettings.FarClippingPlane)
            End If
            If lIdx <> -1 Then
                Dim oTmp As frmCacheSelect = CType(goUILib.GetWindow("frmCacheSelect"), frmCacheSelect)
                If oTmp Is Nothing Then oTmp = New frmCacheSelect(goUILib)
                oTmp.Left = CInt(mlMouseX - (oTmp.Width / 2))
                oTmp.Top = mlMouseY + 10

                If oTmp.Left < 0 Then oTmp.Left = 0
                If oTmp.Top < 0 Then oTmp.Top = 0
                If oTmp.Left + oTmp.Width > goUILib.oDevice.PresentationParameters.BackBufferWidth Then
                    oTmp.Left = (goUILib.oDevice.PresentationParameters.BackBufferWidth - oTmp.Width)
                End If
                If oTmp.Top + oTmp.Height > goUILib.oDevice.PresentationParameters.BackBufferHeight Then
                    oTmp.Top = (goUILib.oDevice.PresentationParameters.BackBufferHeight - oTmp.Height)
                End If

                If oTmp.SetProps(oEnvir.oCache(lIdx)) = False Then
                    moMsgSys.RequestEntityDetails(oEnvir.oCache(lIdx), oEnvir)
                End If
                oTmp = Nothing
            Else
                goUILib.RemoveWindow("frmCacheSelect")
                If glCurrentEnvirView = CurrentView.eSystemView Then
                    Dim oTmpSphere As BoundingSphere
                    Dim bFound As Boolean = False
                    For lIdx = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).WormholeUB
                        With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(lIdx)
                            oTmpSphere.SphereRadius = 200.0F
                            If .System1 Is Nothing = False AndAlso .System1.ObjectID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID Then
                                oTmpSphere.SphereCenter.X = .LocX1
                                oTmpSphere.SphereCenter.Y = -100
                                oTmpSphere.SphereCenter.Z = .LocY1
                            ElseIf .System2 Is Nothing = False AndAlso .System2.ObjectID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID Then
                                oTmpSphere.SphereCenter.X = .LocX2
                                oTmpSphere.SphereCenter.Y = -100
                                oTmpSphere.SphereCenter.Z = .LocY2
                            End If
                            If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                                'Ok... set our tooltip
                                'goUILib.SetToolTip(.PlanetName, e.X, e.Y) 
                                Dim ofrmPDetails As frmPlanetDetails = CType(goUILib.GetWindow("frmPlanetDetails"), frmPlanetDetails)
                                If ofrmPDetails Is Nothing Then ofrmPDetails = New frmPlanetDetails(goUILib)
                                If ofrmPDetails.PlanetID <> .ObjectID Then
                                    ofrmPDetails.SetFromWormhole(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(lIdx), goGalaxy.moSystems(goGalaxy.CurrentSystemIdx), mlMouseX + 10, mlMouseY)
                                    ofrmPDetails.Visible = True
                                Else : ofrmPDetails.UpdateLoc(mlMouseX + 10, mlMouseY)
                                End If
                                ofrmPDetails = Nothing
                                bFound = True

                                Exit For
                            End If
                        End With
                    Next lIdx

                    If bFound = False Then
                        goUILib.RemoveWindow("frmPlanetDetails")
                    End If
                End If

                'Now our distance stuff
                If mbAltKeyDown = True Then
                    DoDistanceDisplayUpdate(pickRay)
                Else
                    ResetDistanceDisplay()
                End If

            End If
        End If
    End Sub

#End Region

    Private Sub ResetDistanceDisplay()
        If goUILib Is Nothing = False Then goUILib.mfDistFromSelection = Single.MinValue
    End Sub
    Private Sub DoDistanceDisplayUpdate(ByVal pickRay As Ray)
        Dim bFindIntOnTerrain As Boolean = False
        Dim vecInt As Vector3
        Dim fTemp As Single

        If glCurrentEnvirView <> CurrentView.ePlanetView AndAlso glCurrentEnvirView <> CurrentView.eSystemView Then Return

        If glCurrentEnvirView = CurrentView.ePlanetView Then
            If goCurrentEnvir Is Nothing = False Then
                If goCurrentEnvir.oGeoObject Is Nothing = False Then
                    'ok, we must be displaying the planet...
                    bFindIntOnTerrain = True
                End If
            End If
        End If

        If bFindIntOnTerrain = True AndAlso pickRay.Direction.Y < 0 Then
            'we only want to do this if the unit can move there, a unit can only move there if the click occurs on
            ' the ground... a click can only occur on the ground if Y is negative
            Dim vecAddRay As Vector3 = pickRay.Direction

            vecAddRay.Scale(Math.Abs(1 / vecAddRay.Y) * 5)

            vecInt = pickRay.Origin

            'This is an attempt to reduce our test count... we get the maximum height and move our ray that much
            fTemp = (CType(goCurrentEnvir.oGeoObject, Planet).GetMaxHeight() - vecInt.Y) / pickRay.Direction.Y
            vecInt.Add(Vector3.Multiply(pickRay.Direction, fTemp))

            While vecInt.Y > 0
                vecInt.Add(vecAddRay)

                If vecInt.X > goCurrentEnvir.lMaxXPos Then vecInt.X -= goCurrentEnvir.lMapWrapAdjustX
                If vecInt.X < goCurrentEnvir.lMinXPos Then vecInt.X += goCurrentEnvir.lMapWrapAdjustX

                If CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(vecInt.X, vecInt.Z, True) > vecInt.Y Then
                    'ok, that's it
                    Exit While
                End If
            End While
        Else
            vecInt = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 20000)))
        End If
        If goUILib Is Nothing = False Then goUILib.UpdateDistanceIndicator(vecInt.X, vecInt.Z)
    End Sub

#Region "  MOUSE UP FUNCTIONS  "

    Private Sub frmMain_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
        mlLastUIEvent = glCurrentCycle
        If mbFormClosing = True Then Return
        If Me.Visible = False Then Return
        If mbChangingEnvirs = True Then Return
        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return
        mbIgnoreMouseMove = False
        If mbIgnoreNextMouseUp = True Then
            mbIgnoreNextMouseUp = False
            Return
        End If
        If goUILib Is Nothing = False Then
            If goUILib.oButtonDown Is Nothing = False Then
                goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                goUILib.oButtonDown.UIButton_OnMouseUp(e.X, e.Y, e.Button)
                goUILib.oButtonDown = Nothing
                Return
            ElseIf goUILib.oOptionDown Is Nothing = False Then
                goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                goUILib.oOptionDown.UIButton_OnMouseUp(e.X, e.Y, e.Button)
                goUILib.oOptionDown = Nothing
                Return
            End If
        End If
        'Check our UI first...
        If goUILib Is Nothing = False Then
            goUILib.oMovingWindow = Nothing
            If goUILib.lUISelectState = UILib.eSelectState.eMoveWindow OrElse goUILib.lUISelectState = UILib.eSelectState.eResizeWindow Then goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        End If

        'first, determine if our interfaces are clicked
        If goUILib Is Nothing = False AndAlso goCamera.ShowSelectionBox = False AndAlso goUILib.PostMessage(UILibMsgCode.eMouseUpMsgCode, e.X, e.Y, e.Button) = True Then
            If goUILib Is Nothing = False AndAlso goUILib.oButtonDown Is Nothing = False Then goUILib.oButtonDown.ResetPressed()
            If goUILib Is Nothing = False AndAlso goUILib.oOptionDown Is Nothing = False Then goUILib.oOptionDown.ResetPressed()
            mbRightDrag = False
            mbRightDown = False
            Return
        End If
        If goUILib.oButtonDown Is Nothing = False Then goUILib.oButtonDown.ResetPressed()
        goUILib.oButtonDown = Nothing
        If goUILib.oOptionDown Is Nothing = False Then goUILib.oOptionDown.ResetPressed()
        goUILib.oOptionDown = Nothing

        'Check for normal release right-click in normal views (cancels drag and down flags)
        If glCurrentEnvirView > CurrentView.eFullScreenInterface Then
            mbRightDrag = False
            mbRightDown = False
            Return
        End If

        'Now, test which button was released
        If e.Button = Windows.Forms.MouseButtons.Right Then
            HandleRightMouseUp(e.X, e.Y)
        ElseIf e.Button = Windows.Forms.MouseButtons.Left Then
            If goCamera.ShowSelectionBox = True Then
                SelectItemsInSelectionBox(e.X, e.Y)
            End If
        End If

        'regardless, ensure that the selection box is invisible
        goCamera.ShowSelectionBox = False
    End Sub
    Private Sub HandleRightMouseUp(ByVal lMouseX As Int32, ByVal lMouseY As Int32)

        Try
            If goCamera Is Nothing = False Then goCamera.ResetRotateCameraThresholds()

            If mbRightDrag = False Then
                If goUILib Is Nothing = False Then
                    If goUILib.lUISelectState = UILib.eSelectState.eSetFleetDest Then
                        mbRightDown = False
                        mbRightDrag = False
                        'mbIgnoreNextMouseUp = True
                        goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                        goUILib.AddNotification("Set Destination Cancelled", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                End If
                'execute an order... need to determine what we are looking at...
                If glCurrentEnvirView = CurrentView.eGalaxyMapView OrElse glCurrentEnvirView > CurrentView.ePlanetView Then Return

                'First, if we are in planet view
                If glCurrentEnvirView = CurrentView.ePlanetView AndAlso NewTutorialManager.TutorialOn = False Then
                    If lMouseX >= muSettings.MiniMapLocX AndAlso lMouseX < muSettings.MiniMapLocX + muSettings.MiniMapWidthHeight Then
                        If lMouseY >= muSettings.MiniMapLocY AndAlso lMouseY < muSettings.MiniMapLocY + muSettings.MiniMapWidthHeight Then
                            'Ok, clicking on the minimap... RIGHT clicking to be exact
                            Dim vecLoc As Vector3 = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).MiniMapClick(lMouseX, lMouseY)
                            If mbSetRallyPoint = True Then
                                moMsgSys.SendRallyPointMsg(CInt(vecLoc.X), CInt(vecLoc.Z), -1S)
                                If goUILib Is Nothing = False Then goUILib.AddNotification("Rally point set.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                mbSetRallyPoint = False
                            Else
                                'Now... send our move request
                                Dim oWin As frmMultiDisplay = CType(goUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
                                Dim bSendFormation As Boolean = False
                                If oWin Is Nothing = False AndAlso oWin.Visible = True Then
                                    If oWin.GetFormationID <> -1 Then
                                        bSendFormation = True
                                    End If
                                End If
                                If bSendFormation = False Then
                                    moMsgSys.SendMoveRequestMsg(CInt(vecLoc.X), CInt(vecLoc.Z), -1S, mbShiftKeyDown, False)
                                Else
                                    Dim iAngle As Int16 = 0S
                                    If moEngine.bInCtrlMove = True Then
                                        iAngle = CShort(LineAngleDegrees(moEngine.lMouseDownX, moEngine.lMouseDownZ, moEngine.lMouseDirX, moEngine.lMouseDirZ) * 10)
                                    End If
                                    moMsgSys.SendFormationMoveRequestMsg(moEngine.lMouseDownX, moEngine.lMouseDownZ, iAngle, mbShiftKeyDown)
                                End If

                            End If

                            mbRightDrag = False
                            mbRightDown = False
                            goCamera.ShowSelectionBox = False
                            'go ahead and break out
                            Return
                        End If
                    End If
                End If

                Dim pickRay As Ray = CalcPickingRay(lMouseX, lMouseY)

                Dim bFindIntOnTerrain As Boolean = False
                Dim vecInt As Vector3

                Dim lClickedItem As Int32
                Dim yRel As Byte

                Dim oEnvir As BaseEnvironment = goCurrentEnvir

                'Get our height at that location
                If glCurrentEnvirView = CurrentView.ePlanetView Then
                    If oEnvir Is Nothing = False Then
                        If oEnvir.oGeoObject Is Nothing = False Then
                            'ok, we must be displaying the planet...
                            bFindIntOnTerrain = True
                        End If
                    End If
                End If

                'Try for wormholes now
                If (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSuperSpecials) And 4) <> 0 Then
                    Dim oTmpSphere As BoundingSphere
                    For X As Int32 = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).WormholeUB
                        If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes Is Nothing = False Then
                            With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moWormholes(X)
                                Dim bRSLocked As Boolean = False
                                If .System1 Is Nothing = False AndAlso goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID = .System1.ObjectID Then
                                    oTmpSphere.SphereCenter.X = .LocX1
                                    oTmpSphere.SphereCenter.Z = .LocY1
                                    oTmpSphere.SphereCenter.Y = -200

                                    If .System2 Is Nothing = False Then
                                        If .System2.SystemType = SolarSystem.elSystemType.RespawnSystem Then
                                            bRSLocked = True
                                        End If
                                    End If
                                Else
                                    oTmpSphere.SphereCenter.X = .LocX2
                                    oTmpSphere.SphereCenter.Z = .LocY2
                                    oTmpSphere.SphereCenter.Y = -200

                                    If .System1 Is Nothing = False Then
                                        If .System1.SystemType = SolarSystem.elSystemType.RespawnSystem Then
                                            bRSLocked = True
                                        End If
                                    End If
                                End If
                                oTmpSphere.SphereRadius = 200

                                If RaySphereIntersectionTest(pickRay, oTmpSphere) = True Then
                                    If bRSLocked = True Then
                                        goUILib.AddNotification("Unable to jump to Locked Respawn systems.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                    Else
                                        Dim bGoodToSend As Boolean = True
                                        Try
                                            If goCurrentEnvir Is Nothing = False Then
                                                If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                                    If goCurrentEnvir.oGeoObject Is Nothing = False Then
                                                        If CType(goCurrentEnvir.oGeoObject, SolarSystem).SystemType = SolarSystem.elSystemType.RespawnSystem Then
                                                            Dim ofrmMsgBox As New frmMsgBox(goUILib, "Units that leave a Locked Respawn system cannot return until the system is no longer locked. Units may be stranded." & vbCrLf & vbCrLf & "Do you wish to proceed?", MsgBoxStyle.YesNo, "Confirm Order")
                                                            bGoodToSend = False

                                                            mlJumpFromRSResult_ID = .ObjectID
                                                            miJumpFromRSResult_TypeID = .ObjTypeID
                                                            ofrmMsgBox.Visible = True
                                                            AddHandler ofrmMsgBox.DialogClosed, AddressOf JumpFromRSResult
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Catch
                                        End Try

                                        If bGoodToSend = True Then moMsgSys.SendJumpTargetMsg(.ObjectID, .ObjTypeID)
                                    End If

                                    mbRightDrag = False
                                    mbRightDown = False
                                    goCamera.ShowSelectionBox = False
                                    Return
                                End If
                            End With
                        End If
                    Next X
                End If

                'Ok, if our view is the same as our current envir
                lClickedItem = -1
                If oEnvir Is Nothing = False Then
                    If (oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso glCurrentEnvirView = CurrentView.ePlanetView) OrElse _
                      (oEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso glCurrentEnvirView = CurrentView.eSystemView) Then
                        'Did we click an object?
                        lClickedItem = oEnvir.PickObject(pickRay, muSettings.FarClippingPlane, True)
                        If lClickedItem <> -1 Then
                            'Yes... ok, depends on what the entity is
                            Dim oEntity As BaseEntity = oEnvir.oEntity(lClickedItem)
                            If oEntity Is Nothing = False Then
                                If oEntity.OwnerID = goCurrentPlayer.ObjectID Then
                                    yRel = elRelTypes.eBloodBrother
                                Else
                                    yRel = goCurrentPlayer.GetPlayerRelScore(oEntity.OwnerID)
                                End If

                                'Now, check the rel
                                If yRel <= elRelTypes.eWar Then
                                    'War target, send a Set Primary Target msg
                                    moMsgSys.SendSetPrimaryTarget(CType(oEntity, Base_GUID), False)
                                ElseIf yRel < elRelTypes.ePeace Then
                                    'Neutral, just approach
                                    'TODO: Finish this, however for now, do nothing
                                    'moMsgSys.SendDockRequest(CType(goCurrentEnvir.oEntity(lClickedItem), Base_GUID))
                                ElseIf oEntity.OwnerID = glPlayerID Then
                                    'Either mine or an ally's, we attempt to Dock, cannot dock with moving target
                                    'If (goCurrentEnvir.oEntity(lClickedItem).CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                                    moMsgSys.SendDockRequest(CType(oEntity, Base_GUID))
                                    'End If
                                End If
                            End If


                        Else
                            'No, did we click a cache?
                            lClickedItem = oEnvir.PickCache(pickRay, muSettings.FarClippingPlane)
                            If lClickedItem <> -1 Then
                                Dim oCache As MineralCache = oEnvir.oCache(lClickedItem)
                                If oCache Is Nothing = False Then
                                    Dim iObjTypeID As Int16 = oCache.ObjTypeID
                                    lClickedItem = oCache.ObjectID
                                    'Yes, ok... Order the unit to Collect or Extract
                                    moMsgSys.SendSetMiningLoc(lClickedItem, iObjTypeID)
                                End If
                            End If
                        End If

                    End If
                End If

                'Did we not find something clicked on???
                If lClickedItem = -1 Then
                    'Nope, okay, then treat it as a move request or set rally point request
                    If bFindIntOnTerrain = False Then
                        If glCurrentEnvirView = CurrentView.ePlanetMapView Then
                            Dim ptPos As Point = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).GetMapClickedCoords(lMouseX, lMouseY)
                            If ptPos.X <> Int32.MinValue And ptPos.Y <> Int32.MinValue Then
                                vecInt = New Vector3(ptPos.X, 0, ptPos.Y)
                            End If
                        Else
                            vecInt = Plane.IntersectLine(moClickPlane, pickRay.Origin, Vector3.Add(pickRay.Origin, Vector3.Scale(pickRay.Direction, 20000)))

                            If glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                                vecInt.X *= 10000 : vecInt.Y *= 10000 : vecInt.Z *= 10000
                            ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                                vecInt.X *= 30 : vecInt.Y *= 30 : vecInt.Z *= 30
                            End If
                        End If
                    ElseIf pickRay.Direction.Y < 0 Then
                        'we only want to do this if the unit can move there, a unit can only move there if the click occurs on
                        ' the ground... a click can only occur on the ground if Y is negative
                        Try
                            Dim vecAddRay As Vector3 = pickRay.Direction

                            vecAddRay.Scale(Math.Abs(1 / vecAddRay.Y) * 5)

                            vecInt = pickRay.Origin

                            'This is an attempt to reduce our test count... we get the maximum height and move our ray that much
                            Dim lTempVar As Int32 = CInt((CType(oEnvir.oGeoObject, Planet).GetMaxHeight() - vecInt.Y) / pickRay.Direction.Y)
                            vecInt.Add(Vector3.Multiply(pickRay.Direction, lTempVar))

                            While vecInt.Y > 0
                                vecInt.Add(vecAddRay)

                                If CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(vecInt.X, vecInt.Z, True) > vecInt.Y Then
                                    'ok, that's it
                                    Exit While
                                End If
                            End While
                        Catch
                            goCamera.ShowSelectionBox = False
                            Return
                        End Try
                    End If

                    If mbSetRallyPoint = True Then
                        moMsgSys.SendRallyPointMsg(CInt(vecInt.X), CInt(vecInt.Z), -1S)
                        If goUILib Is Nothing = False Then goUILib.AddNotification("Rally point set.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        mbSetRallyPoint = False
                    Else
                        'Move request
                        Dim oWin As frmMultiDisplay = CType(goUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
                        Dim bSendFormation As Boolean = False
                        If oWin Is Nothing = False AndAlso oWin.Visible = True Then
                            If oWin.GetFormationID <> -1 Then
                                bSendFormation = True
                            End If
                        End If
                        If bSendFormation = False Then
                            moMsgSys.SendMoveRequestMsg(CInt(vecInt.X), CInt(vecInt.Z), -1S, mbShiftKeyDown, False)
                        Else
                            Dim iAngle As Int16 = 0S
                            If moEngine.bInCtrlMove = True Then
                                iAngle = CShort(LineAngleDegrees(moEngine.lMouseDownX, moEngine.lMouseDownZ, moEngine.lMouseDirX, moEngine.lMouseDirZ) * 10)
                            End If
                            If mbCtrlKeyDown = True Then
                                moMsgSys.SendFormationMoveRequestMsg(moEngine.lMouseDownX, moEngine.lMouseDownZ, iAngle, mbShiftKeyDown)
                            Else
                                moMsgSys.SendFormationMoveRequestMsg(CInt(vecInt.X), CInt(vecInt.Z), iAngle, mbShiftKeyDown)
                            End If

                        End If

                    End If

                End If

                vecInt = Nothing
                pickRay = Nothing
            End If
        Catch
        End Try

        mbRightDrag = False
        mbRightDown = False
    End Sub
    Private mlJumpFromRSResult_ID As Int32 = -1
    Private miJumpFromRSResult_TypeID As Int16 = -1
    Private Sub JumpFromRSResult(ByVal yRes As MsgBoxResult)
        If yRes = MsgBoxResult.Yes Then
            moMsgSys.SendJumpTargetMsg(mlJumpFromRSResult_ID, miJumpFromRSResult_TypeID)
        End If
    End Sub
    Private Sub SelectItemsInSelectionBox(ByVal lMouseX As Int32, ByVal lMouseY As Int32)
        'ok... mark this as the end
        goCamera.SelectionBoxEnd.X = lMouseX
        goCamera.SelectionBoxEnd.Y = lMouseY

        Dim p1 As System.Drawing.Point = goCamera.SelectionBoxStart
        Dim p2 As System.Drawing.Point = goCamera.SelectionBoxEnd
        Dim lTemp As Int32

        If p1.X > p2.X Then
            lTemp = p1.X : p1.X = p2.X : p2.X = lTemp
        End If
        If p1.Y > p2.Y Then
            lTemp = p1.Y : p1.Y = p2.Y : p2.Y = lTemp
        End If

        Dim rcSel As Rectangle = New Rectangle(p1.X, p1.Y, p2.X - p1.X, p2.Y - p1.Y)
        'Clear our selection
        If mbShiftKeyDown = False OrElse mbAltKeyDown = True Then
            goCurrentEnvir.DeselectAll()
        End If
        If HasAliasedRights(AliasingRights.eViewUnitsAndFacilities) = False Then Return
        'Now... figure out what's selected
        Try
            Dim ptTemp As Point
            Dim oRefDev As Device = GFXEngine.moDevice
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                Dim bCheckGuild As Boolean = goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 255
                Dim lCurUB As Int32 = -1
                If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
                'UIEnable:GroupBoxSelect<-1,-1,0>
                'if alt key isn't pressed, then select our units
                If mbAltKeyDown = False Then
                    For X As Int32 = 0 To lCurUB
                        If oEnvir.lEntityIdx(X) <> -1 Then
                            Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                            If oEntity Is Nothing = False AndAlso (oEntity.OwnerID = glPlayerID OrElse ((oEntity.CurrentStatus And elUnitStatus.eGuildAsset) <> 0) AndAlso bCheckGuild = True AndAlso goCurrentPlayer.oGuild.MemberInGuild(oEntity.OwnerID) = True) Then
                                'Ok, check if its within bounds
                                With oEntity
                                    ptTemp = .GetScreenPos(oRefDev)
                                    If ptTemp.IsEmpty = False AndAlso rcSel.Contains(ptTemp) Then

                                        If NewTutorialManager.TutorialOn = True Then
                                            Dim sParms() As String = {.ObjectID.ToString, .ObjTypeID.ToString, .yProductionType.ToString}
                                            If goUILib Is Nothing = False AndAlso goUILib.CommandAllowedWithParms(True, "GroupBoxSelect", sParms, False) = False Then Continue For
                                        End If

                                        .bSelected = True
                                        moMsgSys.RequestEntityHPUpdate(.ObjectID, .ObjTypeID)
                                    End If
                                End With
                            End If
                        End If
                    Next X
                End If
                'if we have no units selected, then select the first foreign unit in our box
                'alt key doesn't matter at this point
                If goUILib.bEntitiesSelected = False Then
                    For X As Int32 = 0 To lCurUB
                        If oEnvir.lEntityIdx(X) <> -1 Then
                            Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                            If oEntity Is Nothing = False AndAlso oEntity.OwnerID <> glPlayerID AndAlso oEntity.yVisibility = eVisibilityType.Visible Then
                                'Ok, check if its within bounds
                                With oEntity
                                    ptTemp = .GetScreenPos(oRefDev)
                                    If ptTemp.IsEmpty = False AndAlso rcSel.Contains(ptTemp) Then

                                        If NewTutorialManager.TutorialOn = True Then
                                            Dim sParms() As String = {.ObjectID.ToString, .ObjTypeID.ToString, .yProductionType.ToString}
                                            If goUILib Is Nothing = False AndAlso goUILib.CommandAllowedWithParms(True, "GroupBoxSelect", sParms, False) = False Then Continue For
                                        End If

                                        .bSelected = True
                                        moMsgSys.RequestEntityHPUpdate(.ObjectID, .ObjTypeID)
                                        Exit For 'we only need the first unit, we're done
                                    End If
                                End With
                            End If
                        End If
                    Next X
                End If
            End If
            oRefDev = Nothing
        Catch
        End Try
    End Sub
#End Region

    Private Function CalcPickingRay(ByVal lX As Int32, ByVal lY As Int32) As Ray
        Dim fX As Single
        Dim fY As Single
        'Dim vP As Viewport
        Dim matProj As Matrix
        Dim oRay As New Ray()

        Dim matView As Matrix

        Dim uParms As PresentParameters
        Dim fNear As Single

        Dim oDev As Device = GFXEngine.moDevice

        matProj = oDev.Transform.Projection
        matView = oDev.Transform.View

        uParms = oDev.PresentationParameters

        fX = (((2.0F * lX) / uParms.BackBufferWidth) - 1) / matProj.M11
        fY = -(((2.0F * lY) / uParms.BackBufferHeight) - 1) / matProj.M22

        matView.Invert()

        If glCurrentEnvirView = CurrentView.ePlanetView Then
            fNear = muSettings.TerrainNearClippingPlane
        Else
            fNear = muSettings.NearClippingPlane
        End If

        oRay.Direction.X = (fX * matView.M11) + (fY * matView.M21) + matView.M31
        oRay.Direction.Y = (fX * matView.M12) + (fY * matView.M22) + matView.M32
        oRay.Direction.Z = (fX * matView.M13) + (fY * matView.M23) + matView.M33
        oRay.Direction.Normalize()

        oRay.Origin.X = matView.M41
        oRay.Origin.Y = matView.M42
        oRay.Origin.Z = matView.M43

        'calc origin as intersection with near frustum
        oRay.Origin = Vector3.Add(oRay.Origin, Vector3.Multiply(oRay.Direction, fNear))

        Return oRay
    End Function

    Public Shared Function RaySphereIntersectionTest(ByVal oRay As Ray, ByVal oSphere As BoundingSphere) As Boolean
        Dim v As New Vector3()
        Dim b As Single
        Dim c As Single
        Dim d As Single
        Dim s0 As Single
        Dim s1 As Single

        v = Vector3.Subtract(oRay.Origin, oSphere.SphereCenter)
        b = 2.0F * Vector3.Dot(oRay.Direction, v)
        c = Vector3.Dot(v, v) - (oSphere.SphereRadius * oSphere.SphereRadius)
        d = (b * b) - (4.0F * c)
        If d < 0 Then Return False

        d = CSng(Math.Sqrt(d))
        s0 = (-b + d) / 2.0F
        s1 = (-b - d) / 2.0F

        v = Nothing

        If s0 >= 0.0F OrElse s1 >= 0.0F Then
            Return True
        End If
        Return False
    End Function

    '#Region " Point in Triangle "
    '    'TODO: should move this to the trig functions area in GlobalVars
    '    Private Function PointInTriangle(ByVal vecPoint As Vector3, ByVal vecA As Vector3, ByVal vecB As Vector3, ByVal vecC As Vector3) As Boolean
    '        If SameSide(vecPoint, vecA, vecB, vecC) AndAlso SameSide(vecPoint, vecB, vecA, vecC) AndAlso _
    '           SameSide(vecPoint, vecC, vecA, vecB) Then Return True
    '        Return False
    '    End Function

    '    Private Function SameSide(ByVal vecPt1 As Vector3, ByVal vecPt2 As Vector3, ByVal vecA As Vector3, ByVal vecB As Vector3) As Boolean
    '        Dim cp1 As Vector3 = Vector3.Cross(Vector3.Subtract(vecB, vecA), Vector3.Subtract(vecPt1, vecA))
    '        Dim cp2 As Vector3 = Vector3.Cross(Vector3.Subtract(vecB, vecA), Vector3.Subtract(vecPt2, vecA))
    '        If Vector3.Dot(cp1, cp2) >= 0 Then Return True Else Return False
    '    End Function
    '#End Region

    Private Sub DeleteSelectionResult(ByVal lResult As MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            'ok, dismantle items
            Dim lCnt As Int32 = 0
            Dim lUB As Int32 = 5
            Dim yMsg(lUB) As Byte
            Dim lPos As Int32 = 0

            System.BitConverter.GetBytes(GlobalMessageCode.eForcefulDismantle).CopyTo(yMsg, lPos) : lPos += 2
            lPos += 4   'for the cnt

            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                For X As Int32 = 0 To oEnvir.lEntityUB
                    If oEnvir.lEntityIdx(X) > -1 Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False Then
                            If oEntity.OwnerID = glPlayerID AndAlso oEntity.bSelected = True Then
                                lCnt += 1
                                lUB += 6
                                ReDim Preserve yMsg(lUB)
                                oEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                            End If
                        End If
                    End If
                Next X
            End If
            If lCnt < 1 Then Return
            System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, 2)

            moMsgSys.SendToPrimary(yMsg)
        End If
    End Sub

    Private mlCubeMapView As Int32 = 0

    Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
    mlLastUIEvent = glCurrentCycle
        Dim bFound As Boolean = False
        Dim X As Int32

        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return

        mbShiftKeyDown = e.Shift
        mbCtrlKeyDown = e.Control
        mbAltKeyDown = e.Alt

        moEngine.bRenderSelectedRanges = mbAltKeyDown
        If mbAltKeyDown = True Then
            DoDistanceDisplayUpdate(CalcPickingRay(mlMouseX, mlMouseY))
        End If

        'first check if the player is trying to skip the DSE Logo...
        If glCurrentEnvirView = CurrentView.eStartupDSELogo Then
            'ok, any key will cause us to go to the next scene...
            mbIgnoreNextKeyUp = True
            moEngine.mbBreakout = True 'If muSettings.bRanBefore = True Then 
            Return
        ElseIf glCurrentEnvirView = CurrentView.eStartupLogin Then
            Dim frmLogin As frmLoginDlg = CType(goUILib.GetWindow("frmLoginDlg"), frmLoginDlg)
            If frmLogin Is Nothing = False AndAlso frmLogin.Visible = False Then
                goUILib.RetainTooltip = False
                goUILib.SetToolTip(False)
                frmLogin.Visible = True
                mbIgnoreNextKeyUp = True
                Return
            End If
        End If

        If e.Control = True Then
            If e.KeyCode = Keys.C Then
                'copy
                Try
                    If goUILib Is Nothing = False Then
                        If goUILib.FocusedControl Is Nothing = False Then
                            If TypeOf goUILib.FocusedControl Is UITextBox = True Then
                                Dim sVal As String = CType(goUILib.FocusedControl, UITextBox).SelectedText
                                My.Computer.Clipboard.Clear()
                                If sVal Is Nothing = False AndAlso sVal <> "" Then My.Computer.Clipboard.SetText(sVal)
                            End If
                        End If
                    End If
                Catch
                End Try
                mbIgnoreNextKeyUp = True
                Return
            ElseIf e.KeyCode = Keys.V Then
                'paste
                Try
                    If My.Computer.Clipboard.ContainsText() = True Then
                        If goUILib Is Nothing = False Then
                            If goUILib.FocusedControl Is Nothing = False Then
                                If TypeOf goUILib.FocusedControl Is UITextBox = True Then
                                    Dim sVal As String = My.Computer.Clipboard.GetText()
                                    CType(goUILib.FocusedControl, UITextBox).PasteText(sVal)
                                End If
                            End If
                        End If
                    End If
                Catch
                End Try
                mbIgnoreNextKeyUp = True
                Return
            ElseIf e.KeyCode = Keys.X Then
                    'cut
                    If goUILib Is Nothing = False Then
                        If goUILib.FocusedControl Is Nothing = False Then
                            If TypeOf goUILib.FocusedControl Is UITextBox = True Then
                                Dim sVal As String = CType(goUILib.FocusedControl, UITextBox).CutText()
                                My.Computer.Clipboard.Clear()
                                If sVal Is Nothing = False AndAlso sVal <> "" Then My.Computer.Clipboard.SetText(sVal)
                            End If
                        End If
                    End If
                    mbIgnoreNextKeyUp = True
                    Return
            End If


            If Debugger.IsAttached = True Then
                If e.KeyCode = Keys.K Then
                    'Dim ofrm As New frmSubmitBug(goUILib)
                    'ofrm.Visible = True
                    'ofrm.FullScreen = True
                    'Dim lBombX As Int32 = 0 '+ CInt((Rnd() * 20000) - 10000)
                    'Dim lBombZ As Int32 = 0 '+ CInt((Rnd() * 20000) - 10000)
                    'Dim lHt As Int32 = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(lBombX, lBombZ, True))
                    'Dim yType As Byte = CByte(WeaponType.eBomb_Green + CInt(Rnd() * 5))

                    'BombMgr.goBombMgr.AddNewBomb(lBombX + CInt((Rnd() * 6000) - 3000), lHt + 7000, lBombZ + CInt((Rnd() * 6000) - 3000), lBombX, lHt, lBombZ, yType, 10) ' CByte(CInt(Rnd() * 200) + 50))

                    'mlCubeMapView += 1
                    'Select Case mlCubeMapView
                    '    Case 1
                    '        goCamera.mlCameraX = goCamera.mlCameraAtX
                    '        goCamera.mlCameraY = goCamera.mlCameraAtY
                    '        goCamera.mlCameraZ = goCamera.mlCameraAtZ - 1000

                    '    Case 2
                    '        goCamera.mlCameraX = goCamera.mlCameraAtX - 1000
                    '        goCamera.mlCameraY = goCamera.mlCameraAtY
                    '        goCamera.mlCameraZ = goCamera.mlCameraAtZ
                    '    Case 3
                    '        goCamera.mlCameraX = goCamera.mlCameraAtX
                    '        goCamera.mlCameraY = goCamera.mlCameraAtY
                    '        goCamera.mlCameraZ = goCamera.mlCameraAtZ + 1000
                    '    Case 4
                    '        goCamera.mlCameraX = goCamera.mlCameraAtX + 1000
                    '        goCamera.mlCameraY = goCamera.mlCameraAtY
                    '        goCamera.mlCameraZ = goCamera.mlCameraAtZ
                    '    Case 5
                    '        goCamera.mlCameraX = goCamera.mlCameraAtX + 1
                    '        goCamera.mlCameraY = goCamera.mlCameraAtY - 1000
                    '        goCamera.mlCameraZ = goCamera.mlCameraAtZ
                    '    Case 6
                    '        goCamera.mlCameraX = goCamera.mlCameraAtX - 1
                    '        goCamera.mlCameraY = goCamera.mlCameraAtY + 1000
                    '        goCamera.mlCameraZ = goCamera.mlCameraAtZ
                    '    Case Else
                    '        goUILib.AddNotification("Done", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    '        e.Handled = True
                    '        Return
                    'End Select
                    'GFXEngine.mbCaptureScreenshot = True
                    'Dim ofrm As New frmTransportManagement(goUILib)
                    'ofrm.Visible = True
                    'ofrm.FullScreen = True
                    'CType(goCurrentEnvir.oGeoObject, Planet).SaveTexToFile()
                    'goUILib.AddNotification("Done", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Dim ofrm As New frmSelectRespawn(goUILib)
                    ofrm.Visible = True
                    ofrm.DoTest()

                    e.Handled = True
                End If
            End If
            


            'Movie-camera modes

            'If e.KeyCode = Keys.K Then
            '    SolarSystem.SaveCosmoSphere()
            'End If

        End If

        'Ok, first, check special characters
        If e.KeyCode = Keys.F1 Then
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF1Key, 0, 0)
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F1") = False Then Return
                End If
            End If

            Dim ofrm As frmTutorialTOC = CType(goUILib.GetWindow("frmTutorialTOC"), frmTutorialTOC)
            If ofrm Is Nothing = False Then
                If ofrm.Visible = True Then
                    goUILib.RemoveWindow(ofrm.ControlName)
                Else : ofrm.Visible = True
                End If
            Else
                ofrm = New frmTutorialTOC(goUILib)
                ofrm.Visible = True
            End If
            ofrm = Nothing
            e.Handled = True
        ElseIf e.KeyCode = Keys.F2 Then

            If goCurrentPlayer Is Nothing Then Return
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF2Key, 0, 0)

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F2") = False Then Return
                End If
            End If

            Dim ofrm As frmEmailMain = CType(goUILib.GetWindow("frmEmailMain"), frmEmailMain)
            If ofrm Is Nothing = False Then
                If ofrm.Visible = True Then
                    goUILib.RemoveWindow(ofrm.ControlName)
                Else : ofrm.Visible = True
                End If
            Else
                ofrm = New frmEmailMain(goUILib)
                ofrm.Visible = True
            End If
            ofrm = Nothing
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            e.Handled = True
        ElseIf e.KeyCode = Keys.F3 Then

            If goCurrentPlayer Is Nothing Then Return
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF3Key, 0, 0)

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F3") = False Then Return
                End If
            End If

            Dim ofrm As frmFleet = CType(goUILib.GetWindow("frmFleet"), frmFleet)
            If ofrm Is Nothing = False Then
                If ofrm.Visible = True Then
                    goUILib.RemoveWindow(ofrm.ControlName)
                Else : ofrm.Visible = True
                End If
            Else
                ofrm = New frmFleet(goUILib, -1)
                ofrm.Visible = True
            End If
            ofrm = Nothing
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            e.Handled = True
        ElseIf e.KeyCode = Keys.F4 Then

            If goCurrentPlayer Is Nothing Then Return

            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF4Key, 0, 0)

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F4") = False Then Return
                End If
            End If

            Dim oTmpWin As UIWindow
            If goUILib Is Nothing = False Then
                oTmpWin = goUILib.GetWindow("frmTradeMain")
                If oTmpWin Is Nothing = False Then
                    CType(oTmpWin, frmTradeMain).ReturnToPreviousViewAndReleaseBackground()
                    goUILib.RemoveWindow(oTmpWin.ControlName)
                Else
                    If glCurrentEnvirView > CurrentView.eFullScreenInterface Then Return
                    oTmpWin = New frmTradeMain(goUILib, -1)
                End If
                oTmpWin = Nothing
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                e.Handled = True
            End If

        ElseIf e.KeyCode = Keys.F5 Then

            If goCurrentPlayer Is Nothing Then Return
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF5Key, 0, 0)

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F5") = False Then Return
                End If
            End If

            Dim oTmpWin As UIWindow
            If goUILib Is Nothing = False Then
                oTmpWin = goUILib.GetWindow("frmDiplomacy")
                If oTmpWin Is Nothing = False Then
                    goUILib.RemoveWindow(oTmpWin.ControlName)
                    If glCurrentEnvirView = CurrentView.eDiplomacyScreen Then ReturnToPreviousView()
                Else
                    oTmpWin = New frmDiplomacy(goUILib)
                    CType(oTmpWin, frmDiplomacy).SetFromCurrentPlayer()
                End If
                oTmpWin = Nothing
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                e.Handled = True
            End If
        ElseIf e.KeyCode = Keys.F6 AndAlso glCurrentEnvirView < CurrentView.eFullScreenInterface Then

            If goCurrentPlayer Is Nothing Then Return
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF6Key, 0, 0)

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F6") = False Then Return
                End If
            End If

            Dim oCStatWin As UIWindow
            Dim sTemp As String = "frmColonyStatsSmall"
            If goUILib Is Nothing = False Then
                If muSettings.ExpandedColonyStatsScreen = True Then sTemp = "frmColonyStats"
                oCStatWin = goUILib.GetWindow(sTemp)
                If oCStatWin Is Nothing = False Then
                    If oCStatWin.Visible = True Then
                        oCStatWin.Visible = False
                        goUILib.RemoveWindow(oCStatWin.ControlName)
                    Else : oCStatWin.Visible = True
                    End If
                ElseIf muSettings.ExpandedColonyStatsScreen = True Then
                    oCStatWin = New frmColonyStats(goUILib)
                Else
                    oCStatWin = New frmColonyStatsSmall(goUILib)
                End If
                oCStatWin = Nothing
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                e.Handled = True
            End If
        ElseIf e.KeyCode = Keys.F7 Then

            If goCurrentPlayer Is Nothing Then Return
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF7Key, 0, 0)

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F7") = False Then Return
                End If
            End If

            Dim oTmpWin As UIWindow
            If goUILib Is Nothing = False Then
                oTmpWin = goUILib.GetWindow("frmBudget")
                If oTmpWin Is Nothing = False Then
                    goUILib.RemoveWindow(oTmpWin.ControlName)
                Else : oTmpWin = New frmBudget(goUILib)
                End If
                oTmpWin = Nothing
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                e.Handled = True
            End If
        ElseIf e.KeyCode = Keys.F8 AndAlso glCurrentEnvirView < CurrentView.eFullScreenInterface Then

            If goCurrentPlayer Is Nothing Then Return
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF8Key, 0, 0)

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F8") = False Then Return
                End If
            End If

            Dim ofrm As frmMining = CType(goUILib.GetWindow("frmMining"), frmMining)
            If ofrm Is Nothing = False Then
                If ofrm.Visible = True Then
                    goUILib.RemoveWindow(ofrm.ControlName)
                Else : ofrm.Visible = True
                End If
            Else
                ofrm = New frmMining(goUILib)
                ofrm.Visible = True
            End If
            ofrm = Nothing
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            e.Handled = True
        ElseIf e.KeyCode = Keys.F9 Then

            If goCurrentPlayer Is Nothing Then Return
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF9Key, 0, 0)

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F9") = False Then Return
                End If
            End If

            Dim ofrm As frmAgentMain = CType(goUILib.GetWindow("frmAgentMain"), frmAgentMain)
            If ofrm Is Nothing = False Then
                If ofrm.Visible = True Then
                    'goUILib.RemoveWindow(ofrm.ControlName)
                    ofrm.CloseForm()
                Else : ofrm.Visible = True
                End If
            Else
                ofrm = New frmAgentMain(goUILib)
                ofrm.Visible = True
            End If
            ofrm = Nothing
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            e.Handled = True
        ElseIf e.KeyCode = Keys.F10 Then

            If goCurrentPlayer Is Nothing Then Return
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF10Key, 0, 0)

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "F10") = False Then Return
                End If
            End If

            Dim ofrm As frmFormations = CType(goUILib.GetWindow("frmFormations"), frmFormations)
            If ofrm Is Nothing = False Then
                If ofrm.Visible = True Then
                    goUILib.RemoveWindow(ofrm.ControlName)
                Else : ofrm.Visible = True
                End If
            Else
                ofrm = New frmFormations(goUILib)
                ofrm.Visible = True
            End If
            ofrm = Nothing
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            e.Handled = True
        ElseIf e.KeyCode = Keys.F11 Then
            OpenWindow_Guild()
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            'If goCurrentPlayer Is Nothing Then Return
            'BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF11Key, 0, 0)
            'If NewTutorialManager.TutorialOn = True Then
            '    If goUILib Is Nothing = False Then
            '        If goUILib.CommandAllowed(False, "F11") = False Then Return
            '    End If
            'End If
            ''Dim ofrm As New frmChannels(goUILib)
            ''ofrm.Visible = True
            ''ofrm = Nothing
            ''For X = 0 To goCurrentEnvir.lEntityUB
            ''	If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
            ''		goEntityDeath.AddNewDeathSequence(goCurrentEnvir.oEntity(X), X)
            ''	End If
            ''Next
            ''Planet.bUseNewFogDistance = Not Planet.bUseNewFogDistance
            ''Dim ofrm As New CalenderTest(goUILib)
            ''ofrm.Visible = True

            ''Dim yTestMsg(5) As Byte
            ''Static xlPirateWave As Int32 = 0
            ''xlPirateWave += 1
            ''System.BitConverter.GetBytes(GlobalMessageCode.ePirateWaveSpawn).CopyTo(yTestMsg, 0)
            ''System.BitConverter.GetBytes(xlPirateWave).CopyTo(yTestMsg, 2)
            ''moMsgSys.HandlePirateWaveSpawn(yTestMsg)


            ''Dim oFrm As frmSenate = New frmSenate(goUILib)
            ''oFrm.Visible = True
            'If goCurrentPlayer Is Nothing = False Then
            '    If goCurrentPlayer.oGuild Is Nothing Then



            '        Dim ofrm As frmGuildSearch = CType(goUILib.GetWindow("frmGuildSearch"), frmGuildSearch)
            '        If ofrm Is Nothing = False Then
            '            If ofrm.Visible = True Then
            '                goUILib.RemoveWindow(ofrm.ControlName)
            '            Else : ofrm.Visible = True
            '            End If
            '        Else
            '            ofrm = New frmGuildSearch(goUILib)
            '            ofrm.Visible = True
            '        End If
            '        ofrm = Nothing
            '    Else
            '        Dim ofrm As frmGuildMain = CType(goUILib.GetWindow("frmGuildMain"), frmGuildMain)
            '        If ofrm Is Nothing = False Then
            '            If ofrm.Visible = True Then
            '                goUILib.RemoveWindow(ofrm.ControlName)
            '            Else : ofrm.Visible = True
            '            End If
            '        Else
            '            ofrm = New frmGuildMain(goUILib)
            '            ofrm.Visible = True
            '        End If
            '        ofrm = Nothing
            '    End If
            'End If


            e.Handled = True
        ElseIf e.KeyCode = Keys.C AndAlso e.Alt = True Then
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eCaptureScreenshot, 0, 0)
            moEngine.CaptureScreenshot()
            e.Handled = True
        ElseIf e.KeyCode = Keys.F12 Then
            Try
                BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eF12Key, 0, 0)

                'REPLACE with Store Site if player is not purchased
                If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 0 AndAlso goCurrentPlayer.lAccountStatus <> AccountStatusType.eActiveAccount AndAlso goCurrentPlayer.lAccountStatus <> AccountStatusType.eMondelisActive Then
                    Process.Start("https://www.beyondprotocol.com/shop~~~/flexicart/_modules/users/F12login.php")
                ElseIf goCurrentPlayer Is Nothing = False Then
                    'Process.Start("http://www.beyondprotocol.com/issuetracker")
                    If goUILib Is Nothing = False Then
                        Dim ofrm As New frmNewEmail(goUILib)
                        ofrm.SetToText("Support")
                        ofrm = Nothing
                    End If
                ElseIf goUILib Is Nothing = False Then
                    Dim oFrm As frmTutorialTOC = CType(goUILib.GetWindow("frmTutorialTOC"), frmTutorialTOC)
                    If oFrm Is Nothing Then oFrm = New frmTutorialTOC(goUILib)
                    oFrm.Visible = True
                    oFrm.ShowForTrigger(TutorialAlert.eTutorialTriggerItem.GettingTechSupport)
                End If
            Catch
            End Try
            'Dim ofrmBug As frmBugMain = New frmBugMain(goUILib)
            'ofrmBug.Visible = True
            'ofrmBug = Nothing
            e.Handled = True
        ElseIf e.KeyCode = Keys.N AndAlso e.Control = True AndAlso glCurrentEnvirView < CurrentView.eFullScreenInterface Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "CtrlN") = False Then Return
                End If
            End If
            If e.Shift = True Then
                SelectAndGotoNextUnit(eSelectNextType.eAnyType Or eSelectNextType.ePreviousBitShift)
            Else
                SelectAndGotoNextUnit(eSelectNextType.eAnyType)
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.E AndAlso e.Control = True Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "CtrlE") = False Then Return
                End If
            End If
            Dim yType As eSelectNextType = eSelectNextType.eEngineer
            If e.Shift = True Then yType = yType Or eSelectNextType.ePreviousBitShift
            If SelectAndGotoNextUnit(yType) > -1 Then
                If NewTutorialManager.TutorialOn = True Then
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eControlEToSelectEngineer, -1, -1, -1, "")
                End If
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.F AndAlso e.Control = True Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "CtrlF") = False Then Return
                End If
            End If
            Dim yType As eSelectNextType = eSelectNextType.eFacility
            If e.Shift = True Then yType = yType Or eSelectNextType.ePreviousBitShift
            SelectAndGotoNextUnit(yType)
            e.Handled = True
        ElseIf e.KeyCode = Keys.I AndAlso e.Control = True Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "CtrlI") = False Then Return
                End If
            End If
            Dim yType As eSelectNextType = eSelectNextType.eIdleEngineer
            If e.Shift = True Then yType = yType Or eSelectNextType.ePreviousBitShift
            Dim lResult As Int32 = SelectAndGotoNextUnit(yType)
            If lResult > -1 AndAlso NewTutorialManager.TutorialOn = True Then
                Try
                    Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(lResult)
                    If oEntity Is Nothing = False Then
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eControlISelect, oEntity.ObjectID, oEntity.ObjTypeID, oEntity.yProductionType, "")
                    End If
                Catch
                End Try
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.M AndAlso e.Control = True Then
            'Ok, toggle rendering mineral caches
            muSettings.RenderMineralCaches = Not muSettings.RenderMineralCaches
            If muSettings.RenderMineralCaches = False Then
                Dim oTmp As frmCacheSelect = CType(goUILib.GetWindow("frmCacheSelect"), frmCacheSelect)
                If Not oTmp Is Nothing Then
                    goUILib.RemoveWindow(oTmp.ControlName)
                    oTmp.Visible = False
                    oTmp = Nothing
                End If
            End If
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eToggleMineralCaches, 0, 0)
            e.Handled = True
        ElseIf e.KeyCode = Keys.P AndAlso e.Control = True Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "CtrlP") = False Then Return
                End If
            End If
            Dim yType As eSelectNextType = eSelectNextType.eUnpoweredFacility
            If e.Shift = True Then yType = yType Or eSelectNextType.ePreviousBitShift
            SelectAndGotoNextUnit(yType)
            e.Handled = True
        ElseIf e.KeyCode = Keys.Q AndAlso e.Control = True Then
            If muSettings.CtrlQExits = True Then
                BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eCtrlQ, 0, 0)
                Me.Close()
                e.Handled = True
            End If
        ElseIf e.KeyCode = Keys.R AndAlso e.Control = True Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "CtrlR") = False Then Return
                End If
            End If
            mbSetRallyPoint = Not mbSetRallyPoint

            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eToggleRallyPoint, 0, 0)

            If goUILib Is Nothing = False Then
                If mbSetRallyPoint = True Then
                    goUILib.AddNotification("Select Rally Point destination...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Else : goUILib.AddNotification("Select Rally Point Cancelled.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.S AndAlso e.Control = True AndAlso e.Shift = True Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "CtrlS") = False Then Return
                End If
            End If
            SelectAllSimilarNamedUnits(False)
        ElseIf e.KeyCode = Keys.S AndAlso e.Control = True Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "CtrlS") = False Then Return
                End If
            End If
            SelectAllSimilarUnits(False)
            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eControlSPressed, -1, -1, -1, "")
            End If
        ElseIf e.KeyCode = Keys.U AndAlso e.Control = True Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "CtrlU") = False Then Return
                End If
            End If
            Dim yType As eSelectNextType = eSelectNextType.eUnit
            If e.Shift = True Then yType = yType Or eSelectNextType.ePreviousBitShift
            Dim lResult As Int32 = SelectAndGotoNextUnit(yType)
            If lResult > -1 Then
                If NewTutorialManager.TutorialOn = True Then
                    Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(lResult)
                    If oEntity Is Nothing = False Then
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eControlUSelection, oEntity.ObjectID, oEntity.ObjTypeID, oEntity.yProductionType, "")
                    End If
                End If
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.T AndAlso e.Control = True Then
            If NewTutorialManager.TutorialOn = True Then Return

            If e.Shift = True Then
                moMsgSys.SendTetherPoint(0)
            Else
                moMsgSys.SendTetherPoint(1)
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.Left AndAlso e.Control = True AndAlso mbChangingEnvirs = False Then
            Dim ofrmED As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
            If ofrmED Is Nothing Then ofrmED = New frmEnvirDisplay(goUILib)
            If Not ofrmED Is Nothing Then
                ofrmED.btnGoLeft_Click(ofrmED.btnGoLeft.ControlName)
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.Right AndAlso e.Control = True AndAlso mbChangingEnvirs = False Then
            Dim ofrmED As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
            If ofrmED Is Nothing Then ofrmED = New frmEnvirDisplay(goUILib)
            If Not ofrmED Is Nothing Then
                ofrmED.btnGoRight_Click(ofrmED.btnGoRight.ControlName)
            End If
            e.Handled = True
        ElseIf e.KeyCode = Keys.Enter AndAlso e.Control = True AndAlso mbChangingEnvirs = False Then
            Dim ofrmED As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
            If ofrmED Is Nothing Then ofrmED = New frmEnvirDisplay(goUILib)
            If Not ofrmED Is Nothing Then
                ofrmED.GotoEnvironmentWrapper_local()
            End If
            e.Handled = True
        End If

        'BREAK OUT POINT!!!!
        If e.Handled = True Then Return

        'Ok, no for the goUILib stuff... before posting to the UILib, check if the BuildGhost is being used
        If goUILib Is Nothing = False Then
            If goUILib.BuildGhost Is Nothing = False Then
                If e.KeyCode = Keys.Left Then
                    goUILib.BuildGhostAngle += 10S
                    If goUILib.BuildGhostAngle >= 3600S Then
                        goUILib.BuildGhostAngle -= 3600S
                    End If
                    Return
                ElseIf e.KeyCode = Keys.Right Then
                    goUILib.BuildGhostAngle -= 10S
                    If goUILib.BuildGhostAngle < 0S Then
                        goUILib.BuildGhostAngle += 3600S
                    End If
                    Return
                ElseIf e.KeyCode = Keys.Escape Then
                    goUILib.BuildGhost = Nothing
                    goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                    e.Handled = True
                    Return
                End If
            End If

            'Ok, we're here, so check the postmessage
            If goUILib.PostMessage(UILibMsgCode.eKeyDownCode, e) = True Then Return
        End If

        'Ok, if we are here then e.Handled = False and goUILIB didn't care about the message...
        If e.KeyCode = Keys.Enter AndAlso e.Alt = True Then
            GFXEngine.ToggleFullScreen() 'GFXEngine.ToggleFullScreen()
            e.Handled = True
            Return
        ElseIf e.KeyCode = Keys.Enter OrElse e.KeyCode = Keys.OemQuestion Then      'Enter and / are special characters as the chat window can do something with them but someone could be typing...
            If goUILib Is Nothing = False Then
                Dim ofrmChat As UIWindow = goUILib.GetWindow("frmChat")
                If ofrmChat Is Nothing = False Then
                    For lCtrlIdx As Int32 = 0 To ofrmChat.ChildrenUB
                        If ofrmChat.moChildren(lCtrlIdx).ControlName = "txtNew" Then
                            If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
                            goUILib.FocusedControl = ofrmChat.moChildren(lCtrlIdx)
                            ofrmChat.moChildren(lCtrlIdx).HasFocus = True
                            e.Handled = True
                            Exit For
                        End If
                    Next
                End If
            End If
        ElseIf e.KeyCode = Keys.Escape Then         'escape is a special character as the UILib may do something with it
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eEscapeKey, 0, 0)

            Dim ofrm As frmOptions = CType(goUILib.GetWindow("frmOptions"), frmOptions)
            If ofrm Is Nothing = False Then
                If ofrm.Visible = True Then
                    goUILib.RemoveWindow(ofrm.ControlName)
                Else : ofrm.Visible = True
                End If
            Else
                ofrm = New frmOptions(goUILib)
                ofrm.Visible = True
            End If
            ofrm = Nothing
            e.Handled = True
        ElseIf e.KeyCode = Keys.Oemtilde Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "`") = False Then Return
                End If
            End If
            If goUILib.yRenderUI = 0 Then goUILib.yRenderUI = 255 Else goUILib.yRenderUI = 0
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eTilda, goUILib.yRenderUI, 0)
        ElseIf e.KeyCode = Keys.Z Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(True, "ZoomCamera") = False Then Return
                End If
            End If
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eZKeyToZoom, 0, 0)
            mbMouseWheelOverride = True
        ElseIf e.KeyCode = Keys.Tab Then
            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(False, "Tab") = False Then Return
                End If
            End If

            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eTabKey, 0, 0)

            Dim ofrm As frmNoteHistory = CType(goUILib.GetWindow("frmNoteHistory"), frmNoteHistory)
            If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                goUILib.RemoveWindow(ofrm.ControlName)
                ofrm = Nothing
            Else
                ofrm = Nothing
                ofrm = New frmNoteHistory(goUILib)
                ofrm.Visible = True
            End If
            e.Handled = True
        Else
            If glCurrentEnvirView < CurrentView.eFullScreenInterface Then
                If e.KeyCode = Keys.Delete Then
                    If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then
                        Dim lFound As Int32 = 0
                        For lIdx As Int32 = 0 To goCurrentEnvir.lEntityUB
                            If goCurrentEnvir.lEntityIdx(lIdx) <> -1 AndAlso goCurrentEnvir.oEntity(lIdx).bSelected = True Then
                                lFound = 1
                                If goCurrentEnvir.oEntity(lIdx).OwnerID <> glPlayerID Then
                                    lFound = -1
                                    Exit For
                                End If
                            End If
                        Next lIdx
                        If lFound = 1 Then
                            Dim oFrm As New frmMsgBox(goUILib, "This will dismantle the currently selected items. Are you sure you wish to do so?", MsgBoxStyle.YesNo, "Confirm Dismantle")
                            oFrm.Visible = True
                            AddHandler oFrm.DialogClosed, AddressOf DeleteSelectionResult
                            e.Handled = True
                        End If
                    End If
                ElseIf e.KeyCode = Keys.Back AndAlso glCurrentEnvirView < CurrentView.eFullScreenInterface Then
                    GoBackAView()
                ElseIf e.KeyCode = Keys.Home Then
                    If NewTutorialManager.TutorialOn = True Then
                        If goUILib Is Nothing = False Then
                            If goUILib.CommandAllowed(False, "Home") = False Then Return
                        End If
                    End If

                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eHomeKey, 0, 0)

                    Dim Y As Int32

                    If goCurrentEnvir Is Nothing = False AndAlso goGalaxy Is Nothing = False Then
                        If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                            'ok, gonna be a bit rough...
                            For X = 0 To goGalaxy.mlSystemUB
                                For Y = 0 To goGalaxy.moSystems(X).PlanetUB
                                    If goCurrentEnvir.ObjectID = goGalaxy.moSystems(X).moPlanets(Y).ObjectID Then
                                        goGalaxy.CurrentSystemIdx = X
                                        goGalaxy.moSystems(X).CurrentPlanetIdx = Y
                                        bFound = True
                                        Exit For
                                    End If
                                Next Y
                                If bFound = True Then Exit For
                            Next X
                            glCurrentEnvirView = CurrentView.ePlanetMapView
                        ElseIf goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                            'easier...
                            For X = 0 To goGalaxy.mlSystemUB
                                If goCurrentEnvir.ObjectID = goGalaxy.moSystems(X).ObjectID Then
                                    goGalaxy.CurrentSystemIdx = X
                                    Exit For
                                End If
                            Next X
                            glCurrentEnvirView = CurrentView.eSystemMapView1
                            With goCamera
                                .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                                .mlCameraX = 0 : .mlCameraY = 1000 : .mlCameraZ = -1000
                            End With
                        End If
                    End If

                    If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
                        goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.HomeKeyPressed)
                    End If
                ElseIf e.KeyCode = Keys.End Then
                    If NewTutorialManager.TutorialOn = True Then
                        If goUILib Is Nothing = False Then
                            If goUILib.CommandAllowed(False, "End") = False Then Return
                        End If
                    End If
                    If goCurrentEnvir Is Nothing = False AndAlso goGalaxy Is Nothing = False Then
                        If glCurrentEnvirView = CurrentView.ePlanetView Then
                            With goCamera
                                Dim lDist As Int32 = CInt(Distance(.mlCameraAtX, .mlCameraAtZ, .mlCameraX, .mlCameraZ))
                                .mlCameraZ = .mlCameraAtZ + lDist
                                .mlCameraX = .mlCameraAtX
                                muSettings.lPlanetViewCameraX = 0
                                muSettings.lPlanetViewCameraZ = lDist
                            End With
                        ElseIf glCurrentEnvirView = CurrentView.eSystemView Then
                            With goCamera
                                Dim lDist As Int32 = CInt(Distance(.mlCameraAtX, .mlCameraAtZ, .mlCameraX, .mlCameraZ))
                                .mlCameraX = .mlCameraAtX
                                .mlCameraZ = .mlCameraAtZ - lDist
                            End With
                        End If
                    End If
                ElseIf e.KeyCode = Keys.A AndAlso goFullScreenBackground Is Nothing = True Then
                    If goUILib Is Nothing = False AndAlso goUILib.FocusedControl Is Nothing = True Then
                        Dim ofrm As frmAdvanceDisplay = CType(goUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
                        If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                            If ofrm.HandleAutoLaunchButtonClick() = True Then
                                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                            End If
                        End If
                        ofrm = Nothing
                    End If
                ElseIf e.KeyCode = Keys.C AndAlso goFullScreenBackground Is Nothing = True Then
                    'Ok, get our advance display
                    If NewTutorialManager.TutorialOn = True Then
                        If goUILib Is Nothing = False Then
                            Dim sParms() As String = {"Contents"}
                            If goUILib.CommandAllowedWithParms(True, "frmAdvanceDisplay", sParms, False) = False Then Return
                        End If
                    End If

                    If goUILib Is Nothing = False Then
                        Dim ofrm As frmAdvanceDisplay = CType(goUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
                        If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                            ofrm.OpenContentsWindow()
                        End If
                        ofrm = Nothing
                    End If
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eCKeyToOpenContents, 0, 0)
                ElseIf e.KeyCode = Keys.O Then
                    If NewTutorialManager.TutorialOn = True Then
                        If goUILib Is Nothing = False Then
                            Dim sParms() As String = {"Orders"}
                            If goUILib.CommandAllowedWithParms(True, "frmAdvanceDisplay", sParms, False) = False Then Return
                        End If
                    End If

                    If goUILib Is Nothing = False Then
                        Dim ofrm As frmAdvanceDisplay = CType(goUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
                        If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                            ofrm.OpenOrdersWindows()
                        Else
                            Dim ofrmBehavior As frmBehavior = CType(goUILib.GetWindow("frmBehavior"), frmBehavior)
                            If ofrmBehavior Is Nothing = False AndAlso ofrmBehavior.Visible = True Then
                                ofrmBehavior.btnMinMax_Click(Nothing)
                            End If
                            ofrmBehavior = Nothing
                        End If
                        ofrm = Nothing
                    End If
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eOKeyToOpenOrders, 0, 0)
                ElseIf e.KeyCode = Keys.P Then
                    If goUILib Is Nothing = False AndAlso goUILib.FocusedControl Is Nothing = True Then
                        Dim ofrm As frmAdvanceDisplay = CType(goUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
                        If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                            If ofrm.HandlePowerButtonClick() = True Then
                                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                            End If
                        End If
                        ofrm = Nothing
                    End If
                ElseIf e.KeyCode = Keys.R Then
                    If goUILib Is Nothing = False Then
                        Dim ofrmRouteConfig As frmRouteConfig = CType(goUILib.GetWindow("frmRouteConfig"), frmRouteConfig)
                        If ofrmRouteConfig Is Nothing = False AndAlso ofrmRouteConfig.Visible = True Then
                            goUILib.RemoveWindow("frmRouteConfig")
                            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                        Else
                            Dim ofrmBehavior As frmBehavior = CType(goUILib.GetWindow("frmBehavior"), frmBehavior)
                            If ofrmBehavior Is Nothing = False AndAlso ofrmBehavior.Visible = True Then
                                ofrmBehavior.OpenRouteConfig()
                                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                            End If
                            ofrmBehavior = Nothing
                        End If
                        ofrmRouteConfig = Nothing
                    End If
                ElseIf e.KeyCode = Keys.U Then
                    If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso NewTutorialManager.TutorialOn = False Then
                        For XX As Int32 = 0 To goCurrentEnvir.lEntityUB
                            If goCurrentEnvir.lEntityIdx(XX) <> -1 AndAlso goCurrentEnvir.oEntity(XX).bSelected = True AndAlso goCurrentEnvir.oEntity(XX).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(XX).ObjTypeID = ObjectType.eUnit AndAlso goCurrentEnvir.oEntity(XX).oMesh.bLandBased = False Then

                                Dim lDX As Int32
                                Dim lDZ As Int32
                                Dim lDID As Int32
                                Dim iDTypeID As Int16
                                With CType(goCurrentEnvir.oGeoObject, Planet)
                                    lDX = .LocX
                                    lDZ = .LocZ
                                    If .ParentSystem Is Nothing Then Return
                                    lDID = .ParentSystem.ObjectID
                                    iDTypeID = .ParentSystem.ObjTypeID
                                End With
                                goUILib.GetMsgSys.SendMoveRequestMsgEx(lDX, lDZ, 0, False, False, lDID, iDTypeID)
                                BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eGotoOrbitClick, 2, 0)
                            End If
                        Next XX
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                    End If
                ElseIf e.KeyCode = Keys.T Then
                    If NewTutorialManager.TutorialOn = True Then
                        If goUILib Is Nothing = False Then
                            If goUILib.CommandAllowed(False, "t") = False Then Return
                        End If
                    End If
                    If goCurrentEnvir Is Nothing = False Then
                        If glCurrentEnvirView = CurrentView.eGalaxyMapView Then
                            glCurrentEnvirView = CurrentView.eSystemView
                        ElseIf glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                            glCurrentEnvirView = CurrentView.eSystemView
                        ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                            glCurrentEnvirView = CurrentView.eSystemView
                        End If

                        Dim lXStart As Int32 = 0
                        If goCamera.TrackingIndex <> -1 Then lXStart = goCamera.TrackingIndex + 1
                        goCamera.TrackingIndex = -1
                        For X = lXStart To goCurrentEnvir.lEntityUB
                            If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                                If goCurrentEnvir.oEntity(X).OwnerID = glPlayerID OrElse goCurrentEnvir.oEntity(X).yVisibility = eVisibilityType.Visible Then
                                    goCamera.TrackingIndex = X
                                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eTKeyToTrack, 0, 0)
                                End If
                                Return
                            End If
                        Next X
                        If goCamera.TrackingIndex = -1 AndAlso lXStart <> 0 Then
                            For X = 0 To lXStart
                                If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                                    goCamera.TrackingIndex = X
                                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eTKeyToTrack, 0, 0)
                                    Exit For
                                End If
                            Next X
                        End If
                    End If
                ElseIf e.KeyCode >= Keys.D0 AndAlso e.KeyCode <= Keys.D9 Then
                    Dim lCtrlGroup As Int32 = e.KeyCode - Keys.D0
                    If e.Control = True Then
                        If NewTutorialManager.TutorialOn = True Then
                            If goUILib Is Nothing = False Then
                                If goUILib.CommandAllowed(False, "Ctrl" & (CInt(e.KeyCode) - CInt(Keys.D0)).ToString) = False Then Return
                            End If
                        End If

                        If e.Shift = True Then
                            'Ok, select the control group (adding it)
                            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eCtrlShiftCtrlGroup, lCtrlGroup, 0)
                            goControlGroups.SelectControlGroup(lCtrlGroup, False)
                        Else
                            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eCtrlGroupAssign, lCtrlGroup, 0)
                        End If
                        'And set the control group
                        goControlGroups.SetControlGroup(lCtrlGroup)
                        goControlGroups.SaveGroups()
                    Else
                        If NewTutorialManager.TutorialOn = True Then
                            If goUILib Is Nothing = False Then
                                If goUILib.CommandAllowed(False, (CInt(e.KeyCode) - CInt(Keys.D0)).ToString) = False Then Return
                            End If
                        End If

                        'Selecting... 
                        BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eRecallCtrlGroup, lCtrlGroup, CInt(e.Shift))
                        goControlGroups.SelectControlGroup(lCtrlGroup, Not e.Shift)
                    End If
                    e.Handled = True

                ElseIf e.KeyCode = Keys.B Then

                    'Ok, get our advance display
                    If NewTutorialManager.TutorialOn = True Then
                        If goUILib Is Nothing = False Then
                            Dim sParms() As String = {"Build"}
                            If goUILib.CommandAllowedWithParms(True, "frmAdvanceDisplay", sParms, False) = False Then Return
                        End If
                    End If

                    If goUILib Is Nothing = False Then
                        Dim ofrm As frmAdvanceDisplay = CType(goUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
                        If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                            ofrm.OpenBuildWindow()
                        Else
                            If goCurrentEnvir Is Nothing = False Then
                                For X = 0 To goCurrentEnvir.lEntityUB
                                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID <> glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eFacility AndAlso goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eMining Then
                                        Dim ofrmMining As frmMining = CType(goUILib.GetWindow("frmMining"), frmMining)
                                        If ofrmMining Is Nothing Then ofrmMining = New frmMining(goUILib)
                                        ofrmMining.Visible = True
                                        ofrmMining.FindAndSelectFacility(X)
                                        Exit For
                                    End If
                                Next X
                            End If
                        End If
                        ofrm = Nothing
                    End If
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eBKeyToOpenBuild, 0, 0)
                ElseIf e.KeyCode = Keys.G Then
                    If goUILib Is Nothing = False Then
                        Dim ofrmAdv As frmAdvanceDisplay = CType(goUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
                        If ofrmAdv Is Nothing = False AndAlso ofrmAdv.Visible = True Then
                            ofrmAdv.HandleUnitGotoClick()
                        Else
                            Dim oMultiSelect As frmMultiDisplay = CType(goUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
                            If oMultiSelect Is Nothing = False AndAlso oMultiSelect.Visible = True Then
                                If oMultiSelect.HandleUnitGotoClick() = True Then
                                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                                End If
                            End If
                            oMultiSelect = Nothing
                        End If
                        ofrmAdv = Nothing
                    End If
                ElseIf e.KeyCode = Keys.F Then
                    Dim oMultiSelect As frmMultiDisplay = CType(goUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
                    If oMultiSelect Is Nothing = False AndAlso oMultiSelect.Visible = True Then
                        oMultiSelect.HandleFilterClick()
                    End If
                    oMultiSelect = Nothing
                ElseIf e.KeyCode = Keys.Left Then
                    If goCamera Is Nothing = False Then
                        BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eArrowKey, CInt(e.KeyCode), 0)
                        mbKeyScrollLeft = True
                        goCamera.bScrollLeft = True
                    End If
                    ElseIf e.KeyCode = Keys.Up Then
                        If goCamera Is Nothing = False Then
                            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eArrowKey, CInt(e.KeyCode), 0)
                            mbKeyScrollUp = True
                            goCamera.bScrollUp = True
                        End If
                    ElseIf e.KeyCode = Keys.Down Then
                        If goCamera Is Nothing = False Then
                            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eArrowKey, CInt(e.KeyCode), 0)
                            mbKeyScrollDown = True
                            goCamera.bScrollDown = True
                        End If
                    ElseIf e.KeyCode = Keys.Right Then
                        If goCamera Is Nothing = False Then
                            BPMetrics.MetricMgr.AddActivity(BPMetrics.eKeyPress.eArrowKey, CInt(e.KeyCode), 0)
                            mbKeyScrollRight = True
                            goCamera.bScrollRight = True
                        End If
                    End If
            End If
        End If
        e.Handled = True
    End Sub


    Public Shared lBackedOutToMapViewCycle As Int32 = 0
    Public Shared Sub GoBackAView()
        If NewTutorialManager.TutorialOn = True Then
            If goUILib Is Nothing = False Then
                If goUILib.CommandAllowed(False, "Backspace") = False Then Return
            End If
        End If
        Dim ofrmED As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
        If ofrmED Is Nothing Then ofrmED = New frmEnvirDisplay(goUILib)
        If Not ofrmED Is Nothing Then
            ofrmED.ChangeEnviroments()
        End If

        If glCurrentEnvirView > CurrentView.eFullScreenInterface Then
            ReturnToPreviousView()
        ElseIf glCurrentEnvirView <> CurrentView.eGalaxyMapView Then

            If NewTutorialManager.TutorialOn = True Then
                If goUILib Is Nothing = False Then
                    If goUILib.CommandAllowed(True, "ChangeView_" & (glCurrentEnvirView - 1).ToString) = False Then Return
                End If
            End If

            glCurrentEnvirView -= 1

            If glCurrentEnvirView = CurrentView.eSystemView Then
                'that means we came from planet map view...
                If goGalaxy Is Nothing = False Then
                    If goGalaxy.CurrentSystemIdx <> -1 Then
                        'really should never happen...
                        Dim lTemp As Int32 = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx
                        If lTemp <> -1 Then
                            'also shouldnt happen...
                            With goCamera
                                .mlCameraAtX = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(lTemp).LocX
                                .mlCameraAtZ = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(lTemp).LocZ
                                .mlCameraAtY = 0
                                .mlCameraX = .mlCameraAtX : .mlCameraY = 1000 : .mlCameraZ = .mlCameraAtZ - 1000
                            End With
                        End If
                    End If
                End If
            ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                With goCamera
                    .mlCameraAtX = CInt(.mlCameraAtX / 30) : .mlCameraAtY = 0 : .mlCameraAtZ = CInt(.mlCameraAtZ / 30)
                    .mlCameraX = .mlCameraAtX : .mlCameraY = 10000 : .mlCameraZ = .mlCameraAtZ - 1000
                End With
            Else
                'For now, reset our viewing area to 0,0,0
                With goCamera
                    .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                    .mlCameraX = 0 : .mlCameraY = 1000 : .mlCameraZ = -1000

                    If glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                        .mlCameraZ = -1200
                    End If

                    If muSettings.lGalaxyViewCameraX = 0 AndAlso muSettings.lGalaxyViewCameraZ = 0 Then
                        muSettings.lGalaxyViewCameraX = 0
                        muSettings.lGalaxyViewCameraZ = -5000
                        muSettings.lGalaxyViewCameraY = 5000
                    End If

                    If glCurrentEnvirView = CurrentView.ePlanetView Then
                        .mlCameraY = muSettings.lPlanetViewCameraY + 1000
                        .mlCameraX += muSettings.lPlanetViewCameraX
                        .mlCameraZ += muSettings.lPlanetViewCameraZ
                    ElseIf glCurrentEnvirView = CurrentView.eGalaxyMapView Then
                        If goGalaxy Is Nothing = False AndAlso goGalaxy.CurrentSystemIdx > -1 Then
                            .mlCameraAtX = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).LocX
                            .mlCameraAtZ = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).LocZ
                            .mlCameraX = .mlCameraAtX
                            .mlCameraZ = .mlCameraAtZ

                            .mlCameraY = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).LocY + muSettings.lGalaxyViewCameraY
                            .mlCameraX += muSettings.lGalaxyViewCameraX
                            .mlCameraZ += muSettings.lGalaxyViewCameraZ
                        Else
                            .mlCameraY = muSettings.lGalaxyViewCameraY
                            .mlCameraX += muSettings.lGalaxyViewCameraX
                            .mlCameraZ += muSettings.lGalaxyViewCameraZ
                        End If
                    End If
                End With
            End If
        End If

        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
            goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.BackspaceKeyPressed)
        End If
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eViewChanged, glCurrentEnvirView, -1, -1, "")
        End If
    End Sub

    Private Sub frmMain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp

        mlLastUIEvent = glCurrentCycle
        If mbFormClosing = True Then Return
        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return

        If mbShiftKeyDown = True AndAlso e.Shift = False Then
            If goUILib.BuildGhost Is Nothing = False Then
                'Bring the build window back now that the placement has occured. Pass only
                Dim ofrm As UIWindow = goUILib.GetWindow("frmBuildWindow")
                If ofrm Is Nothing = False Then ofrm.Visible = True

                'and remove our buildghost
                goUILib.BuildGhost = Nothing
                goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
            Else
                Dim oFrm As frmPrototypeBuilder = CType(goUILib.GetWindow("frmPrototypeBuilder"), frmPrototypeBuilder)
                If oFrm Is Nothing = False Then oFrm.SetWeaponPlaceState(False, False)
                oFrm = Nothing
            End If
        End If
        mbShiftKeyDown = e.Shift
        mbCtrlKeyDown = e.Control
        If mbCtrlKeyDown = False Then moEngine.bInCtrlMove = False
        mbAltKeyDown = e.Alt
        moEngine.bRenderSelectedRanges = mbAltKeyDown

        If mbAltKeyDown = False Then ResetDistanceDisplay()

        If mbIgnoreNextKeyUp = True Then
            mbIgnoreNextKeyUp = False
            Return
        End If
        If glCurrentEnvirView = CurrentView.eStartupDSELogo Then Return

        If e.KeyCode = Keys.Left AndAlso e.Control = False Then
            If goCamera Is Nothing = False Then
                mbKeyScrollLeft = False
                goCamera.bScrollLeft = False
            End If
        ElseIf e.KeyCode = Keys.Up Then
            If goCamera Is Nothing = False Then
                mbKeyScrollUp = False
                goCamera.bScrollUp = False
            End If
        ElseIf e.KeyCode = Keys.Down Then
            If goCamera Is Nothing = False Then
                mbKeyScrollDown = False
                goCamera.bScrollDown = False
            End If
        ElseIf e.KeyCode = Keys.Right AndAlso e.Control = False Then
            If goCamera Is Nothing = False Then
                mbKeyScrollRight = False
                goCamera.bScrollRight = False
            End If
        End If

        If goUILib Is Nothing = False AndAlso goUILib.PostMessage(UILibMsgCode.eKeyUpCode, e) = True Then Exit Sub

        If e.KeyCode = Keys.F Then
            mbFKeyDown = False
        ElseIf e.KeyCode = Keys.Z Then
            mbMouseWheelOverride = False
        End If

        e.Handled = True
    End Sub

    Private Sub frmMain_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress

        mlLastUIEvent = glCurrentCycle
        If mbFormClosing = True Then Return
        If mbIgnoreNextKeyUp = True Then Return
        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return
        If goUILib Is Nothing = False AndAlso goUILib.PostMessage(UILibMsgCode.eKeyPressCode, e) = True Then Exit Sub
    End Sub

    Public Sub ForceChangeEnvironment(ByVal lID As Int32, ByVal iTypeID As Int16)

        muSettings.gbPlanetMapDontRenderTerrain = False
        muSettings.gbSolarMapDontShowPlanets = False
        muSettings.gbSolarMapDontShowStars = False
        muSettings.gbSolarMapDontShowWormholes = False

        mbEnvirReady = False
        mbChangingEnvirs = True

        If goUILib Is Nothing = False Then goUILib.ClearSelection()

        If goCurrentEnvir Is Nothing = False Then
            If goSound Is Nothing = False Then
                If goCurrentEnvir.mlPrimaryAmbienceID <> -1 Then goSound.StopSound(goCurrentEnvir.mlPrimaryAmbienceID)
            End If
        End If

        Dim lLoopCnt As Int32 = 0
        While gb_InHandleMovement
            Threading.Thread.Sleep(10)
            lLoopCnt += 1
            If lLoopCnt > 1000 Then Exit While
        End While

        goCurrentEnvir = Nothing

        'TODO: This is a hack because it was breaking on me all the time...
        If moMsgSys Is Nothing Then moMsgSys = goUILib.GetMsgSys()
        moMsgSys.SendChangeEnvironment(lID, iTypeID)

    End Sub

End Class

