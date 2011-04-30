Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class Galaxy
    Inherits Base_GUID

    Private Const mf_SYSTEM_SPHERE_RADIUS As Single = 64

    Public GalaxyName As String

    Private mlCurrentSystemIdx As Int32 = -1

    Public moSystems() As SolarSystem
    Public mlSystemUB As Int32 = -1
    'TODO: Include Nebulae

    'Private moSpr As Sprite
    Private moBPSpr As BPSprite

    Private mbCurrentStarBlinkState As Boolean
    Private mlLastFlashStateUpdate As Int32 = 0

    Private mbGalaxySelectionChange As Boolean
    Private mlGalaxySelectionIdx As Int32 = -1

    Public Property GalaxySelectionChange() As Boolean
        Get
            GalaxySelectionChange = mbGalaxySelectionChange
        End Get
        Set(ByVal value As Boolean)
            mbGalaxySelectionChange = value
        End Set
    End Property

    Public Property GalaxySelectionIdx() As Int32
        Get
            GalaxySelectionIdx = mlGalaxySelectionIdx
        End Get
        Set(ByVal value As Int32)
            mlGalaxySelectionIdx = value
        End Set
    End Property

    Public Function GetSystemName(ByVal lID As Int32) As String
        For lSysIdx As Int32 = 0 To mlSystemUB
            If moSystems(lSysIdx).ObjectID = lID Then Return moSystems(lSysIdx).SystemName
        Next lSysIdx

        Return ""
    End Function

    Public Property CurrentSystemIdx() As Int32
        Get
            Return mlCurrentSystemIdx
        End Get
        Set(ByVal Value As Int32)
            If Value <> mlCurrentSystemIdx Then
                If mlCurrentSystemIdx > -1 Then
                    moSystems(mlCurrentSystemIdx).CleanResources(False)
                End If
            End If
            mlCurrentSystemIdx = Value
        End Set
    End Property

    'MSC - 08.18.08 - remarked out to see if removes the nagging crash on close
    'Private Sub SpriteDispose(ByVal sender As Object, ByVal e As EventArgs)
    '    moSpr = Nothing
    'End Sub

    'Private Sub SpriteLost(ByVal sender As Object, ByVal e As EventArgs)
    '    If moSpr Is Nothing = False Then moSpr.Dispose()
    '    moSpr = Nothing
    'End Sub

    'Private Sub RenderWormholes()
    '    Dim moDevice As Device = GFXEngine.moDevice

    '    moDevice.VertexFormat = CustomVertex.PositionTextured.Format
    '    If WpnFXManager.moTexture Is Nothing OrElse WpnFXManager.moTexture.Disposed = True Then WpnFXManager.moTexture = goResMgr.GetTexture("WpnFire.dds", GFXResourceManager.eGetTextureType.NoSpecifics)
    '    moDevice.SetTexture(0, Nothing) 'WpnFXManager.moTexture)
    '    moDevice.Transform.World = Matrix.Identity


    '    Dim oMat As Material
    '    With oMat
    '        '.Ambient = Color.Black
    '        .Emissive = Color.White
    '        .Diffuse = System.Drawing.Color.FromArgb(128, 255, 255, 255)
    '    End With

    '    moDevice.Material = oMat

    '    Dim fTuLow As Single = 0
    '    Dim fTuHi As Single = 0.5F
    '    Dim fTvLow As Single = 0.25F
    '    Dim fTvHi As Single = 0.5F

    '    For X As Int32 = 0 To mlSystemUB
    '        If moSystems(X) Is Nothing = False Then
    '            For Y As Int32 = 0 To moSystems(X).WormholeUB
    '                If moSystems(X).moWormholes(Y) Is Nothing = False Then
    '                    With moSystems(X).moWormholes(Y)
    '                        If .System1 Is Nothing = False AndAlso .System1.ObjectID = moSystems(X).ObjectID AndAlso .System2 Is Nothing = False Then

    '                            Dim uVerts(3) As CustomVertex.PositionTextured

    '                            Dim fSize As Single = 1
    '                            Dim fCurrentX As Single = .System1.LocX
    '                            Dim fCurrentY As Single = .System1.LocY + 8
    '                            Dim fCurrentZ As Single = .System1.LocZ
    '                            Dim fDestX As Single = .System2.LocX
    '                            Dim fDestY As Single = .System2.LocY + 8
    '                            Dim fDestZ As Single = .System2.LocZ

    '                            uVerts(0) = New CustomVertex.PositionTextured(fCurrentX - fSize, fCurrentY, fCurrentZ + (fSize / 2), fTuLow, fTvLow)
    '                            uVerts(1) = New CustomVertex.PositionTextured(fCurrentX + fSize, fCurrentY, fCurrentZ - (fSize / 2), fTuLow, fTvHi)
    '                            uVerts(2) = New CustomVertex.PositionTextured(fDestX - fSize, fDestY, fDestZ + (fSize / 2), fTuHi, fTvLow)
    '                            uVerts(3) = New CustomVertex.PositionTextured(fDestX + fSize, fDestY, fDestZ - (fSize / 2), fTuHi, fTvHi)
    '                            moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)

    '                            uVerts(0) = New CustomVertex.PositionTextured(fCurrentX + (fSize / 2), fCurrentY + fSize, fCurrentZ, fTuLow, fTvLow)
    '                            uVerts(1) = New CustomVertex.PositionTextured(fCurrentX - (fSize / 2), fCurrentY - fSize, fCurrentZ, fTuLow, fTvHi)
    '                            uVerts(2) = New CustomVertex.PositionTextured(fDestX + (fSize / 2), fDestY + fSize, fDestZ, fTuHi, fTvLow)
    '                            uVerts(3) = New CustomVertex.PositionTextured(fDestX - (fSize / 2), fDestY - fSize, fDestZ, fTuHi, fTvHi)
    '                            moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)

    '                        End If
    '                    End With
    '                End If
    '            Next Y
    '        End If
    '    Next X
    'End Sub


    Public Sub ReleaseSprite()
        'Try
        '    If moSpr Is Nothing = False Then moSpr.Dispose()
        '    moSpr = Nothing
        'Catch
        'End Try
    End Sub

    Public Sub Render(ByVal lEnvirType As Int32)

        Dim moDevice As Device = GFXEngine.moDevice

        'If moSpr Is Nothing Then
        '    Device.IsUsingEventHandlers = False
        '    moSpr = New Sprite(moDevice)
        '    'AddHandler moSpr.Disposing, AddressOf SpriteDispose
        '    'AddHandler moSpr.Lost, AddressOf SpriteLost
        '    'AddHandler moSpr.Reset, AddressOf SpriteLost
        '    Device.IsUsingEventHandlers = True
        'End If
        If moBPSpr Is Nothing Then moBPSpr = New BPSprite()

        If lEnvirType = CurrentView.eGalaxyMapView Then
            'Set the camera on the current environment
            If mbGalaxySelectionChange OrElse (goCamera.mlCameraY = 1000 AndAlso goCamera.mlCameraZ = -1000) Then
                mbGalaxySelectionChange = False
                If mlGalaxySelectionIdx = -1 Then
                    mlGalaxySelectionIdx = goGalaxy.CurrentSystemIdx
                End If
                CenterOnStarID(goGalaxy.moSystems(mlGalaxySelectionIdx).ObjectID)
                'Dim x As String = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).SystemName
                'Dim y As Int32 = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID

                'If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                ''CenterOnStarID(goCurrentEnvir.ObjectID)
                'CenterOnStarID(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID)
                'ElseIf goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                'CenterOnStarID(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID)
                'End If
            End If

            Dim Z As Int32
            Dim oVecCenter As Vector3 = New Vector3(16, 0, 16)
            Dim oVecLoc As Vector3
            Dim lColor As System.Drawing.Color
            Dim fMult As Single

            With moDevice
                .Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.Black, 1, 0)

                'We disable our lighting for this section because these sprits have their own lighting...
                .RenderState.Lighting = False
                .Transform.World = Matrix.Identity

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
                .RenderState.SourceBlend = Blend.SourceAlpha
                .RenderState.DestinationBlend = Blend.One
                .RenderState.AlphaBlendEnable = True
                .RenderState.ZBufferEnable = False
            End With

            'moSpr.SetWorldViewLH(moDevice.Transform.World, moDevice.Transform.View)
            'moSpr.Begin(SpriteFlags.Billboard Or SpriteFlags.AlphaBlend Or SpriteFlags.ObjectSpace Or SpriteFlags.SortDepthFrontToBack Or SpriteFlags.DoNotModifyRenderState)
            Try
                If StarType.StarMapTex Is Nothing Then
                    StarType.StarMapTex = goResMgr.GetTexture("whitestar.dds", GFXResourceManager.eGetTextureType.NoSpecifics)
                End If
            Catch ex As Exception
            End Try
            moBPSpr.BeginRender(mlSystemUB + 1, StarType.StarMapTex, moDevice)
            For Z = 0 To mlSystemUB

                Dim fSysX As Single = moSystems(Z).LocX '* 2
                Dim fSysY As Single = moSystems(Z).LocY '* 2
                Dim fSysZ As Single = moSystems(Z).LocZ '* 2
                If goCamera.CullObject_NoMaxDistance(CullBox.GetCullBox(fSysX, fSysY, fSysZ, -100, -100, -100, 100, 100, 100)) = True Then Continue For

                If moSystems(Z).StarType1Idx <> -1 Then
                    With goStarTypes(moSystems(Z).StarType1Idx)
                        lColor = .StarMaterial.Emissive
                        If (.StarAttributes And 1) <> 0 Then
                            fMult = CSng((Rnd() * 0.2) + 0.8)
                            lColor = System.Drawing.Color.FromArgb(255, CInt(lColor.R * fMult), CInt(lColor.G * fMult), CInt(lColor.B * fMult))
                        End If
                        oVecLoc.X = fSysX
                        oVecLoc.Y = fSysY + 5
                        oVecLoc.Z = fSysZ
                        'moSpr.Draw(StarType.StarMapTex, StarType.StarMapRect, oVecCenter, oVecLoc, lColor)
                        moBPSpr.Draw(oVecLoc, 24, lColor, 0)
                    End With
                End If
                If moSystems(Z).StarType2Idx <> -1 Then
                    '2nd is always upright of original
                    With goStarTypes(moSystems(Z).StarType2Idx)
                        lColor = .StarMaterial.Emissive
                        If (.StarAttributes And 1) <> 0 Then
                            fMult = CSng((Rnd() * 0.2) + 0.8)
                            lColor = System.Drawing.Color.FromArgb(255, CInt(lColor.R * fMult), CInt(lColor.G * fMult), CInt(lColor.B * fMult))
                        End If
                        oVecLoc.X = fSysX + 5
                        oVecLoc.Y = fSysY + 10
                        oVecLoc.Z = fSysZ
                        'moSpr.Draw(StarType.StarMapTex, StarType.StarMapRect, oVecCenter, oVecLoc, lColor)
                        moBPSpr.Draw(oVecLoc, 24, lColor, 0)
                    End With
                End If
                If moSystems(Z).StarType3Idx <> -1 Then
                    '3rd is always up left of original
                    With goStarTypes(moSystems(Z).StarType3Idx)
                        lColor = .StarMaterial.Emissive
                        If (.StarAttributes And 1) <> 0 Then
                            fMult = CSng((Rnd() * 0.2) + 0.8)
                            lColor = System.Drawing.Color.FromArgb(255, CInt(lColor.R * fMult), CInt(lColor.G * fMult), CInt(lColor.B * fMult))
                        End If
                        oVecLoc.X = fSysX - 5
                        oVecLoc.Y = fSysY + 10
                        oVecLoc.Z = fSysZ
                        'moSpr.Draw(StarType.StarMapTex, StarType.StarMapRect, oVecCenter, oVecLoc, lColor)
                        moBPSpr.Draw(oVecLoc, 24, lColor, 0)
                    End With
                End If
            Next Z
            'moSpr.End()
            moBPSpr.EndRender()

            'With moDevice
            '    .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            '    .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
            '    .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
            '    .RenderState.SourceBlend = Blend.SourceAlpha
            '    .RenderState.DestinationBlend = Blend.One
            '    .RenderState.AlphaBlendEnable = True
            '    .RenderState.ZBufferEnable = False
            'End With

            'If Not muSettings.gbGalaxyControlHideWormholes Then RenderWormholes()

            'Show our Galaxy Control window...
            If goUILib Is Nothing = False Then
                If goUILib.GetWindow("frmGalaxyControl") Is Nothing Then
                    Dim frmGalaxyControl As New frmGalaxyControl(goUILib)
                    frmGalaxyControl.Visible = True
                    frmGalaxyControl = Nothing
                End If
            End If

            'Reset our render states
            With moDevice
                .RenderState.Lighting = True

                .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                .RenderState.DestinationBlend = Blend.InvSourceAlpha ' Blend.Zero
                .RenderState.AlphaBlendEnable = True

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                .RenderState.ZBufferEnable = True
            End With

            Try
                moDevice.Transform.World = Matrix.Identity

                Dim fCurX As Single = goCamera.mlCameraAtX
                Dim fCurY As Single = goCamera.mlCameraAtY
                Dim fCurZ As Single = goCamera.mlCameraAtZ

                If Me.GalaxySelectionIdx > -1 Then
                    With moSystems(Me.GalaxySelectionIdx)
                        fCurX = .LocX
                        fCurY = .LocY
                        fCurZ = .LocZ
                    End With
                End If

                Using oFont As New Font(moDevice, New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
                    For X As Int32 = 0 To mlSystemUB
                        Dim fTmpX As Single = moSystems(X).LocX
                        Dim fTmpY As Single = moSystems(X).LocY
                        Dim fTmpZ As Single = moSystems(X).LocZ
                        If goCamera.CullObject_NoMaxDistance(CullBox.GetCullBox(fTmpX, fTmpY, fTmpZ, -100, -100, -100, 100, 100, 100)) = True Then Continue For

                        Dim vec As Vector3 = New Vector3(fTmpX, fTmpY, fTmpZ)
                        vec.Project(moDevice.Viewport, moDevice.Transform.Projection, moDevice.Transform.View, moDevice.Transform.World)

                        If X = GalaxySelectionIdx Then
                            If glCurrentCycle - mlLastFlashStateUpdate > 15 Then
                                mlLastFlashStateUpdate = glCurrentCycle
                                mbCurrentStarBlinkState = Not mbCurrentStarBlinkState
                            End If
                            If mbCurrentStarBlinkState = True Then
                                oFont.DrawText(Nothing, moSystems(X).SystemName, CInt(vec.X), CInt(vec.Y), System.Drawing.Color.FromArgb(128, 0, 255, 0))
                            End If
                        Else
                            Dim lClrVal As Int32 = 255
                            If muSettings.gbGalaxyControlStarLabelFalloff = True Then
                                Dim fDX As Single = fTmpX - fCurX
                                Dim fDY As Single = fTmpY - fCurY
                                Dim fDZ As Single = fTmpZ - fCurZ
                                fDX *= fDX
                                fDY *= fDY
                                fDZ *= fDZ
                                Dim fDist As Single = fDX + fDY + fDZ
                                fDist = fDist * 0.00001F

                                lClrVal = CInt((1.0F - fDist) * 200)       '155 because we want a base of 100 to 255
                                If lClrVal < 0 Then lClrVal = 0
                                If lClrVal > 200 Then lClrVal = 200
                                lClrVal += 55

                                If fDist > 2 Then lClrVal = 0
                            End If
                            oFont.DrawText(Nothing, moSystems(X).SystemName, CInt(vec.X), CInt(vec.Y), System.Drawing.Color.FromArgb(192, lClrVal, lClrVal, lClrVal))
                        End If
                    Next X
                End Using
            Catch
            End Try

        ElseIf lEnvirType < CurrentView.eFullScreenInterface Then
            If NewTutorialManager.TutorialOn = True AndAlso goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
                Dim oPlanet As Planet = Planet.GetTutorialPlanet()
                If oPlanet Is Nothing = False Then oPlanet.Render(lEnvirType)
            Else
                If moSystems Is Nothing Then Return

                Dim lCurrSysIdx As Int32 = CurrentSystemIdx
                If lCurrSysIdx = -1 OrElse lCurrSysIdx > mlSystemUB Then
                    'ok, need to determine current system idx
                    Dim oEnvir As BaseEnvironment = goCurrentEnvir
                    If oEnvir Is Nothing = False Then
                        If oEnvir.ObjTypeID = ObjectType.ePlanet Then
                            Dim bFound As Boolean = False
                            Dim lCurUB As Int32 = Math.Min(moSystems.GetUpperBound(0), mlSystemUB)
                            For X As Int32 = 0 To lCurUB
                                If moSystems(X) Is Nothing = False Then
                                    For Y As Int32 = 0 To moSystems(X).PlanetUB
                                        If moSystems(X).moPlanets(Y) Is Nothing = False AndAlso moSystems(X).moPlanets(Y).ObjectID = oEnvir.ObjectID Then
                                            CurrentSystemIdx = X
                                            moSystems(X).CurrentPlanetIdx = Y
                                            bFound = True
                                            Exit For
                                        End If
                                    Next Y
                                    If bFound = True Then Exit For
                                End If
                            Next X
                            If bFound = False Then Return
                        ElseIf oEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                            Dim bFound As Boolean = False
                            Dim lCurUB As Int32 = Math.Min(moSystems.GetUpperBound(0), mlSystemUB)
                            For X As Int32 = 0 To lCurUB
                                If moSystems(X) Is Nothing = False AndAlso moSystems(X).ObjectID = oEnvir.ObjectID Then
                                    bFound = True
                                    CurrentSystemIdx = X
                                    Exit For
                                End If
                            Next X
                            If bFound = False Then Return
                        Else : Return
                        End If
                    Else : Return
                    End If
                End If
                moSystems(CurrentSystemIdx).Render(lEnvirType)
            End If
        End If
    End Sub

    Public Sub AddSystem(ByVal oSystem As SolarSystem)
        mlSystemUB += 1
        ReDim Preserve moSystems(mlSystemUB)
        moSystems(mlSystemUB) = oSystem
    End Sub

    'Public Function PickSystem(ByVal oRay As Ray, ByVal fMaxDist As Single) As Int32
    '    Dim bHit As Boolean = False
    '    Dim X As Int32

    '    Dim vNear As Microsoft.DirectX.Vector3
    '    Dim vDir As Microsoft.DirectX.Vector3

    '    Dim matSys As Microsoft.DirectX.Matrix
    '    Dim oIntInfo As Microsoft.DirectX.Direct3D.IntersectInformation

    '    Dim lCurrentIdx As Int32 = -1
    '    Dim fCurrentDist As Single = fMaxDist


    '    If moSysMesh Is Nothing Then
    '        Device.IsUsingEventHandlers = False
    '        moSysMesh = Mesh.Sphere(moDevice, mf_SYSTEM_SPHERE_RADIUS, 8, 8)
    '        Device.IsUsingEventHandlers = True
    '    End If

    '    For X = 0 To mlSystemUB

    '        'Convert the ray to model space
    '        matSys = Microsoft.DirectX.Matrix.Identity
    '        matSys.Translate(moSystems(X).LocX, moSystems(X).LocY, moSystems(X).LocZ)
    '        matSys.Invert()
    '        vNear = Microsoft.DirectX.Vector3.TransformCoordinate(oRay.Origin, matSys)
    '        vDir = Microsoft.DirectX.Vector3.TransformNormal(oRay.Direction, matSys)

    '        'Now, test the intersection...
    '        bHit = moSysMesh.Intersect(vNear, vDir, oIntInfo)

    '        If bHit = True Then
    '            If oIntInfo.Dist < fMaxDist Then
    '                'we've hit it... and its close enough to be hit... but is it closer then what we have?
    '                If lCurrentIdx <> -1 OrElse oIntInfo.Dist < fCurrentDist Then
    '                    'yes, it is closer
    '                    fCurrentDist = oIntInfo.Dist
    '                    lCurrentIdx = X
    '                End If
    '            End If
    '        End If

    '    Next X

    '    Return lCurrentIdx
    'End Function 

    Public Function GetPlanetName(ByVal lID As Int32) As String

        For X As Int32 = 0 To mlSystemUB
            For Y As Int32 = 0 To moSystems(X).PlanetUB
                If moSystems(X).moPlanets(Y).ObjectID = lID Then Return moSystems(X).moPlanets(Y).PlanetName
            Next Y
        Next X
        Return ""
    End Function

    Public Function CenterOnStar(ByVal sStarName As String, ByVal lId As Int32) As Int32
        If lId >= mlSystemUB Then lId = 0

        sStarName = sStarName.ToUpper.Trim
        If sStarName = "" Then Exit Function
        For X As Int32 = lId + 1 To mlSystemUB
            If moSystems(X) Is Nothing Then Exit For
            If moSystems(X).SystemName.ToUpper.StartsWith(sStarName) = True AndAlso moSystems(X).ObjectID <> 36 AndAlso moSystems(X).ObjectID <> 88 Then
                mlGalaxySelectionIdx = X
                mbGalaxySelectionChange = True
                'CenterOnStarID(moSystems(X).ObjectID)
                Return X
            End If
        Next X
        For X As Int32 = 0 To lId
            If moSystems(X) Is Nothing Then Exit For
            If moSystems(X).SystemName.ToUpper.StartsWith(sStarName) = True AndAlso moSystems(X).ObjectID <> 36 AndAlso moSystems(X).ObjectID <> 88 Then
                mlGalaxySelectionIdx = X
                mbGalaxySelectionChange = True
                'CenterOnStarID(moSystems(X).ObjectID)
                Return X
            End If
        Next X
        Return 0

    End Function

    Public Sub CenterOnStarID(ByVal lId As Int32)
        For X As Int32 = 0 To mlSystemUB
            If moSystems(X) Is Nothing Then Exit For
            If moSystems(X).ObjectID = lId Then
                goCamera.SimplyPlaceCamera(moSystems(X).LocX, moSystems(X).LocY, moSystems(X).LocZ)
                Return
            End If
        Next
    End Sub

    Private mbJumpToPlanet As Boolean = False
    Private mlJumpToSolarObjectID As Int32 = -1
    Private mlJumpToPlanetNumber As Int32 = -1
    Private mlJumpToPlanetObjectID As Int32 = -1
    Private miEnvironmentReceived As Boolean = False

    Public Function GotoStar(ByVal sStarName As String) As Boolean
        mbJumpToPlanet = False
        mlJumpToSolarObjectID = -1
        mlJumpToPlanetNumber = -1
        mlJumpToPlanetObjectID = -1

        If sStarName.Length = 0 Then Return False
        sStarName = Replace(sStarName.ToUpper, "(S)", "").Trim
        If sStarName = "" Then Exit Function
        Dim iPlanetNumber As Integer = 0

        'Do we need to lookup a planet?
        If sStarName.StartsWith("GLIESE 581") = True Then 'Special name
            Dim sVals() As String = Split(sStarName, " ")
            iPlanetNumber = CInt(Val(GetRomanNumeralSortStr(Replace(sStarName, "GLIESE ", ""))))
            sStarName = "GLIESE 581"
        ElseIf sStarName.Contains(" ") Then
            Dim sVals() As String = Split(sStarName, " ")
            iPlanetNumber = CInt(Val(GetRomanNumeralSortStr(sVals(sVals.GetUpperBound(0)))))
            If iPlanetNumber = 0 AndAlso sVals.GetUpperBound(0) > 1 Then '3 part string (New Nomjga New) would be trigger this.
                goUILib.AddNotification("Invalid Entry.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            ElseIf iPlanetNumber = 0 Then
                goUILib.AddNotification("Invalid Planet Entry.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            sStarName = Replace(sStarName.ToUpper, " " & sVals(sVals.GetUpperBound(0)), "").Trim
        End If

        If sStarName = "" Then Exit Function
        For X As Int32 = 0 To mlSystemUB
            If moSystems(X) Is Nothing Then Continue For
            If Replace(moSystems(X).SystemName, "(S)", "").ToUpper.Trim = sStarName AndAlso moSystems(X).ObjectID <> 36 AndAlso moSystems(X).ObjectID <> 88 Then
                If iPlanetNumber > 0 Then
                    If moSystems(X).PlanetUB = -1 Then
                        'Found the system, get the details
                        mlJumpToSolarObjectID = moSystems(X).ObjectID
                        mlJumpToPlanetNumber = iPlanetNumber
                        mbJumpToPlanet = True
                        Dim yData(5) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestSystemDetails).CopyTo(yData, 0)
                        System.BitConverter.GetBytes(moSystems(X).ObjectID).CopyTo(yData, 2)
                        goUILib.SendMsgToPrimary(yData)
                        Return True
                    Else
                        'We have the details, lets go there.
                        For Y As Int32 = 0 To moSystems(X).PlanetUB
                            If GetPlanetNameValue(moSystems(X).moPlanets(Y).PlanetName) = iPlanetNumber Then
                                goUILib.AddNotification("Planet " & moSystems(X).moPlanets(Y).PlanetName.ToString & " found.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                If GotoEnvironmentWrapper(moSystems(X).moPlanets(Y).ObjectID, ObjectType.ePlanet) = True Then Return True
                                Exit For
                            End If
                        Next Y
                        'Something bad happened, why didn't we find the planet?
                        goUILib.AddNotification("Planet number " & iPlanetNumber.ToString & " not found.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                        Return False
                    End If
                Else
                    'We just want to goto a star, lets go.
                    GotoEnvironment(moSystems(X).ObjectID)
                    Return True
                End If
            End If
        Next X
        goUILib.AddNotification("Star not found", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
    End Function

    Public Sub JumpToPlanet(ByVal lSystemID As Int32)
        If miEnvironmentReceived = True Then Exit Sub
        If mbJumpToPlanet = False Then Exit Sub
        If lSystemID <> mlJumpToSolarObjectID Then Exit Sub
        If goCurrentEnvir.ObjectID = lSystemID Then Exit Sub
        miEnvironmentReceived = True
        For X As Int32 = 0 To mlSystemUB
            If moSystems(X) Is Nothing Then Continue For
            If moSystems(X).ObjectID = lSystemID Then
                For Y As Int32 = 0 To moSystems(X).PlanetUB
                    If GetPlanetNameValue(moSystems(X).moPlanets(Y).PlanetName) = mlJumpToPlanetNumber Then
                        mbJumpToPlanet = False
                        mlJumpToSolarObjectID = -1
                        mlJumpToPlanetNumber = -1
                        mlJumpToPlanetObjectID = moSystems(X).moPlanets(Y).ObjectID
                        goUILib.AddNotification("Planet " & moSystems(X).moPlanets(Y).PlanetName.ToString & " found.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Exit Sub
                    End If
                Next Y
                goUILib.AddNotification("Planet number " & mlJumpToPlanetNumber.ToString & " not found.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                Exit For
            End If
        Next X
        mbJumpToPlanet = False
        mlJumpToSolarObjectID = -1
        mlJumpToPlanetNumber = -1
    End Sub

    Public Sub ExecuteJumpToPlanet()
        If mlJumpToPlanetObjectID = -1 Then Exit Sub
        Dim lID As Int32 = mlJumpToPlanetObjectID

        'Reset mlJumpToPlanetObjectID before executing the change.  Heck reset everything just in case.  We don't want an endless loop
        mlJumpToPlanetObjectID = -1
        miEnvironmentReceived = False
        mbJumpToPlanet = False
        mlJumpToSolarObjectID = -1
        mlJumpToPlanetNumber = -1
        frmMain.ForceChangeEnvironment(lID, ObjectType.ePlanet)
    End Sub

    Public Sub GotoEnvironment(ByVal lID As Int32)
        If goCurrentEnvir Is Nothing = False Then
            Dim bAlreadyHere As Boolean = False
            If goCurrentEnvir.ObjectID = lID AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                'we're already here... nothing to do...
                bAlreadyHere = True
                goUILib.AddNotification("You're already there.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Else : frmMain.ForceChangeEnvironment(lID, ObjectType.eSolarSystem)
            End If

            glCurrentEnvirView = CurrentView.eSystemMapView1
        End If
    End Sub

End Class
