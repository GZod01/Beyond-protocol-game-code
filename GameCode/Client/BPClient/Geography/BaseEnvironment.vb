
Public Class BaseEnvironment
    Inherits Base_GUID         'ObjTypeID tells us what sort of environment we are looking at

    Private Const mf_GEO_PLASTIC_EVENT_CHANCE As Single = 0.3       '30%
    Private Const ml_MAX_LINEAR_VOLUME As Int32 = 1000

    Private Structure EventAmbience
        Public lSoundID As Int32
        Public yStateID As Byte         '0 is fade in, 1 is full vol, 2 is fade out (when 2 and vol = 0, sound ends)
        Public fLinearVolume As Single
        Public lLastUpdate As Int32     'from TimeGetTime
        Public bActive As Boolean
    End Structure

    Public oEntity() As BaseEntity
    Public lEntityUB As Int32 = -1
    Public lEntityIdx() As Int32

    Public oCache() As MineralCache
    Public lCacheUB As Int32 = -1
    Public lCacheIdx() As Int32

    Public DomainServerIP As String
    Public DomainServerPort As Int16

    Public oGeoObject As Object     'reference to the geography object this environment is controlling

    Public moPoints() As Microsoft.DirectX.Direct3D.CustomVertex.PositionColoredTextured

    Public lCPUsage As Int32

    'These ID's contain the environment's sound FX
    Public mlPrimaryAmbienceID As Int32 = -1
    Private muEventAmbiences() As EventAmbience
    Private mlEventAmbienceUB As Int32 = -1
    Private Shared msw_Ambience As Stopwatch

    Public lMinXPos As Int32
    Public lMaxXPos As Int32
    Public lMinZPos As Int32
    Public lMaxZPos As Int32
	Public lMapWrapAdjustX As Int32

	Public lIronCurtainPlayers() As Int32
	Public lIronCurtainPlayerUB As Int32 = -1

    Private mlLastCheck As Int32 = Int32.MinValue
    Private mbHasUnitsHere As Boolean = False
    Public Function HasUnitsHere() As Boolean
        If mlLastCheck = Int32.MinValue OrElse glCurrentCycle - mlLastCheck > 90 Then
            mbHasUnitsHere = False
            For X As Int32 = 0 To Me.lEntityUB
                If Me.lEntityIdx(X) <> -1 AndAlso Me.oEntity(X).OwnerID = glPlayerID Then
                    mbHasUnitsHere = True
                    Exit For
                End If
            Next X
            mlLastCheck = glCurrentCycle
        End If
        Return mbHasUnitsHere
    End Function

    Public Sub SetExtents()
        If Me.ObjTypeID = ObjectType.ePlanet Then
            If CType(oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                With CType(oGeoObject, Planet)
                    Dim lEnvirSize As Int32 = .CellSpacing * TerrainClass.Width
                    lMapWrapAdjustX = .CellSpacing * (TerrainClass.Width - 1)
                    Dim lHalfSize As Int32 = lEnvirSize \ 2
                    lMinXPos = -lHalfSize + 1
                    lMinZPos = -lHalfSize + 1
                    lMaxXPos = lHalfSize - 1
                    lMaxZPos = lHalfSize - 1
                End With
            End If
            
        Else
            lMinXPos = -5000000
            lMaxXPos = 5000000
            lMinZPos = lMinXPos
            lMaxZPos = lMaxXPos
            lMapWrapAdjustX = 0
        End If
    End Sub

    Public Function PickObject(ByVal oRay As Ray, ByVal fMaxDist As Single, Optional ByVal bMustBeVisible As Boolean = False) As Int32
        Dim bHit As Boolean = False
        Dim X As Int32

        Dim vNear As Microsoft.DirectX.Vector3
        Dim vDir As Microsoft.DirectX.Vector3

        Dim matUnit As Microsoft.DirectX.Matrix
        Dim oIntInfo As Microsoft.DirectX.Direct3D.IntersectInformation

        Dim lCurrentIdx As Int32 = -1
        Dim fCurrentDist As Single = fMaxDist
		Try
			For X = 0 To lEntityUB
				If lEntityIdx(X) <> -1 Then
					Dim oTmpEntity As BaseEntity = oEntity(X)
					If oTmpEntity Is Nothing = False AndAlso oTmpEntity.oMesh Is Nothing = False AndAlso oTmpEntity.oMesh.oMesh Is Nothing = False Then
						If bMustBeVisible = False OrElse oTmpEntity.yVisibility = eVisibilityType.Visible Then
							'Convert the ray to model space
                            matUnit = oTmpEntity.GetWorldMatrix ' Microsoft.DirectX.Matrix.Identity

                            'matUnit.Multiply(Microsoft.DirectX.Matrix.RotationY(-1.57079637F))
                            'matUnit.Multiply(Microsoft.DirectX.Matrix.Translation(oTmpEntity.fMapWrapLocX, oTmpEntity.LocY, oTmpEntity.LocZ))

							matUnit.Invert()
							vNear = Microsoft.DirectX.Vector3.TransformCoordinate(oRay.Origin, matUnit)
							vDir = Microsoft.DirectX.Vector3.TransformNormal(oRay.Direction, matUnit)

							'Now, test the intersection...
							If oTmpEntity.yVisibility = eVisibilityType.Visible Then
								bHit = oTmpEntity.oMesh.oMesh.Intersect(vNear, vDir, oIntInfo)
							Else
								'TODO: use the Radar Mesh instead of the actual mesh of the object
							End If

							If bHit = True Then
								If oIntInfo.Dist < fMaxDist Then
									'we've hit it... and its close enough to be hit... but is it closer then what we have?
									If lCurrentIdx = -1 OrElse oIntInfo.Dist < fCurrentDist Then
										'yes, it is closer
										fCurrentDist = oIntInfo.Dist
										lCurrentIdx = X
									End If
								End If
							End If
						End If

					End If
				End If
			Next X
		Catch
		End Try

		If NewTutorialManager.TutorialOn = True AndAlso lCurrentIdx <> -1 Then
			Dim oTmpEntity As BaseEntity = oEntity(lCurrentIdx)
			If oTmpEntity Is Nothing = False Then
				NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eClickOnSelect, oTmpEntity.ObjectID, oTmpEntity.ObjTypeID, oTmpEntity.yProductionType, "")
			End If
		End If

		Return lCurrentIdx
	End Function

    Public Function PickCache(ByVal oRay As Ray, ByVal fMaxDist As Single) As Int32
        Dim bHit As Boolean = False
        Dim X As Int32

        Dim vNear As Microsoft.DirectX.Vector3
        Dim vDir As Microsoft.DirectX.Vector3

        Dim matUnit As Microsoft.DirectX.Matrix
        Dim oIntInfo As Microsoft.DirectX.Direct3D.IntersectInformation

        Dim lCurrentIdx As Int32 = -1
        Dim fCurrentDist As Single = fMaxDist

        'Ok, if we hit a unit... then we don't check our caches...
        If bHit = False Then
            Try
                For X = 0 To lCacheUB
                    If lCacheIdx(X) <> -1 Then
                        Dim oTmpCache As MineralCache = oCache(X)
                        If oTmpCache Is Nothing = False AndAlso MineralCache.CacheMesh Is Nothing = False Then
                            matUnit = Microsoft.DirectX.Matrix.Identity
                            matUnit.Translate(oTmpCache.lMapWrapLocX, oTmpCache.LocY, oTmpCache.LocZ)
                            matUnit.Invert()
                            vNear = Microsoft.DirectX.Vector3.TransformCoordinate(oRay.Origin, matUnit)
                            vDir = Microsoft.DirectX.Vector3.TransformNormal(oRay.Direction, matUnit)

                            If oTmpCache.CacheTypeID = MineralCacheType.eMineable Then
                                Dim lModelIdx As Int32 = lCacheIdx(X) Mod 3
                                bHit = MineralCache.CacheMesh(lModelIdx).Intersect(vNear, vDir, oIntInfo)
                            Else
                                bHit = MineralCache.DebrisMesh.Intersect(vNear, vDir, oIntInfo)
                            End If

                            If bHit = True Then
                                If oIntInfo.Dist < fMaxDist Then
                                    If lCurrentIdx = -1 OrElse oIntInfo.Dist < fCurrentDist Then
                                        'yes, it is closer
                                        fCurrentDist = oIntInfo.Dist
                                        lCurrentIdx = X
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next X
            Catch
            End Try
        End If

        Return lCurrentIdx
    End Function

    Public Function UnitInRadarRange(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal fSphereRadius As Single) As Byte
        'returns 0 if not visible, 1 if in max radar range but not in opt radar range, 2 if in opt radar range
        Dim X As Int32

        Dim fMaxRadiiSum As Single
        Dim fOptRadiiSum As Single

        Dim yRes As Byte = eVisibilityType.NoVisibility

        Dim fDist As Single
        Dim fDX As Single
        Dim fDZ As Single

        'Range offset is passed in from Models.dat, we multiply it by gl_FINAL_GRID_SQUARE_SIZE
        fSphereRadius *= gl_FINAL_GRID_SQUARE_SIZE

        Dim bCheckGuild As Boolean = False
        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
            If (goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.ShareUnitVision) <> 0 Then bCheckGuild = True
        End If

        Try
            For X = 0 To lEntityUB
				If lEntityIdx(X) <> -1 Then
					'If oEntity(X).OwnerID = glPlayerID AndAlso oEntity(X).oUnitDef Is Nothing = False AndAlso _
					'(oEntity(X).CurrentStatus And elUnitStatus.eRadarOperational) <> 0 AndAlso _
					'(oEntity(X).ObjTypeID <> ObjectType.eFacility OrElse (oEntity(X).CurrentStatus And elUnitStatus.eFacilityPowered) <> 0) Then
					Dim oTmpEntity As BaseEntity = oEntity(X)
                    If oTmpEntity Is Nothing = False AndAlso oTmpEntity.oUnitDef Is Nothing = False Then
                        If oTmpEntity.OwnerID = glPlayerID OrElse (bCheckGuild = True AndAlso oTmpEntity.bGuildMember = True) Then
                            If (oTmpEntity.CurrentStatus And elUnitStatus.eRadarOperational) = 0 OrElse (oTmpEntity.ObjTypeID = ObjectType.eFacility AndAlso (oEntity(X).CurrentStatus And elUnitStatus.eFacilityPowered) = 0) Then
                                fMaxRadiiSum = 0
                                fOptRadiiSum = fSphereRadius + (oTmpEntity.oMesh.RangeOffset * gl_FINAL_GRID_SQUARE_SIZE)
                            Else
                                fMaxRadiiSum = fSphereRadius + (oTmpEntity.oUnitDef.FOW_MaxRadarRange)
                                fOptRadiiSum = fSphereRadius + (oTmpEntity.oUnitDef.FOW_OptRadarRange)
                            End If

                            fDX = lLocX - oTmpEntity.LocX
                            fDZ = lLocZ - oTmpEntity.LocZ
                            fDX *= fDX
                            fDZ *= fDZ

                            fDist = CSng(Math.Sqrt(fDX + fDZ))

                            If fDist < fOptRadiiSum Then
                                'immediately return Visible
                                Return eVisibilityType.Visible
                            ElseIf fDist < fMaxRadiiSum Then
                                yRes = eVisibilityType.InMaxRange        'store the value for later, there mightr be an optimum range
                            End If
                        End If

                    End If
				End If
			Next X
        Catch
            yRes = eVisibilityType.NoVisibility
        End Try

        Return yRes
    End Function

    Public Function ReadySysView2Points(ByVal bRenderSelectedRngs As Boolean) As ArrayList
        Dim X As Int32

        Dim lR As Int32
        Dim lG As Int32
        Dim lB As Int32

        Dim lInvis As Int32 = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

        lR = Math.Min(255, Math.Max(0, CInt(muSettings.MyAssetColor.X * 255)))
        lG = Math.Min(255, Math.Max(0, CInt(muSettings.MyAssetColor.Y * 255)))
        lB = Math.Min(255, Math.Max(0, CInt(muSettings.MyAssetColor.Z * 255)))
        Dim lPlayer As Int32 = System.Drawing.Color.FromArgb(255, lR, lG, lB).ToArgb

        lR = Math.Min(255, Math.Max(0, CInt(muSettings.EnemyAssetColor.X * 255)))
        lG = Math.Min(255, Math.Max(0, CInt(muSettings.EnemyAssetColor.Y * 255)))
        lB = Math.Min(255, Math.Max(0, CInt(muSettings.EnemyAssetColor.Z * 255)))
        Dim lEnemy As Int32 = System.Drawing.Color.FromArgb(255, lR, lG, lB).ToArgb

        lR = Math.Min(255, Math.Max(0, CInt(muSettings.AllyAssetColor.X * 255)))
        lG = Math.Min(255, Math.Max(0, CInt(muSettings.AllyAssetColor.Y * 255)))
        lB = Math.Min(255, Math.Max(0, CInt(muSettings.AllyAssetColor.Z * 255)))
        Dim lAlly As Int32 = System.Drawing.Color.FromArgb(255, lR, lG, lB).ToArgb

        lR = Math.Min(255, Math.Max(0, CInt(muSettings.NeutralAssetColor.X * 255)))
        lG = Math.Min(255, Math.Max(0, CInt(muSettings.NeutralAssetColor.Y * 255)))
        lB = Math.Min(255, Math.Max(0, CInt(muSettings.NeutralAssetColor.Z * 255)))
        Dim lNeutral As Int32 = System.Drawing.Color.FromArgb(255, lR, lG, lB).ToArgb

        lR = Math.Min(255, Math.Max(0, CInt(muSettings.GuildAssetColor.X * 255)))
        lG = Math.Min(255, Math.Max(0, CInt(muSettings.GuildAssetColor.Y * 255)))
        lB = Math.Min(255, Math.Max(0, CInt(muSettings.GuildAssetColor.Z * 255)))
        Dim lGuild As Int32 = System.Drawing.Color.FromArgb(255, lR, lG, lB).ToArgb

        Dim yRel As Byte
        Dim alList As New ArrayList()

        If moPoints Is Nothing OrElse moPoints.Length - 1 <> lEntityUB Then
            ReDim moPoints(lEntityUB)
        End If

        Dim bRenderDots As Boolean = goCurrentPlayer.ShowSpaceDots

        Dim bCheckGuild As Boolean = False
        Dim oGuild As Guild = Nothing
        If goCurrentPlayer Is Nothing = False Then
            oGuild = goCurrentPlayer.oGuild
            If oGuild Is Nothing = False Then
                bCheckGuild = (goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.ShareUnitVision) <> 0
            End If
        End If

        'Now, go through and do it to it
        Try
            For X = 0 To lEntityUB
                If lEntityIdx(X) <> -1 Then

                    Dim oTmpEntity As BaseEntity = oEntity(X)

                    If oTmpEntity Is Nothing = False Then
                        With oTmpEntity
                            If .OwnerID = glPlayerID Then 'OrElse True = True Then '
                                If .bSelected = True AndAlso bRenderSelectedRngs = True Then alList.Add(X)
                                .yVisibility = eVisibilityType.Visible
                                If muSettings.gbPlanetMapDontShowMyUnits = False Then
                                    moPoints(X).Color = lPlayer
                                Else : moPoints(X).Color = lInvis
                                End If

                            ElseIf (bCheckGuild = True AndAlso .bGuildMember = True) Then
                                .yVisibility = eVisibilityType.Visible
                                If muSettings.gbPlanetMapDontShowGuilds = False Then
                                    moPoints(X).Color = lGuild
                                Else : moPoints(X).Color = lInvis
                                End If
                            ElseIf bRenderDots = True AndAlso .oMesh Is Nothing = False Then '.oMesh Is Nothing = False Then '
                                .yVisibility = UnitInRadarRange(CInt(.LocX), CInt(.LocZ), .oMesh.RangeOffset)
                                '.yVisibility = eVisibilityType.Visible

                                If .yVisibility = eVisibilityType.Visible Then
                                    If .yRelID = Byte.MaxValue OrElse .yRelID = 0 Then
                                        yRel = goCurrentPlayer.GetPlayerRelScore(.OwnerID)
                                        .yRelID = yRel
                                    Else : yRel = .yRelID
                                    End If

                                    If yRel <= elRelTypes.eWar Then
                                        If muSettings.gbPlanetMapDontShowEnemy = False Then
                                            moPoints(X).Color = lEnemy
                                        Else : moPoints(X).Color = lInvis
                                        End If
                                    ElseIf yRel <= elRelTypes.ePeace Then   'elRelTypes.eNeutral
                                        If muSettings.gbPlanetMapDontShowNeutral = False Then
                                            moPoints(X).Color = lNeutral
                                        Else : moPoints(X).Color = lInvis
                                        End If
                                    Else
                                        If muSettings.gbPlanetMapDontShowAllied = False Then
                                            moPoints(X).Color = lAlly
                                        Else : moPoints(X).Color = lInvis
                                        End If
                                    End If
                                ElseIf .yVisibility = eVisibilityType.InMaxRange Then
                                    If muSettings.gbPlanetMapDontShowUnknown = False Then
                                        moPoints(X).Color = lNeutral
                                    Else : moPoints(X).Color = lInvis
                                    End If
                                Else : moPoints(X).Color = lInvis

                                End If
                            Else : moPoints(X).Color = lInvis
                            End If
                            moPoints(X).Position = New Microsoft.DirectX.Vector3(.LocX / 30.0F, .LocY / 30.0F, .LocZ / 30.0F)
                        End With
                    End If

                Else
                    moPoints(X).Color = lInvis
                    moPoints(X).Position = New Microsoft.DirectX.Vector3(Single.MaxValue, Single.MaxValue, Single.MaxValue)
                End If
            Next X
        Catch
        End Try

        Return alList
    End Function

    Public Sub ReadySysView1Points()
        Dim X As Int32
        Dim lPlayer As Int32 = System.Drawing.Color.FromArgb(255, 15, 255, 15).ToArgb

        Dim ptVals() As System.Drawing.Point = Nothing
        Dim yType() As Byte = Nothing
        Dim lPtValUB As Int32 = -1

        Dim lX As Int32
        Dim lZ As Int32
        Dim bFound As Boolean = False

        Dim bCheckGuild As Boolean = False
        Dim oGuild As Guild = Nothing
        If goCurrentPlayer Is Nothing = False Then
            oGuild = goCurrentPlayer.oGuild
            If oGuild Is Nothing = False Then
                bCheckGuild = (goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.ShareUnitVision) <> 0
            End If
        End If

        'Now, go through and do it to it
        For X = 0 To lEntityUB
            If lEntityIdx(X) <> -1 Then
                If oEntity(X).OwnerID = glPlayerID OrElse (bCheckGuild = True AndAlso oEntity(X).bGuildMember = True) Then 'If True = True OrElse oEntity(X).OwnerID = glPlayerID Then '
                    lX = CInt(oEntity(X).LocX) \ 10000
                    lZ = CInt(oEntity(X).LocZ) \ 10000
                    bFound = False
                    For lIdx As Int32 = 0 To lPtValUB
                        If ptVals(lIdx).X = lX AndAlso ptVals(lIdx).Y = lZ Then
                            bFound = True
                        End If
                    Next lIdx

                    If bFound = False Then
                        lPtValUB += 1
                        ReDim Preserve ptVals(lPtValUB)
                        ReDim Preserve yType(lPtValUB)
                        ptVals(lPtValUB).X = lX : ptVals(lPtValUB).Y = lZ

                        If oEntity(X).OwnerID <> glPlayerID Then
                            If muSettings.gbPlanetMapDontShowGuilds = False Then
                                If yType(lPtValUB) <> 1 Then yType(lPtValUB) = 2
                            Else : yType(lPtValUB) = 0
                            End If
                        Else
                            If muSettings.gbPlanetMapDontShowMyUnits = False Then
                                yType(lPtValUB) = 1
                            Else : yType(lPtValUB) = 0
                            End If
                        End If
                    End If
                End If
            End If
        Next X

        If lPtValUB > -1 Then
            Dim lR As Int32 = CInt(muSettings.GuildAssetColor.X * 255)
            Dim lG As Int32 = CInt(muSettings.GuildAssetColor.Y * 255)
            Dim lB As Int32 = CInt(muSettings.GuildAssetColor.Z * 255)
            lR = Math.Min(255, Math.Max(0, lR))
            lG = Math.Min(255, Math.Max(0, lG))
            lB = Math.Min(255, Math.Max(0, lB))

            Dim lGuildClr As Int32 = System.Drawing.Color.FromArgb(255, lR, lG, lB).ToArgb
            lR = CInt(muSettings.MyAssetColor.X * 255)
            lG = CInt(muSettings.MyAssetColor.Y * 255)
            lB = CInt(muSettings.MyAssetColor.Z * 255)
            lR = Math.Min(255, Math.Max(0, lR))
            lG = Math.Min(255, Math.Max(0, lG))
            lB = Math.Min(255, Math.Max(0, lB))
            Dim lPlayerClr As Int32 = System.Drawing.Color.FromArgb(255, lR, lG, lB).ToArgb
            Dim lInvis As Int32 = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

            If moPoints Is Nothing OrElse moPoints.GetUpperBound(0) <> lPtValUB Then
                ReDim moPoints(lPtValUB)
            End If
            For X = 0 To lPtValUB
                moPoints(X).Position = New Microsoft.DirectX.Vector3(ptVals(X).X, 0, ptVals(X).Y)
                If yType(X) = 1 Then
                    moPoints(X).Color = lPlayerClr
                ElseIf yType(X) = 2 Then
                    moPoints(X).Color = lGuildClr
                Else
                    moPoints(X).Color = lInvis
                End If
            Next X
        Else
            ReDim moPoints(-1)
        End If

    End Sub

    'Private Sub InitializeFOWGrid()
    '    Erase oFOWGrid          'remove it if it exists...
    '    ReDim oFOWGrid(610, 610)        'redim it
    '    Dim X As Int32
    '    Dim Z As Int32

    '    Dim lCenterX As Int32 = CInt(Math.Floor(611 / 2))
    '    Dim lCenterZ As Int32 = CInt(Math.Floor(611 / 2))

    '    For Z = 0 To 610
    '        For X = 0 To 610
    '            oFOWGrid(X, Z) = New FOWGridSquare()
    '            With oFOWGrid(X, Z)
    '                .lSquareX = (X - lCenterX) * 16384
    '                .lSquareZ = (Z - lCenterZ) * 16384
    '            End With
    '        Next X
    '    Next Z
    'End Sub

    'Public Sub New()
    '    InitializeFOWGrid()
    'End Sub

    Public Sub SetPrimaryAmbience()
		Dim sName As String

		If goSound Is Nothing Then Return

        If mlPrimaryAmbienceID <> -1 Then
            'tell the sound engine to stop that ambient sound, it will be collected by the garbage collector
            goSound.StopSound(mlPrimaryAmbienceID)
        End If

        'Now, lookup our new sound
        Dim oRand As New Random
        If Me.ObjTypeID = ObjectType.ePlanet Then
            If oGeoObject Is Nothing = False Then
                'TODO: Right now, I assume only 1 sound FX for each planet type, we may want to change that
                Dim yMapTypeID As Byte = CType(oGeoObject, Planet).MapTypeID
                Select Case yMapTypeID
                    Case PlanetType.eAcidic : sName = "Acid1.wav" '"Ambience\AcidPlanetAmbience1.wav"
                    Case PlanetType.eAdaptable
                        If oRand.Next(0, 101) < 50 Then
                            sName = "Adapt1.wav"
                        Else
                            sName = "Adapt2.wav"
                        End If
                        'sName = "Ambience\AdaptablePlanetAmbience1.wav"
                    Case PlanetType.eBarren : sName = "Barren1.wav"
                    Case PlanetType.eDesert : sName = "Desert1.wav"
                    Case PlanetType.eGeoPlastic
                        'sName = "Ambience\GeoPlasticPlanetAmbience1.wav"
                        Dim lRoll As Int32 = oRand.Next(0, 101)
                        If lRoll < 50 Then
                            sName = "Lava1.wav"
                            'ElseIf lRoll < 50 Then
                            '    sName = "Lava2.wav"
                            'ElseIf lRoll < 75 Then
                            '    sName = "Lava3.wav"
                            'Else : sName = "Lava4.wav"
                        Else : sName = "Lava3.wav"
                        End If
                    Case PlanetType.eTerran : sName = "Terran1.wav" '"Ambience\TerranPlanetAmbience1.wav"
                    Case PlanetType.eTundra : sName = "Tundra1.wav" '"Ambience\TundraPlanetAmbience1.wav"
                    Case PlanetType.eWaterWorld : sName = "Ocean1.wav" ' "Ambience\WaterworldPlanetAmbience1.wav"
                    Case Else
                        sName = ""
                End Select
                goSound.UpdateSoundVolumes()
                If sName <> "" Then mlPrimaryAmbienceID = goSound.StartSound(sName, True, SoundMgr.SoundUsage.eAmbience, Nothing, Nothing)
            End If
        Else
            'TODO: currently, we do not use an ambience effect for the other environments
        End If

    End Sub

    Public Sub UpdateEventAmbience()
        'Dim X As Int32
        Dim lTemp As Int32
        'Dim fElapsed As Single

        'Limit updates a bit to better distribute processing power
        If msw_Ambience Is Nothing Then msw_Ambience = Stopwatch.StartNew
        If msw_Ambience.IsRunning = False Then
            msw_Ambience.Reset()
            msw_Ambience.Start()
        End If
        If msw_Ambience.ElapsedMilliseconds < 5000 Then Return
        msw_Ambience.Reset()

        If goSound Is Nothing Then Return

        'Ok, right now, only GeoPlastic planets create random events on the client (unsupervised by the server)
        If Me.ObjTypeID = ObjectType.ePlanet Then
            If oGeoObject Is Nothing = False Then
                If CType(oGeoObject, Planet).MapTypeID = PlanetType.eGeoPlastic Then
                    'Ok, check for chance to create random event
                    If Rnd() < mf_GEO_PLASTIC_EVENT_CHANCE Then
                        'Create a new event... we don't particularly care about tracking this event, so run it as a 1 time
                        Dim oRand As New Random
                        lTemp = oRand.Next(1, 7)  'CInt(Int(Rnd() * 3) + 1)      '3 types of plateshift events
                        goSound.StartSound("Plate" & lTemp & ".wav", False, SoundMgr.SoundUsage.eWeather, Nothing, Nothing)
                    End If
                Else
                    ''Ok, now, go thru any events currently in progress...
                    'For X = 0 To mlEventAmbienceUB
                    '    If muEventAmbiences(X).bActive = True Then
                    '        'Case 1      'At Full Volume, do nothing, server will tell us when to quit
                    '        With muEventAmbiences(X)
                    '            If .lLastUpdate = 0 Then fElapsed = 1 Else fElapsed = CSng((timeGetTime - .lLastUpdate) / 30)

                    '            Select Case .yStateID
                    '                Case 0      'Fading In
                    '                    .fLinearVolume += fElapsed
                    '                    If .fLinearVolume >= ml_MAX_LINEAR_VOLUME Then
                    '                        .fLinearVolume = ml_MAX_LINEAR_VOLUME
                    '                        .yStateID = 1    'at full volume
                    '                    End If
                    '                    lTemp = goSound.LinearVolumeToDirectX(CInt(Math.Floor(.fLinearVolume)), ml_MAX_LINEAR_VOLUME)
                    '                    goSound.SetSoundVolume(.lSoundID, lTemp)
                    '                Case 2      'Fading out
                    '                    .fLinearVolume -= fElapsed
                    '                    If .fLinearVolume <= 0 Then
                    '                        .fLinearVolume = 0
                    '                        .bActive = False
                    '                        goSound.StopSound(.lSoundID)
                    '                    Else
                    '                        lTemp = goSound.LinearVolumeToDirectX(CInt(Math.Floor(.fLinearVolume)), ml_MAX_LINEAR_VOLUME)
                    '                        goSound.SetSoundVolume(.lSoundID, lTemp)
                    '                    End If
                    '            End Select

                    '            .lLastUpdate = timeGetTime
                    '        End With
                    '    End If
                    'Next X

                End If
            End If
        End If

        goSound.UpdateSoundVolumes()
    End Sub

    Public Sub DisposeMe()
		Dim X As Int32

		If goUILib Is Nothing = False Then
            'goUILib.RemoveWindow("frmCommand")
			goUILib.ClearSelection()
		End If
        If goEntityDeath Is Nothing = False Then goEntityDeath.ClearAll()
        'If goRewards Is Nothing = False Then goRewards.ClearAll()

        Erase oEntity
        Erase oCache
        oGeoObject = Nothing

        Erase moPoints

        If goSound Is Nothing = False Then
            If mlPrimaryAmbienceID <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlPrimaryAmbienceID)
            For X = 0 To mlEventAmbienceUB
                If muEventAmbiences(X).lSoundID <> -1 AndAlso muEventAmbiences(X).bActive Then goSound.StopSound(muEventAmbiences(X).lSoundID)
            Next X
        End If
        
        Erase muEventAmbiences
    End Sub

    Public Function GetEnvirSpeedMod(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal iAngle As Int16) As Single
        If Me.ObjTypeID = ObjectType.ePlanet Then
            If oGeoObject Is Nothing = False Then
                Return CType(oGeoObject, Planet).GetSpeedMod(lLocX, lLocZ, iAngle)
            Else : Return 1.0F
            End If
        Else
            'Ok, must be solar system

            'TODO: Here is where you would put the nebula speed modifier, for example return 0.8

            Return 1.0F
        End If
    End Function

    Public Function LocateClosestMineralCache(ByVal vecLoc As Microsoft.DirectX.Vector3) As Microsoft.DirectX.Vector3
        'ok, set up a rect first...
        Dim lHalfMin As Int32 = gl_MINIMUM_MINING_FAC_SNAP_TO_DISTANCE \ 2
        If NewTutorialManager.TutorialOn = True Then
            If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
                If goCurrentPlayer.lTutorialStep < 200 Then
                    lHalfMin *= 4
                End If
            End If
        End If
        Dim rcTemp As Rectangle = Rectangle.FromLTRB(CInt(vecLoc.X - lHalfMin), CInt(vecLoc.Z - lHalfMin), CInt(vecLoc.X + lHalfMin), CInt(vecLoc.Z + lHalfMin))

        Dim fDist As Single
        Dim lCurrIdx As Int32 = -1
        Dim fTemp As Single

        'Now, check the mienral caches
        For X As Int32 = 0 To lCacheUB
            If lCacheIdx(X) <> -1 AndAlso oCache(X).CacheTypeID = MineralCacheType.eMineable Then
                If rcTemp.Contains(oCache(X).LocX, oCache(X).LocZ) = True Then
                    'ok, contains...
                    If lCurrIdx = -1 Then
                        'Ok, this one is good
                        lCurrIdx = X
                        fDist = Distance(CInt(vecLoc.X), CInt(vecLoc.Z), oCache(X).LocX, oCache(X).LocZ)
                    Else
                        fTemp = Distance(CInt(vecLoc.X), CInt(vecLoc.Z), oCache(X).LocX, oCache(X).LocZ)
                        If fTemp < fDist Then
                            lCurrIdx = X
                            fDist = fTemp
                        End If
                    End If
                End If
            End If
        Next X

        If lCurrIdx = -1 Then
            Return vecLoc
        Else
            Return New Microsoft.DirectX.Vector3(oCache(lCurrIdx).LocX, oCache(lCurrIdx).LocY, oCache(lCurrIdx).LocZ)
        End If
	End Function

    Public Function PositionOnPlanetRing(ByVal vecLoc As Microsoft.DirectX.Vector3) As Boolean
        Dim bResult As Boolean = False
        Try
            'Ok, first, are we a system
            Dim oObj As Object = Me.oGeoObject
            If oObj Is Nothing Then Return bResult
            If CType(oObj, Base_GUID).ObjTypeID <> ObjectType.eSolarSystem Then Return bResult

            Dim oSystem As SolarSystem = CType(oObj, SolarSystem)
            If oSystem Is Nothing Then Return bResult

            Dim fDist As Single = Single.MaxValue
            Dim oPlanet As Planet = Nothing

            For X As Int32 = 0 To oSystem.PlanetUB
                Dim oTmp As Planet = oSystem.moPlanets(X)
                If oTmp Is Nothing = False Then
                    Dim fDX As Single = vecLoc.X - oTmp.LocX
                    Dim fDZ As Single = vecLoc.Z - oTmp.LocZ
                    Dim fTemp As Single = CSng(Math.Sqrt((fDX * fDX) + (fDZ * fDZ)))

                    If fTemp < fDist Then
                        oPlanet = oTmp
                        fDist = fTemp
                    End If
                End If
            Next X

            If oPlanet Is Nothing = False Then
                'now, determine if the planet is ok
                If oPlanet.mbHasRing = True Then
                    If fDist < oPlanet.OuterRingRadius + gl_MINIMUM_MINING_FAC_SNAP_TO_DISTANCE Then
                        bResult = True
                    End If
                End If
            End If

        Catch
        End Try
        Return bResult
    End Function
    Public Function PlaceAlongPlanetRing(ByVal vecLoc As Microsoft.DirectX.Vector3) As Microsoft.DirectX.Vector3
        Try
            'Ok, first, are we a system
            Dim oObj As Object = Me.oGeoObject
            If oObj Is Nothing Then Return vecLoc
            If CType(oObj, Base_GUID).ObjTypeID <> ObjectType.eSolarSystem Then Return vecLoc

            Dim oSystem As SolarSystem = CType(oObj, SolarSystem)
            If oSystem Is Nothing Then Return vecLoc

            Dim fDist As Single = Single.MaxValue
            Dim oPlanet As Planet = Nothing

            For X As Int32 = 0 To oSystem.PlanetUB
                Dim oTmp As Planet = oSystem.moPlanets(X)
                If oTmp Is Nothing = False Then
                    Dim fDX As Single = vecLoc.X - oTmp.LocX
                    Dim fDZ As Single = vecLoc.Z - oTmp.LocZ
                    Dim fTemp As Single = CSng(Math.Sqrt((fDX * fDX) + (fDZ * fDZ)))

                    If fTemp < fDist Then
                        oPlanet = oTmp
                        fDist = fTemp
                    End If
                End If
            Next X

            If oPlanet Is Nothing = False Then
                'now, determine if the planet is ok
                If oPlanet.mbHasRing = True Then
                    If fDist < oPlanet.OuterRingRadius + gl_MINIMUM_MINING_FAC_SNAP_TO_DISTANCE Then
                        Dim fTargetDist As Single = ((oPlanet.OuterRingRadius - oPlanet.InnerRingRadius) \ 2) + oPlanet.InnerRingRadius
                        'Ok, determine the angle
                        Dim fAngle As Single = LineAngleDegrees(oPlanet.LocX, oPlanet.LocZ, CInt(vecLoc.X), CInt(vecLoc.Z))
                        'Now, set up our points
                        Dim fX As Single = fTargetDist
                        Dim fZ As Single = 0
                        RotatePoint(0, 0, fX, fZ, fAngle)

                        Return New Microsoft.DirectX.Vector3(fX + oPlanet.LocX, 0, fZ + oPlanet.LocZ)
                    End If
                End If 
            End If

        Catch
        End Try
        Return vecLoc

 
    End Function

	Public Function IsPlayerIronCurtain(ByVal lPlayerID As Int32) As Boolean
		For X As Int32 = 0 To lIronCurtainPlayerUB
			If lIronCurtainPlayers(X) = lPlayerID Then Return True
		Next X
		Return False
	End Function

	Public Sub DeselectAll()
		Try
			Dim lCurUB As Int32 = -1

			If Me.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(lEntityUB, lEntityIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If lEntityIdx(X) <> -1 Then
					If oEntity(X) Is Nothing = False Then oEntity(X).bSelected = False
				End If
			Next X
			If goUILib Is Nothing = False Then goUILib.ClearSelection()
		Catch
		End Try
		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eDeselectSelection, -1, -1, -1, "")
		End If
    End Sub

    Public Sub DeselectNonPlayerSelected()
        Try
            Dim lCurUB As Int32 = -1

            If Me.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(lEntityUB, lEntityIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If lEntityIdx(X) <> -1 Then
                    If oEntity(X) Is Nothing = False Then
                        If oEntity(X).bSelected = True Then
                            If oEntity(X).OwnerID <> glPlayerID Then
                                oEntity(X).bSelected = False
                                If goUILib Is Nothing = False Then goUILib.RemoveSelection(X)
                            End If
                        End If
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub

	Public Sub LoadEnvironmentTacticalData()
		Dim sFilePath As String = AppDomain.CurrentDomain.BaseDirectory
		If sFilePath.EndsWith("\") = False Then sFilePath &= "\"
		sFilePath &= "Data\" & Hex(glPlayerID) & "\"
		Dim lTmp As Int32 = Me.ObjectID
		If Me.ObjTypeID = ObjectType.eSolarSystem Then lTmp = -lTmp
		sFilePath &= Hex(lTmp) & ".dat"

        If Exists(sFilePath) = False Then Return

        Dim oFS As New System.IO.FileStream(sFilePath, IO.FileMode.Open)
        Try
            'Ok, get the tactical data's version #
            Dim lVerNum As Int32 = oFS.ReadByte()

            If lVerNum = 2 Then
                While oFS.Position + 63 < oFS.Length
                    Dim yData(63) As Byte
                    Dim lBytesRead As Int32 = oFS.Read(yData, 0, 64)
                    If lBytesRead = 0 Then Exit While

                    Dim lPos As Int32 = 0
                    Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim lX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim lZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    Dim yArmorHP(3) As Byte
                    Array.Copy(yData, lPos, yArmorHP, 0, 4) : lPos += 4
                    Dim yShieldHP As Byte = yData(lPos) : lPos += 1
                    Dim yStructure As Byte = yData(lPos) : lPos += 1

                    Dim lEngineFXClr As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim lShieldFXClr As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    Dim sPName As String = GetStringFromBytes(yData, lPos, 30) : lPos += 30
                    sPName = sPName.Trim()
                    sPName = Base64.DecodeToString(sPName)

                    Dim oTmp As BaseEntity = Nothing
                    Dim iTypeID As Int16
                    If lID < 0 Then iTypeID = ObjectType.eFacility Else iTypeID = ObjectType.eUnit
                    lID = Math.Abs(lID)
                    For X As Int32 = 0 To lEntityUB
                        If lEntityIdx(X) = lID AndAlso oEntity(X).ObjTypeID = iTypeID Then
                            oTmp = oEntity(X)
                            Exit For
                        End If
                    Next X
                    Dim bAdded As Boolean = False
                    If oTmp Is Nothing Then
                        oTmp = New BaseEntity
                        oTmp.OwnerID = -1
                        bAdded = True
                    End If

                    With oTmp
                        .ObjTypeID = iTypeID
                        .ObjectID = Math.Abs(lID)
                        .LocX = lX
                        .LocZ = lZ
                        .LocAngle = iA
                        If .OwnerID < 1 Then .OwnerID = lOwnerID
                        .sOverrideOwnerName = sPName
                        If .oUnitDef Is Nothing Then .oUnitDef = New EntityDef
                        .oUnitDef.ModelID = iModelID
                        .oMesh = goResMgr.GetMesh(iModelID)
                        .oUnitDef.Maneuver = 0
                        .oUnitDef.MaxSpeed = 0
                        .yArmorHP = yArmorHP
                        .yShieldHP = yShieldHP
                        .yStructureHP = yStructure
                        .clrEngineFX = System.Drawing.Color.FromArgb(lEngineFXClr)
                        .clrShieldFX = System.Drawing.Color.FromArgb(lShieldFXClr)
                        .CellLocX = CInt(Math.Floor(.LocX / muSettings.EntityClipPlane))
                        .CellLocZ = CInt(Math.Floor(.LocZ / muSettings.EntityClipPlane))
                        .yVisibility = eVisibilityType.Visible
                        .yRelID = 0

                        If bAdded = True Then .bObjectDestroyed = True

                        .bForceGetWorldMatrix = True
                        '.GetWorldMatrix()		'do this to set the current matrix					
                    End With

                    If bAdded = True Then
                        oTmp.bObjectDestroyed = True
                        oTmp.yVisibility = eVisibilityType.FacilityIntel
                        Dim lIdx As Int32 = goCurrentEnvir.lEntityUB + 1
                        ReDim Preserve lEntityIdx(lIdx)
                        ReDim Preserve oEntity(lIdx)
                        oEntity(lIdx) = oTmp
                        lEntityIdx(lIdx) = oTmp.ObjectID
                        goCurrentEnvir.lEntityUB = lIdx
                    End If
                End While
            End If
        Catch
        End Try

		oFS.Close()
		oFS.Dispose()
        oFS = Nothing

        Try
            For X As Int32 = 0 To glItemIntelUB
                If glItemIntelIdx(X) > -1 Then
                    Dim oII As PlayerItemIntel = goItemIntel(X)
                    If oII Is Nothing = False AndAlso oII.iItemTypeID = ObjectType.eFacility Then
                        If oII.EnvirID = Me.ObjectID AndAlso oII.EnvirTypeID = Me.ObjTypeID Then

                            Dim oTmp As BaseEntity = Nothing
                            For Y As Int32 = 0 To lEntityUB
                                If lEntityIdx(Y) = oII.lItemID AndAlso oEntity(Y).ObjTypeID = ObjectType.eFacility Then
                                    oTmp = oEntity(Y)
                                    Exit For
                                End If
                            Next Y
                            Dim bAdded As Boolean = False
                            If oTmp Is Nothing Then
                                oTmp = New BaseEntity
                                oTmp.OwnerID = -1
                                bAdded = True
                            End If


                            With oTmp
                                .ObjTypeID = ObjectType.eFacility
                                .ObjectID = oII.lItemID
                                .LocX = oII.LocX
                                .LocZ = oII.LocZ
                                .LocAngle = 0
                                If .OwnerID < 1 Then .OwnerID = oII.lOtherPlayerID
                                
                                .CellLocX = CInt(Math.Floor(.LocX / muSettings.EntityClipPlane))
                                .CellLocZ = CInt(Math.Floor(.LocZ / muSettings.EntityClipPlane))
                                .yVisibility = eVisibilityType.Visible

                                If bAdded = True Then
                                    .yArmorHP(0) = 100 : .yArmorHP(1) = 100 : .yArmorHP(2) = 100 : .yArmorHP(3) = 100
                                    .yShieldHP = 100
                                    .yStructureHP = 100
                                    .bObjectDestroyed = True
                                    If .oUnitDef Is Nothing = False Then .oUnitDef = New EntityDef
                                    .oUnitDef.ModelID = CShort(oII.lValue)
                                    .oMesh = goResMgr.GetMesh(.oUnitDef.ModelID)
                                    .oUnitDef.Maneuver = 0
                                    .oUnitDef.MaxSpeed = 0
                                End If

                                .bForceGetWorldMatrix = True
                            End With

                            If bAdded = True Then
                                oTmp.bObjectDestroyed = True
                                oTmp.yVisibility = eVisibilityType.FacilityIntel
                                Dim lIdx As Int32 = goCurrentEnvir.lEntityUB + 1
                                ReDim Preserve lEntityIdx(lIdx)
                                ReDim Preserve oEntity(lIdx)
                                oEntity(lIdx) = oTmp
                                lEntityIdx(lIdx) = oTmp.ObjectID
                                goCurrentEnvir.lEntityUB = lIdx
                            End If
                        End If
                    End If
                End If
            Next X
        Catch
        End Try

    End Sub

	Public Sub SaveEnvironmentTacticalData()

		'On a per line basis...
		'{
		'	EntityID (4)	---- negative if facility
		'	LocX (4)
		'	LocZ (4)
		'	LocA (2)
		'	OwnerID (4)
		'	ModelID (2)
		'	HPData (6)
		'	enginefx clr (4)
		'	shieldfx clr (4)
		'}

		Dim lCurUB As Int32 = -1
		If lEntityIdx Is Nothing = False Then lCurUB = Math.Min(lEntityUB, lEntityIdx.GetUpperBound(0))
		If oEntity Is Nothing = False Then lCurUB = Math.Min(lCurUB, oEntity.GetUpperBound(0))

		Dim sFilePath As String = AppDomain.CurrentDomain.BaseDirectory
		If sFilePath.EndsWith("\") = False Then sFilePath &= "\"

		Dim sPlayerFolder As String = Hex(glPlayerID)
		sFilePath &= "Data"
		If Exists(sFilePath) = False Then MkDir(sFilePath)
		sFilePath &= "\" & sPlayerFolder
		If Exists(sFilePath) = False Then MkDir(sFilePath)
		sFilePath &= "\"

		Dim lTmp As Int32 = Me.ObjectID
		If Me.ObjTypeID = ObjectType.eSolarSystem Then lTmp = -lTmp
		sFilePath &= Hex(lTmp) & ".dat"

		Dim oFS As New System.IO.FileStream(sFilePath, IO.FileMode.Create)

        Try
            'export the version number
            oFS.WriteByte(2)

            For X As Int32 = 0 To lCurUB
                If lEntityIdx(X) > -1 Then
                    Dim oTmp As BaseEntity = oEntity(X)
                    If oTmp Is Nothing = False AndAlso oTmp.OwnerID <> glPlayerID AndAlso (oTmp.yVisibility = eVisibilityType.FacilityIntel OrElse oTmp.yVisibility = eVisibilityType.Visible) AndAlso oTmp.ObjTypeID = ObjectType.eFacility Then
                        If oTmp.yProductionType = ProductionType.eMining Then Continue For
                        With oTmp
                            Dim yRow(63) As Byte
                            Dim lID As Int32 = .ObjectID
                            Dim lPos As Int32 = 0
                            If .ObjTypeID = ObjectType.eFacility Then lID = -lID
                            System.BitConverter.GetBytes(lID).CopyTo(yRow, lPos) : lPos += 4
                            System.BitConverter.GetBytes(CInt(.LocX)).CopyTo(yRow, lPos) : lPos += 4
                            System.BitConverter.GetBytes(CInt(.LocZ)).CopyTo(yRow, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.LocAngle).CopyTo(yRow, lPos) : lPos += 2
                            System.BitConverter.GetBytes(.OwnerID).CopyTo(yRow, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.oUnitDef.ModelID).CopyTo(yRow, lPos) : lPos += 2
                            .yArmorHP.CopyTo(yRow, lPos) : lPos += 4
                            yRow(lPos) = .yShieldHP : lPos += 1
                            yRow(lPos) = .yStructureHP : lPos += 1

                            System.BitConverter.GetBytes(.clrEngineFX.ToArgb).CopyTo(yRow, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.clrShieldFX.ToArgb).CopyTo(yRow, lPos) : lPos += 4

                            If .sOverrideOwnerName Is Nothing OrElse .sOverrideOwnerName = "" Then
                                .sOverrideOwnerName = GetCacheObjectValue(.OwnerID, ObjectType.ePlayer)
                            End If

                            Dim sEnc As String = Base64.EncodeString(.sOverrideOwnerName)
                            System.Text.ASCIIEncoding.ASCII.GetBytes(sEnc).CopyTo(yRow, lPos) : lPos += 30

                            oFS.Write(yRow, 0, yRow.Length)
                        End With
                    End If
                End If
            Next X
        Catch
        End Try

		oFS.Close()
		oFS.Dispose()
		oFS = Nothing
	End Sub

	Protected Overrides Sub Finalize()
		'SaveEnvironmentTacticalData()
		MyBase.Finalize()
    End Sub

#Region "  Unit Production Queue Management  "
    Public oUnitQueueItem() As UnitProdQueue = Nothing
    Public lUnitQueueItemUB As Int32 = -1

    Public Sub AddUnitProdQueueItem(ByRef oItem As UnitProdQueue)
        Try
            Dim X As Int32 = lUnitQueueItemUB + 1
            For X = X To 1 Step -1
                If oUnitQueueItem(X - 1) Is Nothing = False Then Exit For
            Next

            If X > lUnitQueueItemUB Then
                lUnitQueueItemUB = X
                ReDim Preserve oUnitQueueItem(lUnitQueueItemUB)
            End If

            oUnitQueueItem(lUnitQueueItemUB) = oItem
        Catch
        End Try
    End Sub
    Public Sub RemoveUnitProdQueueItem(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lProdID As Int32, ByVal iProdTypeID As Int16)
        Try
            Dim X As Int32
            Dim bFound As Boolean
            For X = 0 To lUnitQueueItemUB
                If oUnitQueueItem(X) Is Nothing = False Then
                    With oUnitQueueItem(X)
                        If .lProdID = lProdID AndAlso .iProdTypeID = iProdTypeID AndAlso .lBuilderID = lEntityID AndAlso .iBuilderTypeID = iEntityTypeID Then
                            oUnitQueueItem(X) = Nothing
                            bFound = True
                            Exit For
                        End If
                    End With
                End If
            Next X
            If bFound = False Then Return

            For X = X To lUnitQueueItemUB - 1
                oUnitQueueItem(X) = oUnitQueueItem(X + 1)
            Next
            lUnitQueueItemUB -= 1
            ReDim Preserve oUnitQueueItem(lUnitQueueItemUB)
        Catch
        End Try
    End Sub
    Public Function ClearUnitsProdQueue(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16) As Boolean
        'return false if the request needs to go to the server
        Dim bResult As Boolean = False
        Try
            For X As Int32 = 0 To lUnitQueueItemUB
                If oUnitQueueItem(X) Is Nothing = False Then
                    With oUnitQueueItem(X)
                        If .lBuilderID = lEntityID AndAlso .iBuilderTypeID = iEntityTypeID Then
                            oUnitQueueItem(X) = Nothing
                            bResult = True
                        End If
                    End With
                End If
            Next X
        Catch
        End Try
        Return bResult
    End Function
#End Region

    Public Sub JumpToEntity(ByVal lID As Int32, ByVal lTypeID As Int32)
        Try
            For X As Int32 = 0 To lEntityUB
                If lEntityIdx(X) = lID Then
                    Dim oCurr As BaseEntity = oEntity(X)
                    If oCurr Is Nothing = False AndAlso oCurr.ObjectID = lID AndAlso oCurr.ObjTypeID = lTypeID Then


                        If Me.ObjTypeID = ObjectType.ePlanet Then
                            glCurrentEnvirView = CurrentView.ePlanetView
                            goCamera.mlCameraY = 1700
                        Else
                            glCurrentEnvirView = CurrentView.eSystemView
                            goCamera.mlCameraY = 3000
                        End If

                        With goCamera
                            .mlCameraAtX = CInt(oCurr.LocX) : .mlCameraAtY = 0 : .mlCameraAtZ = CInt(oCurr.LocZ)
                            .mlCameraX = .mlCameraAtX : .mlCameraZ = .mlCameraAtZ - 500

                            Try
                                If ObjTypeID = ObjectType.ePlanet Then
                                    goCamera.mlCameraY = CInt(CType(oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

                                    goCamera.mlCameraX += muSettings.lPlanetViewCameraX
                                    goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
                                End If
                            Catch
                            End Try
                        End With
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub
End Class
