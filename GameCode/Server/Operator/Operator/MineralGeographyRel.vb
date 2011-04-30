Option Strict On

Public Class MineralGeographyRel
	Public lMineralID As Int32 = -1
	Public lGeoID As Int32 = -1
	Public iGeoTypeID As Int16 = -1

	Public lRedistMaxQty As Int32 = 0
	Public lRedistMaxConc As Int32 = 0


	Private Shared moRels() As MineralGeographyRel
    Private Shared mlRelUB As Int32 = -1
    Public Shared Sub SpawnNewMineralCache(ByVal lMineralID As Int32, ByVal lOriginalEnvirID As Int32, ByVal iOriginalEnvirTypeID As Int16)
        Dim lCnt As Int32 = 0

        If lMineralID = 41991 Then Return

        LogEvent(LogEventType.Informational, "SpawnNewMineralCache Request Received (" & lOriginalEnvirID & ", " & iOriginalEnvirTypeID & ") for " & lMineralID)
        If lMineralID = 4 OrElse lMineralID = 6 Then
            Dim lPrevMinID As Int32 = lMineralID
            lMineralID = CInt(Rnd() * 104) + 1
            LogEvent(LogEventType.Informational, "  Changing " & lPrevMinID & " to " & lMineralID)
        End If

        For X As Int32 = 0 To mlRelUB
            If moRels(X) Is Nothing = False AndAlso moRels(X).lMineralID = lMineralID AndAlso (moRels(X).lGeoID <> lOriginalEnvirID OrElse moRels(X).iGeoTypeID <> iOriginalEnvirTypeID) Then
                lCnt += 1
            End If
        Next X

        Dim lEnvirID As Int32 = -1
        Dim iEnvirTypeID As Int16 = -1
        Dim lMaxConc As Int32 = 0
        Dim lMaxQty As Int32 = 0
        If lCnt = 0 Then
            'Ok, roll 100
            Dim lRoll As Int32 = CInt(Rnd() * 100)
            '20% of the time, they go to spawns
            '15% of the time, they go to respawns
            '30% of the time, they go to hubs
            '35% of the time, they go to hub hubs
            Dim yType As Byte = 0
            Dim bUnlockedOK As Boolean = False
            If lRoll < 20 Then
                yType = CByte(GeoSpawner.elSystemType.SpawnSystem)
                bUnlockedOK = True
            ElseIf lRoll < 35 Then
                yType = CByte(GeoSpawner.elSystemType.RespawnSystem)
                bUnlockedOK = True
            ElseIf lRoll < 65 Then
                yType = CByte(GeoSpawner.elSystemType.HubSystem)
            Else
                yType = CByte(GeoSpawner.elSystemType.HubHubSystem)
            End If

            For X As Int32 = 0 To glPlanetUB
                If glPlanetIdx(X) > -1 AndAlso goPlanet(X).ParentSystem Is Nothing = False AndAlso (goPlanet(X).ParentSystem.SystemType = yType OrElse (bUnlockedOK = True AndAlso goPlanet(X).ParentSystem.SystemType = GeoSpawner.elSystemType.UnlockedSystem)) Then
                    lCnt += 1
                End If
            Next X

            Dim lVal As Int32 = CInt(Rnd() * (lCnt - 1)) + 1
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To glPlanetUB
                If glPlanetIdx(X) > -1 AndAlso goPlanet(X).ParentSystem Is Nothing = False AndAlso goPlanet(X).ParentSystem.SystemType = yType Then
                    lVal -= 1
                    lIdx = X
                    If lVal = 0 Then Exit For
                End If
            Next X
            If lIdx > -1 Then
                lEnvirID = goPlanet(lIdx).ObjectID
                iEnvirTypeID = goPlanet(lIdx).ObjTypeID
                '50% of time...
                Dim lQty As Int32 = CInt(((Rnd() * 10) + 2) * 400000)
                Dim lCon As Int32 = CInt(Rnd() * 90) + 10
                lQty \= 100000
                lQty -= CInt(Rnd() * 10)
                lQty = Math.Abs(lQty)
                lQty += 1
                lQty *= 100000
                lCon -= CInt(Rnd() * 10)
                lCon = Math.Abs(lCon)
                lCon += 1
                lQty \= 2

                If lMineralID < 7 Then
                    lQty \= 100
                    lQty *= 3
                ElseIf lMineralID = 57 OrElse lMineralID = 64 OrElse lMineralID = 56 OrElse lMineralID = 80 Then
                    lQty \= 4
                End If
                lCon = CInt(Rnd() * 50) + 50

                lMaxQty = lQty
                lMaxConc = lCon
            End If
        Else
            Dim lVal As Int32 = CInt(Rnd() * (lCnt - 1)) + 1
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To mlRelUB
                If moRels(X) Is Nothing = False AndAlso moRels(X).lMineralID = lMineralID AndAlso (moRels(X).lGeoID <> lOriginalEnvirID OrElse moRels(X).iGeoTypeID <> iOriginalEnvirTypeID) Then
                    lVal -= 1
                    lIdx = X
                    If lVal = 0 Then Exit For
                End If
            Next X

            If lIdx > -1 Then
                'ok, found it...
                lEnvirID = moRels(lIdx).lGeoID
                iEnvirTypeID = moRels(lIdx).iGeoTypeID
                lMaxConc = moRels(lIdx).lRedistMaxConc
                lMaxQty = moRels(lIdx).lRedistMaxQty

                Dim lQrtrVal As Int32 = lMaxConc \ 4
                Dim lConc As Int32 = CInt(Rnd() * lQrtrVal) + (lQrtrVal * 3)
                lQrtrVal = lMaxQty \ 4
                Dim lQty As Int32 = CInt(Rnd() * lQrtrVal) + (lQrtrVal * 3)

                lMaxConc = lConc
                lMaxQty = lQty

                If iEnvirTypeID <> ObjectType.ePlanet AndAlso iEnvirTypeID <> ObjectType.eSolarSystem Then
                    lIdx = CInt(Rnd() * glPlanetUB)
                    If goPlanet(lIdx) Is Nothing = False Then
                        lEnvirID = goPlanet(lIdx).ObjectID
                        iEnvirTypeID = ObjectType.ePlanet
                    End If
                End If
            End If
        End If

        If lEnvirID > 0 AndAlso iEnvirTypeID > 0 AndAlso lMaxConc > 0 AndAlso lMaxQty > 0 Then

            LogEvent(LogEventType.Informational, "Spawning new mineral cache from " & lOriginalEnvirID & ", " & iOriginalEnvirTypeID & " to " & lEnvirID & ", " & iEnvirTypeID & " (" & lMaxConc & "/" & lMaxQty & ")")

            Dim yMsg(21) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(ObjectType.eMineral).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lMaxQty).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lMaxConc).CopyTo(yMsg, lPos) : lPos += 4

            'Now, determine which server is in charge of that environment...
            If iEnvirTypeID = ObjectType.ePlanet Then
                Dim oPlanet As Planet = GetEpicaPlanet(lEnvirID)
                If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then
                    If oPlanet.ParentSystem.oPrimaryServer Is Nothing = False Then
                        oPlanet.ParentSystem.oPrimaryServer.oSocket.SendData(yMsg)
                    End If
                End If
                'Dim oplanet As Planet
                'oplanet.ParentSystem.oPrimaryServer()
            ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
                Dim oSystem As SolarSystem = GetEpicaSystem(lEnvirID)
                If oSystem Is Nothing = False AndAlso oSystem.oPrimaryServer Is Nothing = False Then
                    oSystem.oPrimaryServer.oSocket.SendData(yMsg)
                End If
            End If
        End If

    End Sub
    'Public Shared Sub SpawnNewMineralCache(ByVal lMineralID As Int32, ByVal lOriginalEnvirID As Int32, ByVal iOriginalEnvirTypeID As Int16)
    '	'ok, a server is telling us that a mineralcache has depleted in the environment passed in... need to respawn it somewhere
    '	Dim lCnt As Int32 = 0
    '	For X As Int32 = 0 To mlRelUB
    '		If moRels(X) Is Nothing = False AndAlso moRels(X).lMineralID = lMineralID AndAlso (moRels(X).lGeoID <> lOriginalEnvirID OrElse moRels(X).iGeoTypeID <> iOriginalEnvirTypeID) Then
    '			lCnt += 1
    '		End If
    '	Next X

    '	Dim lVal As Int32 = CInt(Rnd() * (lCnt - 1)) + 1
    '	Dim lIdx As Int32 = -1
    '	For X As Int32 = 0 To mlRelUB
    '		If moRels(X) Is Nothing = False AndAlso moRels(X).lMineralID = lMineralID AndAlso (moRels(X).lGeoID <> lOriginalEnvirID OrElse moRels(X).iGeoTypeID <> iOriginalEnvirTypeID) Then
    '			lVal -= 1
    '			lIdx = X
    '			If lVal = 0 Then Exit For
    '		End If
    '	Next X

    '	If lIdx > -1 Then
    '		'ok, found it...
    '		Dim lEnvirID As Int32 = moRels(lIdx).lGeoID
    '		Dim iEnvirTypeID As Int16 = moRels(lIdx).iGeoTypeID
    '		Dim lMaxConc As Int32 = moRels(lIdx).lRedistMaxConc
    '		Dim lMaxQty As Int32 = moRels(lIdx).lRedistMaxQty

    '		If iEnvirTypeID <> ObjectType.ePlanet AndAlso iEnvirTypeID <> ObjectType.eSolarSystem Then
    '			lIdx = CInt(Rnd() * glPlanetUB)
    '			If goPlanet(lIdx) Is Nothing = False Then
    '				lEnvirID = goPlanet(lIdx).ObjectID
    '				iEnvirTypeID = ObjectType.ePlanet
    '			End If
    '		End If

    '		Dim yMsg(21) As Byte
    '		Dim lPos As Int32 = 0
    '           System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
    '		System.BitConverter.GetBytes(lMineralID).CopyTo(yMsg, lPos) : lPos += 4
    '		System.BitConverter.GetBytes(ObjectType.eMineral).CopyTo(yMsg, lPos) : lPos += 2
    '		System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
    '		System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
    '		System.BitConverter.GetBytes(lMaxQty).CopyTo(yMsg, lPos) : lPos += 4
    '		System.BitConverter.GetBytes(lMaxConc).CopyTo(yMsg, lPos) : lPos += 4

    '		'Now, determine which server is in charge of that environment...
    '		If iEnvirTypeID = ObjectType.ePlanet Then
    '			Dim oPlanet As Planet = GetEpicaPlanet(lEnvirID)
    '			If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then
    '				If oPlanet.ParentSystem.oPrimaryServer Is Nothing = False Then
    '					oPlanet.ParentSystem.oPrimaryServer.oSocket.SendData(yMsg)
    '				End If
    '			End If
    '			'Dim oplanet As Planet
    '			'oplanet.ParentSystem.oPrimaryServer()
    '		ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
    '			Dim oSystem As SolarSystem = GetEpicaSystem(lEnvirID)
    '			If oSystem Is Nothing = False AndAlso oSystem.oPrimaryServer Is Nothing = False Then
    '				oSystem.oPrimaryServer.oSocket.SendData(yMsg)
    '			End If
    '		End If
    '	End If

    'End Sub

	Public Shared Sub AddMineralGeographyRel(ByVal plMineralID As Int32, ByVal plGeoID As Int32, ByVal piGeoTypeID As Int16, ByVal plMaxQty As Int32, ByVal plMaxConc As Int32)
		mlRelUB += 1
		ReDim Preserve moRels(mlRelUB)
		moRels(mlRelUB) = New MineralGeographyRel
		With moRels(mlRelUB)
			.lGeoID = plGeoID
			.iGeoTypeID = piGeoTypeID
			.lMineralID = plMineralID
			.lRedistMaxConc = plMaxConc
			.lRedistMaxQty = plMaxQty
		End With
	End Sub

End Class
