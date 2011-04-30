Option Strict On

Public Class GeoSpawner

	Private Shared mlToSpawnCount As Int32 = 0
	Private Shared mbInSpawn As Boolean = False
	Private Shared moPrimary As ServerObject = Nothing
	Private Shared moRegion As ServerObject = Nothing

	Private Shared moCurrHHSys As SolarSystem = Nothing			'TODO: need to load this when the server comes up

    Public Shared mfSystemSpawnRadius As Single = 0.0F
	Public Shared mfSystemSpawnRadiusChange As Single = 0.1F
	Private Const gdRadPerDegree As Single = Math.PI / 180.0F
	Private Const gdHalfPie As Single = Math.PI / 2.0F
	Private Const gdPieAndAHalf As Single = Math.PI * 1.5F
	Private Const gdPi As Single = Math.PI
	Private Const gdTwoPie As Single = Math.PI * 2.0F
    Private Const gdDegreePerRad As Single = 180.0F / gdPi

    Private Shared myLastCheckPrimaryMsg() As Byte

    Public Shared Sub IncrementGenerateCounter()
        If mbInSpawn = False Then mlToSpawnCount += 1
    End Sub

    Private Shared Sub RotatePoint(ByVal fAxisX As Single, ByVal fAxisY As Single, ByRef fEndX As Single, ByRef fEndY As Single, ByVal fDegree As Single)
        Dim fDX As Single
        Dim fDY As Single
        Dim fRads As Single

        fRads = fDegree * gdRadPerDegree 'CSng(Math.PI / 180.0F)
        fDX = fEndX - fAxisX
        fDY = fEndY - fAxisY

        Dim fCosRads As Single = CSng(Math.Cos(fRads))
        Dim fSinRads As Single = CSng(Math.Sin(fRads))

        fEndX = fAxisX + ((fDX * fCosRads) + (fDY * fSinRads))
        fEndY = fAxisY + -((fDX * fSinRads) - (fDY * fCosRads))
    End Sub

	Private Shared Function LineAngleDegrees(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Single
		Dim dDeltaX As Single
		Dim dDeltaY As Single
		Dim dAngle As Single

		dDeltaX = lX2 - lX1
		dDeltaY = lY2 - lY1

		If dDeltaX = 0 Then		'vertical
			If dDeltaY < 0 Then
				dAngle = gdHalfPie
			Else
				dAngle = gdPieAndAHalf
			End If
		ElseIf dDeltaY = 0 Then		'horizontal
			If dDeltaX < 0 Then
				dAngle = gdPi
			Else
				dAngle = 0
			End If
		Else	'angled
			dAngle = CSng(System.Math.Atan(System.Math.Abs(dDeltaY / dDeltaX)))
			'Correct for VB's reversed Y... VB Upper Right is ok... but the other quads are not
			If dDeltaX > -1 And dDeltaY > -1 Then		'VB Lower Right
				dAngle = gdTwoPie - dAngle
			ElseIf dDeltaX < 0 And dDeltaY > -1 Then	'VB Lower Left
				dAngle = gdPi + dAngle
			ElseIf dDeltaX < 0 And dDeltaY < 0 Then		'VB Upper Left
				dAngle = gdPi - dAngle
			End If
		End If

		'Not sure this is suppose to be CINT
		'Return CInt(dAngle * gdDegreePerRad)
		Return dAngle * gdDegreePerRad

	End Function

	Public Enum elSystemType As Int32
		SpawnSystem = 0
		HubSystem = 1
		RespawnSystem = 2
        HubHubSystem = 3
        UnlockedSystem = 4
		TutorialSystem = 255
	End Enum

	Public Shared Sub GenerateNewNeighborhood()
		If mbInSpawn = True Then Return
		If mlToSpawnCount = 0 Then Return
		mbInSpawn = True
		mlToSpawnCount -= 1

		moPrimary = Nothing
		moRegion = Nothing

        LogEvent(LogEventType.Informational, "Generating new neighborhood")
		'ok, this method is called because the operator has detected by reference of the collection of primaries that 
		' the galaxy is shrinking and we need to expand it. So, we are going to add a new solar system neighborhood.

		'First, however, we need to ensure we have the server architecture in place to handle the new neighborhood
		Dim oPrimary As ServerObject = goMsgSys.FindNeighborhoodSuitor(ConnectionType.ePrimaryServerApp, Nothing)
		If oPrimary Is Nothing Then
            'TODO: ok, we need to spawn a primary server
            Dim oEmail As ServerObject = Nothing 'GET THE EMAIL SERVER OBJECT
            'Doesn't work because we have not spawned the geometry yet!
            'spawnserver(oemail, ConnectionType.ePrimaryServerApp, 
		Else
			PrimarySuitorSpawned(oPrimary)
		End If

        'mbInSpawn = False
	End Sub

    Public Shared Sub PrimarySuitorSpawned(ByRef oPrimary As ServerObject)
        moPrimary = oPrimary

        'Ok, now, we need to get the region
        Dim oRegion As ServerObject = goMsgSys.FindNeighborhoodSuitor(ConnectionType.eRegionServerApp, oPrimary)
        If oRegion Is Nothing Then
            'TODO: ok, we need to spawn a region server and assign it to the primary server passed in
        Else
            RegionSuitorSpawned(oRegion)
        End If
    End Sub

	Public Shared Sub RegionSuitorSpawned(ByRef oRegion As ServerObject)
		moRegion = oRegion
		'OK, the primary and region should be set
		If moPrimary Is Nothing OrElse moRegion Is Nothing Then
			LogEvent(LogEventType.CriticalError, "In RegionSuitorSpawn, and either Region or Primary is nothing.")
			'ok, increment our numbers
			mlToSpawnCount += 1
			'and clear our flag of being in spawn
			mbInSpawn = False
			Return
		End If

		'Ok, set up our neighborhood loc
		Dim upt3Neighborhood As Point3 = GetGalacticLoc()

		'the app MessAround has all of the system placement/generation data needed
		'To do this, we create a series of systems... 2 spawn systems, 2 respawn systems, 2 hub system and then connect it to a hub hub
		'  if a hub hub is unavailable (because they are full), we spawn a new hub hub and connect it to an existing hub hub
		Dim oS1 As SolarSystem = CreateNewSystem(elSystemType.SpawnSystem, upt3Neighborhood)
        Dim oS2 As SolarSystem = CreateNewSystem(elSystemType.SpawnSystem, upt3Neighborhood)
        Dim oR1 As SolarSystem = CreateNewSystem(elSystemType.SpawnSystem, upt3Neighborhood)
        Dim oR2 As SolarSystem = CreateNewSystem(elSystemType.SpawnSystem, upt3Neighborhood)
		Dim oH1 As SolarSystem = CreateNewSystem(elSystemType.HubSystem, upt3Neighborhood)
		Dim oH2 As SolarSystem = CreateNewSystem(elSystemType.HubSystem, upt3Neighborhood)
		Dim bNewHH As Boolean = False
        'Now, add our wormhole links...
        Dim oRandom As New Random()
        Dim oHH As SolarSystem = GetOrAddHubHub(bNewHH, upt3Neighborhood, oRandom)

        AddWormhole(oS1, oH1, oRandom)
        AddWormhole(oS2, oH1, oRandom)
        AddWormhole(oR1, oH2, oRandom)
        AddWormhole(oR2, oH2, oRandom)
        AddWormhole(oH1, oH2, oRandom)
        AddWormhole(oH1, oHH, oRandom)
        AddWormhole(oH2, oHH, oRandom)

		'Ok, our solar system has been generated and saved... so now, we update all Primaries with the new system
		Dim yMsg(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, 0)
        If bNewHH = True Then
            oHH.GetGUIDAsString.CopyTo(yMsg, 2)
            goMsgSys.SendToPrimaryServers(yMsg)
        End If
        oS1.GetGUIDAsString.CopyTo(yMsg, 2)
        goMsgSys.SendToPrimaryServers(yMsg)
        oS2.GetGUIDAsString.CopyTo(yMsg, 2)
        goMsgSys.SendToPrimaryServers(yMsg)
        oR1.GetGUIDAsString.CopyTo(yMsg, 2)
        goMsgSys.SendToPrimaryServers(yMsg)
        oR2.GetGUIDAsString.CopyTo(yMsg, 2)
        goMsgSys.SendToPrimaryServers(yMsg)
        oH1.GetGUIDAsString.CopyTo(yMsg, 2)
        goMsgSys.SendToPrimaryServers(yMsg)
        oH2.GetGUIDAsString.CopyTo(yMsg, 2)
        goMsgSys.SendToPrimaryServers(yMsg)

        'With uSpawnRequests(oRegion.lSpawnRequestIdx)
        With oRegion
            Dim lFirstIdx As Int32 = .lSpawnUB + 1

            If bNewHH = True Then .lSpawnUB += 7 Else .lSpawnUB += 6

            ReDim Preserve .lSpawnID(.lSpawnUB)
            ReDim Preserve .iSpawnTypeID(.lSpawnUB)

            .lSpawnID(lFirstIdx) = oS1.ObjectID : .iSpawnTypeID(lFirstIdx) = oS1.ObjTypeID : lFirstIdx += 1
            .lSpawnID(lFirstIdx) = oS2.ObjectID : .iSpawnTypeID(lFirstIdx) = oS2.ObjTypeID : lFirstIdx += 1
            .lSpawnID(lFirstIdx) = oR1.ObjectID : .iSpawnTypeID(lFirstIdx) = oR1.ObjTypeID : lFirstIdx += 1
            .lSpawnID(lFirstIdx) = oR2.ObjectID : .iSpawnTypeID(lFirstIdx) = oR2.ObjTypeID : lFirstIdx += 1
            .lSpawnID(lFirstIdx) = oH1.ObjectID : .iSpawnTypeID(lFirstIdx) = oH1.ObjTypeID : lFirstIdx += 1
            .lSpawnID(lFirstIdx) = oH2.ObjectID : .iSpawnTypeID(lFirstIdx) = oH2.ObjTypeID : lFirstIdx += 1
            If bNewHH = True Then
                .lSpawnID(lFirstIdx) = oHH.ObjectID : .iSpawnTypeID(lFirstIdx) = oHH.ObjTypeID : lFirstIdx += 1
            End If
        End With
        'End With

        Dim lNewCnt As Int32 = 6
        If bNewHH = True Then lNewCnt += 1

        Dim lPos As Int32 = 0
        ReDim yMsg((lNewCnt * 4) + 5)
        System.BitConverter.GetBytes(GlobalMessageCode.eCheckPrimaryReady).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lNewCnt).CopyTo(yMsg, lPos) : lPos += 4
        If bNewHH = True Then
            System.BitConverter.GetBytes(oHH.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        End If
        System.BitConverter.GetBytes(oS1.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oS2.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oR1.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oR2.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oH1.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oH2.ObjectID).CopyTo(yMsg, lPos) : lPos += 4

        myLastCheckPrimaryMsg = yMsg

        moPrimary.oSocket.SendData(yMsg)

    End Sub

    Public Shared Sub SetStartHubHub(ByRef oSys As SolarSystem)
        moCurrHHSys = oSys
    End Sub
    Public Shared Sub HandleCheckPrimaryReady(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4

        If lUB = Int32.MinValue Then
            If myLastCheckPrimaryMsg Is Nothing = False Then moPrimary.oSocket.SendData(myLastCheckPrimaryMsg)
            Return
        End If

        Dim lID(lUB) As Int32
        For X As Int32 = 0 To lUB
            lID(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Next X

        If moPrimary Is Nothing Then moPrimary = goMsgSys.FindNeighborhoodSuitor(ConnectionType.ePrimaryServerApp, Nothing)
        If moRegion Is Nothing Then moRegion = goMsgSys.FindNeighborhoodSuitor(ConnectionType.eRegionServerApp, moPrimary)

        'Ok, we are here, so we have all that we need...
        Dim yMsg(27) As Byte

        Dim lSysGuidPos As Int32
        lPos = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eEnvironmentDomain).CopyTo(yMsg, lPos) : lPos += 2
        lSysGuidPos = lPos
        lPos += 4
        System.BitConverter.GetBytes(ObjectType.eSolarSystem).CopyTo(yMsg, lPos) : lPos += 2
        StringToBytes(moRegion.sIPAddress).CopyTo(yMsg, lPos) : lPos += 20

        For X As Int32 = 0 To lUB
            System.BitConverter.GetBytes(lID(X)).CopyTo(yMsg, lSysGuidPos)
            moPrimary.oSocket.SendData(yMsg)
        Next X

        mbInSpawn = False
    End Sub

    Private Shared msNames() As String = Nothing
    Private Shared msNameStart() As String = Nothing
    Private Shared Function GetSystemName() As String

        Dim sFinal As String = ""
        Try
            If msNames Is Nothing Then
                Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory()
                If sFile.EndsWith("\") = False Then sFile &= "\"
                sFile &= "SystemNames.txt"
                Dim oFSIn As New IO.FileStream(sFile, IO.FileMode.Open)
                Dim oRead As New IO.StreamReader(oFSIn)

                ReDim msNames(-1)
                Dim lIdx As Int32 = -1
                While oRead.EndOfStream = False
                    lIdx += 1
                    ReDim Preserve msNames(lIdx)
                    msNames(lIdx) = oRead.ReadLine()
                End While
                oRead.Close()
                oFSIn.Close()

                lIdx = -1
                For X As Int32 = 0 To msNames.GetUpperBound(0)
                    Dim sStart As String = msNames(X).Substring(0, 1)
                    Dim sMid As String = ""
                    If msNames(X).Length > 1 Then sMid = msNames(X).Substring(1, 1)
                    If "AEIOU".IndexOf(sStart) <> -1 OrElse (sMid <> "" AndAlso "AEIOU".IndexOf(sMid) <> -1) Then
                        lIdx += 1
                        ReDim Preserve msNameStart(lIdx)
                        msNameStart(lIdx) = msNames(X)
                    End If
                Next X
            End If

            Dim oRandom As New Random()

            Dim lPartCnt As Int32 = oRandom.Next(1, 2)
            sFinal = msNameStart(oRandom.Next(0, msNameStart.GetUpperBound(0) + 1))  '""
            For X As Int32 = 0 To lPartCnt - 1
                Dim lPart As Int32 = oRandom.Next(0, msNames.GetUpperBound(0) + 1)
                sFinal &= msNames(lPart)
            Next X

            sFinal = sFinal.ToLower
            sFinal = sFinal.Substring(0, 1).ToUpper & sFinal.Substring(1)
        Catch
        End Try

        Return sFinal

    End Function
    Private Shared Function CreateNewSystem(ByVal lType As elSystemType, ByVal uNeighrborhoodLoc As Point3) As SolarSystem

        Dim lRndElem As Int32
        Dim lAddElem As Int32

        Dim oSystem As New SolarSystem()

        Dim sSystemName As String = GetSystemName()
        If lType = elSystemType.SpawnSystem Then sSystemName &= "(S)"

        With oSystem
            Const lNormalRndMult As Int32 = 150
            Const lNormalRndAdd As Int32 = 75
            Const lHHRndMult As Int32 = 300
            Const lHHRndAdd As Int32 = 150

            If lType = elSystemType.HubHubSystem Then
                .LocX = uNeighrborhoodLoc.X + CInt((Rnd() * lHHRndMult) - lHHRndAdd)
                .LocY = uNeighrborhoodLoc.Y + CInt((Rnd() * lHHRndMult) - lHHRndAdd)
                .LocZ = uNeighrborhoodLoc.Z + CInt((Rnd() * lHHRndMult) - lHHRndAdd)
            Else
                .LocX = uNeighrborhoodLoc.X + CInt((Rnd() * lNormalRndMult) - lNormalRndAdd)
                .LocY = uNeighrborhoodLoc.Y + CInt((Rnd() * lNormalRndMult) - lNormalRndAdd)
                .LocZ = uNeighrborhoodLoc.Z + CInt((Rnd() * lNormalRndMult) - lNormalRndAdd)
            End If

            .LocX *= 2
            .LocY *= 2
            .LocZ *= 2

            .ObjectID = -1
            .ObjTypeID = ObjectType.eSolarSystem
            .ParentGalaxy = goGalaxy(0)
            .SystemName = StringToBytes(sSystemName)
            .SystemType = CByte(lType)

            'used oRandom here cuz it is easier to manage
            Dim oRandom As New Random()
            Select Case lType
                Case elSystemType.SpawnSystem
                    'Ok, normal colorization - 90% unary, 10% binary - 7 - 13 (smaller stars)
                    If oRandom.Next(0, 100) < 90 Then
                        'unary 
                        .StarType1ID = CByte(oRandom.Next(7, 14))
                        .StarType2ID = 0
                        .StarType3ID = 0
                    Else
                        'binary
                        .StarType1ID = CByte(oRandom.Next(7, 14))
                        .StarType2ID = CByte(oRandom.Next(10, 14))
                        .StarType3ID = 0
                    End If
                Case elSystemType.RespawnSystem
                    'smaller stars 7-13, 60% unary, 30% binary, 10% trinary
                    Dim lVal As Int32 = oRandom.Next(0, 100)
                    If lVal < 60 Then
                        'unary
                        .StarType1ID = CByte(oRandom.Next(7, 14))
                        .StarType2ID = 0
                        .StarType3ID = 0
                    ElseIf lVal < 90 Then
                        'binary
                        .StarType1ID = CByte(oRandom.Next(7, 14))
                        .StarType2ID = CByte(oRandom.Next(10, 14))
                        .StarType3ID = 0
                    Else
                        'trinary
                        .StarType1ID = CByte(oRandom.Next(7, 14))
                        .StarType2ID = CByte(oRandom.Next(10, 14))
                        .StarType3ID = CByte(oRandom.Next(10, 14))
                    End If
                Case elSystemType.HubSystem
                    '50% unary, 30% binary, 20% trinary - full range 4-14
                    Dim lVal As Int32 = oRandom.Next(0, 100)
                    If lVal < 50 Then
                        'unary
                        .StarType1ID = CByte(oRandom.Next(4, 15))
                        .StarType2ID = 0
                        .StarType3ID = 0
                    ElseIf lVal < 30 Then
                        'binary
                        .StarType1ID = CByte(oRandom.Next(4, 15))
                        .StarType2ID = CByte(oRandom.Next(7, 14))
                        .StarType3ID = 0
                    Else
                        'trinary
                        .StarType1ID = CByte(oRandom.Next(7, 14))
                        .StarType2ID = CByte(oRandom.Next(10, 14))
                        .StarType3ID = CByte(oRandom.Next(10, 14))
                    End If
                Case elSystemType.HubHubSystem
                    If oRandom.Next(0, 100) < 50 Then
                        'special stars - 1, 2, 3, 13, 14, 15 or in some cases, none
                        Dim lVal As Int32 = oRandom.Next(0, 7)
                        Select Case lVal
                            Case 0 : .StarType1ID = 1
                            Case 1 : .StarType1ID = 2
                            Case 2 : .StarType1ID = 3
                            Case 3 : .StarType1ID = 13
                            Case 4 : .StarType1ID = 14
                            Case 5 : .StarType1ID = 15
                            Case Else
                                .StarType1ID = 11
                        End Select
                        .StarType2ID = 0
                        .StarType3ID = 0
                    Else
                        'random full spectrum
                        Dim lVal As Int32 = oRandom.Next(0, 100)
                        If lVal < 50 Then
                            'unary
                            .StarType1ID = CByte(oRandom.Next(4, 15))
                            .StarType2ID = 0
                            .StarType3ID = 0
                        ElseIf lVal < 30 Then
                            'binary
                            .StarType1ID = CByte(oRandom.Next(4, 15))
                            .StarType2ID = CByte(oRandom.Next(7, 14))
                            .StarType3ID = 0
                        Else
                            'trinary
                            .StarType1ID = CByte(oRandom.Next(4, 15))
                            .StarType2ID = CByte(oRandom.Next(7, 14))
                            .StarType3ID = CByte(oRandom.Next(10, 14))
                        End If
                    End If
            End Select
            '.FleetJumpPointX = CInt((4000 * Rnd()) - 2000)
            '.FleetJumpPointZ = CInt((4000 * Rnd()) - 2000)
            Dim dDX As Single = goStarType(.StarType1ID).StarRadius
            Dim dDY As Single = 0
            Dim oRandom2 As New Random
            Dim lAngle As Single = oRandom2.Next(0, 360)
            RotatePoint(0, 0, dDX, dDY, lAngle)
            .FleetJumpPointX = CInt(dDX)
            .FleetJumpPointZ = CInt(dDY)

            '.StarType1ID = 11
            '.StarType2ID = 0
            '.StarType3ID = 0
        End With

        Dim lHeatIndex As Int32 = GetHeatIntensityScore(oSystem.StarType1ID, oSystem.StarType2ID, oSystem.StarType3ID)
        Dim lCenterRadius As Int32 = GetCenterRadius(oSystem.StarType1ID, oSystem.StarType2ID, oSystem.StarType3ID)
        Dim lTinyPlanetPool As Int32 = 3
        Dim lSmallPlanetPool As Int32 = 3
        Dim lMediumPlanetPool As Int32 = 8
        Dim lLargePlanetPool As Int32 = 7
        Dim lHugePlanetPool As Int32 = 5

        Select Case lType
            Case elSystemType.HubHubSystem
                lRndElem = 6 '5
                lAddElem = 5 '18
                lTinyPlanetPool = 25
                lSmallPlanetPool = 25
                lMediumPlanetPool = 25
                lLargePlanetPool = 25
                lHugePlanetPool = 25
            Case elSystemType.HubSystem
                lRndElem = 4 '7
                lAddElem = 5
                lTinyPlanetPool = 15
                lSmallPlanetPool = 15
                lMediumPlanetPool = 15
                lLargePlanetPool = 15
                lHugePlanetPool = 15
            Case elSystemType.RespawnSystem
                lRndElem = 4 '6
                lAddElem = 5 '14
                lTinyPlanetPool = 2
                lSmallPlanetPool = 4
                lMediumPlanetPool = 5
                lLargePlanetPool = 7
                lHugePlanetPool = 7
            Case elSystemType.SpawnSystem
                lRndElem = 8 '5
                lAddElem = 9 '12
                lTinyPlanetPool = 3
                lSmallPlanetPool = 3
                lMediumPlanetPool = 8
                lLargePlanetPool = 7
                lHugePlanetPool = 5
            Case Else
                Return Nothing
        End Select

        Dim lCnt As Int32 = CInt(Rnd() * lRndElem) + lAddElem
        If lCnt Mod 2 = 0 Then lCnt += 1
        Dim oPlanets(lCnt - 1) As Planet

        Dim fDistFromSun(lCnt - 1) As Single

        For X As Int32 = 0 To lCnt - 1
            Dim lX As Int32 = CInt(Rnd() * 9000000) - 4500000
            Dim lZ As Int32 = CInt(Rnd() * 9000000) - 4500000

            Dim lLoopCnt As Int32 = 0
            While Math.Abs(lX) < lCenterRadius AndAlso Math.Abs(lZ) < lCenterRadius AndAlso lLoopCnt < 1000
                lX = CInt(Rnd() * 9000000) - 4500000
                lZ = CInt(Rnd() * 9000000) - 4500000
                lLoopCnt += 1
            End While

            oPlanets(X) = New Planet()
            With oPlanets(X)
                .LocX = lX
                .LocZ = lZ

                'ok, determine the type of planet
                .AxisAngle = 700 + CInt((Rnd() * 400) - 200)
                .ParentSystem = oSystem
                .ObjectID = -1
                .ObjTypeID = ObjectType.ePlanet

                'Determine the size of the planet....
                Dim lTmpCnt As Int32 = lTinyPlanetPool + lSmallPlanetPool + lMediumPlanetPool + lLargePlanetPool + lHugePlanetPool
                Dim lIdx As Int32 = CInt(Rnd() * lTmpCnt)

                lIdx -= lTinyPlanetPool
                If lIdx < 0 Then
                    .PlanetSizeID = 0
                Else
                    lIdx -= lSmallPlanetPool
                    If lIdx < 0 Then
                        .PlanetSizeID = 1
                    Else
                        lIdx -= lMediumPlanetPool
                        If lIdx < 0 Then
                            .PlanetSizeID = 2
                        Else
                            lIdx -= lLargePlanetPool
                            If lIdx < 0 Then
                                .PlanetSizeID = 3
                            Else : .PlanetSizeID = 4
                            End If
                        End If
                    End If
                End If
                Select Case .PlanetSizeID
                    Case 0
                        lTinyPlanetPool -= 1
                        .PlanetRadius = 1000
                        .Gravity = CByte(CInt(Rnd() * 70) + 20)
                    Case 1
                        lSmallPlanetPool -= 1
                        .PlanetRadius = 1500
                        .Gravity = CByte(CInt(Rnd() * 70) + 40)
                    Case 2
                        lMediumPlanetPool -= 1
                        .PlanetRadius = 2550
                        .Gravity = CByte(CInt(Rnd() * 50) + 60)
                    Case 3
                        lLargePlanetPool -= 1
                        .PlanetRadius = 4000
                        .Gravity = CByte(CInt(Rnd() * 50) + 70)
                    Case Else
                        lHugePlanetPool -= 1
                        .PlanetRadius = 6000
                        .Gravity = CByte(CInt(Rnd() * 60) + 80)
                End Select

                .RotationDelay = CShort(CInt(Rnd() * 300) + 100)

                'ok, determine heat and density
                Dim fDX As Single = .LocX
                fDX *= fDX
                Dim fDZ As Single = .LocZ
                fDZ *= fDZ
                Dim fDist As Single = CSng(Math.Sqrt(fDX + fDZ))
                fDist -= lCenterRadius
                fDist /= lCenterRadius

                .PlanetTypeID = CByte(Math.Floor(Rnd() * 7))
                If .PlanetTypeID = 7 AndAlso lType = elSystemType.SpawnSystem Then .PlanetTypeID = 1
                .LocY = -(.PlanetRadius + 2000)

                If .PlanetTypeID = 1 Then
                    .Atmosphere = 0
                ElseIf .PlanetTypeID = 4 Then
                    .Atmosphere = 10
                Else : .Atmosphere = 100
                End If

                .Hydrosphere = 100
                If .PlanetTypeID = 6 OrElse .PlanetTypeID = 7 Then
                    .Vegetation = 100
                Else : .Vegetation = 0
                End If

                If .PlanetTypeID = 1 Then .Hydrosphere = 0

                Dim lValue As Int32 = CInt(fDist * lHeatIndex)
                lValue \= 10
                If lValue > 255 Then lValue \= 2
                If lValue > 255 Then lValue = 255
                If lValue < 10 Then lValue = 10

                'ok, based on our heat value, we may need to shift our planet
                Select Case .PlanetTypeID
                    Case PlanetType.eAcidic
                        'Ok, acidic causes the temperature to rise, no shifting
                        If lValue < 30 Then
                            lValue = 30
                            lValue += CInt(Rnd() * 25)
                        End If
                    Case PlanetType.eAdaptable
                        'Ok, adaptable can be any value less than 40, if the value exceeds 40, we shift to desert
                        If lValue > 40 Then .PlanetTypeID = PlanetType.eDesert
                    Case PlanetType.eBarren
                        'barren is always barren
                    Case PlanetType.eDesert
                        If lValue < 20 Then .PlanetTypeID = PlanetType.eAdaptable
                    Case PlanetType.eGeoPlastic
                        If lValue < 30 Then
                            Dim lRoll As Int32 = CInt(Rnd() * 100)
                            If lRoll < 30 Then
                                lValue = 30 + CInt(Rnd() * 25)
                            ElseIf lRoll < 50 Then
                                .PlanetTypeID = PlanetType.eAcidic
                                lValue = 30 + CInt(Rnd() * 25)
                            ElseIf lRoll < 70 Then
                                .PlanetTypeID = PlanetType.eAdaptable
                            ElseIf lRoll < 80 Then
                                .PlanetTypeID = PlanetType.eTerran
                                lValue += CInt(Rnd() * 10)
                            ElseIf lRoll < 95 Then
                                .PlanetTypeID = PlanetType.eTundra
                                If lValue > 27 Then
                                    lValue = 27
                                    lValue -= CInt(Rnd() * 10)
                                End If
                            Else
                                .PlanetTypeID = PlanetType.eWaterWorld
                            End If
                        End If
                    Case PlanetType.eTerran
                        If lValue < 28 Then
                            If Rnd() * 100 < 70 Then
                                lValue = 28 + CInt(Rnd() * 5)
                            Else
                                .PlanetTypeID = PlanetType.eTundra
                            End If
                        ElseIf lValue > 32 Then
                            If Rnd() * 100 < 70 Then
                                lValue = 32 - CInt(Rnd() * 5)
                            Else
                                .PlanetTypeID = PlanetType.eAdaptable
                            End If
                        End If
                    Case PlanetType.eTundra
                        If lValue > 27 Then
                            lValue = 27
                            lValue -= CInt(Rnd() * 10)
                        End If
                    Case PlanetType.eWaterWorld
                        lValue = CInt(Rnd() * 5) + 27
                End Select
                .SurfaceTemperature = CByte(lValue)

                If .PlanetSizeID = 3 Then
                    If CInt(Rnd() * 100) < 40 Then
                        .RingDiffuse = System.Drawing.Color.FromArgb(255, 32, 128, 255).ToArgb
                        .InnerRingRadius = .PlanetRadius + 1500 + CInt(Rnd() * 1000)
                        .OuterRingRadius = .InnerRingRadius + CInt(Rnd() * 3000) + 1500
                        .RingMineralID = Math.Min(105, Math.Max(1, CInt(Rnd() * 105) + 1))
                        .RingConcentration = CInt(Rnd() * 25) + 1
                    End If
                ElseIf .PlanetSizeID = 4 Then
                    If CInt(Rnd() * 100) < 60 Then
                        .RingDiffuse = System.Drawing.Color.FromArgb(255, 32, 128, 255).ToArgb
                        .InnerRingRadius = .PlanetRadius + 1500 + CInt(Rnd() * 1500)
                        .OuterRingRadius = .InnerRingRadius + CInt(Rnd() * 4000) + 2000
                        .RingMineralID = Math.Min(105, Math.Max(1, CInt(Rnd() * 105) + 1))
                        .RingConcentration = CInt(Rnd() * 25) + 1
                    End If
                End If
                fDistFromSun(X) = CSng(Math.Sqrt((CSng(.LocX) * CSng(.LocX)) + (CSng(.LocZ) * CSng(.LocZ))))

            End With
        Next X

        'Now, name our planets
        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1

        For X As Int32 = 0 To lCnt - 1
            Dim lIdx As Int32 = -1

            For Y As Int32 = 0 To lSortedUB
                If fDistFromSun(lSorted(Y)) > fDistFromSun(X) Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedUB += 1
            ReDim Preserve lSorted(lSortedUB)
            If lIdx = -1 Then
                lSorted(lSortedUB) = X
            Else
                For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                    lSorted(Y) = lSorted(Y - 1)
                Next Y
                lSorted(lIdx) = X
            End If
        Next X

        If oSystem.SaveObject() = False Then
            Return Nothing
        Else
            Dim lNewIdx As Int32 = glSystemUB + 1
            ReDim Preserve goSystem(lNewIdx)
            ReDim Preserve glSystemIdx(lNewIdx)
            goSystem(lNewIdx) = oSystem
            glSystemIdx(lNewIdx) = oSystem.ObjectID
            glSystemUB = lNewIdx
        End If

        Dim lPlanetsSinceRare As Int32 = CInt(Rnd() * 10)
        PrepareMineralData()

        Dim sNoSName As String = sSystemName.Replace("(S)", "").Trim
        For X As Int32 = 0 To lSortedUB
            Dim lNewIdx As Int32 = glPlanetUB + 1
            ReDim Preserve goPlanet(lNewIdx)
            ReDim Preserve glPlanetIdx(lNewIdx)

            oPlanets(lSorted(X)).PlanetName = StringToBytes(sNoSName & " " & GetRomanNumeral(X + 1))
            oPlanets(lSorted(X)).ParentSystem = oSystem

            goPlanet(lNewIdx) = oPlanets(lSorted(X))

            If oPlanets(lSorted(X)).SaveObject() = False Then Continue For

            glPlanetIdx(lNewIdx) = oPlanets(lSorted(X)).ObjectID
            glPlanetUB = lNewIdx
            oSystem.AddPlanetIndex(lNewIdx)


            SpawnPlanetsMinerals(oPlanets(lSorted(X)), lPlanetsSinceRare)
        Next X

        'TODO: Create our mineralgeographyrel entries... dont worry with these for now since nothing else uses them
        'TODO: Spawn our mineralcaches based on the mineralgeographyrel - use the random generator as it is in the cache  generator
        

        Return oSystem
    End Function

    Public Shared Sub SpawnSystemsMinerals(ByVal lSystemID As Int32)
        Dim oSystem As SolarSystem = GetEpicaSystem(lSystemID)
        If oSystem Is Nothing Then Return

        Dim lPlanetsSinceRare As Int32 = CInt(Rnd() * 10)
        PrepareMineralData()

        For X As Int32 = 0 To oSystem.mlPlanetUB

            Dim oPlanet As Planet = oSystem.GetPlanet(X)
            If oPlanet Is Nothing Then Continue For

            SpawnPlanetsMinerals(oPlanet, lPlanetsSinceRare)

            oPlanet.SaveObject()
        Next X
    End Sub

    Private Shared mlMineralID() As Int32 = Nothing
    Private Shared mlRarity() As Int32 = Nothing
    Private Shared mlComplexity() As Int32 = Nothing
    Private Shared mlMaxRarity As Int32 = 0
    Private Shared Sub PrepareMineralData()
        If mlMineralID Is Nothing = False Then Return

        ReDim mlMineralID(-1)
        ReDim mlRarity(-1)

        Dim oComm As New OleDb.OleDbCommand("SELECT tblMineral.MineralID, tblMineral.Rarity, tblMineralPropertyValue.PropertyValue " & _
          " FROM tblMineral LEFT OUTER JOIN tblMineralPropertyValue ON tblMineral.MineralID = tblMineralPropertyValue.MineralID WHERE " & _
          " TblMineralPropertyValue.MineralPropertyID = 20 and tblMineral.AlloyTechID < 1 and tblmineral.mineralid <> 157 and tblMineral.MineralID <> 41991", goCN)

        Try
            Dim oRS As OleDb.OleDbDataReader = oComm.ExecuteReader()
            Dim lMinUB As Int32 = -1

            While oRS.Read
                lMinUB += 1
                ReDim Preserve mlMineralID(lMinUB)
                ReDim Preserve mlRarity(lMinUB)
                ReDim Preserve mlComplexity(lMinUB)

                mlMineralID(lMinUB) = CInt(oRS("MineralID"))
                mlRarity(lMinUB) = CInt(oRS("Rarity"))
                mlComplexity(lMinUB) = CInt(oRS("PropertyValue"))

                If mlRarity(lMinUB) = 0 Then mlRarity(lMinUB) = 100
                mlMaxRarity += mlRarity(lMinUB)
            End While
            oRS.Close()
            oRS = Nothing
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "PrepareMineralData: " & ex.Message)
            mlMineralID = Nothing
        Finally
            oComm.Dispose()
            oComm = Nothing
        End Try
        
    End Sub

    Public Shared Sub LegalizeMineralLocs()
        For X As Int32 = 0 To glPlanetUB
            If glPlanetIdx(X) > -1 Then
                Dim oPlanet As Planet = goPlanet(X)
                If oPlanet Is Nothing = False Then

                    Dim oTerrain As TerrainClass = New TerrainClass(oPlanet.ObjectID)
                    oTerrain.MapType = oPlanet.PlanetTypeID
                    Select Case oPlanet.PlanetSizeID
                        Case 0 : oTerrain.CellSpacing = gl_TINY_PLANET_CELL_SPACING
                        Case 1 : oTerrain.CellSpacing = gl_SMALL_PLANET_CELL_SPACING
                        Case 2 : oTerrain.CellSpacing = gl_MEDIUM_PLANET_CELL_SPACING
                        Case 3 : oTerrain.CellSpacing = gl_LARGE_PLANET_CELL_SPACING
                        Case 4 : oTerrain.CellSpacing = gl_HUGE_PLANET_CELL_SPACING
                    End Select

                    Dim lExtent As Int32 = oTerrain.CellSpacing * (TerrainClass.Width - 4)
                    Dim lHalfExtent As Int32 = (lExtent \ 2) - 1

                    oTerrain.PopulateData()

                    Dim sSQL As String = "SELECT * FROM tblMineralCache WHERE ParentTypeID = 3 AND ParentID = " & goPlanet(X).ObjectID & _
                      " AND (LocX > " & lHalfExtent & " OR LocX < " & -lHalfExtent & " OR LocY > " & lHalfExtent & " OR LocY < " & -lHalfExtent & ")"
                    Dim oComm As OleDb.OleDbCommand = Nothing
                    Dim oData As OleDb.OleDbDataReader = Nothing

                    Dim ptLoc() As Point = Nothing
                    Dim lID() As Int32 = Nothing
                    Dim lUB As Int32 = -1

                    Try
                        oComm = New OleDb.OleDbCommand(sSQL, goCN)
                        oData = oComm.ExecuteReader()
                        While oData.Read = True
                            Dim lCacheID As Int32 = CInt(oData("CacheID"))

                            Dim bAdded As Boolean = False
                            While bAdded = False
                                Dim lX As Int32 = CInt(Int(Rnd() * lExtent) - lHalfExtent)
                                Dim lZ As Int32 = CInt(Int(Rnd() * lExtent) - lHalfExtent)

                                If AreaValid(oTerrain, lX, lZ) = True Then
                                    If oTerrain.TerrainGradeBuildable(lX, lZ) = True Then

                                        bAdded = True

                                        lUB += 1
                                        ReDim Preserve lID(lUB)
                                        ReDim Preserve ptLoc(lUB)
                                        lID(lUB) = lCacheID
                                        ptLoc(lUB) = New Point(lX, lZ)
                                    End If
                                End If
                            End While

                        End While
                    Catch ex As Exception
                        LogEvent(LogEventType.CriticalError, "LegalizeMineralLocs: " & ex.Message)
                    End Try
                    If oData Is Nothing = False Then oData.Dispose()
                    oData = Nothing
                    If oComm Is Nothing = False Then oComm.Dispose()
                    oComm = Nothing

                    For Y As Int32 = 0 To lUB
                        sSQL = "UPDATE tblMineralCache SET LocX = " & ptLoc(Y).X & ", LocY = " & ptLoc(Y).Y & " WHERE CacheID = " & lID(Y)
                        Try
                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                            oComm.ExecuteNonQuery()
                        Catch ex As Exception
                            LogEvent(LogEventType.CriticalError, "LegalizeMineralLocs: " & ex.Message)
                        End Try
                        If oComm Is Nothing = False Then oComm.Dispose()
                        oComm = Nothing
                    Next Y
                End If
            End If
        Next X

        LogEvent(LogEventType.Informational, "Finished LegalizeMineralLocs")

        'Dim sSQL As String = "select * from tblmineralcache where parenttypeid = 3 and parentid in " & _
        '    "(select planetid from tblplanet where planetsizeid = 1 and parentid > 36) and (locx < -35850 or locy < -35850 or locx > 35850 or locy > 35850)"
        'Dim oComm As OleDb.OleDbCommand = Nothing
        'Dim oData As OleDb.OleDbDataReader = Nothing

        'Try
        '    oComm = New OleDb.OleDbCommand(sSQL, goCN)
        '    oData = oComm.ExecuteReader()

        '    Dim lCaches() As Int32 = Nothing
        '    Dim lParents() As Int32 = Nothing
        '    Dim lCacheUB As Int32 = -1

        '    Dim lPlanets() As Int32 = Nothing
        '    Dim lPlanetUB As Int32 = -1

        '    While oData.Read = True
        '        lCacheUB += 1
        '        ReDim Preserve lCaches(lCacheUB)
        '        ReDim Preserve lParents(lCacheUB)
        '        lCaches(lCacheUB) = CInt(oData("CacheID"))
        '        lParents(lCacheUB) = CInt(oData("ParentID"))

        '        Dim bFound As Boolean = False
        '        For X As Int32 = 0 To lPlanetUB
        '            If lPlanets(X) = lParents(lCacheUB) Then
        '                bFound = True
        '                Exit For
        '            End If
        '        Next X
        '        If bFound = False Then
        '            lPlanetUB += 1
        '            ReDim Preserve lPlanets(lPlanetUB)
        '            lPlanets(lPlanetUB) = lParents(lCacheUB)
        '        End If
        '    End While

        '    oData.Close()
        '    oComm.Dispose()
        '    oData = Nothing
        '    oComm = Nothing

        '    For X As Int32 = 0 To lPlanetUB
        '        Dim oPlanet As Planet = GetEpicaPlanet(lPlanets(X))
        '        If oPlanet Is Nothing = False Then
        '            Dim oTC As New TerrainClass(oPlanet.ObjectID)
        '            oTC.MapType = oPlanet.PlanetTypeID
        '            Select Case oPlanet.PlanetSizeID
        '                Case 0 : oTC.CellSpacing = gl_TINY_PLANET_CELL_SPACING
        '                Case 1 : oTC.CellSpacing = gl_SMALL_PLANET_CELL_SPACING
        '                Case 2 : oTC.CellSpacing = gl_MEDIUM_PLANET_CELL_SPACING
        '                Case 3 : oTC.CellSpacing = gl_LARGE_PLANET_CELL_SPACING
        '                Case 4 : oTC.CellSpacing = gl_HUGE_PLANET_CELL_SPACING
        '            End Select
        '            Dim lExtent As Int32 = oTC.CellSpacing * (TerrainClass.Width - 1)
        '            Dim lHalfExtent As Int32 = (lExtent \ 2) - 1

        '            oTC.PopulateData()

        '            For Y As Int32 = 0 To lCacheUB
        '                If lParents(Y) = lPlanets(X) Then
        '                    Dim bAdded As Boolean = False
        '                    While bAdded = False
        '                        Dim lX As Int32 = CInt(Int(Rnd() * lExtent) - lHalfExtent)
        '                        Dim lZ As Int32 = CInt(Int(Rnd() * lExtent) - lHalfExtent)

        '                        If AreaValid(oTC, lX, lZ) = True Then
        '                            If oTC.TerrainGradeBuildable(lX, lZ) = True Then
        '                                sSQL = "UPDATE tblMineralCache SET LocX = " & lX & ", LocY = " & lZ & " WHERE CacheID = " & lCaches(Y)
        '                                oComm = New OleDb.OleDbCommand(sSQL, goCN)
        '                                oComm.ExecuteNonQuery()
        '                                oComm = Nothing
        '                                bAdded = True
        '                            End If
        '                        End If
        '                    End While

        '                End If
        '            Next Y

        '        End If
        '    Next X

        'Catch ex As Exception
        '    LogEvent(LogEventType.CriticalError, "LegalizeMineralLocs: " & ex.Message)
        'Finally
        '    If oData Is Nothing = False Then oData.Close()
        '    oData = Nothing
        '    If oComm Is Nothing = False Then oComm.Dispose()
        '    oComm = Nothing
        'End Try
    End Sub

    Private Shared Sub SpawnPlanetsMinerals(ByRef oPlanet As Planet, ByRef lPlanetsSinceRare As Int32)
        lPlanetsSinceRare += 1
        Dim oTerrain As TerrainClass = New TerrainClass(oPlanet.ObjectID)
        oTerrain.MapType = oPlanet.PlanetTypeID
        Select Case oPlanet.PlanetSizeID
            Case 0 : oTerrain.CellSpacing = gl_TINY_PLANET_CELL_SPACING
            Case 1 : oTerrain.CellSpacing = gl_SMALL_PLANET_CELL_SPACING
            Case 2 : oTerrain.CellSpacing = gl_MEDIUM_PLANET_CELL_SPACING
            Case 3 : oTerrain.CellSpacing = gl_LARGE_PLANET_CELL_SPACING
            Case 4 : oTerrain.CellSpacing = gl_HUGE_PLANET_CELL_SPACING
        End Select

        Dim lExtent As Int32 = oTerrain.CellSpacing * (TerrainClass.Width - 4)
        Dim lHalfExtent As Int32 = (lExtent \ 2) - 1

        oTerrain.PopulateData()

        Dim lMaxX As Int32 = 150        'change this when needed
        If oPlanet.PlanetTypeID = 7 Then
            lMaxX = CInt(lMaxX * (1.0F - (oTerrain.mlWaterPerc / 100.0F))) * (oPlanet.PlanetSizeID + 2)
        End If
        If oPlanet.PlanetSizeID = 0 Then
            lMaxX = CInt(lMaxX * 0.5F)
        ElseIf oPlanet.PlanetSizeID = 1 Then
            lMaxX = CInt(lMaxX * 0.75F)
        End If
        Dim lMinUB As Int32 = mlMineralID.GetUpperBound(0)

        For X As Int32 = 0 To lMaxX
            Dim lVal As Int32 = CInt(Rnd() * mlMaxRarity)
            Dim lCon As Int32 = mlRarity(0)
            Dim lQty As Int32 = CInt(((Rnd() * 10) + 2) * 400000)

            Dim lMineralID As Int32 = 1
            Dim lComp As Int32 = 100

            If Rnd() * 100 < 50 Then
                For lTmpIdx As Int32 = 0 To lMinUB
                    If lVal > 100 Then
                        lVal -= mlRarity(lTmpIdx)
                    ElseIf lVal < mlRarity(lTmpIdx) Then
                        lComp = mlComplexity(lTmpIdx)
                        If lComp < 61 Then
                            If lPlanetsSinceRare < 10 Then Continue For
                            lPlanetsSinceRare = 0
                        End If

                        lMineralID = mlMineralID(lTmpIdx)
                        lQty = CInt(((Rnd() * 10) + 2) * 400000)
                        lCon = mlRarity(lTmpIdx) 'lCVal(lTmpIdx)
                        Exit For
                    Else
                        lVal -= mlRarity(lTmpIdx)
                    End If
                Next lTmpIdx
            Else
                Dim lTmpIdx As Int32 = CInt(Rnd() * 5)
                lMineralID = mlMineralID(lTmpIdx)
                lComp = mlComplexity(lTmpIdx)
                lQty = CInt(((Rnd() * 10) + 2) * 400000)
                lCon = mlRarity(lTmpIdx)
                If lComp < 61 Then lPlanetsSinceRare = 0
            End If

            'MSC - 1/7/09 - added to remove rakuron and manganon from the galaxy
            If lMineralID = 4 OrElse lMineralID = 6 Then
                lMineralID += CInt(Rnd() * 20)
            End If

            Dim bAdded As Boolean = False

            While bAdded = False
                Dim lX As Int32 = CInt(Int(Rnd() * lExtent) - lHalfExtent)
                Dim lZ As Int32 = CInt(Int(Rnd() * lExtent) - lHalfExtent)

                If AreaValid(oTerrain, lX, lZ) = True Then
                    If oTerrain.TerrainGradeBuildable(lX, lZ) = True Then

                        bAdded = True

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

                        Dim oComm As OleDb.OleDbCommand = Nothing
                        Try
                            Dim sSQL As String = "INSERT INTO tblMineralCache (CacheTypeID, ParentID, ParentTypeID, LocX, LocY, MineralID, Quantity, Concentration, OriginalConcentration) VALUES " & _
                              "(0, " & oPlanet.ObjectID & ", 3, " & lX & ", " & lZ & ", " & lMineralID & ", " & lQty & ", " & lCon & ", " & lCon & ")"
                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                            oComm.ExecuteNonQuery()
                        Catch ex As Exception
                            LogEvent(LogEventType.CriticalError, "Unable to SpawnPlanetsMinerals item: " & ex.Message)
                        Finally
                            If oComm Is Nothing = False Then oComm.Dispose()
                            oComm = Nothing
                        End Try
                    End If
                End If
            End While
        Next X

    End Sub

    Private Shared Function AreaValid(ByRef oTerrain As TerrainClass, ByVal lX As Int32, ByVal lZ As Int32) As Boolean
        Dim lWH As Int32 = CInt(oTerrain.WaterHeight * oTerrain.ml_Y_Mult)
        Return oTerrain.GetHeightAtLocation(lX, lZ) > lWH AndAlso _
               oTerrain.GetHeightAtLocation(lX - oTerrain.CellSpacing, lZ - oTerrain.CellSpacing) > lWH AndAlso _
               oTerrain.GetHeightAtLocation(lX, lZ - oTerrain.CellSpacing) > lWH AndAlso _
               oTerrain.GetHeightAtLocation(lX + oTerrain.CellSpacing, lZ - oTerrain.CellSpacing) > lWH AndAlso _
               oTerrain.GetHeightAtLocation(lX - oTerrain.CellSpacing, lZ) > lWH AndAlso _
               oTerrain.GetHeightAtLocation(lX + oTerrain.CellSpacing, lZ) > lWH AndAlso _
               oTerrain.GetHeightAtLocation(lX - oTerrain.CellSpacing, lZ + oTerrain.CellSpacing) > lWH AndAlso _
               oTerrain.GetHeightAtLocation(lX, lZ + oTerrain.CellSpacing) > lWH AndAlso _
               oTerrain.GetHeightAtLocation(lX + oTerrain.CellSpacing, lZ + oTerrain.CellSpacing) > lWH
    End Function

#Region " Terrain Class "
    Private Class TerrainClass

        Public Structure Vector3
            Public X As Single
            Public Y As Single
            Public Z As Single

            Public Sub New(ByVal fX As Single, ByVal fY As Single, ByVal fZ As Single)
                X = fX
                Y = fY
                Z = fZ
            End Sub

            Public Sub Add(ByVal v1 As Vector3)
                X += v1.X
                Y += v1.Y
                Z += v1.Z
            End Sub

            Public Shared Function Add(ByVal v1 As Vector3, ByVal v2 As Vector3) As Vector3
                Dim v As Vector3
                v.X = v1.X + v2.X
                v.Y = v1.Y + v2.Y
                v.Z = v1.Z + v2.Z
                Return v
            End Function

            Public Sub Subtract(ByVal v1 As Vector3)
                X -= v1.X
                Y -= v1.Y
                Z -= v1.Z
            End Sub

            Public Shared Function Subtract(ByVal v1 As Vector3, ByVal v2 As Vector3) As Vector3
                Dim v As Vector3
                v.X = v1.X - v2.X
                v.Y = v1.Y - v2.Y
                v.Z = v1.Z - v2.Z
                Return v
            End Function

            Public Shared Function CrossProduct(ByVal v1 As Vector3, ByVal v2 As Vector3) As Vector3
                Dim v As Vector3
                v.X = (v1.Y * v2.Z) - (v1.Z * v2.Y)
                v.Y = (v1.Z * v2.X) - (v1.X * v2.Z)
                v.Z = (v1.X * v2.Y) - (v1.Y * v2.X)
                Return v
            End Function

            Public Sub Normalize()
                Dim fTemp As Single = CSng(Math.Sqrt((X * X) + (Y * Y) + (Z * Z)))
                If fTemp > 0 Then
                    fTemp = 1 / fTemp
                Else : fTemp = 1
                End If
                X *= fTemp
                Y *= fTemp
                Z *= fTemp
            End Sub

            Public Shared Function Empty() As Vector3
                Dim v As Vector3
                v.X = 0.0F
                v.Y = 0.0F
                v.Z = 0.0F
                Return v
            End Function

            Public Sub Multiply(ByVal fValue As Single)
                X *= fValue
                Y *= fValue
                Z *= fValue
            End Sub

            Public Shared Function Multiply(ByVal v1 As Vector3, ByVal fValue As Single) As Vector3
                Dim v As Vector3
                v.X = v1.X * fValue
                v.Y = v1.Y * fValue
                v.Z = v1.Z * fValue
                Return v
            End Function
        End Structure
#Region "Constant Expressions"
        Public Const Width As Int32 = 240 '256
        Public Const Height As Int32 = 240 '256
        Private Const mlPASSES As Int32 = 5
        Private Const mlQuads As Int32 = 24 '8      '8x8 quad = 64 quads total, this is to segment the total vertex buffer further
        Private Const mlVertsPerQuad As Int32 = CInt(Width / mlQuads)
        Private Const mlHalfHeight As Int32 = CInt(Height / 2)
        Private Const mlHalfWidth As Int32 = CInt(Width / 2)
        Private Const VertsTotal As Int32 = Width * Height
        Private Const QuadsX As Int32 = Width - 1
        Private Const QuadsZ As Int32 = Height - 1
        Private Const TrisX As Int32 = CInt(QuadsX / 2)
#End Region

        Public CellSpacing As Int32 = 200
        Public ml_Y_Mult As Single = 15.0       'TODO: does this need to be a single???
        Public MapType As Int32
        Public WaterHeight As Byte      'height of the water level
        Public mlWaterPerc As Int32
#Region "Private Variables"
        'Our heightmap array, we need to keep this for... getting heights at locations
        Private HeightMap() As Byte
        'Indicates if the Heightmap has been generated yet
        Private mbHMReady As Boolean = False
        'the Parent Planet ID... set in New()
        Private mlSeed As Int32

        Private HMNormals() As Vector3      'TODO: We should not need to store these permanently, figure another way
#End Region

        Public Sub New(ByVal lSeed As Int32)
            mlSeed = lSeed
        End Sub

        Public Function GetVertexHeight(ByVal lVertX As Int32, ByVal lVertY As Int32) As Int32
            Dim lIdx As Int32 = lVertY * Width + lVertX
            If lIdx > -1 AndAlso lIdx < HeightMap.Length Then Return CInt(HeightMap(lIdx)) Else Return 0
        End Function

        'NOTE: I changed X and Z to Int32 for performance... also, this returned a Single before, now it is an Int32
        Public Function GetHeightAtLocation(ByVal lLocX As Int32, ByVal lLocZ As Int32) As Int32
            Dim fTX As Single   'translated X 
            Dim fTZ As Single   'translated Z 
            Dim lCol As Int32
            Dim lRow As Int32

            Dim yA As Int32
            Dim yB As Int32
            Dim yC As Int32
            Dim yD As Int32

            Dim lIdx As Int32

            fTX = mlHalfWidth + CSng(lLocX / CellSpacing)
            fTZ = mlHalfHeight + CSng(lLocZ / CellSpacing)

            lCol = CInt(Math.Floor(fTX))
            lRow = CInt(Math.Floor(fTZ))
            fTX -= lCol
            fTZ -= lRow

            lIdx = (lRow * Width) + lCol
            If lIdx < 0 Or lIdx > HeightMap.Length - 1 Then yA = 0 Else yA = HeightMap(lIdx)
            lIdx = (lRow * Width) + lCol + 1
            If lIdx < 0 Or lIdx > HeightMap.Length - 1 Then yB = 0 Else yB = HeightMap(lIdx)
            lIdx = ((lRow + 1) * Width) + lCol
            If lIdx < 0 Or lIdx > HeightMap.Length - 1 Then yC = 0 Else yC = HeightMap(lIdx)
            lIdx = ((lRow + 1) * Width) + lCol + 1
            If lIdx < 0 Or lIdx > HeightMap.Length - 1 Then yD = 0 Else yD = HeightMap(lIdx)

            Dim fV1 As Single = yA * (1 - fTX) + yB * fTX
            Dim fV2 As Single = yC * (1 - fTX) + yD * fTX
            Return CInt(Math.Max((fV1 * (1 - fTZ) + fV2 * fTZ) * ml_Y_Mult, WaterHeight * ml_Y_Mult))
        End Function

        'NOTE: Changed Singles to Int32
        Public Function HasLineOfSight(ByVal lX1 As Int32, ByVal lZ1 As Int32, ByVal lX2 As Int32, ByVal lZ2 As Int32, ByVal lAttackerHeight As Int32) As Boolean
            Dim lY1 As Int32
            Dim lY2 As Int32

            lY1 = GetHeightAtLocation(lX1, lZ1)
            lY2 = GetHeightAtLocation(lX2, lZ2)

            Return HasLineOfSight(lX1, lY1, lZ1, lX2, lY2, lZ2, lAttackerHeight)
        End Function

        'NOTE: Changed Singles to Int32
        Public Function HasLineOfSight(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lZ1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32, ByVal lZ2 As Int32, ByVal lAttackerHt As Int32) As Boolean
            Dim lPt1Col As Int32
            Dim lPt1Row As Int32

            Dim lPt2Col As Int32
            Dim lPt2Row As Int32

            Dim fTmp As Single

            Dim lStep As Int32
            Dim lStepCnt As Int32

            Dim fXMod As Single
            Dim fZMod As Single

            Dim fX As Single
            Dim fZ As Single

            Dim fMaxHt As Single

            Dim lHtDiff As Int32

            Dim lAttackerHeight As Int32

            Dim lIdx As Int32

            fTmp = mlHalfWidth + CSng(lX1 / CellSpacing)
            lPt1Col = CInt(Math.Floor(fTmp))
            fTmp = mlHalfHeight + CSng(lZ1 / CellSpacing)
            lPt1Row = CInt(Math.Floor(fTmp))

            fTmp = mlHalfWidth + CSng(lX2 / CellSpacing)
            lPt2Col = CInt(Math.Floor(fTmp))
            fTmp = mlHalfHeight + CSng(lZ2 / CellSpacing)
            lPt2Row = CInt(Math.Floor(fTmp))

            lAttackerHeight = lY1 + lAttackerHt

            If lY1 < lY2 Then
                'in this case, we need to do an angle algorithm, but to reduce computations... we do it bro style
                fXMod = lPt2Col - lPt1Col
                fZMod = lPt2Row - lPt1Row
                lStepCnt = CInt(Math.Max(Math.Abs(fXMod), Math.Abs(fZMod)))
                fXMod /= lStepCnt
                fZMod /= lStepCnt

                fX = lPt1Col
                fZ = lPt1Row

                lHtDiff = lY2 - lY1
                For lStep = 0 To lStepCnt - 1
                    fMaxHt = (CSng(lStep / lStepCnt) * lHtDiff) + lAttackerHeight
                    lIdx = CInt((fZ * Width) + fX)
                    If lIdx < 0 OrElse lIdx > HeightMap.GetUpperBound(0) OrElse HeightMap(lIdx) * ml_Y_Mult > fMaxHt Then
                        Return False
                    End If
                    fX += fXMod
                    fZ += fZMod
                Next lStep
            Else
                'in this case, we need to do the reverse
                fXMod = lPt1Col - lPt2Col
                fZMod = lPt1Row - lPt2Row
                lStepCnt = CInt(Math.Max(Math.Abs(fXMod), Math.Abs(fZMod)))
                fXMod /= lStepCnt
                fZMod /= lStepCnt

                fX = lPt2Col
                fZ = lPt2Row

                lHtDiff = lY1 - lY2
                For lStep = 0 To lStepCnt - 1
                    fMaxHt = lAttackerHeight - (CSng(lStep / lStepCnt) * lHtDiff)
                    lIdx = CInt((fZ * Width) + fX)
                    If HeightMap(lIdx) * ml_Y_Mult > fMaxHt Then
                        Return False
                    End If
                    fX += fXMod
                    fZ += fZMod
                Next lStep
            End If
            Return True
        End Function

        Public Sub CleanResources()
            'A call to this sub means that the program wants to release all but the bare essential data...
            mbHMReady = False
            Erase HeightMap
        End Sub
        Public Function TerrainGradeBuildable(ByVal lLocX As Int32, ByVal lLocZ As Int32) As Boolean
            'Need to convert locX to Vert
            Dim fX As Single = CSng(lLocX / CellSpacing) + (TerrainClass.Width \ 2)
            Dim fZ As Single = CSng(lLocZ / CellSpacing) + (TerrainClass.Height \ 2)
            Dim vecTemp As Vector3 = GetTerrainNormalEx(fX, fZ)
            Return Not (Math.Abs(vecTemp.X) > 0.4F OrElse Math.Abs(vecTemp.Z) > 0.4F)
        End Function

        Private Shared mlPercLookup() As Int32 = Nothing
        Private Shared Sub SetupPercLookup()
            If mlPercLookup Is Nothing = False Then Return

            ReDim mlPercLookup(100)
            mlPercLookup(0) = 235
            mlPercLookup(1) = 576
            mlPercLookup(2) = 1152
            mlPercLookup(3) = 1728
            mlPercLookup(4) = 2304
            mlPercLookup(5) = 2880
            mlPercLookup(6) = 3456
            mlPercLookup(7) = 4032
            mlPercLookup(8) = 4608
            mlPercLookup(9) = 5184
            mlPercLookup(10) = 5760
            mlPercLookup(11) = 6336
            mlPercLookup(12) = 6912
            mlPercLookup(13) = 7488
            mlPercLookup(14) = 8064
            mlPercLookup(15) = 8640
            mlPercLookup(16) = 9216
            mlPercLookup(17) = 9792
            mlPercLookup(18) = 10368
            mlPercLookup(19) = 10944
            mlPercLookup(20) = 11520
            mlPercLookup(21) = 12096
            mlPercLookup(22) = 12672
            mlPercLookup(23) = 13248
            mlPercLookup(24) = 13824
            mlPercLookup(25) = 14400
            mlPercLookup(26) = 14976
            mlPercLookup(27) = 15552
            mlPercLookup(28) = 16128
            mlPercLookup(29) = 16704
            mlPercLookup(30) = 17280
            mlPercLookup(31) = 17856
            mlPercLookup(32) = 18432
            mlPercLookup(33) = 19008
            mlPercLookup(34) = 19584
            mlPercLookup(35) = 20160
            mlPercLookup(36) = 20736
            mlPercLookup(37) = 21312
            mlPercLookup(38) = 21888
            mlPercLookup(39) = 22464
            mlPercLookup(40) = 23040
            mlPercLookup(41) = 23616
            mlPercLookup(42) = 24192
            mlPercLookup(43) = 24768
            mlPercLookup(44) = 25344
            mlPercLookup(45) = 25920
            mlPercLookup(46) = 26496
            mlPercLookup(47) = 27072
            mlPercLookup(48) = 27648
            mlPercLookup(49) = 28224
            mlPercLookup(50) = 28800
            mlPercLookup(51) = 29376
            mlPercLookup(52) = 29952
            mlPercLookup(53) = 30528
            mlPercLookup(54) = 31104
            mlPercLookup(55) = 31680
            mlPercLookup(56) = 32256
            mlPercLookup(57) = 32832
            mlPercLookup(58) = 33408
            mlPercLookup(59) = 33984
            mlPercLookup(60) = 34560
            mlPercLookup(61) = 35136
            mlPercLookup(62) = 35712
            mlPercLookup(63) = 36288
            mlPercLookup(64) = 36864
            mlPercLookup(65) = 37440
            mlPercLookup(66) = 38016
            mlPercLookup(67) = 38592
            mlPercLookup(68) = 39168
            mlPercLookup(69) = 39744
            mlPercLookup(70) = 40320
            mlPercLookup(71) = 40896
            mlPercLookup(72) = 41472
            mlPercLookup(73) = 42048
            mlPercLookup(74) = 42624
            mlPercLookup(75) = 43200
            mlPercLookup(76) = 43776
            mlPercLookup(77) = 44352
            mlPercLookup(78) = 44928
            mlPercLookup(79) = 45504
            mlPercLookup(80) = 46080
            mlPercLookup(81) = 46656
            mlPercLookup(82) = 47232
            mlPercLookup(83) = 47808
            mlPercLookup(84) = 48384
            mlPercLookup(85) = 48960
            mlPercLookup(86) = 49536
            mlPercLookup(87) = 50112
            mlPercLookup(88) = 50688
            mlPercLookup(89) = 51264
            mlPercLookup(90) = 51840
            mlPercLookup(91) = 52416
            mlPercLookup(92) = 52992
            mlPercLookup(93) = 53568
            mlPercLookup(94) = 54144
            mlPercLookup(95) = 54720
            mlPercLookup(96) = 55296
            mlPercLookup(97) = 55872
            mlPercLookup(98) = 56448
            mlPercLookup(99) = 57024
            mlPercLookup(100) = 57600
        End Sub


        Private Sub GenerateTerrain(ByVal lSeed As Int32)
            SetupPercLookup()

            Dim lWidthSpans() As Int32
            Dim lHeightSpans() As Int32
            Dim X As Int32
            Dim lVal As Int32
            Dim lWaterPerc As Int32
            Dim bDone As Boolean
            Dim bLandmassLoop As Boolean
            Dim lStartX As Int32
            Dim lStartY As Int32
            Dim Y As Int32
            Dim lIdx As Int32
            Dim lPass As Int32

            Dim lSubXMax As Int32
            Dim lSubYMax As Int32
            Dim lSubX As Int32
            Dim lSubY As Int32

            Dim lTemp As Int32

            Dim lTotalLand As Int32

            Dim HM_Type() As Byte

            Call Rnd(-1)
            Randomize(lSeed)

            'fill our spans
            ReDim lWidthSpans(mlPASSES - 1)
            For X = 0 To mlPASSES - 1
                lVal = CInt(2 ^ X)
                If lVal > (Width / mlPASSES) Then
                    lVal = lWidthSpans(X - 1) + CInt((Width / mlPASSES) * (X / mlPASSES))
                End If
                lWidthSpans(X) = lVal
            Next X
            ReDim lHeightSpans(mlPASSES - 1)
            For X = 0 To mlPASSES - 1
                lVal = CInt(2 ^ X)
                If lVal > (Height / mlPASSES) Then
                    lVal = lHeightSpans(X - 1) + CInt((Height / mlPASSES) * (X / mlPASSES))
                End If
                lHeightSpans(X) = lVal
            Next X

            ReDim HeightMap(Width * Height)
            ReDim HM_Type(Width * Height)

            'Get our map type
            'MapType = -1
            'lTemp = Int(Rnd() * 100) + 1
            'Select Case lTemp
            '    Case Is < me_Planet_Type.ePT_Acid
            '        MapType = me_Planet_Type.ePT_Acid
            '    Case Is < me_Planet_Type.ePT_Barren
            '        MapType = me_Planet_Type.ePT_Barren
            '    Case Is < me_Planet_Type.ePT_GeoPlastic
            '        MapType = me_Planet_Type.ePT_GeoPlastic
            '    Case Is < me_Planet_Type.ePT_Desert
            '        MapType = me_Planet_Type.ePT_Desert
            '    Case Is < me_Planet_Type.ePT_Adaptable
            '        MapType = me_Planet_Type.ePT_Adaptable
            '    Case Is < me_Planet_Type.ePT_Tundra
            '        MapType = me_Planet_Type.ePT_Tundra
            '    Case Is < me_Planet_Type.ePT_Terran
            '        MapType = me_Planet_Type.ePT_Terran
            '    Case Else
            '        MapType = me_Planet_Type.ePT_Waterworld
            'End Select

            lTemp = CInt(Int(Rnd() * 100) + 1)


            'Get our water percentage based on map type
            lWaterPerc = GetWaterPerc(MapType)

            'Now, set our values to water unless there is none
            If lWaterPerc = 0 Then
                'set em all to land
                WaterHeight = 0
                For X = 0 To HeightMap.Length - 1
                    HeightMap(X) = 100
                    HM_Type(X) = 1
                Next X
            Else
                For X = 0 To HeightMap.Length - 1
                    HeightMap(X) = 0
                    HM_Type(X) = 0
                Next X

                'Base the waterheight off of the waterperc, if our water perc is practically nothing, then
                '  we don't want waterheight to be 256 :)
                'WaterHeight = CByte(Int(Rnd() * (lWaterPerc + 20)) + 1)

                If MapType = PlanetType.eTerran Then
                    WaterHeight = 70
                ElseIf MapType = PlanetType.eWaterWorld Then
                    WaterHeight = 160
                Else : WaterHeight = 40
                End If

                'Generate landmasses
                bDone = False
                While bDone = False
                    lStartX = CInt(Int(Rnd() * Width) + 1)
                    lStartY = CInt(Int(Rnd() * Height) + 1)
                    X = lStartX
                    Y = lStartY

                    bLandmassLoop = False
                    While bLandmassLoop = False
                        'Get our next movement...
                        Select Case Int(Rnd() * 8) + 1
                            Case 1 : Y -= 1
                            Case 2 : Y -= 1 : X += 1
                            Case 3 : X += 1
                            Case 4 : Y += 1 : X += 1
                            Case 5 : Y += 1
                            Case 6 : Y += 1 : X -= 1
                            Case 7 : X -= 1
                            Case 8 : X -= 1 : Y -= 1
                        End Select
                        'Validate ranges
                        If X < 0 Then X = Width - 1
                        If X >= Width Then X = 0
                        If Y < 0 Then Y = 1
                        If Y >= Height Then Y = Height - 2

                        lIdx = (Y * Width) + X
                        If HM_Type(lIdx) = 0 Then
                            HeightMap(lIdx) = CByte(WaterHeight + 32)      'to ensure that we are way above water
                            HM_Type(lIdx) = 1
                            lTotalLand += 1
                        End If
                        'Ok, this fixes a very nasty rounding bug that occurs on 32-bit systems.
                        ' basically, it would calculate the equation as being equal because one value
                        ' would be stored in memory identically to another. Therefore, we do some voodoo here
                        ' to see if the rounding error is happening....
                        'Dim dblVal1 As Double = (CDbl(lTotalLand) / CDbl(Width * Height)) 'get our errored value
                        'Dim lCompareVal As Int32 = 0
                        'If (dblVal1 * 100) > (100 - lWaterPerc) - 4 Then                ' are we close enough to test, if not, don't waste time
                        '    Dim sPerc1 As String = dblVal1.ToString()                   'store the errored value as a string
                        '    If sPerc1.Contains("E") = True Then                         'if the string contains an E, we're not even close
                        '        dblVal1 = 0
                        '        lCompareVal = 0
                        '    Else
                        '        sPerc1 = Mid$(sPerc1, 1, 5)                             'Ok, lop off some digits
                        '        lCompareVal = CInt(Val(sPerc1) * 100)                             'now, multiply the results accordingly and we have a valid test
                        '    End If
                        'Else : lCompareVal = CInt(dblVal1 * 100)
                        'End If
                        Dim lLookupIdx As Int32 = 100 - lWaterPerc

                        If X = lStartX AndAlso Y = lStartY Then
                            bLandmassLoop = True
                        ElseIf lTotalLand >= mlPercLookup(lLookupIdx) Then
                            'ElseIf lCompareVal >= 100 - lWaterPerc Then 'ElseIf CInt(Math.Floor((lTotalLand / (Width * Height)) * 100)) >= 100 - lWaterPerc Then
                            bLandmassLoop = True
                            bDone = True
                            'goUILib.AddNotification(Me.mlSeed & ": " & lTotalLand, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        End If
                        'MSC - 10/10/08 - remarked this out and added the above to change the double to an int32
                        'Dim dblVal1 As Double = (CDbl(lTotalLand) / CDbl(Width * Height)) 'get our errored value
                        'If (dblVal1 * 100) > (100 - lWaterPerc) - 2 Then                ' are we close enough to test, if not, don't waste time
                        '    Dim sPerc1 As String = dblVal1.ToString()                   'store the errored value as a string
                        '    If sPerc1.Contains("E") = True Then                         'if the string contains an E, we're not even close
                        '        dblVal1 = 0
                        '    Else
                        '        sPerc1 = Mid$(sPerc1, 1, 5)                             'Ok, lop off some digits
                        '        dblVal1 = Val(sPerc1) * 100                             'now, multiply the results accordingly and we have a valid test
                        '    End If
                        'End If

                        'If X = lStartX AndAlso Y = lStartY Then
                        '    bLandmassLoop = True
                        'ElseIf dblVal1 >= 100 - lWaterPerc Then 'ElseIf CInt(Math.Floor((lTotalLand / (Width * Height)) * 100)) >= 100 - lWaterPerc Then
                        '    bLandmassLoop = True
                        '    bDone = True
                        '    'goUILib.AddNotification(Me.mlSeed & ": " & lTotalLand, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        'End If
                    End While
                End While
            End If
            mlWaterPerc = lWaterPerc

            'Generate the terrain on the map...
            For lPass = mlPASSES - 1 To 0 Step -1
                For X = 0 To Width - 1 Step lWidthSpans(lPass)
                    For Y = 0 To Height - 1 Step lHeightSpans(lPass)
                        lVal = CInt(Int(Rnd() * 255) + 1)

                        lSubXMax = X + lWidthSpans(lPass) - 1
                        If lSubXMax > Width Then lSubXMax = Width - 1
                        lSubYMax = Y + lHeightSpans(lPass) - 1
                        If lSubYMax > Height Then lSubYMax = Height - 1

                        For lSubX = X To lSubXMax
                            For lSubY = Y To lSubYMax
                                lIdx = lSubY * Width + lSubX
                                If HM_Type(lIdx) = 0 Then
                                    lTemp = CInt((lVal / 255) * WaterHeight)
                                Else
                                    lTemp = lVal
                                End If
                                HeightMap(lIdx) = SmoothTerrainVal(lSubX, lSubY, lTemp)
                            Next lSubY
                        Next lSubX
                    Next Y
                Next X
            Next lPass

            'Now, assuming we are lunar...
            If MapType = PlanetType.eBarren Then
                DoLunar()
            End If

            'Now, apply final filter to smooth everything over
            For lPass = 1 To 0 Step -1
                For X = 0 To Width - 1 Step lWidthSpans(lPass)
                    For Y = 0 To Height - 1 Step lHeightSpans(lPass)
                        lSubXMax = X + lWidthSpans(lPass) - 1
                        If lSubXMax > Width Then lSubXMax = Width - 1
                        lSubYMax = Y + lHeightSpans(lPass) - 1
                        If lSubYMax > Height Then lSubYMax = Height - 1

                        For lSubX = X To lSubXMax
                            For lSubY = Y To lSubYMax
                                lIdx = lSubY * Width + lSubX

                                lTemp = HeightMap(lIdx)
                                HeightMap(lIdx) = SmoothTerrainVal(lSubX, lSubY, lTemp)
                            Next lSubY
                        Next lSubX
                    Next Y
                Next X
            Next lPass

            If MapType = PlanetType.eGeoPlastic OrElse MapType = PlanetType.eWaterWorld Then
                CreateVolcanoes()
            End If

            'Now, accentuate
            lTemp = 0 : lVal = 0
            For X = 0 To Width - 1
                For Y = 0 To Height - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > lTemp Then lTemp = HeightMap(lIdx)
                    If HeightMap(lIdx) > WaterHeight Then lVal += HeightMap(lIdx)
                Next Y
            Next X
            lVal = lVal \ (Width * Height) 'CInt(lVal / (Width * Height))
            'lVal = lVal + CInt((lTemp - lVal) / 1.8)        'what is 1.8?
            lVal += CInt((lTemp - lVal) / 1.8F)
            For X = 0 To Width - 1
                For Y = 0 To Height - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > lVal Then HeightMap(lIdx) = CByte(HeightMap(lIdx) * (255 / lTemp))
                Next Y
            Next X

            If MapType = PlanetType.eBarren Then
                For Y = 0 To Height - 1
                    For X = 0 To Width - 1
                        lIdx = Y * Width + X
                        lTemp = HeightMap(lIdx)
                        lTemp += (CInt(Rnd() * 20) - 10)
                        If lTemp < 0 Then lTemp = 0
                        If lTemp > 255 Then lTemp = 255
                        HeightMap(lIdx) = CByte(lTemp)
                    Next X
                Next Y
                MaximizeTerrain()
            ElseIf MapType = PlanetType.eDesert Then
                DoDesert()

                'and an additional smooth over
                For X = 0 To Width - 1
                    For Y = 0 To Height - 1
                        lIdx = Y * Width + X
                        HeightMap(lIdx) = SmoothTerrainVal(X, Y, HeightMap(lIdx))
                    Next Y
                Next X
            ElseIf MapType = PlanetType.eAcidic Then
                DoAcidPlateaus()
            ElseIf MapType = PlanetType.eTerran Then
                DoPeakAccents()
            End If

            'One last soften...
            For X = 0 To Width - 1
                For Y = 0 To Height - 1
                    lIdx = Y * Width + X
                    HeightMap(lIdx) = SmoothTerrainVal(X, Y, HeightMap(lIdx))
                Next Y
            Next X

            If MapType = PlanetType.eGeoPlastic Then
                DoGeoPlastic()
            ElseIf MapType = PlanetType.eAcidic Then
                DoAcidic()
            ElseIf MapType = PlanetType.eWaterWorld Then
                If mptVolcanoes Is Nothing = False Then
                    For X = 0 To mptVolcanoes.GetUpperBound(0)
                        lIdx = mptVolcanoes(X).Y * Width + mptVolcanoes(X).X
                        HeightMap(lIdx) = 0
                    Next
                End If
                MaximizeTerrain()
            ElseIf MapType = PlanetType.eTerran Then
                MaximizeTerrain()
            End If

            'To ensure map wrapping lines up perfectly
            For Y = 0 To TerrainClass.Height - 1
                lIdx = (Y * TerrainClass.Width)
                Dim lOtherIdx As Int32 = (Y * TerrainClass.Width) + (TerrainClass.Width - 1)
                HeightMap(lOtherIdx) = HeightMap(lIdx)
            Next Y

            mbHMReady = True

        End Sub
        Private mptVolcanoes() As Point
        Private Sub CreateVolcanoes()
            Dim lCnt As Int32 '= CInt(Rnd() * 5) + 1

            If MapType = PlanetType.eWaterWorld Then
                If Rnd() * 100 > 60 Then Return

                ReDim mptVolcanoes(0)

                Dim lYRng As Int32 = Height \ 2
                Dim lYOffset As Int32 = Height \ 4
                Dim lIdx As Int32

                Dim lVolX As Int32
                Dim lVolY As Int32

                'Get our max height
                Dim lVolcanoVal As Int32
                For Y As Int32 = 0 To Height - 1
                    For X As Int32 = 0 To Width - 1
                        lIdx = Y * Width + X
                        If HeightMap(lIdx) > lVolcanoVal Then
                            lVolcanoVal = HeightMap(lIdx)
                            lVolX = X
                            lVolY = Y
                        End If
                    Next X
                Next Y
                lVolcanoVal += 5
                If lVolcanoVal > 255 Then lVolcanoVal = 255

                mptVolcanoes(0).X = lVolX
                mptVolcanoes(0).Y = lVolY

                Dim lSize As Int32 = CInt(Rnd() * 4) + 6
                For lLocY As Int32 = lVolY - lSize To lVolY + lSize
                    For lLocX As Int32 = lVolX - lSize To lVolX + lSize

                        Dim lTempX As Int32 = lLocX
                        Dim lTempY As Int32 = lLocY
                        If lTempX > Width - 1 Then lTempX -= Width
                        If lTempX < 0 Then lTempX += Width
                        If lTempY > Height - 1 Then lTempY -= Height
                        If lTempY < 0 Then lTempY += Height

                        lIdx = lTempY * Width + lTempX

                        If lLocX <> lVolX OrElse lLocY <> lVolY Then
                            If HeightMap(lIdx) > WaterHeight Then
                                If (Math.Abs(lLocX - lVolX) = 1 AndAlso Math.Abs(lLocY - lVolY) < 2) OrElse (Math.Abs(lLocY - lVolY) = 1 AndAlso Math.Abs(lLocX - lVolX) < 2) Then
                                    HeightMap(lIdx) = CByte(lVolcanoVal)
                                Else
                                    Dim lXVal As Int32 = lLocX - lVolX
                                    lXVal *= lXVal
                                    Dim lYVal As Int32 = lLocY - lVolY
                                    lYVal *= lYVal

                                    Dim fTotalVal As Single = CSng(Math.Sqrt(lXVal + lYVal))
                                    Dim lOriginal As Int32 = HeightMap(lIdx)
                                    Dim lVal As Int32 = WaterHeight + CInt((lVolcanoVal - WaterHeight) * (1.0F - (fTotalVal / (lSize * 3))))

                                    If lVal > lOriginal Then lVal = (lVal + lOriginal + lVal) \ 3 Else lVal = lOriginal
                                    If lVal > lVolcanoVal Then lVal = lVolcanoVal
                                    If lVal < lOriginal Then lVal = lOriginal

                                    HeightMap(lIdx) = CByte(lVal)
                                End If
                            End If
                        Else : HeightMap(lIdx) = CByte(lVolcanoVal)
                        End If
                    Next lLocX
                Next lLocY

            Else
                lCnt = CInt(Rnd() * 5) + 1
                ReDim mptVolcanoes(lCnt - 1)

                Dim lYRng As Int32 = Height \ 2
                Dim lYOffset As Int32 = Height \ 4
                Dim lIdx As Int32

                'Get our max height
                Dim lVolcanoVal As Int32
                For Y As Int32 = 0 To Height - 1
                    For X As Int32 = 0 To Width - 1
                        lIdx = Y * Width + X
                        If HeightMap(lIdx) > lVolcanoVal Then lVolcanoVal = HeightMap(lIdx)
                    Next X
                Next Y
                lVolcanoVal += 5
                If lVolcanoVal > 255 Then lVolcanoVal = 255

                While lCnt > 0
                    Dim Y As Int32 = CInt(Rnd() * lYRng) + lYOffset
                    Dim X As Int32 = CInt(Rnd() * (Width - 1))

                    If HeightMap(Y * Width + X) > WaterHeight Then
                        mptVolcanoes(lCnt - 1).X = X
                        mptVolcanoes(lCnt - 1).Y = Y
                        lCnt -= 1

                        Dim lSize As Int32 = CInt(Rnd() * 4) + 6
                        For lLocY As Int32 = Y - lSize To Y + lSize
                            For lLocX As Int32 = X - lSize To X + lSize

                                Dim lTempX As Int32 = lLocX
                                Dim lTempY As Int32 = lLocY
                                If lTempX > Width - 1 Then lTempX -= Width
                                If lTempX < 0 Then lTempX += Width
                                If lTempY > Height - 1 Then lTempY -= Height
                                If lTempY < 0 Then lTempY += Height

                                lIdx = lTempY * Width + lTempX

                                If lLocX <> X OrElse lLocY <> Y Then
                                    If HeightMap(lIdx) > WaterHeight Then
                                        If (Math.Abs(lLocX - X) = 1 AndAlso Math.Abs(lLocY - Y) < 2) OrElse (Math.Abs(lLocY - Y) = 1 AndAlso Math.Abs(lLocX - X) < 2) Then
                                            HeightMap(lIdx) = CByte(lVolcanoVal)
                                        Else
                                            Dim lXVal As Int32 = lLocX - X
                                            lXVal *= lXVal
                                            Dim lYVal As Int32 = lLocY - Y
                                            lYVal *= lYVal

                                            Dim fTotalVal As Single = CSng(Math.Sqrt(lXVal + lYVal))
                                            Dim lOriginal As Int32 = HeightMap(lIdx)
                                            Dim lVal As Int32 = WaterHeight + CInt((lVolcanoVal - WaterHeight) * (1.0F - (fTotalVal / (lSize * 3))))

                                            If lVal > lOriginal Then lVal = (lVal + lOriginal + lVal) \ 3 Else lVal = lOriginal
                                            If lVal > lVolcanoVal Then lVal = lVolcanoVal
                                            If lVal < lOriginal Then lVal = lOriginal

                                            HeightMap(lIdx) = CByte(lVal)
                                        End If
                                    End If
                                Else : HeightMap(lIdx) = CByte(lVolcanoVal)
                                End If
                            Next lLocX
                        Next lLocY
                    End If

                End While
            End If

        End Sub

        Private Sub DoPeakAccents()
            Dim X As Int32
            Dim Y As Int32
            Dim lIdx As Int32
            Dim ptPeaks() As Point = Nothing
            Dim lPeakHt() As Int32 = Nothing
            Dim lPeakUB As Int32 = -1
            Dim lMinPeakHt As Int32 = 0

            For Y = 0 To Height - 1
                For X = 0 To Width - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > lMinPeakHt Then lMinPeakHt = HeightMap(lIdx)
                Next X
            Next Y
            lMinPeakHt -= CInt((Rnd() * 15) + 5)

            For Y = 0 To Height - 1
                For X = 0 To Width - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > lMinPeakHt Then
                        Dim bSkip As Boolean = False
                        For lTmpIdx As Int32 = 0 To lPeakUB
                            If Math.Abs(ptPeaks(lTmpIdx).X - X) < 3 AndAlso Math.Abs(ptPeaks(lTmpIdx).Y - Y) < 3 Then
                                If lPeakHt(lTmpIdx) < HeightMap(lIdx) Then
                                    ptPeaks(lTmpIdx).X = X
                                    ptPeaks(lTmpIdx).Y = Y
                                    lPeakHt(lTmpIdx) = HeightMap(lIdx)
                                End If
                                bSkip = True
                            End If
                        Next lTmpIdx

                        If bSkip = False Then
                            lPeakUB += 1
                            ReDim Preserve ptPeaks(lPeakUB)
                            ReDim Preserve lPeakHt(lPeakUB)
                            ptPeaks(lPeakUB).X = X
                            ptPeaks(lPeakUB).Y = Y
                            lPeakHt(lPeakUB) = HeightMap(lIdx)
                        End If
                    End If
                Next X
            Next Y

            Dim mfMults(5) As Single
            mfMults(0) = 1.0F
            mfMults(1) = 0.667F
            mfMults(2) = 0.475F
            mfMults(3) = 0.365F
            mfMults(4) = 0.304F
            mfMults(5) = 0.276F

            'Now, check our UB
            If lPeakUB <> -1 Then
                For lPkIdx As Int32 = 0 To lPeakUB

                    Dim lVolX As Int32 = ptPeaks(lPkIdx).X
                    Dim lVolY As Int32 = ptPeaks(lPkIdx).Y

                    For lLocY As Int32 = lVolY - 6 To lVolY + 6
                        For lLocX As Int32 = lVolX - 6 To lVolX + 6

                            Dim lTempX As Int32 = lLocX
                            Dim lTempY As Int32 = lLocY
                            If lTempX > Width - 1 Then lTempX -= Width
                            If lTempX < 0 Then lTempX += Width
                            If lTempY > Height - 1 Then lTempY -= Height
                            If lTempY < 0 Then lTempY += Height

                            lIdx = lTempY * Width + lTempX

                            If lLocX <> lVolX OrElse lLocY <> lVolY Then
                                If HeightMap(lIdx) > WaterHeight Then
                                    If (Math.Abs(lLocX - lVolX) = 1 AndAlso Math.Abs(lLocY - lVolY) < 2) OrElse (Math.Abs(lLocY - lVolY) = 1 AndAlso Math.Abs(lLocX - lVolX) < 2) Then
                                        HeightMap(lIdx) = 255
                                    Else
                                        Dim lXVal As Int32 = lLocX - lVolX
                                        lXVal *= lXVal
                                        Dim lYVal As Int32 = lLocY - lVolY
                                        lYVal *= lYVal

                                        Dim fTotalVal As Single = lXVal + lYVal 'CSng(Math.Sqrt(lXVal + lYVal))
                                        Dim lOriginal As Int32 = HeightMap(lIdx)
                                        Dim lVal As Int32 = WaterHeight + CInt((255 - WaterHeight) * (1.0F - (fTotalVal / 18.0F)))
                                        If lVal > lOriginal Then lVal = (lVal + lOriginal + lVal) \ 3 Else lVal = lOriginal
                                        If lVal > 255 Then lVal = 255
                                        If lVal < lOriginal \ 2 Then lVal = lOriginal \ 2

                                        HeightMap(lIdx) = CByte(lVal)
                                    End If
                                End If
                            Else : HeightMap(lIdx) = CByte(255)
                            End If
                        Next lLocX
                    Next lLocY
                Next lPkIdx

            End If
        End Sub

        Private Sub DoGeoPlastic()
            'GeoPlastic - 
            Dim lIdx As Int32
            Dim lTemp As Int32
            Dim lAboveWaterHeight As Int32 = 255 - WaterHeight

            'short dropoff cliffs where rivers of lava flow. 
            For Y As Int32 = 0 To Height - 1
                For X As Int32 = 0 To Width - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > WaterHeight + 10 Then
                        lTemp = CInt(HeightMap(lIdx)) + 10
                    Else : lTemp = 0
                    End If
                    If lTemp > 255 Then lTemp = 255
                    HeightMap(lIdx) = CByte(lTemp)
                Next X
            Next Y

            'Magma regions (water's beach area). 
            'Mountainous regions with no sharp inclines (rolling but with peaks). 

            '"Rips" in the ground for steam vents.
            Dim lCnt As Int32 = CInt(Rnd() * 10) + 15
            While lCnt <> 0
                Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
                Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

                lIdx = lSZ * Width + lSX
                If HeightMap(lIdx) > WaterHeight AndAlso HeightMap(lIdx) < CInt(lAboveWaterHeight * 0.6F) Then
                    Dim lSize As Int32 = CInt(Rnd() * 10) + 10
                    Dim lDirX As Int32
                    Dim lDirZ As Int32

                    If Rnd() * 100 < 50 Then lDirX = -1 Else lDirX = 1
                    If Rnd() * 100 < 50 Then lDirZ = -1 Else lDirZ = 1

                    lCnt -= 1

                    While lSize > 0
                        If Rnd() * 100 < 50 Then
                            lSX += lDirX
                        Else : lSZ += lDirZ
                        End If

                        If lSX < 0 Then
                            lSX += Width
                        ElseIf lSX > Width - 1 Then
                            lSX -= Width
                        End If
                        If lSZ < 0 Then
                            lSZ += Height
                        ElseIf lSZ > Height - 1 Then
                            lSZ -= Height
                        End If

                        lIdx = lSZ * Width + lSX

                        HeightMap(lIdx) = 0

                        lSize -= 1
                    End While

                End If
            End While


            'Reset the volcano center's to 0
            If mptVolcanoes Is Nothing = False Then
                For X As Int32 = 0 To mptVolcanoes.GetUpperBound(0)
                    lIdx = mptVolcanoes(X).Y * Width + mptVolcanoes(X).X
                    HeightMap(lIdx) = 0
                Next
            End If

            For Y As Int32 = 0 To Height - 1
                For X As Int32 = 0 To Width - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) < WaterHeight Then
                        If Rnd() * 100 < 1 Then
                            lIdx = 1        'needed to do this to force the compiler not to wax this method... we need to call rnd to sync with client
                        End If
                    End If
                Next X
            Next Y

        End Sub

        Private Sub DoAcidPlateaus()
            'Ok, at this point, everything is accentuated... we are almost done
            Dim fPlateauStrength() As Single
            ReDim fPlateauStrength((Width * Height) - 1)

            'Ok, now... let's create our plateaus
            Dim lCnt As Int32 = CInt(Rnd() * 3) + 3
            Dim lIdx As Int32 ''

            While lCnt > 0
                Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
                Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

                lIdx = lSZ * Width + lSX
                If HeightMap(lIdx) > WaterHeight Then
                    lCnt -= 1

                    Select Case CInt(Rnd() * 100)
                        Case Is < 50
                            'EAST / WEST
                            'determine which side is the high side
                            If CInt(Rnd() * 100) < 50 Then
                                'East
                                Dim lEndVal As Int32 = Math.Max(lSX - (CInt(Rnd() * 30) + 30), 0)
                                For X As Int32 = lSX To lEndVal Step -1    '0 step -1

                                    Dim fHtMult As Single = 0
                                    If lSX - lEndVal <> 0 Then fHtMult = CSng((X - lEndVal) / (lSX - lEndVal))

                                    'For Y As Int32 = 0 To Height - 1
                                    For Y As Int32 = Math.Max(lSZ - CInt(Rnd() * 30 + 15), 0) To Math.Min(Height - 1, lSZ + CInt(Rnd() * 30 + 15))
                                        lIdx = Y * Width + X
                                        Dim lTemp As Int32 = HeightMap(lIdx)
                                        If lTemp > WaterHeight Then
                                            Dim fTemp As Single = fPlateauStrength(lIdx)
                                            If fTemp = 0 Then
                                                fTemp = fHtMult
                                            Else : fTemp = (fTemp + fHtMult) / 2.0F
                                            End If
                                            fPlateauStrength(lIdx) = fTemp
                                        End If
                                    Next Y
                                Next X
                            Else
                                'West
                                Dim lEndVal As Int32 = Math.Min(Width - 1, lSX + (CInt(Rnd() * 30) + 30))
                                For X As Int32 = lSX To lEndVal 'Width - 1

                                    Dim fHtMult As Single = 0
                                    If lEndVal - lSX <> 0 Then fHtMult = CSng((X - lSX) / (lEndVal - lSX))

                                    'For Y As Int32 = 0 To Height - 1
                                    For Y As Int32 = Math.Max(lSZ - CInt(Rnd() * 30 + 15), 0) To Math.Min(Height - 1, lSZ + CInt(Rnd() * 30 + 15))
                                        lIdx = Y * Width + X
                                        Dim lTemp As Int32 = HeightMap(lIdx)
                                        If lTemp > WaterHeight Then
                                            Dim fTemp As Single = fPlateauStrength(lIdx)
                                            If fTemp = 0 Then
                                                fTemp = fHtMult
                                            Else : fTemp = (fTemp + fHtMult) / 2.0F
                                            End If
                                            fPlateauStrength(lIdx) = fTemp
                                        End If
                                    Next Y
                                Next X
                            End If
                        Case Else ' Is < 40
                            'NORTH / SOUTH
                            'determine which side is the high side
                            If CInt(Rnd() * 100) < 50 Then
                                'NORTH
                                Dim lEndVal As Int32 = Math.Max(lSZ - (CInt(Rnd() * 30) + 30), 0)
                                For Y As Int32 = lSZ To lEndVal Step -1 '0 Step -1
                                    Dim fHtMult As Single = 0
                                    If lSZ - lEndVal <> 0 Then fHtMult = CSng((Y - lEndVal) / (lSZ - lEndVal))
                                    'For X As Int32 = 0 To Width - 1
                                    For X As Int32 = Math.Max(lSX - CInt(Rnd() * 30 + 15), 0) To Math.Min(Width - 1, lSX + CInt(Rnd() * 30 + 15))
                                        lIdx = Y * Width + X
                                        Dim lTemp As Int32 = HeightMap(lIdx)
                                        If lTemp > WaterHeight Then
                                            Dim fTemp As Single = fPlateauStrength(lIdx)
                                            If fTemp = 0 Then
                                                fTemp = fHtMult
                                            Else : fTemp = (fTemp + fHtMult) / 2.0F
                                            End If
                                            fPlateauStrength(lIdx) = fTemp
                                        End If
                                    Next X
                                Next Y
                            Else
                                'SOUTH
                                Dim lEndVal As Int32 = Math.Min(Height - 1, lSZ + (CInt(Rnd() * 30) + 30))
                                For Y As Int32 = lSZ To lEndVal 'Height - 1
                                    Dim fHtMult As Single = 0
                                    If lEndVal - lSZ <> 0 Then fHtMult = CSng((Y - lSZ) / (lEndVal - lSZ))

                                    'For X As Int32 = 0 To Width - 1
                                    For X As Int32 = Math.Max(lSX - CInt(Rnd() * 30 + 15), 0) To Math.Min(Width - 1, lSX + CInt(Rnd() * 30 + 15))
                                        lIdx = Y * Width + X
                                        Dim lTemp As Int32 = HeightMap(lIdx)
                                        If lTemp > WaterHeight Then
                                            Dim fTemp As Single = fPlateauStrength(lIdx)
                                            If fTemp = 0 Then
                                                fTemp = fHtMult
                                            Else : fTemp = (fTemp + fHtMult) / 2.0F
                                            End If
                                            fPlateauStrength(lIdx) = fTemp
                                        End If
                                    Next X
                                Next Y
                            End If
                    End Select

                End If
            End While

            'Now, go through all of our values
            Dim lTmpValWH As Int32 = 255 - WaterHeight
            For Y As Int32 = 0 To Height - 1
                For X As Int32 = 0 To Width - 1
                    lIdx = Y * Width + X
                    Dim lVal As Int32 = HeightMap(lIdx)

                    If lVal > WaterHeight Then
                        Dim fVal As Single = fPlateauStrength(lIdx) * lTmpValWH

                        'Now... I want to affect this by fVal
                        fVal = ((lVal + fVal) / 2.0F) + WaterHeight
                        If fVal > 255 Then
                            HeightMap(lIdx) = 255
                        ElseIf fVal < WaterHeight + 1 Then
                            HeightMap(lIdx) = CByte(WaterHeight + 1)
                        Else
                            HeightMap(lIdx) = CByte(fVal)
                        End If

                    End If
                Next X
            Next Y

            MaximizeTerrain()

        End Sub

        Private Sub MaximizeTerrain()
            Dim lIdx As Int32

            Dim lMinVal As Int32 = 255
            Dim lMaxVal As Int32 = 0

            For Y As Int32 = 0 To Height - 1
                For X As Int32 = 0 To Width - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > WaterHeight Then
                        If HeightMap(lIdx) > lMaxVal Then lMaxVal = HeightMap(lIdx)
                        If HeightMap(lIdx) < lMinVal Then lMinVal = HeightMap(lIdx)
                    End If
                Next X
            Next Y

            Dim lDiff As Int32 = lMaxVal - lMinVal
            Dim lDesiredDiff As Int32 = 255 - (WaterHeight + 1)

            For Y As Int32 = 0 To Height - 1
                For X As Int32 = 0 To Width - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > WaterHeight Then
                        Dim lVal As Int32 = HeightMap(lIdx)
                        lVal = CInt(((lVal - lMinVal) / lDiff) * lDesiredDiff) + WaterHeight + 1
                        If lVal < WaterHeight + 1 Then lVal = WaterHeight + 1
                        If lVal > 255 Then lVal = 255
                        HeightMap(lIdx) = CByte(lVal)
                    End If
                Next X
            Next Y
        End Sub

        Private Sub DoAcidic()
            Dim lIdx As Int32
            Dim lAboveWaterHeight As Int32 = 255 - WaterHeight

            'Deep chasms with acidic rivers flowing through them. 
            Dim lCnt As Int32 = CInt(Rnd() * 10) + 15
            While lCnt <> 0
                Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
                Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

                lIdx = lSZ * Width + lSX
                If HeightMap(lIdx) > WaterHeight AndAlso HeightMap(lIdx) < CInt(lAboveWaterHeight * 0.6F) Then
                    Dim lSize As Int32 = CInt(Rnd() * 10) + 10
                    Dim lDirX As Int32
                    Dim lDirZ As Int32

                    If Rnd() * 100 < 50 Then lDirX = -1 Else lDirX = 1
                    If Rnd() * 100 < 50 Then lDirZ = -1 Else lDirZ = 1

                    lCnt -= 1

                    While lSize > 0

                        Dim bXChg As Boolean = False
                        If Rnd() * 100 < 50 Then
                            lSX += lDirX
                            bXChg = True
                        Else : lSZ += lDirZ
                        End If


                        If lSX < 0 Then
                            lSX += Width
                        ElseIf lSX > Width - 1 Then
                            lSX -= Width
                        End If
                        If lSZ < 0 Then
                            lSZ += Height
                        ElseIf lSZ > Height - 1 Then
                            lSZ -= Height
                        End If

                        lIdx = lSZ * Width + lSX
                        HeightMap(lIdx) = 0

                        If Rnd() * 100 < 35 Then
                            Dim lTempX As Int32 = lSX
                            Dim lTempZ As Int32 = lSZ

                            If bXChg = True Then
                                lTempZ += lDirZ
                                If lTempZ < 0 Then
                                    lTempZ += Height
                                ElseIf lTempZ > Height - 1 Then
                                    lTempZ -= Height
                                End If
                            Else
                                lTempX += lDirX
                                If lTempX < 0 Then
                                    lTempX += Width
                                ElseIf lTempX > Width - 1 Then
                                    lTempX -= Width
                                End If
                            End If
                            lIdx = lTempZ * Width + lTempX
                            HeightMap(lIdx) = 0
                        End If

                        lSize -= 1
                    End While

                End If
            End While

            'One last soften...
            For X As Int32 = 0 To Width - 1
                For Y As Int32 = 0 To Height - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > WaterHeight Then
                        HeightMap(lIdx) = SmoothTerrainVal(X, Y, HeightMap(lIdx))
                    End If
                Next Y
            Next X
        End Sub

        Private Sub DoDesert()
            Dim lIdx As Int32

            'smooth mountain regions
            Dim lCnt As Int32 = CInt(Rnd() * 5) + 3
            While lCnt > 0
                Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
                Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

                lIdx = lSZ * Width + lSX
                If HeightMap(lIdx) > WaterHeight Then
                    lCnt -= 1

                    Dim lSizeW As Int32 = CInt(Rnd() * 40) + 15
                    Dim lHalfSizeW As Int32 = lSizeW \ 2
                    Dim lSizeH As Int32 = CInt(Rnd() * 40) + 15
                    Dim lHalfSizeH As Int32 = lSizeH \ 2

                    For Y As Int32 = -(CInt(lHalfSizeH * Math.Max(Rnd, 0.3))) To (CInt(lHalfSizeH * Math.Max(Rnd, 0.3)))
                        For X As Int32 = -(CInt(lHalfSizeW * Math.Max(Rnd, 0.3))) To (CInt(lHalfSizeW * Math.Max(Rnd, 0.3)))

                            Dim lMapX As Int32 = lSX + X
                            Dim lMapY As Int32 = lSZ + Y

                            If lMapX < 0 Then lMapX += Width
                            If lMapX > Width - 1 Then lMapX -= Width
                            If lMapY < 0 Then lMapY += Height
                            If lMapY > Height - 1 Then lMapY -= Height

                            lIdx = lMapY * Width + lMapX

                            If HeightMap(lIdx) > WaterHeight Then
                                Dim lTemp As Int32 = HeightMap(lIdx)
                                lTemp += CInt(Rnd() * 20) + 100
                                If lTemp > 255 Then lTemp = 255
                                If lTemp < 0 Then lTemp = 0

                                HeightMap(lIdx) = CByte(lTemp)
                            End If
                        Next X
                    Next Y

                    'Now, add nobs to the top and bottoms
                    For X As Int32 = -lHalfSizeW To lHalfSizeW
                        For Y As Int32 = CInt(Rnd() * 10) To 0 Step -1
                            Dim lMapX As Int32 = lSX + X
                            Dim lMapY As Int32 = lSZ - Y - lHalfSizeH

                            If lMapX < 0 Then lMapX += Width
                            If lMapX > Width - 1 Then lMapX -= Width
                            If lMapY < 0 Then lMapY += Height
                            If lMapY > Height - 1 Then lMapY -= Height

                            lIdx = lMapY * Width + lMapX

                            If HeightMap(lIdx) > WaterHeight Then
                                Dim lTemp As Int32 = HeightMap(lIdx)
                                lTemp += CInt(Rnd() * 20) + 100
                                If lTemp > 255 Then lTemp = 255
                                If lTemp < 0 Then lTemp = 0

                                HeightMap(lIdx) = CByte(lTemp)
                            End If
                        Next Y

                        For Y As Int32 = CInt(Rnd() * 10) To 0 Step -1
                            Dim lMapX As Int32 = lSX + X
                            Dim lMapY As Int32 = lSZ + Y + lHalfSizeH

                            If lMapX < 0 Then lMapX += Width
                            If lMapX > Width - 1 Then lMapX -= Width
                            If lMapY < 0 Then lMapY += Height
                            If lMapY > Height - 1 Then lMapY -= Height

                            lIdx = lMapY * Width + lMapX

                            If HeightMap(lIdx) > WaterHeight Then
                                Dim lTemp As Int32 = HeightMap(lIdx)
                                lTemp += CInt(Rnd() * 20) + 100
                                If lTemp > 255 Then lTemp = 255
                                If lTemp < 0 Then lTemp = 0

                                HeightMap(lIdx) = CByte(lTemp)
                            End If
                        Next Y
                    Next X
                End If
            End While

        End Sub

        Private Sub DoLunar()
            Dim lPass As Int32
            Dim lPassMax As Int32
            Dim lStartX As Int32
            Dim lStartY As Int32
            Dim lSubX As Int32
            Dim lSubY As Int32
            Dim lSubXTmp As Int32
            Dim lSubYTmp As Int32
            Dim lSubYMax As Int32

            Dim lRadius As Int32
            Dim lImpactPower As Int32
            Dim lCraterBasinHeight As Int32
            Dim lDist As Int32

            'alrighty, let's do some crazy stuff...
            lPassMax = CInt(Rnd() * 10) + 20      'lpass is the number
            For lPass = 0 To lPassMax
                'ok, get ground-zero
                lStartX = CInt(Int(Rnd() * (Width - 1)))
                lStartY = CInt(Int(Rnd() * (Height - 1)))
                'and our radius
                lRadius = CInt(Int(Rnd() * 32) + 1)
                lImpactPower = CInt(Int(Rnd() * 10) + 1) 'impactpower
                lCraterBasinHeight = HeightMap(lStartY * Width + lStartX) - 60       'crater basin height
                If lCraterBasinHeight < 0 Then lCraterBasinHeight = 0

                lSubYMax = (lStartY + (lRadius + lImpactPower))
                If lSubYMax > Height - 1 Then lSubYMax = Height - 1

                For lSubY = (lStartY - (lRadius + lImpactPower)) To lSubYMax
                    For lSubX = (lStartX - (lRadius + lImpactPower)) To (lStartX + (lRadius + lImpactPower))
                        lSubXTmp = lSubX
                        lSubYTmp = lSubY

                        If lSubYTmp < 0 Then lSubYTmp = 0
                        If lSubXTmp > Width - 1 Then lSubXTmp -= Width
                        If lSubXTmp < 0 Then lSubXTmp += Width

                        lDist = CInt(Math.Floor(Distance(lStartX, lStartY, lSubX, lSubY)))
                        Dim lTemp As Int32 = HeightMap(lSubYTmp * Width + lSubXTmp)

                        If lDist < lRadius Then
                            lTemp = CInt(lCraterBasinHeight + Int(Rnd() * 5) + 1)
                        ElseIf lDist = lRadius Then
                            lTemp = CInt(lCraterBasinHeight + 128)
                        ElseIf lDist - lRadius < lImpactPower Then
                            lTemp = CInt((lCraterBasinHeight + (12 * lImpactPower)) - (((lDist - lRadius) / lImpactPower) * lCraterBasinHeight))
                        End If
                        If lTemp < 0 Then lTemp = 0
                        If lTemp > 255 Then lTemp = 255
                        HeightMap(lSubYTmp * Width + lSubXTmp) = CByte(lTemp)
                    Next lSubX
                Next lSubY
            Next lPass
        End Sub

        Private Function Distance(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Single
            Dim dX As Single
            Dim dY As Single

            dX = lX2 - lX1
            dY = lY2 - lY1
            dX = dX * dX
            dY = dY * dY
            Distance = CSng(Math.Sqrt(dX + dY))
        End Function

        Private Function GetWaterPerc(ByVal lType As Int32) As Int32
            Select Case lType
                Case PlanetType.eBarren
                    Return 0
                Case PlanetType.eDesert
                    Return CInt(Int(Rnd() * 15) + 1)
                Case PlanetType.eGeoPlastic
                    Return CInt(Int(Rnd() * 40) + 15)
                Case PlanetType.eWaterWorld
                    Return CInt(80 + (Int(Rnd() * 20) + 1))
                Case PlanetType.eAdaptable
                    Return CInt(5 + CInt(Rnd() * 40))
                Case PlanetType.eTundra
                    Return 20 + CInt(Rnd() * 40)
                Case PlanetType.eTerran
                    Return 30 + CInt(Rnd() * 20)
                Case Else 'PlanetType.eTerran, PlanetType.eAcidic
                    Return CInt(30 + (Int(Rnd() * 50) + 1))
            End Select
        End Function

        Private Function SmoothTerrainVal(ByVal X As Int32, ByVal Y As Int32, ByVal lVal As Int32) As Byte
            Dim fCorners As Single
            Dim fSides As Single
            Dim fCenter As Single
            Dim fTotal As Single

            Dim lBackX As Int32
            Dim lBackY As Int32
            Dim lForeX As Int32
            Dim lForeY As Int32

            If X = 0 Then
                lBackX = Width - 1
            Else : lBackX = X - 1
            End If
            If X = Width - 1 Then
                lForeX = 0
            Else : lForeX = X + 1
            End If
            If Y = 0 Then
                lBackY = 0
            Else : lBackY = Y - 1
            End If
            If Y = Height - 1 Then
                lForeY = Height - 2
            Else : lForeY = Y + 1
            End If

            fCorners = 0
            fCorners = fCorners + HeightMap((lBackY * Width) + lBackX) 'muTiles(lBackX, lBackY).lHeight
            fCorners = fCorners + HeightMap((lBackY * Width) + lForeX) 'muTiles(lForeX, lBackY).lHeight
            fCorners = fCorners + HeightMap((lForeY * Width) + lBackX) 'muTiles(lBackX, lForeY).lHeight
            fCorners = fCorners + HeightMap((lForeY * Width) + lForeX) 'muTiles(lForeX, lForeY).lHeight
            fCorners = fCorners / 16

            fSides = 0
            fSides = fSides + HeightMap((Y * Width) + lBackX) 'muTiles(lBackX, Y).lHeight
            fSides = fSides + HeightMap((Y * Width) + lForeX) 'muTiles(lForeX, Y).lHeight
            fSides = fSides + HeightMap((lBackY * Width) + X) 'muTiles(X, lBackY).lHeight
            fSides = fSides + HeightMap((lForeY * Width) + X) 'muTiles(X, lForeY).lHeight
            fSides = fSides / 8

            fCenter = lVal / 4.0F

            fTotal = fCorners + fSides + fCenter
            If fTotal < 0 Then fTotal = 0
            If fTotal > 255 Then fTotal = 255

            Return CByte(fTotal)
        End Function

        Public Sub PopulateData()
            'this sub takes the place of everything that I removed from Client, we need to do everything to get it ready...
            GenerateTerrain(mlSeed)
            ComputeNormals()
        End Sub

        'NOTE: DO NOT CALL THIS NORMALLY, THIS ROUTINE IS FOR SPEED MOD CALCULATIONS!!!
        'NOTE: Called from Planet.TerrainGradeBuildable
        'NOTE: Called from Planet.SetupStartLocGrid
        Public Function GetTerrainNormalEx(ByVal fVertX As Single, ByVal fVertY As Single) As Vector3
            Dim fTX As Single   'translated X 
            Dim fTZ As Single   'translated Z 
            Dim lCol As Int32
            Dim lRow As Int32

            Dim dZ As Single
            Dim dX As Single

            Dim vN1 As vector3
            Dim vN2 As Vector3
            Dim vN3 As Vector3
            Dim vN4 As Vector3

            Dim lIdx As Int32

            fTX = fVertX
            fTZ = fVertY

            lCol = CInt(Math.Floor(fTX))
            lRow = CInt(Math.Floor(fTZ))
            dX = fTX - lCol
            dZ = fTZ - lRow

            lIdx = (lRow * Width) + lCol
            If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN1 = Vector3.Empty Else vN1 = HMNormals(lIdx)
            lIdx = (lRow * Width) + lCol + 1
            If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN2 = Vector3.Empty Else vN2 = HMNormals(lIdx)
            lIdx = ((lRow + 1) * Width) + lCol
            If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN3 = Vector3.Empty Else vN3 = HMNormals(lIdx)
            lIdx = ((lRow + 1) * Width) + lCol + 1
            If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN4 = Vector3.Empty Else vN4 = HMNormals(lIdx)

            vN1.Multiply(1 - dX)
            vN2.Multiply(dX)
            Dim vA As Vector3 = Vector3.Add(vN1, vN2)
            vN3.Multiply(1 - dX)
            vN4.Multiply(dX)
            Dim vB As Vector3 = Vector3.Add(vN3, vN4)
            vA.Multiply(1 - dZ)
            vB.Multiply(dZ)
            Dim vF As Vector3 = Vector3.Add(vA, vB)
            vF.Normalize()

            Return vF
        End Function

        Private Sub ComputeNormals()
            Dim Z As Int32
            Dim X As Int32

            Dim vecX As Vector3 = New Vector3()
            Dim vecZ As Vector3 = New Vector3()
            Dim vecN As Vector3

            Dim lHeight As Int32
            Dim oVerts() As Vector3

            ReDim HMNormals(VertsTotal - 1)
            ReDim oVerts(VertsTotal - 1)

            For Z = 0 To Height - 1
                For X = 0 To Width - 1
                    lHeight = HeightMap((Z * Height) + X)

                    oVerts((Z * Height) + X) = New Vector3((X - (mlHalfWidth)) * CellSpacing, lHeight * ml_Y_Mult, (Z - (mlHalfHeight)) * CellSpacing)
                Next X
            Next Z

            For Z = 1 To QuadsZ - 1
                For X = 1 To QuadsX - 1
                    vecX = Vector3.Subtract(oVerts(Z * Width + X + 1), oVerts(Z * Width + X - 1))
                    vecZ = Vector3.Subtract(oVerts((Z + 1) * Width + X), oVerts((Z - 1) * Width + X))
                    vecN = Vector3.CrossProduct(vecZ, vecX)
                    vecN.Normalize()

                    HMNormals(Z * Width + X) = vecN
                Next X
            Next Z

            vecX = Nothing
            vecZ = Nothing
            vecN = Nothing

            Erase oVerts

        End Sub
    End Class

#End Region



    Private Shared Function GetOrAddHubHub(ByRef bNew As Boolean, ByVal uNeighborhoodPt As Point3, ByRef oRandom As Random) As SolarSystem
        bNew = False
        Dim bNeedToSpawn As Boolean = False
        If moCurrHHSys Is Nothing Then
            bNeedToSpawn = True
        Else
            If moCurrHHSys.mlWormholeUB > 2 Then bNeedToSpawn = True
        End If
        If bNeedToSpawn = True Then
            Dim oPrev As SolarSystem = moCurrHHSys
            moCurrHHSys = CreateNewSystem(elSystemType.HubHubSystem, uNeighborhoodPt)

            'if a new hh is made, link the new hh to the previous hh via wormhole
            If oPrev Is Nothing = False Then AddWormhole(oPrev, moCurrHHSys, oRandom)

            'if a new hh is made, set bNew = true else, set to false
            bNew = True
        End If

        moCurrHHSys.lHubHubLinks += 1
        Return moCurrHHSys
    End Function
    Private Shared Function Distance(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Double
        Dim dX As Double = lX1 - lX2
        Dim dY As Double = lY1 - lY2
        dX *= dX
        dY *= dY
        Return Math.Sqrt(dX + dY)
    End Function
    Private Shared Sub AddWormhole(ByRef oSys1 As SolarSystem, ByRef oSys2 As SolarSystem, ByRef oRandom As Random)
        Dim oWormhole As New Wormhole()
        With oWormhole

            .ObjectID = -1
            .ObjTypeID = ObjectType.eWormhole
            .StartCycle = 0
            .System1 = oSys1
            .System2 = oSys2

            If .System1.SystemType = elSystemType.SpawnSystem OrElse .System1.SystemType = elSystemType.RespawnSystem OrElse .System1.SystemType = elSystemType.HubSystem Then
                .WormholeFlags = elWormholeFlag.eSystem1Jumpable Or elWormholeFlag.eSystem2Jumpable Or elWormholeFlag.eSystem1Detectable
            Else
                .WormholeFlags = 0
            End If


            Dim lCenterRadius1 As Int32 = GetCenterRadius(oSys1.StarType1ID, oSys1.StarType2ID, oSys1.StarType3ID)
            Dim lCenterRadius2 As Int32 = GetCenterRadius(oSys2.StarType1ID, oSys2.StarType2ID, oSys2.StarType3ID)

            'Ok, let's place them...
            Dim lCnt As Int32 = 0
            Dim lX As Int32 = oRandom.Next(-4500000, 4500000) ' CInt(Rnd() * 9000000) - 4500000
            Dim lZ As Int32 = oRandom.Next(-4500000, 4500000) 'CInt(Rnd() * 9000000) - 4500000
            While ((Math.Abs(lX) < lCenterRadius1 AndAlso Math.Abs(lZ) < lCenterRadius1) OrElse Distance(0, 0, lX, lZ) > 4500000) AndAlso lCnt < 100
                lX = oRandom.Next(-4500000, 4500000) 'CInt(Rnd() * 9000000) - 4500000
                lZ = oRandom.Next(-4500000, 4500000) 'CInt(Rnd() * 9000000) - 4500000
                lCnt += 1
            End While
            .LocX1 = lX
            .LocY1 = lZ

            lX = oRandom.Next(-4500000, 4500000) 'CInt(Rnd() * 9000000) - 4500000
            lZ = oRandom.Next(-4500000, 4500000) 'CInt(Rnd() * 9000000) - 4500000
            lCnt = 0
            While ((Math.Abs(lX) < lCenterRadius2 AndAlso Math.Abs(lZ) < lCenterRadius2) OrElse Distance(0, 0, lX, lZ) > 4500000) AndAlso lCnt < 100
                lX = oRandom.Next(-4500000, 4500000) 'CInt(Rnd() * 9000000) - 4500000
                lZ = oRandom.Next(-4500000, 4500000) 'CInt(Rnd() * 9000000) - 4500000
                lCnt += 1
            End While
            .LocX2 = lX
            .LocY2 = lZ

            'Dim lCnt As Int32 = oSys1.mlPlanetUB

            'Dim lVal1 As Int32 = CInt(Rnd() * lCnt)
            'Dim lVal2 As Int32 = CInt(Rnd() * lCnt)
            'While lVal2 = lVal1 AndAlso lCnt > 0
            '    lVal2 = CInt(Rnd() * lCnt)
            'End While

            'Dim oP1 As Planet = oSys1.GetPlanet(lVal1)
            'Dim oP2 As Planet = oSys1.GetPlanet(lVal2)
            'If oP1 Is Nothing OrElse oP2 Is Nothing Then Return

            'Dim fDiffX As Single = oP1.LocX - oP2.LocX
            'Dim fDiffY As Single = oP1.LocZ - oP2.LocZ
            'fDiffX *= fDiffX
            'fDiffY *= fDiffY
            'Dim fDist As Single = CSng(Math.Sqrt(fDiffX + fDiffY))
            'fDist *= CInt((Rnd() * 0.8F) + 0.1F)

            'fDiffX = fDist
            'fDiffY = 0
            'Dim fAngle As Single = LineAngleDegrees(oP1.LocX, oP1.LocZ, oP2.LocX, oP2.LocZ)
            'RotatePoint(0, 0, fDiffX, fDiffY, fAngle)
            '.LocX1 = CInt(fDiffX)
            '.LocY1 = CInt(fDiffY)

            ''=== new planet ===
            'lCnt = oSys2.mlPlanetUB
            'lVal1 = CInt(Rnd() * lCnt)
            'lVal2 = CInt(Rnd() * lCnt)
            'While lVal2 = lVal1 AndAlso lCnt > 0
            '    lVal2 = CInt(Rnd() * lCnt)
            'End While
            'oP1 = oSys2.GetPlanet(lVal1)
            'oP2 = oSys2.GetPlanet(lVal2)

            'fDiffX = oP1.LocX - oP2.LocX
            'fDiffY = oP1.LocZ - oP2.LocZ
            'fDiffX *= fDiffX
            'fDiffY *= fDiffY
            'fDist = CSng(Math.Sqrt(fDiffX + fDiffY))
            'fDist *= CInt((Rnd() * 0.8F) + 0.1F)

            'fDiffX = fDist
            'fDiffY = 0
            'fAngle = LineAngleDegrees(oP1.LocX, oP1.LocZ, oP2.LocX, oP2.LocZ)
            'RotatePoint(0, 0, fDiffX, fDiffY, fAngle)
            '.LocX2 = CInt(fDiffX)
            '.LocY2 = CInt(fDiffY)

            If .SaveObject() = False Then Return
        End With
        glWormholeUB += 1
        ReDim Preserve goWormhole(glWormholeUB)
        goWormhole(glWormholeUB) = oWormhole
    End Sub

    Private Structure Point3
        Public X As Int32
        Public Y As Int32
        Public Z As Int32
    End Structure
    Private Shared Function GetGalacticLoc() As Point3

        Dim lIDCnt As Int32 = 0

        If mfSystemSpawnRadius < 10 Then
            mfSystemSpawnRadiusChange = 0.1F
            mfSystemSpawnRadius = 10.0F
            Dim lTmpCnt As Int32 = 0
            For X As Int32 = 0 To glSystemUB
                If glSystemIdx(X) > -1 Then
                    If goSystem(X).SystemType = elSystemType.HubSystem Then
                        lTmpCnt += 1
                        If lTmpCnt = 2 Then
                            lIDCnt += 1
                            lTmpCnt = 0

                            If lIDCnt > 5 Then
                                mfSystemSpawnRadius += mfSystemSpawnRadiusChange
                                mfSystemSpawnRadiusChange += 0.2F
                                If mfSystemSpawnRadiusChange > 1.0F Then mfSystemSpawnRadiusChange = 0.1F
                            Else
                                mfSystemSpawnRadius += mfSystemSpawnRadiusChange
                                mfSystemSpawnRadiusChange += 0.001F
                                If mfSystemSpawnRadiusChange > 0.3F Then mfSystemSpawnRadiusChange = 0.1F
                            End If

                            
                        End If
                    End If
                End If
            Next X
        End If

        mfSystemSpawnRadius += mfSystemSpawnRadiusChange
        mfSystemSpawnRadiusChange += 0.2F
        If mfSystemSpawnRadiusChange > 1.0F Then mfSystemSpawnRadiusChange = 0.1F

        Randomize()

        Dim fX As Single = mfSystemSpawnRadius * 10
        Dim fZ As Single = 0

        RotatePoint(0, 0, fX, fZ, Rnd() * 360)
        Dim lLocX As Int32 = CInt(fX)
        Dim lLocZ As Int32 = CInt(fZ)

        Dim fTemp As Single = Math.Min(1, Math.Max(1, Math.Abs(fX) + Math.Abs(fZ)) / 500.0F) '1800.0F)      'MSC - 12/19/08 - changed to 500 to remove the height of the galaxy some
        fTemp = 600 - (550 * fTemp)
        fTemp = (Rnd() * fTemp) - (fTemp * 0.5F)
        Dim lLocY As Int32 = CInt(fTemp)

        Dim uPt3 As Point3
        With uPt3
            .X = lLocX
            .Y = lLocY
            .Z = lLocZ
        End With
        Return uPt3
    End Function

    Private Shared Function GetHeatIntensityScore(ByVal lST1 As Int32, ByVal lST2 As Int32, ByVal lST3 As Int32) As Int32
        Dim lTemp As Int32
        Dim lResult As Int32 = 0

        For X As Int32 = 0 To 2
            If X = 0 Then
                lTemp = lST1
            ElseIf X = 1 Then
                lTemp = lST2
            Else : lTemp = lST3
            End If

            Select Case lTemp
                Case 1
                    lResult += 45
                Case 2
                    lResult += 55
                Case 3
                    lResult += 60
                Case 4
                    lResult += 24
                Case 5
                    lResult += 32
                Case 6
                    lResult += 40
                Case 7
                    lResult += 12
                Case 8
                    lResult += 16
                Case 9
                    lResult += 20
                Case 10
                    lResult += 5
                Case 11
                    lResult += 7
                Case 12
                    lResult += 10
                Case 13
                    lResult += 1
            End Select
        Next X

        Return lResult

    End Function
    Private Shared Function GetCenterRadius(ByVal lST1 As Int32, ByVal lST2 As Int32, ByVal lST3 As Int32) As Int32
        Dim lResult As Int32 = 0
        If lST1 > 0 Then
            For X As Int32 = 0 To glStarTypeUB 'mlStarTypeUB
                If glStarTypeIdx(X) = lST1 Then
                    lResult += goStarType(X).StarRadius
                End If
            Next X
        End If
        If lST2 > 0 Then
            For X As Int32 = 0 To glStarTypeUB
                If glStarTypeIdx(X) = lST2 Then
                    lResult += goStarType(X).StarRadius
                End If
            Next X
        End If
        If lST3 > 0 Then
            For X As Int32 = 0 To glStarTypeUB
                If glStarTypeIdx(X) = lST3 Then
                    lResult += goStarType(X).StarRadius
                End If
            Next X
        End If
        Return (lResult * 2)
    End Function
    Private Shared Function GetRomanNumeral(ByVal lVal As Int32) As String
        Select Case lVal
            Case 1
                Return "I"
            Case 2
                Return "II"
            Case 3
                Return "III"
            Case 4
                Return "IV"
            Case 5
                Return "V"
            Case 6
                Return "VI"
            Case 7
                Return "VII"
            Case 8
                Return "VIII"
            Case 9
                Return "IX"
            Case 10
                Return "X"
            Case 11
                Return "XI"
            Case 12
                Return "XII"
            Case 13
                Return "XIII"
            Case 14
                Return "XIV"
            Case 15
                Return "XV"
            Case 16
                Return "XVI"
            Case 17
                Return "XVII"
            Case 18
                Return "XVIII"
            Case 19
                Return "XIX"
            Case 20
                Return "XX"
            Case 21
                Return "XXI"
            Case 22
                Return "XXII"
            Case 23
                Return "XXIII"
            Case Else
                Stop
                Return ""
        End Select
    End Function



End Class

