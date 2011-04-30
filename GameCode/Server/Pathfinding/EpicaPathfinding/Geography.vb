Option Strict On

'Copied straight out of RegionServer's code

Public Class StarType
    Public StarTypeID As Byte
    Public StarTypeAttrs As Int32       'bit-wise attributes
    Public StarRadius As Int32
    Public HeatIndex As Byte

	Public Shared Function GetStarTypeIdx(ByVal lStarTypeID As Int32) As Int32
		For X As Int32 = 0 To glStarTypeUB
			If goStarTypes(X) Is Nothing = False AndAlso goStarTypes(X).StarTypeID = lStarTypeID Then
				Return X
			End If
		Next X
		Return -1
	End Function
End Class

Public Class Galaxy
    Inherits Base_GUID

    Public GalaxyName As String

    Public moSystems() As SolarSystem
    Public mlSystemUB As Int32 = -1
    'TODO: Include Nebulae

    Public CurrentSystemIdx As Int32 = -1

    Public Sub AddSystem(ByVal oSystem As SolarSystem)
        mlSystemUB += 1
        ReDim Preserve moSystems(mlSystemUB)
        moSystems(mlSystemUB) = oSystem
	End Sub

	Public Function GetSystem(ByVal lId As Int32) As SolarSystem
		For X As Int32 = 0 To mlSystemUB
			If moSystems(X) Is Nothing = False AndAlso moSystems(X).ObjectID = lId Then
				Return moSystems(X)
			End If
		Next X
		Return Nothing
	End Function
End Class

Public Class SolarSystem
    Inherits Base_GUID

    Public SystemName As String
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32

    Public EnvirIdx As Int32 = -1       'index in the goEnvirs array

    Public StarType1Idx As Int32 = -1   'index in the StarType array
    Public StarType2Idx As Int32 = -1   'index in the StarType array
    Public StarType3Idx As Int32 = -1   'index in the StarType array

    Public SystemType As Byte = 0

    Public FleetJumpPointX As Int32
    Public FleetJumpPointZ As Int32

    Public moPlanets() As Planet
    Public PlanetUB As Int32 = -1

    Public Sub AddPlanet(ByVal oPlanet As Planet)
        PlanetUB += 1
        ReDim Preserve moPlanets(PlanetUB)
        moPlanets(PlanetUB) = oPlanet
    End Sub

End Class

Public Class Planet
    Inherits Base_GUID

    Public PlanetName As String
    Public MapTypeID As Byte        'the planet typeid
    Public PlanetSizeID As Byte     'used for determining map size, 0-tiny, 1-small, 2-medium, 3-large, 4-huge. Maybe able to remove and base off of Radius
    Public PlanetRadius As Int16    'might be able to remove sizeID and base the map size off of radius
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32
    Public Vegetation As Byte
    Public Atmosphere As Byte
    Public Hydrosphere As Byte
    Public Gravity As Byte
    Public SurfaceTemperature As Byte
    Public RotationDelay As Int16   'cycles between incrementing rotation angle

    Public ParentSystem As SolarSystem

    Private moTerrain As TerrainClass

    Public AxisAngle As Int32       'axis angle (yaw)
    Public RotateAngle As Int32     'rotation angle
 

    Public oPather As Pather
	Public lHalfExtent As Int32

	Private Shared moInstancePlanet As Planet = Nothing
	Public Shared Function GetInstancePlanet() As Planet
		If moInstancePlanet Is Nothing Then
			moInstancePlanet = New Planet(0, 0, PlanetType.eGeoPlastic)
			With moInstancePlanet
				.Atmosphere = 0
				.AxisAngle = 680
				.Gravity = 40
				.Hydrosphere = 0
				.LocX = 100000
				.LocY = 0
				.LocZ = 0
				.ObjTypeID = ObjectType.ePlanet
				.PlanetName = "Tutorial One"
				.PlanetRadius = 1000
				.RotateAngle = 0
				.RotationDelay = 10
				.SurfaceTemperature = 100
				.Vegetation = 0
				.PopulateInstanceData()
			End With
		End If
		Return moInstancePlanet
	End Function

	Public Sub New(ByVal lID As Int32, ByVal ySizeID As Byte, ByVal yMapTypeID As Byte)
		ObjectID = lID
		PlanetSizeID = ySizeID
		MapTypeID = yMapTypeID

		moTerrain = New TerrainClass(lID)
		If PlanetSizeID = 0 Then moTerrain.ml_Y_Mult = 9.0F
		moTerrain.MapType = yMapTypeID

		Select Case PlanetSizeID
			Case 0 : moTerrain.CellSpacing = gl_TINY_PLANET_CELL_SPACING
			Case 1 : moTerrain.CellSpacing = gl_SMALL_PLANET_CELL_SPACING
			Case 2 : moTerrain.CellSpacing = gl_MEDIUM_PLANET_CELL_SPACING
			Case 3 : moTerrain.CellSpacing = gl_LARGE_PLANET_CELL_SPACING  '550
			Case 4 : moTerrain.CellSpacing = gl_HUGE_PLANET_CELL_SPACING  '700
		End Select
	End Sub

	'NOTE: Changed Singles to Int32s
	Public Function GetHeightAtPoint(ByVal lX As Int32, ByVal lZ As Int32) As Int32
		Return moTerrain.GetHeightAtLocation(lX, lZ)
	End Function

	Public Function HasLineOfSight(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lZ1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32, ByVal lZ2 As Int32, ByVal lAttackerHeight As Int32) As Boolean
		Return moTerrain.HasLineOfSight(lX1, lY1, lZ1, lX2, lY2, lZ2, lAttackerHeight)
	End Function

	Public Sub PopulateData()
		'this sub takes the place of everything that I removed from Client, we need to do everything to get it ready...
		moTerrain.PopulateData()

		'Now... our data is populated... let's send that to the pather
		oPather = New Pather(moTerrain)

		lHalfExtent = (TerrainClass.Width \ 2) * moTerrain.CellSpacing
	End Sub

	Public Sub PopulateInstanceData()
		'this sub takes the place of everything that I removed from Client, we need to do everything to get it ready...
		moTerrain.PopulateInstanceData()

		'Now... our data is populated... let's send that to the pather
		oPather = New Pather(moTerrain)

		lHalfExtent = (TerrainClass.Width \ 2) * moTerrain.CellSpacing
	End Sub

    Public Function GetCellSpacing() As Int32
        Return moTerrain.CellSpacing
    End Function

    Public Function WaterRealHeight() As Single
        Return CSng(moTerrain.WaterHeight) * moTerrain.ml_Y_Mult
    End Function

    Public Sub SaveToFile()
        moTerrain.SaveToFile(Me.ObjectID)
    End Sub
End Class

Public Class TerrainClass

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

#Region "Private Variables"
    'Our heightmap array, we need to keep this for... getting heights at locations
    Public HeightMap() As Byte
    'Indicates if the Heightmap has been generated yet
    Private mbHMReady As Boolean = False
    'the Parent Planet ID... set in New()
    Private mlSeed As Int32
#End Region

    Public Sub New(ByVal lSeed As Int32)
        mlSeed = lSeed
    End Sub

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
                lIdx = CInt(fZ * 256 + fX)
                If HeightMap(lIdx) * ml_Y_Mult > fMaxHt Then
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
                lIdx = CInt(fZ * 256 + fX)
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
        Return CSng(Math.Sqrt(dX + dY))
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
    End Sub

    Public Sub SaveToFile(ByVal lPlanetID As Int32)
		Dim oFS As New IO.FileStream("C:\terrain_" & lPlanetID & ".txt", IO.FileMode.Create)
		Dim oWrite As New IO.StreamWriter(oFS)
		oWrite.WriteLine(Me.WaterHeight)

		For Y As Int32 = 0 To Height - 1
			Dim sLine As String = ""
			For X As Int32 = 0 To Width - 1
				If sLine <> "" Then sLine &= ","
				Dim lIdx As Int32 = Y * Width + X
				sLine &= HeightMap(lIdx).ToString
			Next X
			oWrite.WriteLine(sLine)
		Next Y
		oWrite.Close()
		oWrite.Dispose()
		oFS.Close()
		oFS.Dispose()
	End Sub

	Public Sub PopulateInstanceData()
		Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
		If sPath.EndsWith("\") = False Then sPath &= "\"
		Dim oFS As IO.FileStream = New IO.FileStream(sPath & "terr.txt", IO.FileMode.Open)
		Dim oRead As IO.StreamReader = New IO.StreamReader(oFS)

		WaterHeight = CByte(Val(oRead.ReadLine))
		ReDim HeightMap(Width * Height)
		For Y As Int32 = 0 To TerrainClass.Height - 1
			Dim sLine As String = oRead.ReadLine
			Dim sValues() As String = Split(sLine, ",")
			For X As Int32 = 0 To TerrainClass.Width - 1
				Dim lIdx As Int32 = (Y * TerrainClass.Width) + X
				HeightMap(lIdx) = CByte(Val(sValues(X)))
			Next X
		Next Y
		oRead.Close()
		oRead.Dispose()
		oFS.Close()
		oFS.Dispose()
		oRead = Nothing
		oFS = Nothing

		mbHMReady = True 
	End Sub
End Class

