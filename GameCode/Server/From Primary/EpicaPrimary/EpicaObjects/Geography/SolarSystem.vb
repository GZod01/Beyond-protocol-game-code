Public Class SolarSystem
    Inherits Epica_GUID

    Public Enum elSystemType As Int32
        SpawnSystem = 0
        HubSystem = 1
        RespawnSystem = 2
        HubHubSystem = 3
        UnlockedSystem = 4
        TutorialSystem = 255
    End Enum

    Public SystemName(19) As Byte
    Public ParentGalaxy As Galaxy
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32

    Public StarType1ID As Byte
    Public StarType2ID As Byte
	Public StarType3ID As Byte

	Public OwnerID As Int32

    Public FleetJumpPointX As Int32
    Public FleetJumpPointZ As Int32

    'System Type is not transmitted or saved... just loaded
    Public SystemType As Byte       '0 = System A (StartLoc/SpawnCapable), 1 = System B (Hub System/Unspawnable), 2 = System C (Respawn system for dead players)

    Private mlPlanetIdx() As Int32  'the index of the planet in the goPlanets array (NOT THE ID!!!)
    Public mlPlanetUB As Int32 = -1

    Public oDomain As DomainServer

    Public moWormholes() As Wormhole        'pointers to wormholes
    Public mlWormholeUB As Int32 = -1

    Public lUnitCount As Int32 = 0
    Public lFacilityCount As Int32 = 0

    Private mySendString() As Byte
	'Public mbDetailsReady As Boolean = False
	'Private myDetailsString() As Byte
    Public Overloads Sub DataChanged()    'call this sub when changing properties of this object
        mbStringReady = False
		'mbDetailsReady = False
    End Sub

    Public Function GetObjAsString() As Byte()
        'here we will return the object populated as a string
        If mbStringReady = False Then
            'Guid - 6
            'name - 20
            'parentid - 4
            'locx - 4
            'locy - 4
            'locz - 4
            'startype1 - 1
            'startype2 - 1
            'startype3 - 1
            ReDim mySendString(53)  '0 to 44 = 45 bytes

            GetGUIDAsString.CopyTo(mySendString, 0)
            SystemName.CopyTo(mySendString, 6)
            System.BitConverter.GetBytes(ParentGalaxy.ObjectID).CopyTo(mySendString, 26)
            System.BitConverter.GetBytes(LocX).CopyTo(mySendString, 30)
            System.BitConverter.GetBytes(LocY).CopyTo(mySendString, 34)
            System.BitConverter.GetBytes(LocZ).CopyTo(mySendString, 38)
            mySendString(42) = StarType1ID
            mySendString(43) = StarType2ID
            mySendString(44) = StarType3ID
            mySendString(45) = SystemType
            System.BitConverter.GetBytes(FleetJumpPointX).CopyTo(mySendString, 46)
            System.BitConverter.GetBytes(FleetJumpPointZ).CopyTo(mySendString, 50)
            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
		If Me.InMyDomain = False Then Return True

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblSolarSystem (SystemName, GalaxyID, LocX, LocY, LocZ, StarType1ID, StarType2ID, " & _
                "StarType3ID, OwnerID) VALUES " & "('" & MakeDBStr(BytesToString(SystemName)) & "', " & ParentGalaxy.ObjectID & ", " & _
                LocX & ", " & LocY & ", " & LocZ & ", " & StarType1ID & ", " & StarType2ID & ", " & StarType3ID & ", " & OwnerID & ", " & FleetJumpPointX & ", " & FleetJumpPointZ & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblSolarSystem SET SystemName = '" & MakeDBStr(BytesToString(SystemName)) & "', GalaxyID = " & _
                 ParentGalaxy.ObjectID & ", LocX = " & LocX & ", LocY = " & LocY & ", LocZ = " & LocZ & _
                 ", StarType1ID = " & StarType1ID & ", StarType2ID = " & StarType2ID & ", StarType3ID = " & StarType3ID & _
                 ", OwnerID = " & OwnerID & _
                 ", FleetJumpPointX = " & FleetJumpPointX & ", FleetJumpPointZ=" & FleetJumpPointZ & _
                 " WHERE SystemID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(SystemId) FROM tblSolarSystem WHERE SystemName = '" & MakeDBStr(BytesToString(SystemName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Public Function GetSystemDetailsResponse() As Byte()
		Dim yMsg(9 + ((mlPlanetUB + 1) * Planet.PLANET_OBJ_MSG_LEN)) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eResponseSystemDetails).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(mlPlanetUB + 1).CopyTo(yMsg, lPos) : lPos += 4
		For X As Int32 = 0 To mlPlanetUB
			If glPlanetIdx(mlPlanetIdx(X)) <> -1 Then
				goPlanet(mlPlanetIdx(X)).GetObjAsString.CopyTo(yMsg, lPos) : lPos += Planet.PLANET_OBJ_MSG_LEN
			End If
		Next X
		Return yMsg
	End Function

	'Public Function GetDetailsString() As Byte()
	'    'ok this is the details string containing the planet objects
	'    Dim X As Int32
	'    Dim lPos As Int32 = 0
	'    If mbDetailsReady = False Then
	'        Dim lMsgLen As Int32 = ((mlPlanetUB + 1) * Planet.PLANET_OBJ_MSG_LEN) - 1
	'        ReDim myDetailsString(lMsgLen)

	'        For X = 0 To mlPlanetUB
	'            If glPlanetIdx(mlPlanetIdx(X)) <> -1 Then
	'                goPlanet(mlPlanetIdx(X)).GetObjAsString.CopyTo(myDetailsString, lPos)
	'                lPos += Planet.PLANET_OBJ_MSG_LEN
	'            End If
	'        Next X

	'        mbDetailsReady = True
	'    Else
	'        'update the message's planet rotate angles
	'        lPos = Planet.PLANET_ROTATE_ANGLE_MSG_LOC
	'        For X = 0 To mlPlanetUB
	'            System.BitConverter.GetBytes(goPlanet(mlPlanetIdx(X)).GetRotateAngle).CopyTo(myDetailsString, lPos)
	'            lPos += Planet.PLANET_OBJ_MSG_LEN
	'        Next X
	'    End If

	'    Return myDetailsString
	'End Function

    Public Function GetDetailsStringLen() As Int32
        Return ((mlPlanetUB + 1) * Planet.PLANET_OBJ_MSG_LEN) - 1
    End Function

    Public Sub AddPlanetIndex(ByVal lPlanetIndex As Int32)
        mlPlanetUB += 1
        ReDim Preserve mlPlanetIdx(mlPlanetUB)
        mlPlanetIdx(mlPlanetUB) = lPlanetIndex
    End Sub

    Public Function StationProximityTest(ByVal lMinDist As Int32, ByVal lLocX As Int32, ByVal lLocZ As Int32) As Boolean
        For X As Int32 = 0 To mlPlanetUB
            If mlPlanetIdx(X) <> -1 AndAlso glPlanetIdx(mlPlanetIdx(X)) <> -1 Then
                With goPlanet(mlPlanetIdx(X))
                    If Distance(.LocX, .LocZ, lLocX, lLocZ) < lMinDist Then Return False
                End With
            End If
        Next X
        Return True
    End Function

    Public Sub AddWormholeReference(ByRef oWormhole As Wormhole)
        For X As Int32 = 0 To mlWormholeUB
            If moWormholes(X) Is Nothing = False AndAlso moWormholes(X).ObjectID = oWormhole.ObjectID Then
                Return
            End If
        Next X

        mlWormholeUB += 1
        ReDim Preserve moWormholes(mlWormholeUB)
        moWormholes(mlWormholeUB) = oWormhole
    End Sub

    Public Sub MarkChildrenInMyDomain(ByVal bValue As Boolean)
        For X As Int32 = 0 To mlPlanetUB
            If mlPlanetIdx(X) <> -1 AndAlso glPlanetIdx(mlPlanetIdx(X)) <> -1 Then goPlanet(mlPlanetIdx(X)).InMyDomain = bValue
        Next X
        For X As Int32 = 0 To mlWormholeUB
            If moWormholes(X) Is Nothing = False Then moWormholes(X).InMyDomain = bValue
        Next X
	End Sub

	Public Function GetNearestPlanet(ByVal lX As Int32, ByVal lZ As Int32) As Planet
		Dim oMin As Planet = Nothing
		Dim lMinDist As Int32 = Int32.MaxValue

		For X As Int32 = 0 To mlPlanetUB
			If mlPlanetIdx(X) <> -1 Then
				With goPlanet(mlPlanetIdx(X))
					Dim lDist As Int32 = Math.Abs(.LocX - lX)
					lDist += Math.Abs(.LocZ - lZ)
					If lDist < lMinDist Then
						lMinDist = lDist
						oMin = goPlanet(mlPlanetIdx(X))
					End If
				End With
			End If
		Next X

		Return oMin
    End Function

    Public Function GetNearestRingedPlanet(ByVal lX As Int32, ByVal lZ As Int32, ByRef lFinalDist As Int32) As Planet
        Dim oMin As Planet = Nothing
        Dim lMinDist As Int32 = Int32.MaxValue
        Try
            For X As Int32 = 0 To mlPlanetUB
                If mlPlanetIdx(X) <> -1 Then
                    With goPlanet(mlPlanetIdx(X))
                        If .RingMineralID > 0 Then
                            Dim fDX As Single = lX - .LocX
                            Dim fDZ As Single = lZ - .LocZ
                            Dim fDist As Single = CSng(Math.Sqrt((fDX * fDX) + (fDZ * fDZ)))
                            If fDist < lMinDist Then
                                lMinDist = CInt(fDist)
                                oMin = goPlanet(mlPlanetIdx(X))
                            End If
                        End If
                    End With
                End If
            Next X
        Catch
        End Try

        lFinalDist = lMinDist
        Return oMin
    End Function

	Public Function GetPlanetIdx(ByVal lIdx As Int32) As Int32
		If lIdx > -1 AndAlso lIdx <= mlPlanetUB Then Return mlPlanetIdx(lIdx) Else Return -1
    End Function

    Private moNPCTradepost As Facility = Nothing
    Public lNPCTradepostID As Int32 = -1
    Public ReadOnly Property NPCTradepost() As Facility
        Get
            If lNPCTradepostID = -1 Then
                Dim lTradepostIdx As Int32 = AureliusAI.SpawnNewFacility(84, eiBehaviorPatterns.eEngagement_Hold_Fire, 0, 0, 0, Me, False)
                If lTradepostIdx > -1 Then moNPCTradepost = goFacility(lTradepostIdx)
            End If
            Return moNPCTradepost
        End Get
    End Property

    Public Sub CheckUnlockSystemWormholes()
        Try
            If Me.SystemType = elSystemType.HubSystem Then
                'From a HUB to another HUB requires that the other HUB have been colonized
                'From a HUB to the HUB HUB requires that all connected HUB's to this HUB have been colonized
                For X As Int32 = 0 To mlWormholeUB
                    Dim oWH As Wormhole = moWormholes(X)
                    If oWH Is Nothing = False Then
                        If oWH.System1 Is Nothing = False AndAlso oWH.System1.ObjectID = Me.ObjectID Then
                            If True = True Then '(oWH.WormholeFlags And elWormholeFlag.eSystem1Colonized) <> 0 Then
                                oWH.WormholeFlags = oWH.WormholeFlags Or elWormholeFlag.eSystem1Colonized
                                'ok, is the wormhole already detectable?
                                If (oWH.WormholeFlags And elWormholeFlag.eSystem1Detectable) = 0 Then
                                    'No, ok, go to system 2
                                    Dim oSys2 As SolarSystem = oWH.System2
                                    If oSys2 Is Nothing = False Then
                                        '  then go through the wormholes of system2 where the system2 of the wormhole is system2
                                        Dim bGood As Boolean = True
                                        For Y As Int32 = 0 To oSys2.mlWormholeUB
                                            Dim oS2WH As Wormhole = oSys2.moWormholes(Y)
                                            If oS2WH Is Nothing = False Then
                                                If oS2WH.System2 Is Nothing = False AndAlso oS2WH.System2.ObjectID = oSys2.ObjectID Then
                                                    '  is the system1 inhabited?
                                                    If (oS2WH.WormholeFlags And elWormholeFlag.eSystem1Colonized) = 0 Then
                                                        '    No: Not unlocked
                                                        bGood = False
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        Next Y

                                        'If we are here without a no, then we can activate this wormhole
                                        If bGood = True Then
                                            oWH.WormholeFlags = oWH.WormholeFlags Or elWormholeFlag.eSystem1Detectable Or elWormholeFlag.eSystem1Jumpable Or elWormholeFlag.eSystem2Jumpable
                                        End If
                                    End If

                                    oWH.SaveObject()
                                    Dim yMsg() As Byte = goMsgSys.GetAddObjectMessage(oWH, GlobalMessageCode.eAddObjectCommand)
                                    If oWH.System1 Is Nothing = False AndAlso oWH.System1.oDomain Is Nothing = False AndAlso oWH.System2.oDomain.DomainSocket Is Nothing = False Then
                                        oWH.System1.oDomain.DomainSocket.SendData(yMsg)
                                    End If
                                    If oWH.System2 Is Nothing = False AndAlso oWH.System2.oDomain Is Nothing = False AndAlso oWH.System2.oDomain.DomainSocket Is Nothing = False Then
                                        oWH.System2.oDomain.DomainSocket.SendData(yMsg)
                                    End If
                                Else
                                    oWH.SaveObject()
                                End If

                            End If

                        End If
                    End If
                Next X
            End If
        Catch
        End Try
    End Sub

End Class
