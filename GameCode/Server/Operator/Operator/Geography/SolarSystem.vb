Public Class SolarSystem
    Inherits Epica_GUID

    Public SystemName(19) As Byte
    Public ParentGalaxy As Galaxy
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32

    Public StarType1ID As Byte
    Public StarType2ID As Byte
    Public StarType3ID As Byte

    Public SystemType As Byte       '0= normal (A), 1 = Hub (B), 2 = Respawn (C)

    Public FleetJumpPointX As Int32
    Public FleetJumpPointZ As Int32

    Private mlPlanetIdx() As Int32  'the index of the planet in the goPlanets array (NOT THE ID!!!)
    Public mlPlanetUB As Int32 = -1

    Public oPrimaryServer As ServerObject

    Public moWormholes() As Wormhole        'pointers to wormholes
    Public mlWormholeUB As Int32 = -1

	'For Performance Monitoring and Managing Spawns
	Public bulPhysicalMemoryFootprint As UInt64
	Public bulVirtualMemoryFootprint As UInt64
    Public fProcessorUsage As Single
    Public lEntityCount As Int32 = 0        'loaded at startup
	'==============================================

	Public lHubHubLinks As Int32 = 0			'indicates how many neighborhoods have linked to this hub hub

    Private mySendString() As Byte
    Public mbDetailsReady As Boolean = False
    Private myDetailsString() As Byte
    Public Overloads Sub DataChanged()    'call this sub when changing properties of this object
        mbStringReady = False
        mbDetailsReady = False
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
            ReDim mySendString(44)  '0 to 44 = 45 bytes

            GetGUIDAsString.CopyTo(mySendString, 0)
            SystemName.CopyTo(mySendString, 6)
            System.BitConverter.GetBytes(ParentGalaxy.ObjectID).CopyTo(mySendString, 26)
            System.BitConverter.GetBytes(LocX).CopyTo(mySendString, 30)
            System.BitConverter.GetBytes(LocY).CopyTo(mySendString, 34)
            System.BitConverter.GetBytes(LocZ).CopyTo(mySendString, 38)
            mySendString(42) = StarType1ID
            mySendString(43) = StarType2ID
            mySendString(44) = StarType3ID
            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblSolarSystem (SystemName, GalaxyID, LocX, LocY, LocZ, StarType1ID, StarType2ID, " & _
                "StarType3ID, SystemType, FleetJumpPointX, FleetJumpPointZ) VALUES " & "('" & MakeDBStr(BytesToString(SystemName)) & "', " & ParentGalaxy.ObjectID & ", " & _
                LocX & ", " & LocY & ", " & LocZ & ", " & StarType1ID & ", " & StarType2ID & ", " & StarType3ID & ", " & SystemType & ", " & FleetJumpPointX & ", " & FleetJumpPointZ & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblSolarSystem SET SystemName = '" & MakeDBStr(BytesToString(SystemName)) & "', GalaxyID = " & _
                 ParentGalaxy.ObjectID & ", LocX = " & LocX & ", LocY = " & LocY & ", LocZ = " & LocZ & _
                 ", StarType1ID = " & StarType1ID & ", StarType2ID = " & StarType2ID & ", StarType3ID = " & StarType3ID & _
                 ", SystemType = " & SystemType & _
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

    Public Sub AddPlanetIndex(ByVal lPlanetIndex As Int32)
        mlPlanetUB += 1
        ReDim Preserve mlPlanetIdx(mlPlanetUB)
        mlPlanetIdx(mlPlanetUB) = lPlanetIndex
    End Sub
 
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

	Public Function GetPlanet(ByVal lIndex As Int32) As Planet
		If lIndex > -1 AndAlso lIndex <= mlPlanetUB Then
			Dim lTmp As Int32 = mlPlanetIdx(lIndex)
			If lTmp > -1 AndAlso lTmp <= glPlanetUB Then Return goPlanet(lTmp)
		End If
		Return Nothing
	End Function
End Class
