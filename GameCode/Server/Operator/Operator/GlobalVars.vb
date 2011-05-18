Option Strict On

Public Enum eyOperatorState As Byte
	LonelyOperator = 0
	MainOperatorWithBackup = 1
	BackupOperator = 2
	EmergencyOperator = 3
End Enum

Module GlobalVars

	Public Const gl_KEEP_ALIVE_TIME As Int32 = 5000

    Public gb_IS_TEST_SERVER As Boolean = False

	Public gsExternalAddress As String = "beyondprotocol.net"

	Public goMsgSys As MsgSystem
	Public uSpawnRequests() As SpawnRequest
	Public lSpawnRequestUB As Int32 = -1

    Public glExpectedBoxCnt As Int32 = 1

	Public gyBackupOperator As eyOperatorState = eyOperatorState.LonelyOperator
	Public goOperatorSW As Stopwatch

    Public gsMainOperatorIP As String = "10.70.5.141"
    Public glMainOperatorPort As Int32 = 7701

	Public oNeighborhoods() As Neighborhood
	Public lNeighborhoodUB As Int32 = -1
	Public Sub PopulateNeighborhoods()
		For X As Int32 = 0 To glSystemUB
			If glSystemIdx(X) <> -1 AndAlso goSystem(X).SystemType = 1 Then
				lNeighborhoodUB += 1
				ReDim Preserve oNeighborhoods(lNeighborhoodUB)
				oNeighborhoods(lNeighborhoodUB) = New Neighborhood()
				With oNeighborhoods(lNeighborhoodUB)
					.lNeighborhoodID = goSystem(X).ObjectID
					.AddSystem(goSystem(X))
				End With
			End If
		Next X
	End Sub

#Region "  Box Operator ID Management  "
	Private Class BoxOperatorIDGenerator
		Private mlBoxOperatorID As Int32 = 0
		Public Function GetNextBoxOperatorID() As Int32
			SyncLock Me
				mlBoxOperatorID += 1
				Return mlBoxOperatorID
			End SyncLock
		End Function
	End Class
	Private oBoxOperatorIDGenerator As BoxOperatorIDGenerator
	Public Function GetNextBoxOperatorID() As Int32
		If oBoxOperatorIDGenerator Is Nothing = True Then oBoxOperatorIDGenerator = New BoxOperatorIDGenerator
		Return oBoxOperatorIDGenerator.GetNextBoxOperatorID()
	End Function
#End Region
#Region "  Object Arrays  "
	Public glGalaxyUB As Int32 = -1
	Public glGalaxyIdx() As Int32
	Public goGalaxy() As Galaxy
	Public glStarTypeUB As Int32 = -1
	Public goStarType() As StarType
	Public glStarTypeIdx() As Int32
	Public goNebula() As Nebula
	Public glNebulaUB As Int32 = -1
	Public glNebulaIdx() As Int32
	Public goSystem() As SolarSystem
	Public glSystemIdx() As Int32
	Public glSystemUB As Int32 = -1
	Public goPlanet() As Planet
	Public glPlanetIdx() As Int32
	Public glPlanetUB As Int32 = -1
	Public goPlayer() As Player
	Public glPlayerIdx() As Int32
	Public glPlayerUB As Int32 = -1
	Public goWormhole() As Wormhole
	Public glWormholeUB As Int32 = -1
	Public glWormholeIdx() As Int32
#End Region
#Region "  Get Object Arrays  "

	Public Function GetEpicaGalaxy(ByVal lObjectID As Int32) As Galaxy
		Dim X As Int32
		For X = 0 To glGalaxyUB
			If glGalaxyIdx(X) = lObjectID Then
				Return goGalaxy(X)
			End If
		Next X
		Return Nothing
	End Function

	Public Function GetEpicaPlanet(ByVal lObjectID As Int32) As Planet
		Dim X As Int32
		For X = 0 To glPlanetUB
			If glPlanetIdx(X) = lObjectID Then
				Return goPlanet(X)
			End If
		Next X
		Return Nothing
	End Function

	Public Function GetEpicaSystem(ByVal lObjectID As Int32) As SolarSystem
		Dim X As Int32
		For X = 0 To glSystemUB
			If glSystemIdx(X) = lObjectID Then
				Return goSystem(X)
			End If
		Next X
		Return Nothing
	End Function

	Public Function GetEpicaWormhole(ByVal lObjectID As Int32) As Wormhole
		Dim X As Int32
		For X = 0 To glWormholeUB
			If glWormholeIdx(X) = lObjectID Then
				Return goWormhole(X)
			End If
		Next X
		Return Nothing
	End Function

	Public Function GetEpicaNebula(ByVal lObjectID As Int32) As Nebula
		Dim X As Int32
		For X = 0 To glNebulaUB
			If glNebulaIdx(X) = lObjectID Then
				Return goNebula(X)
			End If
		Next X
		Return Nothing
	End Function

	Public Function GetEpicaPlayer(ByVal lObjectID As Int32) As Player
		Dim X As Int32
		For X = 0 To glPlayerUB
			If glPlayerIdx(X) = lObjectID Then
				Return goPlayer(X)
			End If
		Next X
		Return Nothing
	End Function

#End Region
#Region "  Event Logging  "
	Public gfrmDisplayForm As frmMain

	Private moFileStream As System.IO.FileStream
	Private moFileWrite As System.IO.StreamWriter
	Public Enum LogEventType
		CriticalError = 1
		PossibleCheat = 2
		Warning = 3
		Informational = 4
	End Enum

    Public Function InitializeEventLogging() As Boolean
        Dim oINI As New InitFile()
        Dim sLog As String

        sLog = oINI.GetString("SETTINGS", "EventLogFilePath", "")
        If sLog = "" Then
            sLog = AppDomain.CurrentDomain.BaseDirectory()
            If Right$(sLog, 1) <> "\" Then sLog = sLog & "\"
            'sLog = sLog & "Events.Log"
            Dim sNewFileName As String = "Events_" & Now.ToString("MMddyyyyHHmmss") & ".log"
            sLog &= sNewFileName
        End If
        moFileStream = System.IO.File.Open(sLog, IO.FileMode.Append)
        moFileWrite = New System.IO.StreamWriter(moFileStream)
        moFileWrite.AutoFlush = True
        Return True
    End Function

	Public Sub LogEvent(ByVal lEventType As LogEventType, ByVal sValue As String)

		'Public Shared Function StringToBytes(ByVal sData As String) As Byte()
		'    Return System.Text.ASCIIEncoding.ASCII.GetBytes(sData)
        'End Function
        sValue = Now.ToString & "|" & sValue
		Select Case lEventType
			Case LogEventType.CriticalError
				sValue = "CRITICAL: " & sValue
			Case LogEventType.PossibleCheat
				sValue = "POSSIBLE CHEAT: " & sValue
			Case LogEventType.Warning
                sValue = "WARNING: " & sValue 
        End Select

		gfrmDisplayForm.AddEventLine(sValue)

		If moFileStream Is Nothing = False Then
			If moFileStream.CanWrite() Then
				moFileWrite.WriteLine(sValue)
			End If
		End If
    End Sub

	Public Sub CloseEventLogger()
		moFileWrite.Close()
		moFileStream.Close()
	End Sub
#End Region
#Region "  DB Connection Stuff  "
    Private m_goCN As OleDb.OleDbConnection = Nothing
    Public ReadOnly Property goCN() As OleDb.OleDbConnection
        Get
            If m_goCN Is Nothing OrElse m_goCN.State = ConnectionState.Closed OrElse m_goCN.State = ConnectionState.Broken Then
                m_goCN = Nothing
                InitializeConnection()
            End If
            Return m_goCN
        End Get
    End Property

    Public Function InitializeConnection() As Boolean
        'go ahead and connect
        Dim sConnStr As String
        Dim oINI As New InitFile()
        Dim sUDL As String
        Dim bNoErrors As Boolean

        bNoErrors = True
        Try
            sUDL = oINI.GetString("SETTINGS", "Perm_Save_UDL", "")
            If sUDL <> "" Then
                sConnStr = "FILE NAME=" & AppDomain.CurrentDomain.BaseDirectory & sUDL
                'sConnStr = "FILE NAME=C:\test.udl"
                m_goCN = New OleDb.OleDbConnection(sConnStr)
                m_goCN.Open()
            End If
        Catch
            LogEvent(LogEventType.CriticalError, "InitializeConnection: " & Err.Description)
        Finally
            If bNoErrors = False Then
                m_goCN = Nothing
            End If
        End Try

        Return bNoErrors
    End Function

    'Private m_goWebCN As Odbc.OdbcConnection = Nothing
    'Public ReadOnly Property goWebCN() As Odbc.OdbcConnection
    '    Get
    '        If m_goWebCN Is Nothing OrElse m_goWebCN.State = ConnectionState.Closed OrElse m_goWebCN.State = ConnectionState.Broken Then
    '            m_goWebCN = Nothing
    '            InitializeWebConnection()
    '        End If
    '        Return m_goWebCN
    '    End Get
    'End Property

    'Public Function InitializeWebConnection() As Boolean
    '    'go ahead and connect
    '    Dim bNoErrors As Boolean

    '    bNoErrors = True
    '    Try
    '        Dim sConnStr As String = "DSN=WebDB"
    '        m_goWebCN = New Odbc.OdbcConnection(sConnStr)
    '        m_goWebCN.Open()
    '    Catch
    '        bNoErrors = False
    '        LogEvent(LogEventType.CriticalError, "InitializeConnection: " & Err.Description)
    '    Finally
    '        If bNoErrors = False Then
    '            m_goWebCN = Nothing
    '        End If
    '    End Try

    '    Return bNoErrors
    'End Function

    'Private m_goSuiteCN As Odbc.OdbcConnection = Nothing
    'Public ReadOnly Property goSuiteCN() As Odbc.OdbcConnection
    '    Get
    '        If m_goSuiteCN Is Nothing OrElse m_goSuiteCN.State = ConnectionState.Closed OrElse m_goSuiteCN.State = ConnectionState.Broken Then
    '            m_goSuiteCN = Nothing
    '            InitializeSuiteConnection()
    '        End If
    '        Return m_goSuiteCN
    '    End Get
    'End Property

    'Public Function InitializeSuiteConnection() As Boolean
    '    'go ahead and connect
    '    Dim bNoErrors As Boolean

    '    bNoErrors = True
    '    Try
    '        Dim sConnStr As String = "DSN=SuiteDB"
    '        m_goSuiteCN = New Odbc.OdbcConnection(sConnStr)
    '        m_goSuiteCN.Open()
    '    Catch
    '        bNoErrors = False
    '        LogEvent(LogEventType.CriticalError, "InitializeSuiteConnection: " & Err.Description)
    '    Finally
    '        If bNoErrors = False Then
    '            m_goSuiteCN = Nothing
    '        End If
    '    End Try

    '    Return bNoErrors
    'End Function

    'Public Sub CloseConn()
    '    If Not m_goWebCN Is Nothing Then
    '        m_goWebCN.Close()
    '    End If
    '    m_goWebCN = Nothing

    '    If Not m_goCN Is Nothing Then
    '        m_goCN.Close()
    '    End If
    '    m_goCN = Nothing

    '    If m_goSuiteCN Is Nothing = False Then m_goSuiteCN.Close()
    '    m_goSuiteCN = Nothing
    'End Sub
#End Region
#Region "  String Management  "
	Public Function BytesToString(ByVal yBytes() As Byte) As String
		Dim lLen As Int32
		Dim X As Int32

		lLen = yBytes.Length
		For X = 0 To yBytes.Length - 1
			If yBytes(X) = 0 Then
				lLen = X
				Exit For
			End If
		Next X
		Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yBytes), 1, lLen)
	End Function

	Public Function StringToBytes(ByVal sVal As String) As Byte()
		Return System.Text.ASCIIEncoding.ASCII.GetBytes(sVal)
	End Function

	Public Function MakeDBStr(ByVal sVal As String) As String
		Return Replace$(sVal, "'", "''")
	End Function

	Public Function GetStringFromBytes(ByRef yData() As Byte, ByVal lPos As Int32, ByVal lDataLen As Int32) As String
		Dim yTemp(lDataLen - 1) As Byte
		Array.Copy(yData, lPos, yTemp, 0, lDataLen)
		Dim lLen As Int32 = yTemp.Length
		For Y As Int32 = 0 To yTemp.Length - 1
			If yTemp(Y) = 0 Then
				lLen = Y
				Exit For
			End If
		Next Y
		Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)
	End Function
#End Region
#Region "  Database Loading  "
	Public Function LoadDataFromDB() As Boolean
		If LoadGeographical() = False Then Return False
        If LoadPlayers() = False Then Return False
        If LoadSenate() = False Then Return False
		Return True
	End Function

	Private Function LoadGeographical() As Boolean
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim oComm As OleDb.OleDbCommand
		Dim sSQL As String
		Dim bResult As Boolean = False

		Try
			LogEvent(LogEventType.Informational, "Loading Galaxy Objects...")
			'Galaxy objects
			sSQL = "SELECT * FROM tblGalaxy"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glGalaxyUB = glGalaxyUB + 1
				ReDim Preserve goGalaxy(glGalaxyUB)
				ReDim Preserve glGalaxyIdx(glGalaxyUB)
				goGalaxy(glGalaxyUB) = New Galaxy()
				With goGalaxy(glGalaxyUB)
					.GalaxyName = StringToBytes(CStr(oData("GalaxyName")))
					.ObjectID = CInt(oData("GalaxyID"))
					.ObjTypeID = ObjectType.eGalaxy
					.CycleNumber = CInt(oData("CycleNumber"))
					glGalaxyIdx(glGalaxyUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Startypes
			LogEvent(LogEventType.Informational, "Loading Star Types...")
			sSQL = "SELECT * FROM tblStarType"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glStarTypeUB += 1
				ReDim Preserve goStarType(glStarTypeUB)
				ReDim Preserve glStarTypeIdx(glStarTypeUB)
				goStarType(glStarTypeUB) = New StarType()
				With goStarType(glStarTypeUB)
					.HeatIndex = CByte(oData("HeatIndex"))
					.LightAmbient = CInt(oData("LightAmbient"))
					.LightAtt1 = CSng(oData("LightAtt1"))
					.LightAtt2 = CSng(oData("LightAtt2"))
					.LightAtt3 = CSng(oData("LightAtt3"))
					.LightDiffuse = CInt(oData("LightDiffuse"))
					.LightRange = CInt(oData("LightRange"))
					.LightSpecular = CInt(oData("LightSpecular"))
					.MatDiffuse = CInt(oData("MatDiffuse"))
					.MatEmissive = CInt(oData("MatEmissive"))
					.StarMapRectIdx = CByte(oData("StarMapRectIdx"))
					.StarRadius = CInt(oData("StarRadius"))
					.StarRarity = CByte(oData("StarRarity"))
					.StarTexture = StringToBytes(CStr(oData("StarTexture")))
					.StarTypeAttrs = CInt(oData("StarTypeAttrs"))
					.StarTypeID = CByte(oData("StarTypeID"))
					.StarTypeName = StringToBytes(CStr(oData("StarTypeName")))
					glStarTypeIdx(glStarTypeUB) = .StarTypeID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Nebula Objects
			LogEvent(LogEventType.Informational, "Loading Nebulae...")
			sSQL = "SELECT * FROM tblNebula"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glNebulaUB = glNebulaUB + 1
				ReDim Preserve goNebula(glNebulaUB)
				ReDim Preserve glNebulaIdx(glNebulaUB)
				goNebula(glNebulaUB) = New Nebula()
				With goNebula(glNebulaUB)
					.Effects = CInt(oData("Effects"))
					.Height = CByte(oData("Height"))
					.LocX = CByte(oData("LocX"))
					.LocY = CByte(oData("LocY"))
					.NebulaDesc = CStr(oData("NebulaDesc"))
					.NebulaName = StringToBytes(CStr(oData("NebulaName")))
					.ObjectID = CInt(oData("NebulaID"))
					.ObjTypeID = ObjectType.eNebula
					.Width = CByte(oData("Width"))
					glNebulaIdx(glNebulaUB) = .ObjectID
					.ParentGalaxy = GetEpicaGalaxy(CInt(oData("galaxyid")))

					'TODO: May want to 'subscribe' the nebula to the galaxy object at this point
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'System Objects
			LogEvent(LogEventType.Informational, "Loading Solar Systems...")
			sSQL = "SELECT * FROM tblSolarSystem"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glSystemUB = glSystemUB + 1
				ReDim Preserve goSystem(glSystemUB)
				ReDim Preserve glSystemIdx(glSystemUB)
				goSystem(glSystemUB) = New SolarSystem()
				With goSystem(glSystemUB)
					.LocX = CInt(oData("LocX"))
					.LocY = CInt(oData("LocY"))
					.LocZ = CInt(oData("LocZ"))
					.ObjectID = CInt(oData("SystemID"))
					.ObjTypeID = ObjectType.eSolarSystem
					.ParentGalaxy = GetEpicaGalaxy(CInt(oData("galaxyid")))
					.StarType1ID = CByte(oData("StarType1ID"))
					.StarType2ID = CByte(oData("StarType2ID"))
					.StarType3ID = CByte(oData("StarType3ID"))
					.SystemName = StringToBytes(CStr(oData("SystemName")))
                    .SystemType = CByte(oData("SystemType"))
                    .FleetJumpPointX = CInt(oData("FleetJumpPointX"))
                    .FleetJumpPointZ = CInt(oData("FleetJumpPointZ"))
					glSystemIdx(glSystemUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
            oComm = Nothing

			'Planet Objects
			LogEvent(LogEventType.Informational, "Loading Planets...")
			sSQL = "SELECT * FROM tblPlanet"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glPlanetUB = glPlanetUB + 1
				ReDim Preserve goPlanet(glPlanetUB)
				ReDim Preserve glPlanetIdx(glPlanetUB)
				goPlanet(glPlanetUB) = New Planet()
				With goPlanet(glPlanetUB)
					.Atmosphere = CByte(oData("Atmosphere"))
					.AxisAngle = CInt(oData("AxisAngle"))
					.Gravity = CByte(oData("Gravity"))
					.Hydrosphere = CByte(oData("Hydrosphere"))
					.InnerRingRadius = CInt(oData("RingInnerRadius"))
					.LocX = CInt(oData("LocX"))
					.LocY = CInt(oData("LocY"))
					.LocZ = CInt(oData("LocZ"))
					.ObjectID = CInt(oData("PlanetID"))
					.ObjTypeID = ObjectType.ePlanet
					.OuterRingRadius = CInt(oData("RingOuterRadius"))
					.ParentSystem = GetEpicaSystem(CInt(oData("parentID")))
					'Set the parent system's Planet Index
					.ParentSystem.AddPlanetIndex(glPlanetUB)
					.PlanetName = StringToBytes(CStr(oData("PlanetName")))
					.PlanetRadius = CShort(oData("PlanetRadius"))
					.PlanetSizeID = CByte(oData("PlanetSizeID"))
					.PlanetTypeID = CByte(oData("PlanetTypeID"))
					.RingDiffuse = CInt(oData("RingDiffuse"))
					.RotationDelay = CShort(oData("RotationDelay"))
					.SurfaceTemperature = CByte(oData("SurfaceTemperature"))
                    .Vegetation = CByte(oData("Vegetation"))
                    .OwnerID = CInt(oData("OwnerID"))
                    .RingConcentration = CInt(oData("RingMineralConcentration"))
                    .RingMineralID = CInt(oData("RingMineralID"))

                    .PlayerSpawns = CInt(oData("PlayerSpawns"))
					.SpawnLocked = CByte(oData("SpawnLocked")) <> 0

					glPlanetIdx(glPlanetUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Wormhole objects
			LogEvent(LogEventType.Informational, "Loading Wormholes...")
			sSQL = "SELECT * FROM tblWormhole"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glWormholeUB = glWormholeUB + 1
				ReDim Preserve goWormhole(glWormholeUB)
				ReDim Preserve glWormholeIdx(glWormholeUB)
				goWormhole(glWormholeUB) = New Wormhole()
				With goWormhole(glWormholeUB)
                    .WormholeFlags = CInt(oData("WormholeFlag"))
					.LocX1 = CInt(oData("LocX1"))
					.LocY1 = CInt(oData("LocY1"))
					.LocX2 = CInt(oData("LocX2"))
					.LocY2 = CInt(oData("LocY2"))
					.ObjectID = CInt(oData("WormholeID"))
					.ObjTypeID = ObjectType.eWormhole
					.StartCycle = CInt(oData("StartCycle"))
					.System1 = GetEpicaSystem(CInt(oData("System1ID")))
					.System2 = GetEpicaSystem(CInt(oData("System2ID")))
					glWormholeIdx(glWormholeUB) = .ObjectID

					If .System1 Is Nothing = False Then .System1.AddWormholeReference(goWormhole(glWormholeUB))
					If .System2 Is Nothing = False Then .System2.AddWormholeReference(goWormhole(glWormholeUB))
				End With
			End While
			oData.Close()
			oData = Nothing
            oComm = Nothing

            LogEvent(LogEventType.Informational, "Setting current Hub Hub...")
            Dim lMaxHHID As Int32 = -1
            Dim oMaxHH As SolarSystem = Nothing
            For X As Int32 = 0 To glSystemUB
                If glSystemIdx(X) > -1 Then
                    Dim oSys As SolarSystem = goSystem(X)
                    If oSys Is Nothing = False AndAlso oSys.SystemType = GeoSpawner.elSystemType.HubHubSystem Then
                        If oSys.ObjectID > lMaxHHID Then
                            lMaxHHID = oSys.ObjectID
                            oMaxHH = oSys
                        End If
                    End If
                End If
            Next X
            GeoSpawner.SetStartHubHub(oMaxHH)

            LogEvent(LogEventType.Informational, "Loading System Entity Counts...")
            sSQL = "SELECT * FROM tblSystemEntityCount"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True
                Dim lSystemID As Int32 = CInt(oData("SystemID"))
                Dim oSys As SolarSystem = GetEpicaSystem(lSystemID)
                If oSys Is Nothing = False Then
                    oSys.lEntityCount = CInt(oData("UnitCount")) + CInt(oData("FacilityCount"))
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing
            'For X As Int32 = 0 To glSystemUB
            '    Dim oSys As SolarSystem = goSystem(X)
            '    If oSys Is Nothing = False Then
            '        Dim sSysID As String = oSys.ObjectID.ToString
            '        sSQL = "SELECT (select count(*) from tblunit where  (parenttypeid = 2 and parentid = " & sSysID & _
            '            ") OR (parenttypeid = 3 and parentid in (select planetid from tblplanet where parentid = " & sSysID & _
            '            "))) + (select count(*) from tblstructure where (parenttypeid = 2 and parentid = " & sSysID & _
            '            ") OR (parenttypeid = 3 and parentid in (select planetid from tblplanet where parentid = " & sSysID & ")))"
            '        oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '        oData = oComm.ExecuteReader()
            '        If oData.Read = True Then
            '            If oData(0) Is DBNull.Value = True Then
            '                oSys.lEntityCount = 0
            '            Else
            '                oSys.lEntityCount = CInt(oData(0))
            '            End If
            '        End If
            '        If oData Is Nothing = False Then oData.Close()
            '        oData = Nothing
            '        If oComm Is Nothing = False Then oComm.Dispose()
            '        oComm = Nothing
            '    End If
            'Next X

            'Unlock T3-T3 wormholes
            LogEvent(LogEventType.Informational, "Unlocking T3-T3 Wormholes...")
            sSQL = "select WormholeID, System1ID, WormholeFlag from tblwormhole where system1id in (select systemid from tblsolarsystem where systemtype = 3)" & _
               " and system2id in (select systemid from tblsolarsystem where systemtype = 3) And ((wormholeflag & 16) = 0)"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            'system1id is the T3 to check, wormholeid is the wormhole to update, wormholeflag is so no data is lost
            Dim lWHID() As Int32 = Nothing
            Dim lSysID() As Int32 = Nothing
            Dim lFlag() As Int32 = Nothing
            Dim lWHUB As Int32 = -1
            While oData.Read = True
                lWHUB += 1
                ReDim Preserve lWHID(lWHUB)
                ReDim Preserve lSysID(lWHUB)
                ReDim Preserve lFlag(lWHUB)
                lWHID(lWHUB) = CInt(oData("WormholeID"))
                lSysID(lWHUB) = CInt(oData("System1ID"))
                lFlag(lWHUB) = CInt(oData("WormholeFlag"))
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Now, go through the list... and get number of HUB systems that do not have colonies for this HUB HUB
            For X As Int32 = 0 To lWHUB
                sSQL = "select count(*) from tblsolarsystem where systemid in (select system1id from tblwormhole where system2id = " & lSysID(X) & ") and " & _
                    " systemtype <> 3 and systemid not in (select parentid from tblplanet where planetid in (select parentid from tblcolony where parenttypeid = 3))"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                Dim bOpen As Boolean = oData.Read = False OrElse CInt(oData(0)) = 0
                oData.Close()
                oData = Nothing
                oComm = Nothing

                If bOpen = True Then
                    LogEvent(LogEventType.Informational, "Unlocking wormholeID " & lWHID(X))
                    Dim lNewFlag As Int32 = lFlag(X)
                    lNewFlag = lNewFlag Or 1 Or 4 Or 8 Or 16        'sys1 detect, sys1 jump, sys2 jump, sys1 colonized
                    sSQL = "UPDATE tblWormhole SET WormholeFlag = " & lNewFlag & " WHERE WormholeID = " & lWHID(X)
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oComm.ExecuteNonQuery()
                    oComm.Dispose()
                    oComm = Nothing
                End If
            Next X
            
            LogEvent(LogEventType.Informational, "Checking for Mineral Regeneration...")
            For X As Int32 = 0 To glSystemUB
                If glSystemIdx(X) > -1 Then
                    If goSystem(X).ObjectID = 88 Then Continue For
                    sSQL = "SELECT COUNT(CacheID) FROM tblMineralCache WHERE ParentID IN (SELECT PlanetID FROM tblPlanet WHERE tblPlanet.ParentID=" & goSystem(X).ObjectID & ")"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader(CommandBehavior.Default)
                    Dim bRegenerate As Boolean = oData.Read = False OrElse CInt(oData(0)) = 0
                    oData.Close()
                    oData = Nothing
                    oComm = Nothing
                    If bRegenerate = True Then
                        LogEvent(LogEventType.Informational, "... Mineral Regeneration " & goSystem(X).ObjectID.ToString & " -" & BytesToString(goSystem(X).SystemName))
                        GeoSpawner.SpawnSystemsMinerals(goSystem(X).ObjectID)
                    End If
                End If
            Next

            bResult = True
        Catch
            LogEvent(LogEventType.CriticalError, "LoadGeographical Error: " & Err.Description)
        Finally
            If oData Is Nothing = False Then
                oData.Close()
            End If
            oData = Nothing
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Private Function LoadPlayers() As Boolean
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim oComm As OleDb.OleDbCommand
		Dim sSQL As String
		Dim bResult As Boolean = False

		Try

			ReDim goPlayer(-1)
			sSQL = "SELECT COUNT(*) FROM tblPlayer"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			If oData.Read = True Then
				Dim lValue As Int32 = CInt(oData(0))
				ReDim goPlayer(lValue - 1)
				ReDim glPlayerIdx(lValue - 1)
			End If

			'And now the player objects
			LogEvent(LogEventType.Informational, "Loading Players...")
			sSQL = "SELECT * FROM tblPlayer"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glPlayerUB = glPlayerUB + 1
				If goPlayer.GetUpperBound(0) < glPlayerUB Then
					ReDim Preserve goPlayer(glPlayerUB)
					ReDim Preserve glPlayerIdx(glPlayerUB)
				End If

				goPlayer(glPlayerUB) = New Player()
				With goPlayer(glPlayerUB)
					.EmpireName = StringToBytes(CStr(oData("EmpireName")))
					.ObjectID = CInt(oData("PlayerID"))
					.ObjTypeID = ObjectType.ePlayer
					.sPlayerName = UCase$(CStr(oData("PlayerName")))
					.sPlayerNameProper = CStr(oData("PlayerName"))
                    .PlayerName = StringToBytes(CStr(oData("PlayerName")))
                    StringToBytes(CStr(oData("PlayerPassword"))).CopyTo(.PlayerPassword, 0)
                    StringToBytes(CStr(oData("PlayerUserName"))).CopyTo(.PlayerUserName, 0)
					.RaceName = StringToBytes(CStr(oData("RaceName")))
					.lLastViewedEnvir = CInt(oData("LastViewedID"))
					.iLastViewedEnvirType = CShort(oData("LastViewedTypeID"))
					.lStartedEnvirID = CInt(oData("StartedEnvirID"))
					.iStartedEnvirTypeID = CShort(oData("StartedEnvirTypeID"))
					glPlayerIdx(glPlayerUB) = .ObjectID
					.blCredits = CLng(oData("Credits"))
					.AccountStatus = CInt(oData("AccountStatus"))
					.yGender = CByte(oData("PlayerGender"))

					.LastLogin = GetDateFromNumber(CInt(oData("LastLogin")))
					.TotalPlayTime = CInt(oData("TotalPlayTime"))

					.lStartLocX = CInt(oData("StartLocX"))
					.lStartLocZ = CInt(oData("StartLocZ"))
					.PirateStartLocX = CInt(oData("PirateStartLocX"))
                    .PirateStartLocZ = CInt(oData("PirateStartLocZ"))

                    .yPlayerTitle = CByte(oData("PlayerTitle"))

					.lPlayerIcon = CInt(oData("PlayerIcon"))

					ReDim .ExternalEmailAddress(254)
					If oData("EmailAddress") Is DBNull.Value = False Then
						StringToBytes(CStr(oData("EmailAddress"))).CopyTo(.ExternalEmailAddress, 0)
					End If
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Player Alias'
			LogEvent(LogEventType.Informational, "Loading Player Aliasing...")
			sSQL = "SELECT * FROM tblAlias"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
				If oPlayer Is Nothing = False Then
					Dim oOther As Player = GetEpicaPlayer(CInt(oData("OtherPlayerID")))
					If oOther Is Nothing = False Then
						Dim yUserName(19) As Byte
						StringToBytes(CStr(oData("sAliasUserName"))).CopyTo(yUserName, 0)
						Dim yPassword(19) As Byte
						StringToBytes(CStr(oData("sAliasPassword"))).CopyTo(yPassword, 0)
						Dim lRights As Int32 = CInt(oData("lAliasRights"))
						oPlayer.AddPlayerAlias(oOther, yUserName, yPassword, lRights)
						oOther.AddPlayerAliasAllowance(oPlayer, yUserName, yPassword, lRights)
					End If
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

            LogEvent(LogEventType.Informational, "Associating Players to Planets...")
            For X As Int32 = 0 To glPlanetUB
                If glPlanetIdx(X) > -1 Then
                    Dim oPlanet As Planet = goPlanet(X)
                    If oPlanet Is Nothing = False AndAlso oPlanet.OwnerID > 0 Then
                        Dim oOwner As Player = GetEpicaPlayer(oPlanet.OwnerID)
                        If oOwner Is Nothing = False Then
                            oOwner.AddPlanetControl(X)
                        End If
                    End If
                End If
            Next X

			bResult = True
		Catch
			LogEvent(LogEventType.CriticalError, "LoadPlayers Error: " & Err.Description)
		Finally
			If oData Is Nothing = False Then
				oData.Close()
			End If
			oData = Nothing
			oComm = Nothing
		End Try
		Return bResult
    End Function

    Private Function LoadSenate() As Boolean
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim oComm As OleDb.OleDbCommand
        Dim sSQL As String
        Dim bResult As Boolean = False

        Try
            'Load our senate stuff
            LogEvent(LogEventType.Informational, "Loading Senate Proposals...")
            sSQL = "SELECT * FROM tblSenateProposal Order By VotingEndDate DESC"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oProposal As New SenateProposal()
                With oProposal
                    .lProposedBy = CInt(oData("ProposedBy"))
                    .lVotesAgainst = 0
                    .lVotesFor = 0
                    .ObjectID = CInt(oData("ProposalID"))
                    .ObjTypeID = ObjectType.eSenateLaw
                    .yDescription = StringToBytes(CStr(oData("ProposalDescription")))
                    .yProposalState = CType(CByte(oData("ProposalState")), eyProposalState)
                    .yTitle = StringToBytes(CStr(oData("ProposalTitle")))

                    .yDefaultVote = CByte(oData("DefaultVote"))
                    .lVotingEndDate = CInt(oData("VotingEndDate"))
                    .lVotingStartDate = CInt(oData("VotingStartDate"))
                    .lDeliveryEstimate = CInt(oData("DeliveryEstimate"))
                End With
                Senate.AddNewProposal(oProposal)
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            LogEvent(LogEventType.Informational, "Loading Senate Votes...")
            sSQL = "SELECT * FROM tblSenateProposalVote"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oProposal As SenateProposal = Senate.GetProposal(CInt(oData("ProposalID")))
                If oProposal Is Nothing = False Then
                    Dim lVoterID As Int32 = CInt(oData("VoterID"))
                    Dim yVoteValue As eyVoteValue = CType(oData("VoteValue"), eyVoteValue)
                    Dim yVotePriority As Byte = CByte(oData("VotePriority"))
                    oProposal.HandlePlayerVote(lVoterID, yVoteValue, True, yVotePriority)
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            LogEvent(LogEventType.Informational, "Recalculating Senate Votes...")
            Senate.RecalculateAllProposals()

            LogEvent(LogEventType.Informational, "Loading Senate Messages...")
            sSQL = "SELECT * FROM tblSenateProposalMsg ORDER BY PostedOn"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oProposal As SenateProposal = Senate.GetProposal(CInt(oData("ProposalID")))
                If oProposal Is Nothing = False Then
                    Dim lPosterID As Int32 = CInt(oData("PosterID"))
                    Dim lPostedOn As Int32 = CInt(oData("PostedOn"))
                    Dim sMsgData As String = CStr(oData("MsgData"))
                    oProposal.AddMsg(lPosterID, lPostedOn, sMsgData)
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing
            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "LoadSenate Error: " & Err.Description)
        Finally
            If oData Is Nothing = False Then
                oData.Close()
            End If
            oData = Nothing
            oComm = Nothing
        End Try
        Return bResult

    End Function

	Public Function LoadSinglePlayer(ByVal lPlayerID As Int32) As Boolean
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim oComm As OleDb.OleDbCommand
		Dim sSQL As String
		Dim bResult As Boolean = False

		Try
			'And now the player objects
			sSQL = "SELECT * FROM tblPlayer WHERE PlayerID = " & lPlayerID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = New Player
				With oPlayer
					.EmpireName = StringToBytes(CStr(oData("EmpireName")))
					.ObjectID = CInt(oData("PlayerID"))
					.ObjTypeID = ObjectType.ePlayer
					.sPlayerName = UCase$(CStr(oData("PlayerName")))
					.sPlayerNameProper = CStr(oData("PlayerName"))
					.PlayerName = StringToBytes(CStr(oData("PlayerName")))
					StringToBytes(CStr(oData("PlayerPassword"))).CopyTo(.PlayerPassword, 0)
					StringToBytes(CStr(oData("PlayerUserName"))).CopyTo(.PlayerUserName, 0)
					.RaceName = StringToBytes(CStr(oData("RaceName")))
					.lLastViewedEnvir = CInt(oData("LastViewedID"))
					.iLastViewedEnvirType = CShort(oData("LastViewedTypeID"))
					.lStartedEnvirID = CInt(oData("StartedEnvirID"))
					.iStartedEnvirTypeID = CShort(oData("StartedEnvirTypeID"))
					.blCredits = CLng(oData("Credits"))
					.AccountStatus = CInt(oData("AccountStatus"))
					.yGender = CByte(oData("PlayerGender"))

					.LastLogin = GetDateFromNumber(CInt(oData("LastLogin")))
					.TotalPlayTime = CInt(oData("TotalPlayTime"))

					.lStartLocX = CInt(oData("StartLocX"))
					.lStartLocZ = CInt(oData("StartLocZ"))
					.PirateStartLocX = CInt(oData("PirateStartLocX"))
					.PirateStartLocZ = CInt(oData("PirateStartLocZ"))

					.lPlayerIcon = CInt(oData("PlayerIcon"))

					ReDim .ExternalEmailAddress(254)
					If oData("EmailAddress") Is DBNull.Value = False Then
						StringToBytes(CStr(oData("EmailAddress"))).CopyTo(.ExternalEmailAddress, 0)
					End If
				End With

				'TODO: Not sure about this synclock
				SyncLock goPlayer
					ReDim Preserve goPlayer(glPlayerUB + 1)
					ReDim Preserve glPlayerIdx(glPlayerUB + 1)
					goPlayer(glPlayerUB + 1) = oPlayer
					glPlayerIdx(glPlayerUB + 1) = oPlayer.ObjectID
					glPlayerUB = glPlayerUB + 1
				End SyncLock

			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			bResult = True
		Catch
			LogEvent(LogEventType.CriticalError, "LoadSinglePlayer Error: " & Err.Description)
		Finally
			If oData Is Nothing = False Then
				oData.Close()
			End If
			oData = Nothing
			oComm = Nothing
		End Try
		Return bResult
	End Function
#End Region

	Public Function GetDateFromNumber(ByVal lVal As Int32) As Date
		If lVal = 0 Then Return Date.MinValue

		Dim sVal As String = lVal.ToString

		'Work from right to left
		Dim lUB As Int32 = sVal.Length ' - 1

		'Ok, the bare minimum for this to work is 8
		If lUB < 8 Then Return Date.MinValue

		'Minute, last two values
		Dim sMin As String = sVal.Substring(lUB - 2)
		'Hour, two less from minute
		Dim sHr As String = sVal.Substring(lUB - 4, 2)
		'etc...
		Dim sDay As String = sVal.Substring(lUB - 6, 2)
		Dim sMon As String = sVal.Substring(lUB - 8, 2)

		Dim sYr As String = ""
		If lUB = 9 Then
			sYr = "0" & sVal.Substring(lUB - 9, 1)
		Else : sYr = sVal.Substring(lUB - 10, 2)
		End If

		Return CDate(sMon & "/" & sDay & "/20" & sYr & " " & sHr & ":" & sMin)
	End Function

	Public Function GetDateAsNumber(ByVal dtDate As Date) As Int32
		Return CInt(Val(dtDate.ToString("yyMMddHHmm")))
	End Function

    Public Sub SpawnNeighborhoods()
        'Get our box count
        Dim lBoxCnt As Int32 = 0
        For X As Int32 = 0 To goMsgSys.mlServerUB
            If goMsgSys.oServerObject(X) Is Nothing = False AndAlso goMsgSys.oServerObject(X).lConnectionType = ConnectionType.eBoxOperator Then lBoxCnt += 1
        Next X

        'Ok, determine number of neighborhoods
        Dim lHoodCnt As Int32 = lNeighborhoodUB + 1
        If lHoodCnt < 1 Then lHoodCnt = 1
        If gb_IS_TEST_SERVER = False Then lHoodCnt = 2 Else lHoodCnt = 1

        'Now, # of neighborhoods per box?
        Dim lHoodPerBox As Int32 = lHoodCnt \ lBoxCnt
        Dim lRemainder As Int32 = lHoodCnt - (lHoodPerBox * lBoxCnt)

        'Now, we have something to indicate what to spawn... check if there is a serverobject for an email server
        Dim oEmailSrvr As ServerObject = Nothing
        For X As Int32 = 0 To goMsgSys.mlServerUB
            If goMsgSys.oServerObject(X) Is Nothing = False AndAlso goMsgSys.oServerObject(X).lConnectionType = ConnectionType.eEmailServerApp Then
                oEmailSrvr = goMsgSys.oServerObject(X)
                Exit For
            End If
        Next X
        If oEmailSrvr Is Nothing = False Then
            If lHoodCnt = 1 Then
                Dim lSpawnID(glSystemUB) As Int32
                Dim iSpawnTypeID(glSystemUB) As Int16
                For X As Int32 = 0 To glSystemUB
                    If glSystemIdx(X) <> -1 Then
                        lSpawnID(X) = goSystem(X).ObjectID
                    Else : lSpawnID(X) = -1
                    End If
                    iSpawnTypeID(X) = ObjectType.eSolarSystem
                Next X
                SpawnServer(oEmailSrvr, ConnectionType.ePrimaryServerApp, lSpawnID, iSpawnTypeID, glSystemUB)
            Else
                'ok, for this test, we are putting 29 and 30 on box 1
                ' and 32 and 33 are on box 2 

                'Dim lSpawnID() As Int32 = {29, 30} ', 32, 33}
                'Dim iSpawnTypeID() As Int16 = {ObjectType.eSolarSystem, ObjectType.eSolarSystem} ', ObjectType.eSolarSystem, ObjectType.eSolarSystem}
                'SpawnServer(oEmailSrvr, ConnectionType.ePrimaryServerApp, lSpawnID, iSpawnTypeID, 1)
                'lSpawnID(0) = 32 : lSpawnID(1) = 33
                'SpawnServer(oEmailSrvr, ConnectionType.ePrimaryServerApp, lSpawnID, iSpawnTypeID, 1)


                Dim lSpawnID(glSystemUB) As Int32
                Dim iSpawnTypeID(glSystemUB) As Int16
                For X As Int32 = 0 To glSystemUB
                    If glSystemIdx(X) <> -1 Then
                        lSpawnID(X) = goSystem(X).ObjectID
                    Else : lSpawnID(X) = -1
                    End If
                    iSpawnTypeID(X) = ObjectType.eSolarSystem
                Next X
                SpawnServer(oEmailSrvr, ConnectionType.ePrimaryServerApp, lSpawnID, iSpawnTypeID, glSystemUB)

            End If
        Else : SpawnServer(Nothing, ConnectionType.eEmailServerApp, Nothing, Nothing, -1)
        End If

    End Sub

    Private Structure uSpawnList
        Public lSpawnID() As Int32
        Public iSpawnTypeID() As Int16
        Public lSpawnUB As Int32
    End Structure
#Region "  OLD SPAWN REGION IDEA  "
    ''TODO: Finish this up

    ''Could determine the memory usage of the entire primary...
    ''  then determine what regions need to be generated from that memory usage
    ''  need to take into account number of boxes, etc... as well

    'Dim lSpawnID() As Int32 = {36} '{29, 30}
    'Dim iSpawnTypeID() As Int16 = {ObjectType.eSolarSystem} ', ObjectType.eSolarSystem} ', ObjectType.eSolarSystem, ObjectType.eSolarSystem}

    ''Put Aurelium on the first box, (36). This also puts all tutorial accounts on the first region processor
    'SpawnServer(oPrimary, ConnectionType.eRegionServerApp, lSpawnID, iSpawnTypeID, 0)

    'Dim lCnt As Int32 = 0
    'For X As Int32 = 0 To uPrimaryRequest.lSpawnID.GetUpperBound(0)
    '    If uPrimaryRequest.lSpawnID(X) <> 36 Then lCnt += 1
    'Next X

    'Dim lPerBox As Int32 = lCnt \ 3
    'Dim lRemain As Int32 = lCnt - (lPerBox * 3)
    'Dim lIdx As Int32 = -1
    'For X As Int32 = 0 To 2
    '    Dim lThisCnt As Int32 = lPerBox
    '    If lRemain <> 0 Then
    '        lThisCnt += 1
    '        lRemain -= 1
    '    End If
    '    ReDim lSpawnID(lThisCnt - 1) : ReDim iSpawnTypeID(lThisCnt - 1)
    '    For Y As Int32 = 0 To lThisCnt - 1
    '        lIdx += 1
    '        If uPrimaryRequest.lSpawnID(lIdx) = 36 Then lIdx += 1
    '        lSpawnID(Y) = uPrimaryRequest.lSpawnID(lIdx)
    '        iSpawnTypeID(Y) = ObjectType.eSolarSystem
    '    Next Y

    '    SpawnServer(oPrimary, ConnectionType.eRegionServerApp, lSpawnID, iSpawnTypeID, lSpawnID.GetUpperBound(0))
    'Next X

    ''lSpawnID(0) = 32 : lSpawnID(1) = 33
    ''SpawnServer(oPrimary, ConnectionType.eRegionServerApp, lSpawnID, iSpawnTypeID, 1)

    ''If oPrimary.lBoxOperatorID = 1 Then
    ''    Dim lSpawnID() As Int32 = {29} ', 30}
    ''    Dim iSpawnTypeID() As Int16 = {ObjectType.eSolarSystem} ', ObjectType.eSolarSystem} ', ObjectType.eSolarSystem, ObjectType.eSolarSystem}
    ''    SpawnServer(oPrimary, ConnectionType.eRegionServerApp, lSpawnID, iSpawnTypeID, 0)

    ''    lSpawnID(0) = 30 ' 32 : lSpawnID(1) = 33
    ''    SpawnServer(oPrimary, ConnectionType.eRegionServerApp, lSpawnID, iSpawnTypeID, 0)
    ''Else
    ''    Dim lSpawnID() As Int32 = {32}
    ''    Dim iSpawnTypeID() As Int16 = {ObjectType.eSolarSystem} ', ObjectType.eSolarSystem} ', ObjectType.eSolarSystem, ObjectType.eSolarSystem}
    ''    SpawnServer(oPrimary, ConnectionType.eRegionServerApp, lSpawnID, iSpawnTypeID, 0)

    ''    lSpawnID(0) = 33
    ''    SpawnServer(oPrimary, ConnectionType.eRegionServerApp, lSpawnID, iSpawnTypeID, 0)
    ''End If
    'TODO: Finish this up

    'Could determine the memory usage of the entire primary...
    '  then determine what regions need to be generated from that memory usage
    '  need to take into account number of boxes, etc... as well
#End Region
    Public Sub SpawnRegions(ByVal oPrimary As ServerObject)
        'ok, check the primary server
        If oPrimary Is Nothing Then Return

        'now, check if the primary server's request idx is -1
        'If oPrimary.lSpawnRequestIdx = -1 Then Return

        'ok, now, set a temp request
        'Dim uPrimaryRequest As SpawnRequest = uSpawnRequests(oPrimary.lSpawnRequestIdx)

        'get the box count
        Dim lBoxCnt As Int32 = 0
        For X As Int32 = 0 To goMsgSys.mlServerUB
            If goMsgSys.oServerObject(X) Is Nothing = False AndAlso goMsgSys.oServerObject(X).lConnectionType = ConnectionType.eBoxOperator Then lBoxCnt += 1
        Next X
        If lBoxCnt = 1 Then
            'only one place to send it
            'SpawnServer(oPrimary, ConnectionType.eRegionServerApp, uPrimaryRequest.lSpawnID, uPrimaryRequest.iSpawnTypeID, uPrimaryRequest.lSpawnUB)
            SpawnServer(oPrimary, ConnectionType.eRegionServerApp, oPrimary.lSpawnID, oPrimary.iSpawnTypeID, oPrimary.lSpawnUB)
        Else

            Dim lCnt As Int32 = oPrimary.lSpawnID.GetUpperBound(0) + 1 'uPrimaryRequest.lSpawnID.GetUpperBound(0) + 1

            Dim lTotalBoxCnt As Int32 = 9       '3 on each box
            Dim uSpawns(lTotalBoxCnt - 1) As uSpawnList
            For X As Int32 = 0 To uSpawns.GetUpperBound(0)
                uSpawns(X).lSpawnUB = -1
                ReDim uSpawns(X).lSpawnID(-1)
            Next X

            'ordered in (App * 3) + Box

            'Dim lBox As Int32 = 0
            Dim lCurBox As Int32 = 0
            Dim lCurApp As Int32 = 0

            'Ok, sort the system ID's...
            Dim oAllSys(lCnt - 1) As SolarSystem
            For X As Int32 = 0 To oPrimary.lSpawnUB
                oAllSys(X) = GetEpicaSystem(oPrimary.lSpawnID(X))
            Next X
            Dim lSorted(lCnt - 1) As Int32
            Dim lSortedUB As Int32 = -1
            Dim bUsed(lCnt - 1) As Boolean
            For X As Int32 = 0 To oPrimary.lSpawnUB
                Dim lIdx As Int32 = -1
                For Y As Int32 = 0 To lSortedUB
                    If oAllSys(lSorted(Y)).lEntityCount > oAllSys(X).lEntityCount Then
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
            For X As Int32 = 0 To lSortedUB
                bUsed(X) = False
            Next X

            'Now, the systems are sorted... the way this works...
            '  we go through the spawn grps...
            '  we assign the highest value and the lowest value to the box
            '  we mark those values as used and move to the next box
            '  we continue this process until all values are marked as used
            Dim bDone As Boolean = False
            While bDone = False
                'Ok, assign the top score
                Dim lHighIdx As Int32 = -1
                Dim lLowIdx As Int32 = -1

                For X As Int32 = lSortedUB To 0 Step -1
                    If bUsed(X) = False Then
                        lHighIdx = X
                        Exit For
                    End If
                Next X
                For X As Int32 = 0 To lSortedUB
                    If bUsed(X) = False Then
                        lLowIdx = X
                        Exit For
                    End If
                Next X

                If lHighIdx = lLowIdx Then lLowIdx = -1

                If lHighIdx = -1 Then
                    bDone = True
                Else
                    bUsed(lHighIdx) = True

                    Dim lListIdx As Int32 = (lCurApp * 3) + lCurBox

                    With uSpawns(lListIdx)
                        .lSpawnUB += 1
                        ReDim Preserve .lSpawnID(.lSpawnUB)
                        ReDim Preserve .iSpawnTypeID(.lSpawnUB)
                        .lSpawnID(.lSpawnUB) = oPrimary.lSpawnID(lSorted(lHighIdx))
                        .iSpawnTypeID(.lSpawnUB) = ObjectType.eSolarSystem
                    End With 
                End If
                If lLowIdx = -1 Then
                    bDone = True
                Else
                    bUsed(lLowIdx) = True

                    Dim lListIdx As Int32 = (lCurApp * 3) + lCurBox

                    With uSpawns(lListIdx)
                        .lSpawnUB += 1
                        ReDim Preserve .lSpawnID(.lSpawnUB)
                        ReDim Preserve .iSpawnTypeID(.lSpawnUB)
                        .lSpawnID(.lSpawnUB) = oPrimary.lSpawnID(lSorted(lLowIdx))
                        .iSpawnTypeID(.lSpawnUB) = ObjectType.eSolarSystem
                    End With
 
                End If

                'lBox += 1
                'If lBox > 2 Then lBox = 0
                lCurBox += 1
                If lCurBox > 2 Then
                    lCurBox = 0
                    lCurApp += 1
                    If lCurApp > 2 Then
                        lCurApp = 0
                    End If
                End If
            End While


            'Ok, go through each item...
            For X As Int32 = 0 To uSpawns.GetUpperBound(0)
                With uSpawns(X)
                    SpawnServer(oPrimary, ConnectionType.eRegionServerApp, .lSpawnID, .iSpawnTypeID, .lSpawnUB)
                End With
            Next X
 
        End If

    End Sub

    Public Sub SpawnServer(ByRef oRelatedServer As ServerObject, ByVal lConnType As ConnectionType, ByVal lSpawnIDs() As Int32, ByVal iSpawnTypeIDs() As Int16, ByVal lSpawnUB As Int32)
        Dim lBoxOperatorID As Int32 = goMsgSys.GetAvailableBoxOperator(lConnType)
        If lBoxOperatorID = -1 Then
            LogEvent(LogEventType.CriticalError, "Unable to spawn Server: BoxOperatorID is -1")
            Return
        End If

        'Now, create our request
        Dim uSpawnRequest As SpawnRequest
        With uSpawnRequest
            .dtRequest = Now
            .lBoxOperatorID = lBoxOperatorID
            .lConnectionType = lConnType
            .lSpawnUB = lSpawnUB
            ReDim .lSpawnID(.lSpawnUB)
            ReDim .iSpawnTypeID(.lSpawnUB)
            For X As Int32 = 0 To .lSpawnUB
                .lSpawnID(X) = lSpawnIDs(X)
                .iSpawnTypeID(X) = iSpawnTypeIDs(X)
            Next X

            .oBoxOperator = Nothing
            .oRelatedServer = oRelatedServer
            .ySpawnRequestState = 1
            '.lSpawnRequestUniqueID = GetSpawnRequestUniqueID()
        End With
        AddSpawnRequest(uSpawnRequest)
    End Sub

	'Public Function GetAvailableBoxOperator(ByVal lConnType As ConnectionType) As Int32
	'	'Region Server Usage
	'	'===================
	'	'assume each system is 300 MB
	'	'assume each planet is 1 MB
	'	'therefore, a system with 32 planets would be 332 MB

	'	'Primary Server Usage
	'	'====================
	'	'assume each primary uses 300 MB

	'	Dim lBoxOperatorID() As Int32 = Nothing
	'	Dim lMemoryUsed() As Int32 = Nothing	  'in MB
	'	Dim lProcesses() As Int32 = Nothing
	'	Dim lUB As Int32 = -1

	'	For X As Int32 = 0 To lSpawnRequestUB
	'		If uSpawnRequests(X).ySpawnRequestState > 0 Then
	'			If uSpawnRequests(X).lConnectionType <> ConnectionType.eBoxOperator Then
	'				'find our Idx for the box operator
	'				Dim lIdx As Int32 = -1
	'				For Y As Int32 = 0 To lUB
	'					If lBoxOperatorID(Y) = uSpawnRequests(X).lBoxOperatorID Then
	'						lIdx = Y
	'						Exit For
	'					End If
	'				Next Y
	'				If lIdx = -1 Then
	'					lUB += 1
	'					ReDim Preserve lBoxOperatorID(lUB)
	'					ReDim Preserve lMemoryUsed(lUB)
	'					ReDim Preserve lProcesses(lUB)
	'					lIdx = lUB
	'					lBoxOperatorID(lIdx) = uSpawnRequests(X).lBoxOperatorID
	'				End If

	'				lProcesses(lIdx) += 1

	'				'spawn request is either unfulfilled or active
	'				Select Case uSpawnRequests(X).lConnectionType
	'					Case ConnectionType.ePrimaryServerApp
	'						'just assume 300mb here
	'						lMemoryUsed(lIdx) += 300
	'					Case ConnectionType.eRegionServerApp
	'						For Y As Int32 = 0 To uSpawnRequests(X).lSpawnUB
	'							If uSpawnRequests(X).iSpawnTypeID(Y) = ObjectType.eSolarSystem Then
	'								lMemoryUsed(lIdx) += 300
	'							ElseIf uSpawnRequests(X).iSpawnTypeID(Y) = ObjectType.ePlanet Then
	'								lMemoryUsed(lIdx) += 1
	'							End If
	'						Next Y
	'				End Select
	'			End If
	'		End If
	'	Next X

	'	'Ok, now, look through our server objects and find all our box operators
	'	For X As Int32 = 0 To goMsgSys.mlServerUB
	'		If goMsgSys.oServerObject(X) Is Nothing = False AndAlso goMsgSys.oServerObject(X).lConnectionType = ConnectionType.eBoxOperator Then
	'			Dim lIdx As Int32 = -1
	'			For Y As Int32 = 0 To lUB
	'				If lBoxOperatorID(Y) = goMsgSys.oServerObject(X).lBoxOperatorID Then
	'					lIdx = Y
	'					Exit For
	'				End If
	'			Next Y
	'			If lIdx = -1 Then
	'				lUB += 1
	'				ReDim Preserve lBoxOperatorID(lUB)
	'				ReDim Preserve lMemoryUsed(lUB)
	'				ReDim Preserve lProcesses(lUB)
	'				lIdx = lUB
	'				lBoxOperatorID(lIdx) = goMsgSys.oServerObject(X).lBoxOperatorID
	'			End If
	'		End If
	'	Next X

	'	'Now, we are ready, we know how many processes are on each box, we know estimated memory usage...
	'	Dim lMinBoxOp As Int32 = -1
	'	Dim lMinMemUsage As Int32 = Int32.MinValue
	'	For X As Int32 = 0 To lUB
	'		If lProcesses(X) = 0 Then Return lBoxOperatorID(X) '0 processes indicates a good box to use
	'		If lMemoryUsed(X) < lMinMemUsage Then
	'			lMinBoxOp = lBoxOperatorID(X)
	'			lMinMemUsage = lMemoryUsed(X)
	'		End If
	'	Next X
	'	If lMinBoxOp = -1 Then
	'		'Ok, odd... return one if we can...
	'		If lUB <> -1 Then Return lBoxOperatorID(0)
	'	Else : Return lMinBoxOp
	'	End If

	'	Return -1
	'End Function

	Private Sub AddSpawnRequest(ByVal uRequest As SpawnRequest)
		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To lSpawnRequestUB
			If uSpawnRequests(X).ySpawnRequestState = 0 Then
				lIdx = X
				Exit For
			End If
		Next X
		If lIdx = -1 Then
			lSpawnRequestUB += 1
			ReDim Preserve uSpawnRequests(lSpawnRequestUB)
			lIdx = lSpawnRequestUB
		End If
		uSpawnRequests(lIdx) = uRequest

		With uSpawnRequests(lIdx)
			For X As Int32 = 0 To goMsgSys.mlServerUB
				If goMsgSys.oServerObject(X) Is Nothing = False AndAlso goMsgSys.oServerObject(X).lConnectionType = ConnectionType.eBoxOperator Then
					If goMsgSys.oServerObject(X).lBoxOperatorID = .lBoxOperatorID Then
						.oBoxOperator = goMsgSys.oServerObject(X)
						Exit For
					End If
				End If
			Next X
			.ySpawnRequestState = 1
			.SendSpawnRequest(lIdx)

			'Now that we found what we are looking for, disconnect from the boxoperator
			'.oBoxOperator.oSocket.Disconnect()
			'.oBoxOperator.bSocketConnected = False
		End With
	End Sub

	Public Function GetPrimaryServerForSpawn() As ServerObject
		For X As Int32 = 0 To goMsgSys.mlServerUB
			If goMsgSys.oServerObject(X) Is Nothing = False AndAlso goMsgSys.oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp Then Return goMsgSys.oServerObject(X)
		Next X
		'TODO: Fix this!!!
		'For X As Int32 = 0 To goMsgSys.mlServerUB
		'    If goMsgSys.oServerObject(X) Is Nothing = False AndAlso goMsgSys.oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp Then
		'        If goMsgSys.oServerObject(X).lPlayerSpawnsAvailable = -1 OrElse goMsgSys.oServerObject(X).lPlayersSpawnedHere = -1 Then
		'            '
		'        End If
		'        If goMsgSys.oServerObject(X).lPlayersSpawnedHere < goMsgSys.oServerObject(X).lPlayerSpawnsAvailable Then
		'            Return goMsgSys.oServerObject(X)
		'        End If
		'    End If
		'Next X

		'LogEvent(LogEventType.CriticalError, "No spawn locations available!!!")
		Return Nothing
	End Function

    Private Sub SaveSixMinuteSnapshot()
        Dim lDT As Int32 = GetDateAsNumber(Now)
        Dim oComm As OleDb.OleDbCommand = Nothing

        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).lConnectedPrimaryIdx > -1 Then
                lCnt += 1
                Try
                    oComm = New OleDb.OleDbCommand("INSERT INTO tblSnapShotPlayer (DateTimeStamp, PlayerID) VALUES (" & lDT & ", " & glPlayerIdx(X) & ")", goCN)
                    oComm.ExecuteNonQuery()
                Catch ex As Exception
                    LogEvent(LogEventType.CriticalError, "Error saving Six Minute Snapshot person: " & ex.Message)
                Finally
                    If oComm Is Nothing = False Then oComm.Dispose()
                    oComm = Nothing
                End Try

            End If
        Next X



        Dim sSQL As String = "INSERT INTO tblSnapShot (DateTimeStamp, NumberPlayersOnline, GalacticCensus, PirateUnitCnt, PirateFacilityCnt, ErrorCount) VALUES (" & _
         lDT & ", " & lCnt & ", 0, 0, 0, 0)"
        Try
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "SaveSnapshot: " & ex.Message)
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try

    End Sub

	Public gbRunning As Boolean = False
	Public Sub MainLoop()
		Dim oSWKeepAlive As Stopwatch
		gbRunning = True

        Dim lLoopCnt As Int32 = 0
        Dim lSixMinCnt As Int32 = 0

        Dim lPlayerUpdateIdx As Int32 = 0
        Dim oSWPlayerUpdate As Stopwatch = Nothing

		oSWKeepAlive = Stopwatch.StartNew()
		While gbRunning = True
			If gyBackupOperator = eyOperatorState.BackupOperator Then
				If oSWKeepAlive.ElapsedMilliseconds > gl_KEEP_ALIVE_TIME Then
                    If goOperatorSW Is Nothing = False AndAlso goOperatorSW.ElapsedMilliseconds > gl_KEEP_ALIVE_TIME Then
                        'ok, bad, operator died...
                        LogEvent(LogEventType.Informational, "Server Outage: Operator!")
                        gyBackupOperator = eyOperatorState.EmergencyOperator
                        goMsgSys.bAcceptNewClients = True

                        gfrmDisplayForm.SetAcceptClientsCheck()

                        SendSupportEmail("Lost contact with Operator. Switching to Emergency Operator. Accepting clients and attempting to rebound.", "DSE Operator Server Down")
                    End If
					oSWKeepAlive.Reset()
					oSWKeepAlive.Start()
				End If
            Else
                If oSWKeepAlive.ElapsedMilliseconds > gl_KEEP_ALIVE_TIME Then
                    goMsgSys.SubmitKeepAlives()

                    lLoopCnt += 1
                    If lLoopCnt * gl_KEEP_ALIVE_TIME > 60000 Then
                        If goMsgSys.bAcceptNewClients = True Then
                            If gb_IS_TEST_SERVER = False Then CheckTransactions()
                        End If
                        lLoopCnt = 0
                    End If
                    lSixMinCnt += 1
                    If lSixMinCnt * gl_KEEP_ALIVE_TIME > 360000 Then
                        lSixMinCnt = 0
                        SaveSixMinuteSnapshot()
                    End If

                    oSWKeepAlive.Reset()
                    oSWKeepAlive.Start()


                    If gyBackupOperator = eyOperatorState.EmergencyOperator Then
                        'Ok, I am the backup operator in emergency mode, attempt to connect to the operator
                        If goMsgSys.ConnectToOperator() = True Then
                            ' 1) Send updates to all servers informing the rebound
                            goMsgSys.SendReboundMsgToServers()

                            goMsgSys.bAcceptNewClients = False
                            gfrmDisplayForm.SetAcceptClientsCheck()

                            ' 2) Send email to support informing server rebound
                            SendSupportEmail("Primary Operator has rebounded. Switching backup operator out of emergency mode.", "DSE Primary Operator Rebound")

                            ' 3) switch our state
                            gyBackupOperator = eyOperatorState.BackupOperator

                            ' 4) Because of the major changes... we will sleep
                            Threading.Thread.Sleep(100)

                            ' 5) and then continue our while loop
                            Continue While
                        End If
                    End If


                End If
                End If

                'For X As Int32 = 0 To goMsgSys.mlServerUB
                '    If goMsgSys.oServerObject(X).lConnectionType = ConnectionType.eRegionServerApp Then
                '        With uSpawnRequests(goMsgSys.oServerObject(X).lSpawnRequestIdx)
                '            Debug.WriteLine("Region at " & goMsgSys.oServerObject(X).sIPAddress)
                '            For Y As Int32 = 0 To .lSpawnID.GetUpperBound(0)
                '                Debug.WriteLine("  " & .lSpawnID(Y))
                '            Next Y
                '        End With
                '    End If
                'Next X

                'For X As Int32 = 0 To glPlayerUB
                '    If glPlayerIdx(X) > -1 Then
                '        If goPlayer(X).ObjectID < 848 Then
                '            If goPlayer(X).AccountStatus = AccountStatusType.eActiveAccount Then
                '                If goPlayer(X).lLastViewedEnvir < 1 Then
                '                    If goPlayer(X).ObjectID <> 25 Then

                '                        With goPlayer(X)
                '                            .AccountStatus = 99

                '                            If .SaveObject() = False Then
                '                            End If

                '                            Dim yMsg(69) As Byte
                '                            Dim lPos As Int32 = 0
                '                            System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerDetails).CopyTo(yMsg, lPos) : lPos += 2
                '                            System.BitConverter.GetBytes(.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                '                            .PlayerUserName.CopyTo(yMsg, lPos) : lPos += 20
                '                            .PlayerPassword.CopyTo(yMsg, lPos) : lPos += 20
                '                            System.BitConverter.GetBytes(.AccountStatus).CopyTo(yMsg, lPos) : lPos += 4
                '                            .sPlayerNameProper = BytesToString(.PlayerName)
                '                            .sPlayerName = .sPlayerNameProper.ToUpper
                '                            StringToBytes(.sPlayerNameProper).CopyTo(yMsg, lPos) : lPos += 20
                '                            goMsgSys.SendToPrimaryServers(yMsg)
                '                            goMsgSys.ForwardToBackupOperator(yMsg, -1)
                '                        End With

                '                    End If
                '                End If
                '            End If
                '        End If
                '    End If
                'Next X

                'goMsgSys.CheckKeepAliveStatus()
                GeoSpawner.GenerateNewNeighborhood()

                Threading.Thread.Sleep(100)
        End While
	End Sub

    Private Sub CheckTransactions()
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand = Nothing ' OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing ' 'OleDb.OleDbDataReader = Nothing

        Dim blTransExecuted As Int64 = 0

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try

            Dim blLastTransID As Int64 = -1

            sSQL = "SELECT * FROM OperatorTransID"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            If oData.Read = True Then
                If oData("LastTransactionID") Is DBNull.Value = False Then blLastTransID = CLng(oData("LastTransactionID"))
            End If
            oData.Close()
            oData = Nothing
            oComm.Dispose()
            oComm = Nothing

            sSQL = "SELECT * FROM WebsiteTrans WHERE TransactionID > " & blLastTransID & " ORDER BY TransactionID"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)

            Dim blInitialLastTransID As Int64 = blLastTransID

            Try
                While oData.Read = True
                    Dim blTransID As Int64 = CLng(oData("TransactionID"))
                    Dim sPreviousUserName As String = ""
                    If oData("PreviousUserName") Is DBNull.Value = False Then sPreviousUserName = CStr(oData("PreviousUserName"))
                    Dim sNewUserName As String = ""
                    If oData("NewUserName") Is DBNull.Value = False Then sNewUserName = CStr(oData("NewUserName"))
                    Dim sNewPassword As String = ""
                    If oData("NewPassword") Is DBNull.Value = False Then sNewPassword = CStr(oData("NewPassword"))
                    Dim lStatus As Int32 = -1
                    If oData("NewStatus") Is DBNull.Value = False Then lStatus = CInt(oData("NewStatus"))
                    Dim lWebUserID As Int32 = -1
                    If oData("WebUserID") Is DBNull.Value = False Then lWebUserID = CInt(oData("WebUserID"))
                    Dim lBBUserID As Int32 = -1
                    If oData("BBUserID") Is DBNull.Value = False Then lBBUserID = CInt(oData("WebUserID"))
                    Dim lTrackerUserID As Int32 = -1
                    If oData("TrackerUserID") Is DBNull.Value = False Then lTrackerUserID = CInt(oData("TrackerUserID"))

                    lStatus = 1

                    blTransExecuted += 1

                    'DO EXECUTION HERE
                    If UpdatePlayerData(sPreviousUserName, sNewUserName, sNewPassword, lStatus, lWebUserID, lBBUserID, lTrackerUserID, blTransID) = False Then
                        'TODO: Add this to a list somewhere that we could not execute it so that an admin console can see it
                    End If

                    'ok, done updating, set our last trans id
                    blLastTransID = blTransID
                End While
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "While checking Transaction (LastTrans: " & blLastTransID & ") - " & ex.Message)
            End Try

            'Ok, kill our recordset
            oData.Close()
            oData = Nothing
            oComm.Dispose()
            oComm = Nothing

            If blInitialLastTransID <> blLastTransID Then
                LogEvent(LogEventType.Informational, "Transactions Executed: " & blTransExecuted.ToString)

                'Now, write our last trans id out
                sSQL = "UPDATE OperatorTransID SET LastTransactionID = " & blLastTransID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    oComm.Dispose()
                    oComm = Nothing
                    sSQL = "INSERT INTO OperatorTransID (LastTransactionID) VALUES (" & blLastTransID & ")"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oComm.ExecuteNonQuery()
                End If
                oComm.Dispose()
                oComm = Nothing
            End If

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "CheckTransactions: " & ex.Message)
        Finally
            oData = Nothing
            oComm = Nothing
        End Try

    End Sub

    Private Function UpdatePlayerData(ByVal sPrevUserName As String, ByVal sNewUserName As String, ByVal sNewPassword As String, ByVal lStatus As Int32, ByVal lWebUserID As Int32, ByVal lBBUserID As Int32, ByVal lTrackerID As Int32, ByVal blTransID As Int64) As Boolean
        'Ok, first, find our user...
        Dim bNewUser As Boolean = False
        Dim oPlayer As Player = Nothing

        If sPrevUserName Is Nothing OrElse sPrevUserName = "" Then
            bNewUser = True

            'Create/Instantiate the user
            oPlayer = New Player
            With oPlayer
                ReDim .EmpireName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewUserName).CopyTo(.EmpireName, 0)
                ReDim .RaceName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewUserName).CopyTo(.RaceName, 0)

                .lLastViewedEnvir = 0
                .iLastViewedEnvirType = 0
                .blCredits = 0
                .lStartedEnvirID = 0
                .iStartedEnvirTypeID = 0
                .TotalPlayTime = 0
                .ObjTypeID = ObjectType.ePlayer
                .yGender = 0
                .AccountStatus = lStatus
            End With

            Dim yTestUserName(19) As Byte
            System.Text.ASCIIEncoding.ASCII.GetBytes(sNewUserName.ToUpper.Trim).CopyTo(yTestUserName, 0)

            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    With goPlayer(X)
                        Dim bUserNameInUse As Boolean = True
                        Dim bPlayerNameInUse As Boolean = True
                        Dim lPlayerNameUB As Int32 = .PlayerName.GetUpperBound(0)
                        For Y As Int32 = 0 To 19
                            If .PlayerUserName(Y) <> yTestUserName(Y) Then
                                bUserNameInUse = False
                            End If
                            If (Y <= lPlayerNameUB AndAlso .PlayerName(Y) <> yTestUserName(Y)) OrElse (Y > lPlayerNameUB AndAlso yTestUserName(Y) <> 0) Then
                                bPlayerNameInUse = False
                            End If
                            If bPlayerNameInUse = False AndAlso bUserNameInUse = False Then Exit For
                        Next Y

                        If bPlayerNameInUse = True OrElse bUserNameInUse = True Then
                            LogEvent(LogEventType.CriticalError, "WebTrans: UserName/PlayerName in use on an add trans. Ignoring trans. Username: " & sNewUserName & ". TransID: " & blTransID.ToString)
                            Return False
                        End If


                    End With
                End If
            Next X

            LogEvent(LogEventType.Informational, "AddNewUser: " & sNewUserName)
        Else
            Dim yUserName(19) As Byte
            System.Text.ASCIIEncoding.ASCII.GetBytes(sPrevUserName.ToUpper.Trim).CopyTo(yUserName, 0)

            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    With goPlayer(X)
                        Dim bFound As Boolean = True
                        For Y As Int32 = 0 To 19
                            If .PlayerUserName(Y) <> yUserName(Y) Then
                                bFound = False
                                Exit For
                            End If
                        Next Y
                        If bFound = True Then
                            oPlayer = goPlayer(X)

                            LogEvent(LogEventType.Informational, "UpdateUser: " & oPlayer.sPlayerNameProper)
                            Exit For
                        End If
                    End With
                End If
            Next X
        End If
        If oPlayer Is Nothing Then
            LogEvent(LogEventType.Warning, "WebTrans: Unable to find PreviousUserName '" & sPrevUserName & "'. Inserting transaction.")

            'Return False
            bNewUser = True

            'Create/Instantiate the user
            oPlayer = New Player
            With oPlayer
                ReDim .EmpireName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewUserName).CopyTo(.EmpireName, 0)
                ReDim .RaceName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewUserName).CopyTo(.RaceName, 0)

                .lLastViewedEnvir = 0
                .iLastViewedEnvirType = 0
                .blCredits = 0
                .lStartedEnvirID = 0
                .iStartedEnvirTypeID = 0
                .TotalPlayTime = 0
                .ObjTypeID = ObjectType.ePlayer
                .yGender = 0
                .AccountStatus = lStatus
            End With

            Dim yTestUserName(19) As Byte
            System.Text.ASCIIEncoding.ASCII.GetBytes(sNewUserName.ToUpper.Trim).CopyTo(yTestUserName, 0)

            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    With goPlayer(X)
                        Dim bUserNameInUse As Boolean = True
                        Dim bPlayerNameInUse As Boolean = True
                        Dim lPlayerNameUB As Int32 = .PlayerName.GetUpperBound(0)
                        For Y As Int32 = 0 To 19
                            If .PlayerUserName(Y) <> yTestUserName(Y) Then
                                bUserNameInUse = False
                            End If
                            If (Y <= lPlayerNameUB AndAlso .PlayerName(Y) <> yTestUserName(Y)) OrElse (Y > lPlayerNameUB AndAlso yTestUserName(Y) <> 0) Then
                                bPlayerNameInUse = False
                            End If
                            If bPlayerNameInUse = False AndAlso bUserNameInUse = False Then Exit For
                        Next Y

                        If bPlayerNameInUse = True OrElse bUserNameInUse = True Then
                            LogEvent(LogEventType.CriticalError, "WebTrans: UserName/PlayerName in use on an add trans. Ignoring trans. Username: " & sNewUserName & ". TransID: " & blTransID.ToString)
                            Return False
                        End If


                    End With
                End If
            Next X
            If oPlayer Is Nothing Then
                LogEvent(LogEventType.CriticalError, "WebTrans.UpdatePlayerData - Tried to insert user but player object is empty. Username: " & sNewUserName & ". TransID: " & blTransID.ToString)
                Return False
            End If
        End If

        'Ok, now, let's update our player..
        With oPlayer
            ReDim .PlayerUserName(19)
            System.Text.ASCIIEncoding.ASCII.GetBytes(sNewUserName.ToUpper.Trim).CopyTo(.PlayerUserName, 0)
            ReDim .PlayerPassword(19)
            System.Text.ASCIIEncoding.ASCII.GetBytes(sNewPassword.ToUpper.Trim).CopyTo(.PlayerPassword, 0)

            Dim lPrevStatus As Int32 = .AccountStatus
            .AccountStatus = lStatus
            'If bNewUser = True Then .PlayerName = .RaceName
            If bNewUser = True Then
                ReDim .PlayerName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewUserName.Trim).CopyTo(.PlayerName, 0)
            End If

            'TODO: We can store the WebID, BBID and TrackerID

            If .SaveObject() = False Then
                LogEvent(LogEventType.CriticalError, "WebTrans: Unable to save changes.")
                Return False
            End If

            If lStatus = AccountStatusType.eActiveAccount OrElse lStatus = AccountStatusType.eMondelisActive Then
                If lPrevStatus <> lStatus AndAlso lStatus <> AccountStatusType.eActiveAccount AndAlso lStatus <> AccountStatusType.eMondelisActive Then
                    Dim sSQL As String = "UPDATE tblPlayer SET DateAccountWentActive = " & GetDateAsNumber(Now) & " WHERE PlayerID = " & .ObjectID
                    Dim oComm As OleDb.OleDbCommand = Nothing
                    Try
                        oComm = New OleDb.OleDbCommand(sSQL, goCN)
                        If oComm.ExecuteNonQuery() = 0 Then Throw New Exception("No Records Affected with update!")
                    Catch ex As Exception
                        LogEvent(LogEventType.CriticalError, "Could not record DateAccountWentActive for player: " & sNewUserName & ". Reason: " & ex.Message)
                    End Try
                    If oComm Is Nothing = False Then oComm.Dispose()
                    oComm = Nothing
                End If
            ElseIf lPrevStatus = AccountStatusType.eActiveAccount OrElse lPrevStatus = AccountStatusType.eMondelisActive Then
                'ok, set the date/time when the account when inactive
                Dim sSQL As String = "UPDATE tblPlayer SET DateAccountWentInactive = " & GetDateAsNumber(Now) & " WHERE PlayerID = " & .ObjectID
                Dim oComm As OleDb.OleDbCommand = Nothing
                Try
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    If oComm.ExecuteNonQuery() = 0 Then Throw New Exception("No Records Affected with update!")
                Catch ex As Exception
                    LogEvent(LogEventType.CriticalError, "Could not record DateAccountWentInactive for player: " & sNewUserName & ". Reason: " & ex.Message)
                End Try
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing
            End If
            

            If bNewUser = True Then
                Dim bFound As Boolean = False
                For X As Int32 = 0 To glPlayerUB
                    If glPlayerIdx(X) = -1 Then
                        bFound = True
                        goPlayer(X) = oPlayer
                        glPlayerIdx(X) = oPlayer.ObjectID
                        Exit For
                    End If
                Next X
                If bFound = False Then
                    ReDim Preserve glPlayerIdx(glPlayerUB + 1)
                    ReDim Preserve goPlayer(glPlayerUB + 1)
                    goPlayer(glPlayerUB + 1) = oPlayer
                    glPlayerIdx(glPlayerUB + 1) = oPlayer.ObjectID
                    glPlayerUB += 1
                End If
            End If

            'Now, let's send our update to all servers
            Dim yMsg(69) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerDetails).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            .PlayerUserName.CopyTo(yMsg, lPos) : lPos += 20
            .PlayerPassword.CopyTo(yMsg, lPos) : lPos += 20
            System.BitConverter.GetBytes(.AccountStatus).CopyTo(yMsg, lPos) : lPos += 4
            .sPlayerNameProper = BytesToString(.PlayerName)
            .sPlayerName = .sPlayerNameProper.ToUpper
            StringToBytes(.sPlayerNameProper).CopyTo(yMsg, lPos) : lPos += 20
            goMsgSys.SendToPrimaryServers(yMsg)
            goMsgSys.ForwardToBackupOperator(yMsg, -1)

        End With
        Return True

    End Function

    'Private mlSynclockObj(-1) As Int32
    'Private mlSpawnRequestIDGen As Int32 = 1
    'Public Function GetSpawnRequestUniqueID() As Int32
    '    Dim lID As Int32 = -1
    '    SyncLock mlSynclockObj
    '        lID = mlSpawnRequestIDGen
    '        mlSpawnRequestIDGen += 1
    '    End SyncLock
    '    Return lID
    'End Function

    Public Sub SendSupportEmail(ByVal sBody As String, ByVal sSubject As String)
        'TODO: Implement this
        Const sOutIP As String = "10.70.5.51"
        Const sTo As String = "jcampbell@darkskyentertainment.com; mcampbell@darkskyentertainment.com; support@darkskyentertainment.com"

        Try
            'create the mail message object
            Dim oMailMsg As New System.Net.Mail.MailMessage("support@darkskyentertainment.com", sTo)
            oMailMsg.BodyEncoding = System.Text.Encoding.Default
            oMailMsg.Subject = sSubject.Trim()
            oMailMsg.Body = sBody.Trim() & vbCrLf
            oMailMsg.Priority = Net.Mail.MailPriority.High
            oMailMsg.IsBodyHtml = False
            oMailMsg.ReplyTo = New Net.Mail.MailAddress("support@darkskyentertainment.com")

            'create Smtpclient to send the mail message
            Dim oSmtpMail As New Net.Mail.SmtpClient
            oSmtpMail.Host = sOutIP
            oSmtpMail.UseDefaultCredentials = False
            oSmtpMail.Credentials = New Net.NetworkCredential("fleetcommander", "nu4tuibezl@pspos")

            oSmtpMail.Send(oMailMsg)

            'sTo = sTo.ToUpper()

            'set keep alive time stamp for this message 
            'NOTE: expiration 1 hour once it is fully implemented
            'lTimeStamp = CInt(Val(Now.ToString("MMddHHmmss")))
        Catch ex As Exception
            'Message Error
            LogEvent(LogEventType.CriticalError, "Unable to sendsupportmail: " & ex.Message)
        End Try
    End Sub

#Region "  Bad Word Filter  "
    Private msWordChecks() As String = Nothing
    Private mlReplacementLen() As Int32 = Nothing
    Private mlWordCheckUB As Int32 = -1
    Private Sub InitializeWordChecks()
        AddWordCheck("fuck")
        AddWordCheck("bitch")
        AddWordCheck("shit")
        AddWordCheck("cunt")
        AddWordCheck("dick")
        AddWordCheck("fagit")
        AddWordCheck("faggot")
        AddWordCheck("nigger")
        AddWordCheck("pussy")
        AddWordCheck("puss")
        AddWordCheck("cock")
        AddWordCheck("fag")
        AddWordCheck("whore")
        AddWordCheck("dike")
        AddWordCheck("kike")
        AddWordCheck("jew")
    End Sub

    Private Sub AddWordCheck(ByVal sWord As String)
        ReDim Preserve msWordChecks(mlWordCheckUB + 2)
        ReDim Preserve mlReplacementLen(mlWordCheckUB + 2)

        Dim lNewIdx As Int32 = mlWordCheckUB + 1
        Dim lOtherIdx As Int32 = mlWordCheckUB + 2
        mlReplacementLen(lNewIdx) = sWord.Length
        mlReplacementLen(lOtherIdx) = sWord.Length
        msWordChecks(lNewIdx) = "\b"
        msWordChecks(lOtherIdx) = ""

        For X As Int32 = 0 To sWord.Length - 1
            Dim sChr As String = sWord(X).ToString
            Dim sCheckVal As String = "[" & sChr.ToLower & "|" & sChr.ToUpper
            If sChr.ToUpper = "S" Then sCheckVal &= "|$"
            If sChr.ToUpper = "I" Then sCheckVal &= "|!|1||"
            If sChr.ToUpper = "A" Then sCheckVal &= "|@"
            If sChr.ToUpper = "O" Then sCheckVal &= "|0"
            If sChr.ToUpper = "H" Then sCheckVal &= "|4"
            sCheckVal &= "]"
            msWordChecks(lNewIdx) &= sCheckVal & "[\W]?"
            msWordChecks(lOtherIdx) &= sCheckVal
        Next X

        msWordChecks(lOtherIdx) &= "\b"

        mlWordCheckUB += 2
    End Sub

    Public Function FilterBadWords(ByVal sInitial As String) As String

        If mlWordCheckUB = -1 Then InitializeWordChecks()

        For X As Int32 = 0 To msWordChecks.GetUpperBound(0)
            sInitial = System.Text.RegularExpressions.Regex.Replace(sInitial, msWordChecks(X), StrDup(mlReplacementLen(X), "*"), System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace)
        Next X

        Return sInitial
    End Function
#End Region

#Region "  IP Parsing  "
    Public Function ParseForIpAddress(ByVal sIpAddress As String, Optional ByRef bRoutable As Boolean = True) As String
        'bRoutable: Only return routable ip's, not 10.x.x.x or 192.168.x.x or 172.16.x.x ~ 172.31.x.x or 5.0.x.x ~ 5.63.0.0
        'Will return Routable as a first priority
        If sIpAddress.ToLower = sIpAddress.ToUpper Then Return sIpAddress
        Dim sTempIP As String = ""
        Dim lPriority As Integer = 0
        Dim lTempPriority As Integer = -1

        If sIpAddress.ToUpper <> sIpAddress.ToLower Then 'Hostname -> Resolve it
            Dim IPList As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(sIpAddress)
            Dim IPaddr As System.Net.IPAddress
            For Each IPaddr In IPList.AddressList
                If (IPaddr.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork) Then
                    If IsPrivateIP(IPaddr.ToString) = False Then
                        Return IPaddr.ToString
                    Else
                        If lPriority > lTempPriority Then
                            sTempIP = IPaddr.ToString
                            lTempPriority = lPriority
                        End If
                    End If
                End If
            Next
        End If
        If bRoutable = False Then Return sTempIP
        Return ""

    End Function

    Private Function IsPrivateIP(ByVal CheckIP As String, Optional ByRef lPriority As Integer = -1) As Boolean
        Dim Quad() As Byte = CType(System.Net.IPAddress.Parse(CheckIP), System.Net.IPAddress).GetAddressBytes
        Select Case Quad(0)
            Case 10
                lPriority = 30
                Return True
            Case 172
                lPriority = 20
                If Quad(1) >= 16 And Quad(1) <= 31 Then Return True
            Case 5
                lPriority = 10
                If Quad(1) >= 0 And Quad(1) <= 63 Then Return True
            Case 192
                lPriority = 30
                If Quad(1) = 168 Then Return True
        End Select
        Return False
    End Function
#End Region

End Module
