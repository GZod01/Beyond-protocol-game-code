Option Strict On
'here, we will contain all of our global variables
Module GlobalVars
    'Essential to practically everything...
    Public Enum AccountStatusType As Integer
        eInactiveAccount = 0
        eActiveAccount
        eBannedAccount
        eSuspendedAccount
        eOpenBetaAccount

        eTrialAccount = 99
        eMondelisInactive = 100
        eMondelisActive = 101
    End Enum

    Public glCurrentCycle As Int32

    Public gbRunning As Boolean
    Public gfrmDisplayForm As Form1

	Public gblAggressions As Int64
	Public gblMovements As Int64
	Public gblEngineLoops As Int64

	Public gl_REL_TINY_MAX As Int32

	Public gbl_Full_Cycles As Int64 = 0
	Public gbl_Full_Cycle_Duration As Int64 = 0

	Private mlLastCPUpdate As Int32

	Public goMsgMonitor As MsgMonitor
	Public Const gb_MONITOR_MSGS As Boolean = True

	Public goMissileMgr As MissileMgr

	Public glBoxOperatorID As Int32
	Public gsOperatorIP As String
    Public glOperatorPort As Int32
    Public gsExternalIP As String
    Public glExternalPort As Int32 = 7710

	Public goSWMvmt As New Stopwatch
	Public goSWCombat As New Stopwatch
	Public goSWMissile As New Stopwatch
	Public goSWCP As New Stopwatch
	Public goSWBomb As New Stopwatch
	Public goSWWarTrack As New Stopwatch

    Public gl_FORCE_AGGRESSION_THRESHOLD As Int32 = 10      '10 cycles 300ms

    'Private mlLastForceInfDistribute As Int32 = 0

	Public Sub Main()
		goMissileMgr = New MissileMgr()

		'the entry point for the program
		Dim swMain As Stopwatch = New Stopwatch
		Dim lAsyncStartIndex As Int32 = 0
		Dim bHadToWait As Boolean
		Dim lPrevAsyncBurst As Int32 = 0
		Dim bSentAsync As Boolean = False

		Dim lPreviousWarTrackCheck As Int32 = 0

		Dim lPreviousReturnToParentCheck As Int32 = 0

		Dim lPreviousPackRegisters As Int32 = 0

		Form.CheckForIllegalCrossThreadCalls = False
		gfrmDisplayForm = New Form1()
		gfrmDisplayForm.Show()

		gfrmDisplayForm.AddEventLine("Parsing Command Line...")
		If ParseCommandLine() = False Then
			gfrmDisplayForm.AddEventLine("Unable to parse command line, exiting...")
			Threading.Thread.Sleep(10000)
			Return
		End If

        goMsgMonitor = New MsgMonitor()

        gfrmDisplayForm.AddEventLine("Loading Model Def Data...")
        ModelDefs.LoadAllModelDefs()

		gfrmDisplayForm.AddEventLine("Killing previous INI data...")
		ClearDomains()
		gfrmDisplayForm.AddEventLine("Initializing Message System...")
		goMsgSys = New MsgSystem()
		gfrmDisplayForm.AddEventLine("Connecting to Operator...")
		If goMsgSys.ConnectToOperator() = False Then
			Application.Exit()
		End If

		While goMsgSys.bReceivedAssignment = False
			Threading.Thread.Sleep(10)
			Application.DoEvents()
		End While

		If NewInitializeServer() = False Then
			MsgBox("Abnormal application startup. Closing application.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Closing")
			Application.Exit()
		End If

		While goMsgSys.AcceptingClients = False
			GeoSpawner.CheckLoadSystemQueue()
			Application.DoEvents()
			Threading.Thread.Sleep(1)
		End While
		gfrmDisplayForm.AddEventLine("Server Initialized!!!")

		'Set everything so that when the game engine starts all units will forcefully aggress
		Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB
            If glEntityIdx(X) > 0 Then
                Dim oEntity As Epica_Entity = goEntity(X)
                If oEntity Is Nothing = False Then
                    If oEntity.ParentEnvir.PotentialAggression = True Then
                        oEntity.bForceAggressionTest = True
                        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, X, oEntity.ObjectID) 'AddEntityMoving(X, oEntity.ObjectID)
                    End If
                End If
            End If
		Next X

		'Now, call handlemovement 1 time
		HandleMovement()
		'Now, clear our movement list
		ResetMovementRegisters()
		'Now, go through our entities, anyone moving OR have a force aggression test flag set, we need to add
		lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB
            If glEntityIdx(X) > 0 Then
                Dim oEntity As Epica_Entity = goEntity(X)
                If oEntity Is Nothing = False Then
                    If (oEntity.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, X, oEntity.ObjectID) 'AddEntityMoving(X, oEntity.ObjectID)
                    ElseIf oEntity.bForceAggressionTest = True Then
                        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, X, oEntity.ObjectID) 'AddEntityMoving(X, oEntity.ObjectID)
                    End If
                End If
            End If
		Next X

		'Indicate to the operator that we are ready
		goMsgSys.SendReadyStateToOperator()

		Dim oSWTimeKeeper As Stopwatch = Stopwatch.StartNew()
        'Dim bSetLastForceInf As Boolean = False
		'at this point, we need to start the real engine...
		gbRunning = True
		While gbRunning
			bHadToWait = False

			swMain.Reset()
			swMain.Start()

			'Do our stuff here
			goSWMvmt.Start()
			HandleMovement()
			goSWMvmt.Stop()

			goSWCombat.Start()
			HandleCombat()
			goSWCombat.Stop()

			goSWMissile.Start()
			goMissileMgr.HandleMissileMovement()
			goSWMissile.Stop()

			goSWBomb.Start()
			HandleBombardment()
			goSWBomb.Stop()

			goSWCP.Start()
			UpdateCommandPoints()
			goSWCP.Stop()

			gblEngineLoops += 1

			goSWWarTrack.Start()
			If glCurrentCycle - lPreviousWarTrackCheck > 900 Then	'once every 30 seconds
				lPreviousWarTrackCheck = glCurrentCycle
				For X As Int32 = 0 To glEnvirUB
					With goEnvirs(X)
						Dim bFoundWar As Boolean = False
						For Y As Int32 = 0 To .lWarUB
							If .lWarIdx(Y) <> -1 Then
								If glCurrentCycle - .oWars(Y).lPreviousMsgSend > WarTracker.ml_WAR_OVER_DELAY Then
									Dim yMsg() As Byte = .oWars(Y).GetNewsItemMsg(.ObjectID, .ObjTypeID)
									If yMsg Is Nothing = False Then
										goMsgSys.SendToPrimary(yMsg)
									End If
									If .oWars(Y).bExpired = True Then
										.lWarIdx(Y) = -1
										.oWars(Y) = Nothing
									End If
								End If
								bFoundWar = True
							End If
						Next Y
						If .lWarUB <> -1 AndAlso bFoundWar = False Then .lWarUB = -1
					End With
				Next X
			End If
			goSWWarTrack.Stop()

			If glCurrentCycle - lPreviousReturnToParentCheck > 9000 Then		'5 minutes
				lPreviousReturnToParentCheck = glCurrentCycle

				lCurUB = -1
				If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
				For X As Int32 = 0 To lCurUB
                    If glEntityIdx(X) > 0 Then
                        Dim oEntity As Epica_Entity = goEntity(X)
                        If oEntity Is Nothing = False Then
                            If oEntity.lLaunchedFromID <> Int32.MinValue AndAlso oEntity.iLaunchedFromTypeID <> Int16.MinValue Then
                                If glCurrentCycle - oEntity.lLastEngagement > 5400 Then     '3 minutes

                                    'Return me home...
                                    goMsgSys.SendAIDockRequestToPathfinding(oEntity, oEntity.lLaunchedFromID, oEntity.iLaunchedFromTypeID)

                                    'Clear out my settings so I am not queried again
                                    oEntity.lLaunchedFromID = Int32.MinValue
                                    oEntity.iLaunchedFromTypeID = Int16.MinValue
                                End If
                            End If
                        End If
                    End If
				Next X

			End If

            'bSetLastForceInf = False
            'If glCurrentCycle - mlLastForceInfDistribute > 30 Then
            '    bSetLastForceInf = True
            '    lCurUB = -1
            '    If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
            '    For X As Int32 = 0 To lCurUB
            '        If glEntityIdx(X) > 0 Then
            '            Dim oEntity As Epica_Entity = goEntity(X)
            '            If oEntity Is Nothing = False AndAlso oEntity.bHasWarpointsToDistribute = True Then
            '                oEntity.DistributeWarpoints(True)
            '            End If
            '        End If
            '    Next X
            'End If

			'gfrmDisplayForm.RefreshUnitLabel()
			'If glCurrentCycle Mod 60 = 0 Then gfrmDisplayForm.RefreshUnitLabel()
			'If gbFastRefreshInterval = True Then
			'    gfrmDisplayForm.RefreshUnitLabel()
			'ElseIf glCurrentCycle Mod 60 = 0 Then
			'    gfrmDisplayForm.RefreshUnitLabel()
			'End If

			'always allow the app to do events here
			Application.DoEvents()

			GeoSpawner.CheckLoadSystemQueue()

			If glCurrentCycle - lPreviousPackRegisters > 9000 Then
				lPreviousPackRegisters = glCurrentCycle
                SyncLockMovementRegisters(MovementCommand.PackMovementRegisters, -1, -1)
            Else
                ClearEnvirGrids()
            End If

			'now, wait for the next cycle if necessary and do events
			bSentAsync = False
			While swMain.ElapsedMilliseconds < 30
				If bHadToWait = False Then
					bHadToWait = True
					Application.DoEvents()
					Threading.Thread.Sleep(0)			'NEW
				Else
					'Now, see if anything needs updated 
					If glCurrentCycle - lPrevAsyncBurst <> 0 Then
						Try
							Dim bFound As Boolean = False
							Dim lStartIdx As Int32 = lAsyncStartIndex
							lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
							For X As Int32 = lStartIdx To lCurUB
                                If glEntityIdx(X) > 0 AndAlso goEntity(X).lUpdatePrimaryWithHPRequest <> Int32.MinValue Then 'AndAlso ((glCurrentCycle - 30) > goEntity(X).lUpdatePrimaryWithHPRequest) Then
                                    bFound = True
                                    goEntity(X).SendAsyncUpdateToPrimary()
                                    goEntity(X).lUpdatePrimaryWithHPRequest = Int32.MinValue
                                    lAsyncStartIndex = X + 1
                                    bSentAsync = True
                                    Exit For
                                End If
							Next X
							If bFound = False Then
								lAsyncStartIndex = 0
								lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
								For X As Int32 = 0 To lCurUB
                                    If glEntityIdx(X) > 0 Then goEntity(X).lUpdatePrimaryWithHPRequest = glCurrentCycle
								Next X
								bSentAsync = True
							End If

							'regardless, set our prevasync
							lPrevAsyncBurst = glCurrentCycle
						Catch
						End Try
					Else
						Application.DoEvents()
                        Threading.Thread.Sleep(0)       'NEW
					End If
				End If
			End While
			'If bSentAsync = True Then lPrevAsyncBurst = glCurrentCycle

			If bHadToWait = False Then
				gbl_Full_Cycles += 1
				'gfrmDisplayForm.AddEventLine("Full Cycle Duration: " & swMain.ElapsedMilliseconds.ToString)
				gbl_Full_Cycle_Duration += swMain.ElapsedMilliseconds
			End If

			swMain.Stop()
            If mbSingleStep = True OrElse glCurrentCycle < 7 Then glCurrentCycle += 1 Else glCurrentCycle = CInt(oSWTimeKeeper.ElapsedMilliseconds \ 30)

            'If bSetLastForceInf = True Then mlLastForceInfDistribute = glCurrentCycle
		End While

		swMain.Reset()
		swMain = Nothing

		If goMsgSys Is Nothing = False Then
			goMsgSys.AcceptingClients = False
			goMsgSys.CloseAllConnections()
		End If
		goMsgSys = Nothing
		gfrmDisplayForm = Nothing
		'Application.Exit()
	End Sub
    Private mbSingleStep As Boolean = False

    Private Function ParseCommandLine() As Boolean
        Try
            Dim sValue As String = My.Application.CommandLineArgs(0)
            If IsNumeric(sValue) = False Then
                MsgBox("Unable to determine Box Operator ID in Command Line")
                Return False
            End If
            glBoxOperatorID = CInt(Val(sValue))

            gsOperatorIP = My.Application.CommandLineArgs(1)
            glOperatorPort = CInt(Val(My.Application.CommandLineArgs(2)))
            gsExternalIP = My.Application.CommandLineArgs(3)
            glExternalPort = CInt(Val(My.Application.CommandLineArgs(4)))

            Return True
        Catch ex As Exception
            MsgBox("ParseCommandLine: " & ex.Message)
            Return False
        End Try
    End Function

	Public Function NewInitializeServer() As Boolean
		Dim bRes As Boolean = False

		GeoSpawner.LoadInitialAssignment()

		gfrmDisplayForm.AddEventLine("Initializing Environment Grids")
		InitializeEnvironmentGrids()

		gfrmDisplayForm.AddEventLine("Connecting to Primary...")
		If goMsgSys.ConnectToPrimary() = True Then
			gfrmDisplayForm.AddEventLine("Connected! Loading Static Geography...")
			If GeoSpawner.LoadStaticGeo() = True Then
				gfrmDisplayForm.AddEventLine("Static Geography Loaded! Connecting to Pathfinding Server...")
				If goMsgSys.RequestPathfindingInfo() AndAlso goMsgSys.ConnectToPathfinding() Then
					gfrmDisplayForm.AddEventLine("Connected! Getting Player Data ...")
					If goMsgSys.GetPlayerListFromPrimary() = True Then
						bRes = True
					Else
						gfrmDisplayForm.AddEventLine("Could not load player data...")
					End If
				End If
			Else
				gfrmDisplayForm.AddEventLine("Could not load static geo...")
			End If
		End If

		Return bRes

	End Function

	'   Public Function InitializeServer() As Boolean
	'       Dim bRes As Boolean = False
	'       Dim X As Int32

	'       gfrmDisplayForm.AddEventLine("Connecting to Primary...")
	'       If goMsgSys.ConnectToPrimary() = True Then
	'           gfrmDisplayForm.AddEventLine("Connected! Loading Geography Objects...")
	'           'Ok, let's do this... here's how this works:
	'           '   2) Load the Grandaddy Object Geography and All Children Geography Objects 
	'           If LoadGeographyObjects() Then
	'               gfrmDisplayForm.AddEventLine("Geography Loaded! Getting Player List...")
	'               '   3) Request the Player List and Player Rels from the Primary
	'               If goMsgSys.GetPlayerListFromPrimary() Then
	'                   gfrmDisplayForm.AddEventLine("Got Player List! Filling Environment Objects...")
	'                   '   4) Go through each object in the geography and request a list of children from the Primary
	'                   '   5) Allocate storage space for all of our objects we have received (and then some)
	'                   For X = 0 To glEnvirUB
	'                       gfrmDisplayForm.AddEventLine("Requesting Environment Objects for " & goEnvirs(X).ObjectID & " of " & goEnvirs(X).ObjTypeID & "... ")
	'                       'send the request to the server... the server will respond asynch...
	'                       goMsgSys.SendRequestEnvirObjs(goEnvirs(X).ObjectID, goEnvirs(X).ObjTypeID)
	'                   Next X

	'                   '   6) Request information from Primary regarding Pathfinding Server (Synchronously)
	'                   '   7) Connect to the Pathfinding Server
	'                   If goMsgSys.RequestPathfindingInfo() AndAlso goMsgSys.ConnectToPathfinding() Then

	'                       '8a) Mark all player objects that started on planets in my domain start position used
	'                       For X = 0 To glPlayerUB
	'                           If glPlayerIdx(X) <> -1 Then
	'                               If goPlayers(X).lStartEnvirID > 0 AndAlso goPlayers(X).lStartLocX <> Int32.MinValue AndAlso goPlayers(X).lStartLocZ <> Int32.MinValue Then
	'                                   For Y As Int32 = 0 To glEnvirUB
	'                                       If goEnvirs(Y).ObjTypeID = ObjectType.ePlanet AndAlso goEnvirs(Y).ObjectID = goPlayers(X).lStartEnvirID Then
	'                                           CType(goEnvirs(Y).oGeoObject, Planet).SetPlayerStartLocationMarked(goPlayers(X).lStartLocX, goPlayers(X).lStartLocZ)
	'                                           CType(goEnvirs(Y).oGeoObject, Planet).SetPirateStartLocationMarked(goPlayers(X).lPirateStartLocX, goPlayers(X).lPirateStartLocZ)
	'                                           Exit For
	'                                       End If
	'                                   Next Y
	'                               End If
	'                           End If
	'                       Next X

	'                       '   8) Send My Domain Information (what I control) to the Pathfinding Server
	'                       For X = 0 To glEnvirUB
	'                           Call goMsgSys.SendRegisterDomainMsg(goEnvirs(X).ObjectID, goEnvirs(X).ObjTypeID)
	'                       Next X

	'                       'And tell the Primary that I am ready...
	'                       goMsgSys.SendServerReady()

	'                       '   9) Begin listening for Clients (they will not connect if the Primary is out)
	'                       goMsgSys.AcceptingClients = True

	'                       'return true
	'                       bRes = True
	'                   End If
	'               End If
	'           Else
	'               MsgBox("Unable to load domain's domain.")
	'               Return False
	'           End If
	'       End If

	'       'now load our arrays
	'       gfrmDisplayForm.AddEventLine("Initializing Environment Grids")
	'       InitializeEnvironmentGrids()

	'       Return bRes
	'End Function

	'   Public Function LoadGeographyObjects() As Boolean
	'       Dim X As Int32 = 1
	'       Dim oINI As InitFile = New InitFile()
	'       Dim iTypeID As Int16
	'       Dim lID As Int32
	'       Dim lIdx As Int32
	'       Dim lPlanet As Int32

	'       gfrmDisplayForm.AddEventLine("Getting Star Types...")
	'       'Get the star types
	'       If goMsgSys.GetStarTypes() = False Then Return False
	'       gfrmDisplayForm.AddEventLine("Getting Galaxy and Systems...")
	'       'get the galaxy and systems
	'       If goMsgSys.GetGalaxyAndSystems() = False Then Return False

	'       Do
	'           lID = CInt(Val(oINI.GetString("DOMAIN", "ID" & X, "-1")))
	'           iTypeID = CShort(Val(oINI.GetString("DOMAIN", "TypeID" & X, "-1")))
	'           gfrmDisplayForm.AddEventLine("Acquiring Domain " & lID & " of " & iTypeID & "...")

	'		If lID <> -1 AndAlso iTypeID <> -1 Then
	'			'Ok, got a domain I need to grab, Now, when grabbing the domains, we get all of the domain's details...
	'			'  all the way down to the smallest child... so, if my domain type id is galaxy, then I get the galaxy
	'			'  the systems, the planets, everything, under the galaxy
	'			Select Case iTypeID
	'				Case ObjectType.eGalaxy
	'					'in charge of the galaxy, so get all of the systems
	'					For lIdx = 0 To goGalaxy.mlSystemUB
	'						goGalaxy.CurrentSystemIdx = lIdx
	'						gfrmDisplayForm.AddEventLine("Getting System Details for " & goGalaxy.moSystems(lIdx).ObjectID & "...")
	'						If goMsgSys.GetSystemDetails(goGalaxy.moSystems(lIdx).ObjectID) = False Then Return False

	'						gfrmDisplayForm.AddEventLine("Creating Environment for System...")
	'						'Create the environment for this system...
	'						glEnvirUB += 1
	'						ReDim Preserve goEnvirs(glEnvirUB)
	'						goEnvirs(glEnvirUB) = New Envir()
	'						With goEnvirs(glEnvirUB)
	'							.ObjectID = goGalaxy.moSystems(lIdx).ObjectID
	'							.ObjTypeID = goGalaxy.moSystems(lIdx).ObjTypeID
	'							.oGeoObject = goGalaxy.moSystems(lIdx)
	'							goGalaxy.moSystems(lIdx).EnvirIdx = glEnvirUB
	'							.SetEnvirGridValues()
	'						End With

	'						gfrmDisplayForm.AddEventLine("Creating Planet Environments...")
	'						'Now, go thru the planets and create the environments
	'						For lPlanet = 0 To goGalaxy.moSystems(lIdx).PlanetUB
	'							glEnvirUB += 1
	'							ReDim Preserve goEnvirs(glEnvirUB)
	'							goEnvirs(glEnvirUB) = New Envir()
	'							With goEnvirs(glEnvirUB)
	'								.ObjectID = goGalaxy.moSystems(lIdx).moPlanets(lPlanet).ObjectID
	'								.ObjTypeID = goGalaxy.moSystems(lIdx).moPlanets(lPlanet).ObjTypeID
	'								.oGeoObject = goGalaxy.moSystems(lIdx).moPlanets(lPlanet)
	'								goGalaxy.moSystems(lIdx).moPlanets(lPlanet).EnvirIdx = glEnvirUB
	'								.SetEnvirGridValues()
	'							End With
	'						Next lPlanet
	'					Next lIdx

	'				Case ObjectType.eSolarSystem
	'					'in charge of a solar system
	'					For lIdx = 0 To goGalaxy.mlSystemUB
	'						If goGalaxy.moSystems(lIdx).ObjectID = lID Then
	'							goGalaxy.CurrentSystemIdx = lIdx
	'							gfrmDisplayForm.AddEventLine("Getting System Details for " & goGalaxy.moSystems(lIdx).ObjectID & "...")
	'							If goMsgSys.GetSystemDetails(goGalaxy.moSystems(lIdx).ObjectID) = False Then Return False

	'							gfrmDisplayForm.AddEventLine("Creating Environment for System...")
	'							'Create the environment for this system...
	'							glEnvirUB += 1
	'							ReDim Preserve goEnvirs(glEnvirUB)
	'							goEnvirs(glEnvirUB) = New Envir()
	'							With goEnvirs(glEnvirUB)
	'								.ObjectID = goGalaxy.moSystems(lIdx).ObjectID
	'								.ObjTypeID = goGalaxy.moSystems(lIdx).ObjTypeID
	'								.oGeoObject = goGalaxy.moSystems(lIdx)
	'								goGalaxy.moSystems(lIdx).EnvirIdx = glEnvirUB
	'								.SetEnvirGridValues()
	'							End With

	'							gfrmDisplayForm.AddEventLine("Creating Planet Environments...")
	'							'Now, go thru the planets and create the environments
	'							For lPlanet = 0 To goGalaxy.moSystems(lIdx).PlanetUB
	'								glEnvirUB += 1
	'								ReDim Preserve goEnvirs(glEnvirUB)
	'								goEnvirs(glEnvirUB) = New Envir()
	'								With goEnvirs(glEnvirUB)
	'									.ObjectID = goGalaxy.moSystems(lIdx).moPlanets(lPlanet).ObjectID
	'									.ObjTypeID = goGalaxy.moSystems(lIdx).moPlanets(lPlanet).ObjTypeID
	'									.oGeoObject = goGalaxy.moSystems(lIdx).moPlanets(lPlanet)
	'									goGalaxy.moSystems(lIdx).moPlanets(lPlanet).EnvirIdx = glEnvirUB
	'									.SetEnvirGridValues()
	'								End With
	'							Next lPlanet

	'							Exit For
	'						End If
	'					Next lIdx

	'				Case Else
	'					'the domain cannot be in charge of anything OTHER than a system or galaxy
	'					Return False
	'			End Select
	'		ElseIf X = 1 Then
	'			Return False
	'		Else
	'			Exit Do
	'		End If

	'           X += 1
	'       Loop

	'       Return True
    '   End Function
    Public Sub NewUpdateCommandPoints()
        On Error Resume Next
        If glCurrentCycle - mlLastCPUpdate > 150 Then
            mlLastCPUpdate = glCurrentCycle
            For X As Int32 = 0 To glEnvirUB
                goEnvirs(X).ResetAllPlayerCP()
            Next X
            Dim lCurUB As Int32 = -1
            If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
            For X As Int32 = 0 To glInstanceUB
                If glInstanceIdx(X) > -1 Then
                    goInstances(X).ResetAllPlayerCP()
                End If
            Next X

            lCurUB = -1
            If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityIdx.GetUpperBound(0), glEntityUB)
            For X As Int32 = 0 To lCurUB
                If glEntityIdx(X) > 0 Then
                    Dim oEntity As Epica_Entity = goEntity(X)
                    If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eUnit Then
                        If oEntity.ParentEnvir Is Nothing = False AndAlso oEntity.Owner Is Nothing = False Then
                            oEntity.ParentEnvir.AdjustPlayerCommandPoints(oEntity.lOwnerID, oEntity.CPUsage + oEntity.Owner.BadWarDecCPIncrease)
                        End If
                    End If
                End If
            Next X

            For X As Int32 = 0 To glEnvirUB
                goEnvirs(X).DoCommandPointUpdates()
            Next X
            lCurUB = -1
            If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
            For X As Int32 = 0 To glInstanceUB
                If glInstanceIdx(X) > -1 Then
                    goInstances(X).DoCommandPointUpdates()
                End If
            Next X
        End If
    End Sub

    Public Sub UpdateCommandPoints()
        On Error Resume Next

        If glCurrentCycle - mlLastCPUpdate > 60 Then
            mlLastCPUpdate = glCurrentCycle
            For X As Int32 = 0 To glEnvirUB
                goEnvirs(X).DoCommandPointUpdates()
			Next X

			Dim lCurUB As Int32 = -1
			If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
			For X As Int32 = 0 To glInstanceUB
				If glInstanceIdx(X) > -1 Then
					goInstances(X).DoCommandPointUpdates()
				End If
			Next X
        End If
    End Sub

    Private mlClearEnvirGridIdx As Int32 = 0
    Private Sub ClearEnvirGrids()
        If goEnvirs Is Nothing = False Then
            If mlClearEnvirGridIdx > goEnvirs.GetUpperBound(0) Then
                mlClearEnvirGridIdx = 0
            Else
                If goEnvirs(mlClearEnvirGridIdx).CheckForGridExpiration() = False Then
                    mlClearEnvirGridIdx += 1
                End If
            End If
        End If
    End Sub

#Region " Messaging System "
    'The single Message System for the entire application
    Public goMsgSys As MsgSystem
#End Region

#Region " Constant Expressions "

    'NOTE: IT IS IMPERATIVE THAT THE GRID ITEMS PER ROW NEVER EQUATE OUT TO A DECIMAL NUMBER...
    '  IE: If my environment is 1,000,000 then the Large value cannot leave a remainder
    '  The result of Envir_Size / Large_Per_Row is then used in the Medium size which cannot leave a remainder...
    '  so on and so forth... do so will cause an error in the Upper Bound area

	Public Const gl_MAX_FIGHTER_HULL As Int32 = 300
	Public Const gl_MAX_TANK_HULL As Int32 = 620
	Public Const gl_MIN_CAPITAL_HULL_SIZE As Int32 = 110000

    'For Geometric Calculations
    Public Const gdPi As Single = Math.PI
    Public Const gdHalfPie As Single = gdPi / 2.0F
    Public Const gdPieAndAHalf As Single = gdPi * 1.5F
    Public Const gdTwoPie As Single = gdPi * 2.0F
	Public Const gdDegreePerRad As Single = 180.0F / gdPi
	Public Const gdRadPerDegree As Single = gdPi / 180.0F

#End Region

#Region " Environment Grid and Facing Processing "
    'Public giPlanetRelativeSmall(,,) As Int16
    'Public glPlanetBaseRelTinyX() As Int32
    'Public glPlanetBaseRelTinyZ() As Int32
    'Public gyPlanetDistances(,) As Byte
    'Public glPlanetFacing(,) As Int32

    'Public giSystemRelativeSmall(,,) As Int16
    'Public glSystemBaseRelTinyX() As Int32
    'Public glSystemBaseRelTinyZ() As Int32
    'Public gySystemDistances(,) As Byte
    'Public glSystemFacing(,) As Int32

    Public giRelativeSmall(,,) As Int16
    Public glBaseRelTinyX() As Int32
    Public glBaseRelTinyZ() As Int32
    Public gyDistances(,) As Byte
    Public glFacing(,) As Int32

    Public Sub InitializeEnvironmentGrids()
        Dim lTempUB As Int32
        Dim lGrid As Int32
        Dim lMySmall As Int32
        Dim lTargetSmall As Int32

        Dim lValue As Int32

        Dim lMyX As Int32
        Dim lMyY As Int32
        Dim lTargetX As Int32
        Dim lTargetY As Int32
        Dim lOffsetX As Int32
        Dim lOffsetY As Int32
        Dim lTempX As Int32
        Dim lTempY As Int32
        Dim lCenterIdx As Int32
        Dim lCenterX As Int32
        Dim lCenterY As Int32

        'Do our redims... planets
        lTempUB = (gl_SMALL_PER_ROW * gl_SMALL_PER_ROW) - 1
        ReDim giRelativeSmall(8, lTempUB, lTempUB)        '0 to 8, 0 to 24, 0 to 24
        lTempUB = ((lTempUB + 1) * 9) - 1
        ReDim glBaseRelTinyX(lTempUB)                     '0 to 224
        ReDim glBaseRelTinyZ(lTempUB)                     '0 to 224
        lTempUB = (gl_SMALL_PER_ROW * gl_FINAL_PER_ROW * 3) - 1
        ReDim gyDistances(lTempUB, lTempUB)               '0 to 299, 0 to 299
        ReDim glFacing(lTempUB, lTempUB)
        gl_REL_TINY_MAX = lTempUB

        'Do our redims... systems
        'lTempUB = (gl_SYSTEM_SMALL_PER_ROW * gl_SYSTEM_SMALL_PER_ROW) - 1
        'ReDim giSystemRelativeSmall(8, lTempUB, lTempUB)        '0 to 8, 0 to 624, 0 to 624
        'lTempUB = ((lTempUB + 1) * 9) - 1
        'ReDim glSystemBaseRelTinyX(lTempUB)                     '0 to 5624
        'ReDim glSystemBaseRelTinyZ(lTempUB)                     '0 to 5624
        'lTempUB = (gl_SYSTEM_SMALL_PER_ROW * gl_FINAL_PER_ROW * 3) - 1
        'ReDim gySystemDistances(lTempUB, lTempUB)               '0 to 1499, 0 to 1499
        'ReDim glSystemFacing(lTempUB, lTempUB)

        'Ok, the relatives are a mess... but they work like this...
        '  assume that my current location is the center of this huge array of small sectors
        '  I want to know the target's location from my location in the huge array
        lTempUB = (gl_SMALL_PER_ROW * gl_SMALL_PER_ROW) - 1
		lCenterIdx = ((lTempUB + 1) * 9) \ 2
		lCenterY = gl_SMALL_PER_ROW \ 2
		lCenterX = gl_SMALL_PER_ROW \ 2
        For lGrid = 0 To 8
            For lMySmall = 0 To lTempUB
                'Calculate my locs
				lMyY = lMySmall \ gl_SMALL_PER_ROW
                lMyX = lMySmall - (lMyY * gl_SMALL_PER_ROW)

                'now ,get the offsets
                lOffsetX = lCenterX - lMyX
                lOffsetY = lCenterY - lMyY

                For lTargetSmall = 0 To lTempUB
                    Select Case lGrid
                        Case 0
                            lTempX = gl_SMALL_PER_ROW
                            lTempY = gl_SMALL_PER_ROW
                        Case 1
                            lTempX = 0
                            lTempY = 0
                        Case 2
                            lTempX = gl_SMALL_PER_ROW
                            lTempY = 0
                        Case 3
                            lTempX = gl_SMALL_PER_ROW * 2
                            lTempY = 0
                        Case 4
                            lTempX = 0
                            lTempY = gl_SMALL_PER_ROW
                        Case 5
                            lTempX = gl_SMALL_PER_ROW * 2
                            lTempY = gl_SMALL_PER_ROW
                        Case 6
                            lTempX = 0
                            lTempY = gl_SMALL_PER_ROW * 2
                        Case 7
                            lTempX = gl_SMALL_PER_ROW
                            lTempY = gl_SMALL_PER_ROW * 2
                        Case 8
                            lTempX = gl_SMALL_PER_ROW * 2
                            lTempY = gl_SMALL_PER_ROW * 2
                    End Select

                    'Ok... this is the situation... lGrid is the grid they are located in...
                    ' 0 to 8 (0 being the same grid as me... 1-8 the others
                    ' lMySmall is My Small Sector ID within the 0 grid
                    ' lTargetSmall is the target's small sector ID within lGrid's grid

                    'I need to calculate the location of lTargetSmall
					lTargetY = lTargetSmall \ gl_SMALL_PER_ROW
                    lTargetX = lTargetSmall - (lTargetY * gl_SMALL_PER_ROW)
                    'then, I need to apply the offset to the calculated location
                    lTargetX += lOffsetX
                    lTargetY += lOffsetY
                    'Finally, using the location, I need to get apply it to the Grid's 0,0 location
                    lTargetX += lTempX
                    lTargetY += lTempY

                    '  and then calculate the resulting index
                    If lTargetX > -1 AndAlso lTargetY > -1 Then
                        giRelativeSmall(lGrid, lMySmall, lTargetSmall) = CShort((lTargetY * (gl_SMALL_PER_ROW * 3)) + lTargetX)
                        If giRelativeSmall(lGrid, lMySmall, lTargetSmall) > ((lTempUB + 1) * 9) - 1 Then
                            giRelativeSmall(lGrid, lMySmall, lTargetSmall) = -1
                        End If
                    Else
                        giRelativeSmall(lGrid, lMySmall, lTargetSmall) = -1
                    End If
                Next lTargetSmall
            Next lMySmall
        Next lGrid

        'Now, for the Base Rel Tiny X and Z...
        'This is equal to the 0,0 coordinate in Tiny squares of the Small Square represented... :)
        lTempUB = ((lTempUB + 1) * 9) - 1
        For lTargetSmall = 0 To lTempUB
            'Ok, need to get the X and Y of the value
			lTargetY = lTargetSmall \ (gl_SMALL_PER_ROW * 3)
            lTargetX = lTargetSmall - CInt(lTargetY * (gl_SMALL_PER_ROW * 3))

            glBaseRelTinyX(lTargetSmall) = lTargetX * gl_FINAL_PER_ROW
            glBaseRelTinyZ(lTargetSmall) = lTargetY * gl_FINAL_PER_ROW
        Next lTargetSmall

        'Now, the distance... and facing
        lTempUB = (gl_SMALL_PER_ROW * gl_FINAL_PER_ROW * 3) - 1
        'lCenterX = (lTempUB + 1) / 2
		lCenterX = lTempUB \ 2
        lCenterY = lCenterX
        For lTempY = 0 To lTempUB
            For lTempX = 0 To lTempUB
                lValue = CInt(Math.Sqrt((lTempX - lCenterX) ^ 2 + (lTempY - lCenterY) ^ 2))
                If lValue > 255 Then
                    lValue = 255
                ElseIf lValue < 0 Then
                    lValue = 0
                End If
                gyDistances(lTempX, lTempY) = CByte(lValue)

                glFacing(lTempX, lTempY) = LineAngleDegrees(lCenterX, lCenterY, lTempX, lTempY) * 10
            Next lTempX
        Next lTempY


        ''Now... do it all again but this time, for systems...
        ''Ok, the relatives are a mess... but they work like this...
        ''  assume that my current location is the center of this huge array of small sectors
        ''  I want to know the target's location from my location in the huge array
        'lTempUB = (gl_SYSTEM_SMALL_PER_ROW * gl_SYSTEM_SMALL_PER_ROW) - 1
        'lCenterIdx = Math.Floor(((lTempUB + 1) * 9) / 2)
        'lCenterY = Math.Floor(gl_SYSTEM_SMALL_PER_ROW / 2)
        'lCenterX = Math.Floor(gl_SYSTEM_SMALL_PER_ROW / 2)
        'For lGrid = 0 To 8
        '    For lMySmall = 0 To lTempUB
        '        'Calculate my locs
        '        lMyY = Math.Floor(lMySmall / gl_SYSTEM_SMALL_PER_ROW)
        '        lMyX = lMySmall - (lMyY * gl_SYSTEM_SMALL_PER_ROW)

        '        'now ,get the offsets
        '        lOffsetX = lCenterX - lMyX
        '        lOffsetY = lCenterY - lMyY

        '        For lTargetSmall = 0 To lTempUB
        '            Select Case lGrid
        '                Case 0
        '                    lTempX = gl_SYSTEM_SMALL_PER_ROW
        '                    lTempY = gl_SYSTEM_SMALL_PER_ROW
        '                Case 1
        '                    lTempX = 0
        '                    lTempY = 0
        '                Case 2
        '                    lTempX = gl_SYSTEM_SMALL_PER_ROW
        '                    lTempY = 0
        '                Case 3
        '                    lTempX = gl_SYSTEM_SMALL_PER_ROW * 2
        '                    lTempY = 0
        '                Case 4
        '                    lTempX = 0
        '                    lTempY = gl_SYSTEM_SMALL_PER_ROW
        '                Case 5
        '                    lTempX = gl_SYSTEM_SMALL_PER_ROW * 2
        '                    lTempY = gl_SYSTEM_SMALL_PER_ROW
        '                Case 6
        '                    lTempX = 0
        '                    lTempY = gl_SYSTEM_SMALL_PER_ROW * 2
        '                Case 7
        '                    lTempX = gl_SYSTEM_SMALL_PER_ROW
        '                    lTempY = gl_SYSTEM_SMALL_PER_ROW * 2
        '                Case 8
        '                    lTempX = gl_SYSTEM_SMALL_PER_ROW * 2
        '                    lTempY = gl_SYSTEM_SMALL_PER_ROW * 2
        '            End Select

        '            'Ok... this is the situation... lGrid is the grid they are located in...
        '            ' 0 to 8 (0 being the same grid as me... 1-8 the others
        '            ' lMySmall is My Small Sector ID within the 0 grid
        '            ' lTargetSmall is the target's small sector ID within lGrid's grid

        '            'I need to calculate the location of lTargetSmall
        '            lTargetY = Math.Floor(lTargetSmall / gl_SYSTEM_SMALL_PER_ROW)
        '            lTargetX = lTargetSmall - (lTargetY * gl_SYSTEM_SMALL_PER_ROW)
        '            'then, I need to apply the offset to the calculated location
        '            lTargetX += lOffsetX
        '            lTargetY += lOffsetY
        '            'Finally, using the location, I need to get apply it to the Grid's 0,0 location
        '            lTargetX += lTempX
        '            lTargetY += lTempY

        '            '  and then calculate the resulting index
        '            If lTargetX > -1 AndAlso lTargetY > -1 Then
        '                giSystemRelativeSmall(lGrid, lMySmall, lTargetSmall) = (lTargetY * (gl_SYSTEM_SMALL_PER_ROW * 3)) + lTargetX
        '                If giSystemRelativeSmall(lGrid, lMySmall, lTargetSmall) > ((lTempUB + 1) * 9) - 1 Then
        '                    giSystemRelativeSmall(lGrid, lMySmall, lTargetSmall) = -1
        '                End If
        '            Else
        '                giSystemRelativeSmall(lGrid, lMySmall, lTargetSmall) = -1
        '            End If
        '        Next lTargetSmall
        '    Next lMySmall
        'Next lGrid

        ''Now, for the Base Rel Tiny X and Z...
        ''This is equal to the 0,0 coordinate in Tiny squares of the Small Square represented... :)
        'lTempUB = ((lTempUB + 1) * 9) - 1
        'For lTargetSmall = 0 To lTempUB
        '    'Ok, need to get the X and Y of the value
        '    lTargetY = Math.Floor(lTargetSmall / (gl_SYSTEM_SMALL_PER_ROW * 3))
        '    lTargetX = lTargetSmall - (lTargetY * (gl_SYSTEM_SMALL_PER_ROW * 3))

        '    glSystemBaseRelTinyX(lTargetSmall) = lTargetX * gl_FINAL_PER_ROW
        '    glSystemBaseRelTinyZ(lTargetSmall) = lTargetY * gl_FINAL_PER_ROW
        'Next lTargetSmall

        ''Now, the distance... and facing
        'lTempUB = (gl_SYSTEM_SMALL_PER_ROW * gl_FINAL_PER_ROW * 3) - 1
        ''lCenterX = (lTempUB + 1) / 2
        'lCenterX = Math.Floor(lTempUB / 2)
        'lCenterY = lCenterX
        'For lTempY = 0 To lTempUB
        '    For lTempX = 0 To lTempUB
        '        lValue = Math.Sqrt((lTempX - lCenterX) ^ 2 + (lTempY - lCenterY) ^ 2)
        '        If lValue > 255 Then
        '            lValue = 255
        '        ElseIf lValue < 0 Then
        '            lValue = 0
        '        End If
        '        gySystemDistances(lTempX, lTempY) = lValue

        '        glSystemFacing(lTempX, lTempY) = LineAngleDegrees(lCenterX, lCenterY, lTempX, lTempY) * 10
        '    Next lTempX
        'Next lTempY

    End Sub

    Public Function gyLargeGridArray(ByVal lGridID1 As Int32, ByVal lGridID2 As Int32, ByVal lGridsPerRow As Int32) As Byte
        Select Case lGridID1 - lGridID2
            Case 0
                Return 0
            Case 1
                Return 5
            Case -1
                Return 4
            Case -lGridsPerRow - 1
                Return 1
            Case -lGridsPerRow
                Return 2
            Case -lGridsPerRow + 1
                Return 3
            Case lGridsPerRow - 1
                Return 6
            Case lGridsPerRow
                Return 7
            Case lGridsPerRow + 1
                Return 8
            Case Else
                Return 255
        End Select
    End Function

    Private Function GetSquareIndex(ByVal X As Int32, ByVal Y As Int32, ByVal Z As Int32, ByVal lGridUB As Int32, ByVal lPerRow As Int32) As Byte

        Dim yResult As Byte
        Dim lLocX As Int32
        Dim lLocY As Int32

        If Z <> 0 Then
            'The way this works is... 0 is the large square that "I" exist in...
            '  if Z is another ID (1, 2, 3, 4, 5, 6, 7, 8) then I am only concerned
            '  with the adjacent edge... A majority of the values will be 255
            'X is "My" small sector and Y is "Targets" small sector
            Select Case Z
                Case 1  'Target is Up Left of me... in this case, I only care if X = 0 and Y = UB
					If X = 0 AndAlso Y = lGridUB Then
						yResult = 1	  'at which point, it would be up left of me
					Else
						yResult = 255
					End If
                Case 2  'Target is above main square...
                    'which means, that only the last row matters and only if X is in the first row
                    If X < lPerRow Then
                        'Now, check if Y is in the bottom row...
                        If Y > lGridUB - lPerRow + 1 Then
                            lLocY = (lPerRow - 1) - (lGridUB - Y)
                            If X = 0 Then
                                If lLocY = 0 Then
                                    yResult = 2
                                ElseIf lLocY = 1 Then
                                    yResult = 3
                                Else
                                    yResult = 255
                                End If
                            ElseIf X = lPerRow - 1 Then
                                If lLocY = X - 1 Then
                                    yResult = 1
                                ElseIf lLocY = X Then
                                    yResult = 2
                                Else
                                    yResult = 255
                                End If
                            Else
                                'in the middle...
                                If lLocY = X - 1 Then
                                    yResult = 1
                                ElseIf lLocY = X Then
                                    yResult = 2
                                ElseIf lLocY = X + 1 Then
                                    yResult = 3
                                Else
                                    yResult = 255
                                End If
                            End If
                        Else
                            yResult = 255
                        End If
                    Else
                        yResult = 255
                    End If
                Case 3  'Target UP RIGHT of main square which means we only care if X = Per_Row -1
                    If X = lPerRow - 1 Then
                        'Now, we only care if Y = GridUB - GridRow - 1
                        If Y = lGridUB - lPerRow - 1 Then
                            yResult = 3
                        Else
                            yResult = 255
                        End If
                    Else
                        yResult = 255
                    End If
                Case 4  'Target is LEFT of main square
                    If X Mod lPerRow = 0 Then
                        If X = 0 Then
                            If Y = X + (lPerRow - 1) Then
                                yResult = 4
                            ElseIf Y = X + ((lPerRow * 2) - 1) Then
                                yResult = 6
                            Else
                                yResult = 255
                            End If
                        ElseIf X = lGridUB - lPerRow - 1 Then
                            If Y = X - 1 Then
                                yResult = 1
                            ElseIf Y = X + (lPerRow - 1) Then
                                yResult = 4
                            Else
                                yResult = 255
                            End If

                        Else
                            'Along the edge but not the first or last...
                            If Y = X - 1 Then
                                yResult = 1
                            ElseIf Y = X + (lPerRow - 1) Then
                                yResult = 4
                            ElseIf Y = X + ((lPerRow * 2) - 1) Then
                                yResult = 6
                            Else
                                yResult = 255
                            End If
                        End If
                    Else
                        yResult = 255
                    End If
                Case 5  'the right edge
                    If (X + 1) Mod lPerRow = 0 Then
                        If X = lPerRow - 1 Then
                            If Y = 0 Then
                                yResult = 5
                            ElseIf Y = X + 1 Then
                                yResult = 8
                            Else
                                yResult = 255
                            End If
                        ElseIf X = lGridUB Then
                            If Y = X - (lPerRow) + 1 Then
                                yResult = 5
                            ElseIf Y = X - (lPerRow * 2) + 1 Then
                                yResult = 3
                            Else
                                yResult = 255
                            End If
                        Else
                            If Y = X - (lPerRow * 2) + 1 Then
                                yResult = 3
                            ElseIf Y = X - (lPerRow) + 1 Then
                                yResult = 5
                            ElseIf Y = X + 1 Then
                                yResult = 8
                            Else
                                yResult = 255
                            End If
                        End If
                    Else
                        yResult = 255
                    End If
                Case 6  'the bottom left, which we only care if Y = PerRow - 1
                    If Y = lPerRow - 1 Then
                        'now, we only care if X = UB - PerRow + 1
                        If X = lGridUB - lPerRow + 1 Then
                            yResult = 6
                        Else
                            yResult = 255
                        End If
                    Else
                        yResult = 255
                    End If
                Case 7  'the bottom, which we only care about if Y < PerRow
                    If Y < lPerRow Then
                        If X > lGridUB - lPerRow + 1 Then
                            lLocX = X - (lGridUB - lPerRow + 1)
                            If lLocX = 0 Then   'First one
                                If Y = 0 Then
                                    yResult = 7
                                ElseIf Y = 1 Then
                                    yResult = 8
                                Else
                                    yResult = 255
                                End If
                            ElseIf lLocX = lPerRow - 1 Then   'last one
                                If Y = lPerRow - 1 Then
                                    yResult = 7
                                ElseIf Y = lPerRow - 2 Then
                                    yResult = 6
                                Else
                                    yResult = 255
                                End If
                            Else    'middle
                                If Y = lLocX - 1 Then
                                    yResult = 6
                                ElseIf Y = lLocX Then
                                    yResult = 7
                                ElseIf Y = lLocX + 1 Then
                                    yResult = 8
                                Else
                                    yResult = 255
                                End If
                            End If
                        Else
                            yResult = 255
                        End If
                    Else
                        yResult = 255
                    End If
                Case 8  'the bottom right, which we only care if X = UB
                    If X = lGridUB Then
                        'now, we only care if y = 0
                        If Y = 0 Then
                            yResult = 8
                        Else
                            yResult = 255
                        End If
                    Else
                        yResult = 255
                    End If
            End Select
        Else    'Z <> 0 = False, and therefore, it equals 0, Same Square
            If X = Y Then
                yResult = 0
            ElseIf Y = X - lPerRow - 1 Then
                yResult = 1
            ElseIf Y = X - lPerRow Then
                yResult = 2
            ElseIf Y = X - lPerRow + 1 Then
                yResult = 3
            ElseIf Y = X - 1 Then
                yResult = 4
            ElseIf Y = X + 1 Then
                yResult = 5
            ElseIf Y = X + lPerRow - 1 Then
                yResult = 6
            ElseIf Y = X + lPerRow Then
                yResult = 7
            ElseIf Y = X + lPerRow + 1 Then
                yResult = 8
            Else
                yResult = 255
            End If
        End If

        Return yResult

    End Function

    Public Function WhatSideDoIHit(ByVal lGridIndex As Int32, ByVal lSmallID As Int32, ByVal lTinyX As Int32, ByVal lTinyZ As Int32, ByVal oTarget As Epica_Entity) As Byte
        Dim lFinal As Int32
        Dim yGridResult As Byte
        Dim yResp As Byte = 255

        Dim lRelTinyX As Int32
        Dim lRelTinyZ As Int32

        'Ok, we want to know what side we are going to HIT... so, we do a reverse lookup
        yGridResult = gyLargeGridArray(lGridIndex, oTarget.lGridIndex, oTarget.ParentEnvir.lGridsPerRow)
        If yGridResult <> 255 Then
            lFinal = giRelativeSmall(yGridResult, oTarget.lSmallSectorID, lSmallID)
            If lFinal <> -1 Then
                lRelTinyX = glBaseRelTinyX(lFinal) + lTinyX + (gl_HALF_FINAL_PER_ROW - oTarget.lTinyX)
                If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                    lRelTinyZ = glBaseRelTinyZ(lFinal) + lTinyZ + (gl_HALF_FINAL_PER_ROW - oTarget.lTinyZ)
                    If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                        lFinal = glFacing(lRelTinyX, lRelTinyZ)
                    Else
                        Return 255
                    End If
                Else
                    Return 255
                End If
            Else
                Return 255
            End If

            'now, do an abs subtraction
            lFinal = lFinal - oTarget.LocAngle
            If lFinal < 0 Then
                lFinal = 3600 + (lFinal Mod 3600)
            Else
                lFinal = lFinal Mod 3600
            End If
			lFinal = lFinal \ 10

            yResp = AngleToQuadrant(lFinal)
        Else
            Return 255      '255 should be safe...
        End If

        Return yResp
    End Function

    Public Function WhatSideDoIHit(ByVal oAttacker As Epica_Entity, ByVal oTarget As Epica_Entity) As Byte
        Dim lFinal As Int32
        Dim yGridResult As Byte
        Dim yResp As Byte = 255

        Dim lRelTinyX As Int32
        Dim lRelTinyZ As Int32

        'Ok, we want to know what side we are going to HIT... so, we do a reverse lookup
        'yGridResult = gyLargeGridArray(oTarget.lGridIndex, oAttacker.lGridIndex, oAttacker.ParentEnvir.lGridsPerRow)
        yGridResult = gyLargeGridArray(oAttacker.lGridIndex, oTarget.lGridIndex, oAttacker.ParentEnvir.lGridsPerRow)
        If yGridResult <> 255 Then
            lFinal = giRelativeSmall(yGridResult, oTarget.lSmallSectorID, oAttacker.lSmallSectorID)
            If lFinal <> -1 Then
                lRelTinyX = glBaseRelTinyX(lFinal) + oAttacker.lTinyX + (gl_HALF_FINAL_PER_ROW - oTarget.lTinyX)
                If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                    lRelTinyZ = glBaseRelTinyZ(lFinal) + oAttacker.lTinyZ + (gl_HALF_FINAL_PER_ROW - oTarget.lTinyZ)
                    If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                        lFinal = glFacing(lRelTinyX, lRelTinyZ)
                    Else
                        Return 255
                    End If
                Else
                    Return 255
                End If
            Else
                Return 255
            End If

            'now, do an abs subtraction
            lFinal = lFinal - oTarget.LocAngle
            If lFinal < 0 Then
                lFinal = 3600 + (lFinal Mod 3600)
            Else
                lFinal = lFinal Mod 3600
            End If
			lFinal = lFinal \ 10

            yResp = AngleToQuadrant(lFinal)
        Else
            Return 255      '255 should be safe...
        End If

        Return yResp
    End Function

    Public Function AngleToQuadrant(ByVal lAngle As Int32) As Byte
        'here, we will return the quadrant from the angle
        Select Case lAngle
            Case Is < 45, Is > 315
                Return UnitArcs.eForwardArc
            Case Is < 135
                'Return UnitArcs.eLeftArc
                Return UnitArcs.eRightArc
            Case Is < 225
                Return UnitArcs.eBackArc
            Case Else
                'Return UnitArcs.eRightArc
                Return UnitArcs.eLeftArc
        End Select
    End Function

    Public Function WhatSideCanFire(ByVal oAttacker As Epica_Entity, ByVal lRelTinyX As Int32, ByVal lRelTinyZ As Int32) As Byte
        Dim yResp As Byte = 255
        Dim lFinal As Int32

        'NOTE: lRelTinyX and lRelTinyZ should have already been checked for extents...

        'If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
        '    lRelTinyZ = glBaseRelTinyZ(lFinal) + oTarget.lTinyZ + (gl_HALF_FINAL_PER_ROW - oAttacker.lTinyZ)
        '    If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
        '        lFinal = glFacing(lRelTinyX, lRelTinyZ)
        '    Else
        '        Return 255
        '    End If
        'Else
        '    Return 255
        'End If

        lFinal = glFacing(lRelTinyX, lRelTinyZ)
        lFinal -= oAttacker.LocAngle
        If lFinal < 0 Then
            lFinal = 3600 + (lFinal Mod 3600)
        Else
            lFinal = lFinal Mod 3600
        End If
		lFinal = lFinal \ 10

        yResp = AngleToQuadrant(lFinal)

        Return yResp
    End Function

    Public Function LineAngleDegrees(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Int32
        Dim dDeltaX As Single
        Dim dDeltaY As Single
        Dim dAngle As Single

        dDeltaX = lX2 - lX1
        dDeltaY = lY2 - lY1

        If dDeltaX = 0 Then     'vertical
            If dDeltaY < 0 Then
                dAngle = gdHalfPie
            Else
                dAngle = gdPieAndAHalf
            End If
        ElseIf dDeltaY = 0 Then     'horizontal
            If dDeltaX < 0 Then
                dAngle = gdPi
            Else
                dAngle = 0
            End If
        Else    'angled
            dAngle = CSng(System.Math.Atan(System.Math.Abs(dDeltaY / dDeltaX)))
            'Correct for VB's reversed Y... VB Upper Right is ok... but the other quads are not
			If dDeltaX > -1 AndAlso dDeltaY > -1 Then		'VB Lower Right
				dAngle = gdTwoPie - dAngle
			ElseIf dDeltaX < 0 AndAlso dDeltaY > -1 Then	'VB Lower Left
				dAngle = gdPi + dAngle
			ElseIf dDeltaX < 0 AndAlso dDeltaY < 0 Then		'VB Upper Left
				dAngle = gdPi - dAngle
			End If
        End If

        Return CInt(dAngle * gdDegreePerRad)
    End Function

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef lEndX As Int32, ByRef lEndY As Int32, ByVal fDegree As Single)
        Dim fDX As Single
        Dim fDY As Single
        Dim fRads As Single

		fRads = fDegree * gdRadPerDegree 'CSng(Math.PI / 180.0F)
        fDX = lEndX - lAxisX
		fDY = lEndY - lAxisY

		Dim fCosRads As Single = CSng(Math.Cos(fRads))
		Dim fSinRads As Single = CSng(Math.Sin(fRads))

		lEndX = lAxisX + CInt((fDX * fCosRads) + (fDY * fSinRads))
		lEndY = lAxisY + -CInt((fDX * fSinRads) - (fDY * fCosRads))
    End Sub

#End Region

#Region " Actual Game Object Arrays "
	Public goEntity(-1) As Epica_Entity
	Public glEntityIdx(-1) As Int32
    Public glEntityUB As Int32 = -1
    Private gcolEntityLookup As New Collection

    Public goPlayers() As Player
    Public glPlayerIdx() As Int32
    Public glPlayerUB As Int32 = -1

    Public goEnvirs() As Envir
	Public glEnvirUB As Int32 = -1

	Public goInstances() As Envir
	Public glInstanceUB As Int32 = -1
	Public glInstanceIdx(-1) As Int32

    Public goEntityDefs() As Epica_Entity_Def
    Public glEntityDefIdx() As Int32
    Public glEntityDefUB As Int32 = -1

    Public goGalaxy As Galaxy
    Public goStarTypes() As StarType
    Public glStarTypeUB As Int32 = -1

    Public goBombRequests() As BombardmentRequest
    Public glBombRequestUB As Int32 = -1
	Public gyBombRequestUsed() As Byte

	Public goFormationDef() As FormationDef
	Public glFormationDefIdx() As Int32
	Public glFormationDefUB As Int32 = -1


    Private Function DoLookupEntityFunction(ByVal yType As Byte, ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lIdx As Int32) As Int32
        SyncLock gcolEntityLookup
            Try
                Dim sKey As String = lID & ";" & iTypeID
                Select Case yType
                    Case 0  'lookup
                        If gcolEntityLookup.Contains(sKey) = True Then
                            Return CInt(gcolEntityLookup(sKey))
                        Else
                            Return -1
                        End If
                    Case 1  'add
                        If gcolEntityLookup.Contains(sKey) = True Then Return -1
                        gcolEntityLookup.Add(lIdx, sKey)
                    Case 2  'remove
                        If gcolEntityLookup.Contains(sKey) = True Then gcolEntityLookup.Remove(sKey)
                End Select
            Catch
            End Try
        End SyncLock
        Return -1
    End Function
    Public Function LookupEntity(ByVal lID As Int32, ByVal iTypeID As Int16) As Int32
        Return DoLookupEntityFunction(0, lID, iTypeID, -1)
        'Dim lIdx As Int32 = -1
        'Try
        '    Dim sKey As String = lID & ";" & iTypeID
        '    lIdx = CInt(gcolEntityLookup(sKey))
        'Catch
        'End Try
        'Return lIdx
    End Function
    Public Sub AddLookupEntity(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lIdx As Int32)
        DoLookupEntityFunction(1, lID, iTypeID, lIdx)
        'Try
        '    RemoveLookupEntity(lID, iTypeID)
        '    Dim sKey As String = lID & ";" & iTypeID
        '    gcolEntityLookup.Add(lIdx, sKey)
        'Catch
        'End Try
    End Sub
    Public Sub RemoveLookupEntity(ByVal lID As Int32, ByVal iTypeID As Int16)
        DoLookupEntityFunction(2, lID, iTypeID, -1)
        'Try
        '    Dim sKey As String = lID & ";" & iTypeID
        '    If gcolEntityLookup.Contains(sKey) = True Then gcolEntityLookup.Remove(sKey)
        'Catch
        'End Try
    End Sub

    Public Sub AddBombRequest(ByVal lPlanetID As Int32, ByVal yType As Byte, ByVal lPlayerID As Int32, ByVal lTargetX As Int32, ByVal lTargetZ As Int32)
        Dim X As Int32
        Dim lIdx As Int32 = -1
        Dim lFirstIndex As Int32 = -1

        For X = 0 To glBombRequestUB
            If gyBombRequestUsed(X) > 0 Then
                If goBombRequests(X).lPlayerID = lPlayerID Then
                    If goBombRequests(X).PlanetID = lPlanetID Then
                        lIdx = X
                        Exit For
                    End If
                End If
            ElseIf lFirstIndex = -1 AndAlso gyBombRequestUsed(X) = 0 Then
                lFirstIndex = X
            End If
		Next X
		If lIdx = -1 Then
			If lFirstIndex = -1 Then
				glBombRequestUB += 1
				ReDim Preserve goBombRequests(glBombRequestUB)
				ReDim Preserve gyBombRequestUsed(glBombRequestUB)
				lIdx = glBombRequestUB
			Else
				goBombRequests(lFirstIndex) = Nothing
				lIdx = lFirstIndex
			End If
		End If

        goBombRequests(lIdx) = New BombardmentRequest(lPlanetID, yType)
        If goBombRequests(lIdx).BombRequestValid = True Then
            With goBombRequests(lIdx)
                .lPlayerID = lPlayerID
                .lTargetX = lTargetX
                .lTargetZ = lTargetZ
            End With
            gyBombRequestUsed(lIdx) = 255
        Else
            'TODO: Respond back that the bombardment request was invalid
        End If

    End Sub

    Public Sub ResyncPlayerCPPenalty(ByVal lPlayerID As Int32, ByVal lPenalty As Int32)
        Try
            For X As Int32 = 0 To glEnvirUB
                Dim oEnvir As Envir = goEnvirs(X)
                If oEnvir Is Nothing = False Then
                    oEnvir.ClearPlayersCP(lPlayerID)
                End If
            Next X

            Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If glEntityIdx(X) > 0 Then
                    Dim oEntity As Epica_Entity = goEntity(X)
                    If oEntity Is Nothing = False Then
                        If oEntity.ObjTypeID = ObjectType.eUnit Then
                            If oEntity.lOwnerID = lPlayerID Then
                                oEntity.ParentEnvir.AdjustPlayerCommandPoints(lPlayerID, lPenalty + oEntity.CPUsage)
                            End If
                        End If
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub
#End Region

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
    Private Sub ClearDomains()
        Dim sFilename As String = System.AppDomain.CurrentDomain.BaseDirectory()
        If Right$(sFilename, 1) <> "\" Then sFilename = sFilename & "\"
        sFilename = sFilename & Replace$(System.AppDomain.CurrentDomain.FriendlyName().ToLower, ".exe", ".ini")
        If Exists(sFilename) = True Then Kill(sFilename)
    End Sub
    Public Function Exists(ByVal sFilename As String) As Boolean
        If Trim(sFilename).Length > 0 Then
            On Error Resume Next
            sFilename = Dir$(sFilename, FileAttribute.Archive Or FileAttribute.Directory Or FileAttribute.Hidden Or FileAttribute.Normal Or FileAttribute.ReadOnly Or FileAttribute.System Or FileAttribute.Volume)
			Return Err.Number = 0 AndAlso sFilename.Length > 0
        Else
            Return False
        End If

	End Function
	Public Function GetDateAsNumber(ByVal dtDate As Date) As Int32
		Return CInt(Val(dtDate.ToString("yyMMddHHmm")))
	End Function

    'Private mbRunNewCode As Boolean = False
    'Public Sub RemovePlayerFromAllEnvirs(ByVal lPlayerID As Int32, ByVal lExceptID As Int32, ByVal iExceptTypeID As Int16)
    '    Dim lCurUB As Int32 = -1
    '    Dim lOtherUB As Int32 = -1
    '    If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))

    '    Try
    '        Dim lNotEnvirIdx As Int32 = -1
    '        If mbRunNewCode = True Then
    '            For X As Int32 = 0 To glPlayerUB
    '                If glPlayerIdx(X) = lPlayerID Then
    '                    lNotEnvirIdx = goPlayers(X).lEnvirIdx
    '                    Exit For
    '                End If
    '            Next X
    '        End If

    '        For X As Int32 = 0 To lCurUB
    '            If glInstanceIdx(X) > 0 Then
    '                Dim oInst As Envir = goInstances(X)
    '                If oInst Is Nothing = False Then
    '                    lOtherUB = -1
    '                    If oInst.lPlayersInEnvirIdx Is Nothing = False Then lOtherUB = Math.Min(oInst.lPlayersInEnvirUB, oInst.lPlayersInEnvirIdx.GetUpperBound(0))
    '                    For Y As Int32 = 0 To lOtherUB
    '                        If oInst.lPlayersInEnvirIdx(Y) = lPlayerID Then
    '                            If mbRunNewCode = True AndAlso oInst.oPlayersInEnvir(Y) Is Nothing = False Then oInst.oPlayersInEnvir(Y).lEnvirIdx = -1
    '                            oInst.lPlayersInEnvirIdx(Y) = -1
    '                            oInst.oPlayersInEnvir(Y) = Nothing
    '                        End If
    '                    Next Y
    '                End If
    '            End If
    '        Next X
    '        lCurUB = glEnvirUB
    '        For X As Int32 = 0 To lCurUB
    '            'MSC 03/17/2009 - removed the remarked out code and added the proper code...
    '            If goEnvirs(X) Is Nothing = False AndAlso (goEnvirs(X).ObjectID <> lExceptID OrElse goEnvirs(X).ObjTypeID <> iExceptTypeID) Then
    '                If X = lNotEnvirIdx Then Continue For

    '                lOtherUB = -1
    '                If goEnvirs(X).lPlayersInEnvirIdx Is Nothing = False Then lOtherUB = Math.Min(goEnvirs(X).lPlayersInEnvirIdx.GetUpperBound(0), goEnvirs(X).lPlayersInEnvirUB)
    '                For Y As Int32 = 0 To lOtherUB
    '                    If goEnvirs(X).lPlayersInEnvirIdx(Y) = lPlayerID Then
    '                        If mbRunNewCode = True AndAlso goEnvirs(X).oPlayersInEnvir(Y) Is Nothing = False Then goEnvirs(X).oPlayersInEnvir(Y).lEnvirIdx = -1
    '                        goEnvirs(X).lPlayersInEnvirIdx(Y) = -1
    '                        goEnvirs(X).oPlayersInEnvir(Y) = Nothing
    '                    End If
    '                Next Y
    '            End If
    '        Next X


    '    Catch ex As Exception
    '        gfrmDisplayForm.AddEventLine("RemovePlayerFromAllEnvirs: " & ex.Message)
    '    End Try

    'End Sub
    Public Sub RemovePlayerFromAllEnvirs(ByVal lPlayerID As Int32, ByVal lExceptID As Int32, ByVal iExceptTypeID As Int16)
        Dim lCurUB As Int32 = -1
        Dim lOtherUB As Int32 = -1
        If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
        Try
            Dim lPUB As Int32 = -1
            If glPlayerIdx Is Nothing = False Then lPUB = Math.Min(glPlayerUB, glPlayerIdx.GetUpperBound(0))
            Dim oPlayer As Player = Nothing
            For X As Int32 = 0 To lPUB
                If glPlayerIdx(X) = lPlayerID Then
                    oPlayer = goPlayers(X)
                    Exit For
                End If
            Next X
            If oPlayer Is Nothing Then Return

            For X As Int32 = 0 To lCurUB
                If glInstanceIdx(X) > 0 Then
                    Dim oInst As Envir = goInstances(X)
                    If oInst Is Nothing = False Then
                        oInst.DoEnvirPlayerChange(Envir.eyChangePlayerEnvirCode.eRemoveFromEnvir, oPlayer)
                    End If
                End If
            Next X
            lCurUB = glEnvirUB
            For X As Int32 = 0 To lCurUB
                'MSC 03/17/2009 - removed the remarked out code and added the proper code...
                If goEnvirs(X) Is Nothing = False AndAlso (goEnvirs(X).ObjectID <> lExceptID OrElse goEnvirs(X).ObjTypeID <> iExceptTypeID) Then
                    goEnvirs(X).DoEnvirPlayerChange(Envir.eyChangePlayerEnvirCode.eRemoveFromEnvir, oPlayer)
                End If
            Next X


        Catch ex As Exception
            gfrmDisplayForm.AddEventLine("RemovePlayerFromAllEnvirs: " & ex.Message)
        End Try
    End Sub

End Module

'Mimics Vector3 of DirectX but is 'leaner' (actually it is about 8 times faster)
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