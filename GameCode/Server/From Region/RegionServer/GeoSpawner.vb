Option Strict On

Public Class GeoSpawner

	Private Shared mbStaticGeoLoaded As Boolean = False

	'Our System Loading Queue...
	Private Shared mlLoadSystemQueueID(-1) As Int32
	Private Shared mbInLoadSystem As Boolean = False
	Private Shared myServerReadySent As Byte = 0			'0 indicates we've not had any systems added, 1 indicates that systems have been added but server ready hasn't been sent, all else indicates server ready sent

	'The system currently being loaded
	Private Shared moCurrentSystem As SolarSystem = Nothing

	'For received environment objects management
	Private Shared mlLastReceivedID As Int32
	Private Shared miLastReceivedTypeID As Int16

	Public Shared Function LoadStaticGeo() As Boolean
		If mbStaticGeoLoaded = True Then Return True
		gfrmDisplayForm.AddEventLine("Getting Star Types...")
		'Get the star types
		If goMsgSys.GetStarTypes() = False Then Return False
		gfrmDisplayForm.AddEventLine("Getting Galaxy and Systems...")
		'get the galaxy and systems
		If goMsgSys.GetGalaxyAndSystems() = False Then Return False

		mbStaticGeoLoaded = True
		Return True
	End Function

	Public Shared Sub LoadInitialAssignment()
		Dim oINI As New InitFile
		Dim bDone As Boolean = False
		Dim X As Int32 = 1
		While bDone = False
			Dim lID As Int32 = CInt(Val(oINI.GetString("DOMAIN", "ID" & X, "-1")))
			If lID > 0 Then
				AddToLoadSystemQueue(lID)
				X += 1
			Else : bDone = True
			End If
		End While
	End Sub

	Public Shared Sub AddToLoadSystemQueue(ByVal lSystemID As Int32)
		If mlLoadSystemQueueID Is Nothing Then ReDim mlLoadSystemQueueID(-1)
		SyncLock mlLoadSystemQueueID
			Dim lIdx As Int32 = -1
			For X As Int32 = 0 To mlLoadSystemQueueID.GetUpperBound(0)
				If lIdx = -1 AndAlso mlLoadSystemQueueID(X) = -1 Then
					lIdx = X
				ElseIf mlLoadSystemQueueID(X) = lSystemID Then
					Return
				End If
			Next X
			If lIdx = -1 Then
				ReDim Preserve mlLoadSystemQueueID(mlLoadSystemQueueID.GetUpperBound(0) + 1)
				mlLoadSystemQueueID(mlLoadSystemQueueID.GetUpperBound(0)) = lSystemID
			Else
				mlLoadSystemQueueID(lIdx) = lSystemID
			End If
		End SyncLock
	End Sub

	'Ok, this method manages our system loading
	Public Shared Sub CheckLoadSystemQueue()
		If mbInLoadSystem = True Then Return 'avoid multiple loads
		If mlLoadSystemQueueID Is Nothing Then Return
		For X As Int32 = 0 To mlLoadSystemQueueID.GetUpperBound(0)
			If mlLoadSystemQueueID(X) > 0 Then
				If myServerReadySent = 0 Then myServerReadySent = 1
				mbInLoadSystem = True
				Dim lSystemID As Int32 = mlLoadSystemQueueID(X)
				mlLoadSystemQueueID(X) = -1
				If LoadSystem(lSystemID) = False Then
					gfrmDisplayForm.AddEventLine("Unable to load system " & lSystemID)
					mbInLoadSystem = False
				Else : Return
				End If
			End If
		Next X

		'Ok, we are here, this means all of our system loading is done, so check if we ever sent our server ready msg
		If myServerReadySent < 2 Then
			myServerReadySent = 255
			goMsgSys.SendServerReady()
			goMsgSys.AcceptingClients = True
		End If
	End Sub

	''' <summary>
	''' Step 1: Call LoadSystem which validates our selection and then do a request to the server for the object
	'''   we return instantly with a true or false indicating whether the request was successfully made 
	''' </summary>
	''' <param name="lID"> SystemID of the system to load </param>
	''' <returns></returns>
	''' <remarks></remarks>
	Private Shared Function LoadSystem(ByVal lID As Int32) As Boolean
		moCurrentSystem = Nothing
		If lID = -1 Then Return False

		For lIdx As Int32 = 0 To goGalaxy.mlSystemUB
			If goGalaxy.moSystems(lIdx).ObjectID = lID Then
				moCurrentSystem = goGalaxy.moSystems(lIdx)
				gfrmDisplayForm.AddEventLine("Getting System Details for " & goGalaxy.moSystems(lIdx).ObjectID & "...")
				If goMsgSys.GetSystemDetails(goGalaxy.moSystems(lIdx).ObjectID) = False Then Return False Else Return True
			End If
		Next lIdx
		Return False
	End Function

	''' <summary>
	''' Step 2: When Get system details response is received and processed, we call this method in a new thread!!!
	'''   this method instantiates our object and grids and planets
	''' </summary> 
	''' <remarks></remarks>
	Public Shared Sub ReceivedSystemDetails()
		gfrmDisplayForm.AddEventLine("Creating Environment for System...")
		'Create the environment for this system...
		glEnvirUB += 1
		ReDim Preserve goEnvirs(glEnvirUB)
		goEnvirs(glEnvirUB) = New Envir()
		With goEnvirs(glEnvirUB)
			.ObjectID = moCurrentSystem.ObjectID
			.ObjTypeID = moCurrentSystem.ObjTypeID
			.oGeoObject = moCurrentSystem
			moCurrentSystem.EnvirIdx = glEnvirUB
			.SetEnvirGridValues()
		End With

		gfrmDisplayForm.AddEventLine("Creating Planet Environments...")
		'Now, go thru the planets and create the environments
		For lPlanet As Int32 = 0 To moCurrentSystem.PlanetUB
			glEnvirUB += 1
			ReDim Preserve goEnvirs(glEnvirUB)
			goEnvirs(glEnvirUB) = New Envir()
			With goEnvirs(glEnvirUB)
				.ObjectID = moCurrentSystem.moPlanets(lPlanet).ObjectID
				.ObjTypeID = moCurrentSystem.moPlanets(lPlanet).ObjTypeID
				.oGeoObject = moCurrentSystem.moPlanets(lPlanet)
				moCurrentSystem.moPlanets(lPlanet).EnvirIdx = glEnvirUB
				.SetEnvirGridValues()
			End With
		Next lPlanet

		gfrmDisplayForm.AddEventLine("Geography Loaded! Getting Environment Objects...")
		'Send a request for environment objects for the system
		goMsgSys.SendRequestEnvirObjs(moCurrentSystem.ObjectID, moCurrentSystem.ObjTypeID)
	End Sub

	''' <summary>
	''' Step 3: Received an environment object list, could be a system or a planet within system
	'''   next, move on to the next entity of the parent system... once at the end of the list, we're done
	''' </summary>
	''' <param name="yData"> The msg received from the primary </param>
	''' <remarks></remarks>
	''' 
	Public Shared Sub ReceivedEnvironmentObjects(ByRef yData() As Byte)
		Dim lPos As Int32 = 2
		mlLastReceivedID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		miLastReceivedTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim oThread As New Threading.Thread(AddressOf FinishReceivedEnvironmentObjects)
		oThread.Start()
	End Sub

	''' <summary>
	''' To be called via ReceivedEnvironmentObjects in a new thread after mlLastReceivedID and miLastReceivedTypeID have been set
	''' </summary>
	''' <remarks></remarks>
	Private Shared Sub FinishReceivedEnvironmentObjects()
		'first, send our msg of the register
		goMsgSys.SendRegisterDomainMsg(mlLastReceivedID, miLastReceivedTypeID)

		'Now, determine our next environment
		If miLastReceivedTypeID = ObjectType.eSolarSystem Then
			If mlLastReceivedID <> moCurrentSystem.ObjectID Then
				gfrmDisplayForm.AddEventLine("ERROR: Invalid System Result from RequestEnvironmentObjects!")
				Return
			End If
			If moCurrentSystem.PlanetUB > -1 Then
				'Send a request for environment objects for this planet
				goMsgSys.SendRequestEnvirObjs(moCurrentSystem.moPlanets(0).ObjectID, moCurrentSystem.moPlanets(0).ObjTypeID)
                Return
            Else
                mbInLoadSystem = False
            End If
		ElseIf miLastReceivedTypeID = ObjectType.ePlanet Then
			Dim oSystem As SolarSystem = moCurrentSystem
			If oSystem Is Nothing = False Then
				Dim bNext As Boolean = False
				For X As Int32 = 0 To oSystem.PlanetUB
					If bNext = True Then
						'Send a request for environment object for this planet
						goMsgSys.SendRequestEnvirObjs(moCurrentSystem.moPlanets(X).ObjectID, moCurrentSystem.moPlanets(X).ObjTypeID)
						Return
					ElseIf oSystem.moPlanets(X).ObjectID = mlLastReceivedID Then
						bNext = True
					End If
				Next X
			End If

			'okay we are done, update our player start loc data
			For X As Int32 = 0 To glPlayerUB
				If glPlayerIdx(X) <> -1 Then
					If goPlayers(X).lStartEnvirID > 0 AndAlso goPlayers(X).lStartLocX <> Int32.MinValue AndAlso goPlayers(X).lStartLocZ <> Int32.MinValue Then
						For Y As Int32 = 0 To glEnvirUB
							If goEnvirs(Y).ObjTypeID = ObjectType.ePlanet AndAlso goEnvirs(Y).ObjectID = goPlayers(X).lStartEnvirID Then
								CType(goEnvirs(Y).oGeoObject, Planet).SetPlayerStartLocationMarked(goPlayers(X).lStartLocX, goPlayers(X).lStartLocZ)
								CType(goEnvirs(Y).oGeoObject, Planet).SetPirateStartLocationMarked(goPlayers(X).lPirateStartLocX, goPlayers(X).lPirateStartLocZ)
								Exit For
							End If
						Next Y
					End If
				End If
			Next X

			'now, set our in load system flag
			mbInLoadSystem = False
		End If
	End Sub

End Class
