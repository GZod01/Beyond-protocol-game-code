Option Strict On

Public Class frmMain
    Private WithEvents moMsgSys As MsgSystem

    Private ConnectionAttempt As Boolean = False
    'Private msOverall As String = ""

    Public bShown As Boolean = False

    Private mlIgnoreListUB As Int32 = -1
    Private msIgnoreList() As String

    Private moCurrentTab As ChatTab = Nothing

    Public Event LoginAttemptResult(ByVal bSuccess As Boolean)

    Private moTabs() As ChatTab
    Private mlTabUB As Int32 = -1

    Public Shared sCommandArg As String

    Public Function DoLogin() As Boolean
        moMsgSys = New MsgSystem()

        If Command() Is Nothing OrElse Command.ToUpper <> "SA" Then
            moMsgSys.msPrimaryIP = "74.113.102.134"
            moMsgSys.mlPrimaryPort = 7779
            sCommandArg = "BP"
        Else
            moMsgSys.msPrimaryIP = "74.113.102.132"
            moMsgSys.mlPrimaryPort = 7759 '7779
            sCommandArg = "SA"
        End If

        
        ConnectionAttempt = True
        Return moMsgSys.ConnectToPrimary()
    End Function

    Private Function GetColorRTF(ByVal yType As ChatMessageType) As String
        Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 0, 0)
        Select Case yType
            Case ChatMessageType.eAliasChatMessage
                clrVal = AliasChatColor
            Case ChatMessageType.eAllianceMessage
                clrVal = GuildChatColor
            Case ChatMessageType.eChannelMessage
                clrVal = ChannelChatColor
            Case ChatMessageType.eLocalMessage
                clrVal = LocalChatColor
            Case ChatMessageType.eNotificationMessage
                clrVal = AlertChatColor
            Case ChatMessageType.ePrivateMessage
                clrVal = PMChatColor
            Case ChatMessageType.eSenateMessage
                clrVal = SenateChatColor
            Case ChatMessageType.eSysAdminMessage
                clrVal = AlertChatColor
        End Select
        Return "\red" & clrVal.R.ToString & "\green" & clrVal.G.ToString & "\blue" & clrVal.B.ToString & ";"
    End Function

    Public Function GetCurrentRTB() As System.Windows.Forms.RichTextBox
        Dim oTab As TabPage = moCurrentTab.oTab 'tbMain.SelectedTab()
        If oTab Is Nothing = False Then
            Return CType(oTab.Controls(0), System.Windows.Forms.RichTextBox)
        End If
        Return Nothing
    End Function

    Private Function IgnoreMessage(ByVal sText As String) As Boolean
        If sText Is Nothing Then Return False
        Dim lIdx As Int32 = sText.IndexOf("tells you,")
        Try
            If lIdx <> -1 Then
                Dim sWho As String = sText.Substring(0, lIdx).Trim.ToUpper

                For X As Int32 = 0 To mlIgnoreListUB
                    If msIgnoreList(X) = sWho Then
                        Return True
                    End If
                Next X
            Else
                sText = sText.ToUpper
                For X As Int32 = 0 To mlIgnoreListUB
                    If sText.ToUpper.StartsWith(msIgnoreList(X)) Then Return True
                Next X
            End If
        Catch
        End Try

        Return False
    End Function
    Public Sub AddTextLine(ByVal lPlayerID As Int32, ByVal sEvent As String, ByVal yType As ChatMessageType)

        Dim lFilter As Int32 = CInt(Math.Pow(2, yType))
        Dim sChannel As String = ""
        If lFilter = ChatTab.ChatFilter.eChannelMessages Then
            'Find our channel
            Dim lTemp As Int32 = sEvent.IndexOf("(")
            If lTemp <> -1 Then
                sChannel = sEvent.Substring(lTemp + 1, sEvent.IndexOf(")", lTemp) - lTemp - 1)
            End If
        End If
        If sChannel = "Emperor" Then yType = ChatMessageType.eSenateMessage
        If IgnoreMessage(sEvent) = True Then Return

        For X As Int32 = 0 To mlTabUB
            If moTabs(X).AcceptsLine(lFilter, sChannel) = True Then
                moTabs(X).AddLine(sEvent, lPlayerID, yType)
            End If
        Next X

        Me.Invoke(New doUpdate(AddressOf delegateUpdateEvents), sEvent)
    End Sub

    Private Delegate Sub doUpdate(ByVal sEvent As String)
    Private Sub delegateUpdateEvents(ByVal sEvent As String)
        If ConnectionAttempt Then
            If InStr(sEvent, "Login Successful") > 0 Then
                ConnectionAttempt = False
                txtNew.Enabled = True
                txtNew.ReadOnly = False
                btnSend.Enabled = True
                btnChannels.Enabled = True
                btnOptions.Enabled = True
                btnColors.Enabled = True
                Me.Text = "Chat Monitor - " & gsUserName
                RaiseEvent LoginAttemptResult(True)
            ElseIf InStr(sEvent, "Failed") > 0 Then
                ConnectionAttempt = False
                RaiseEvent LoginAttemptResult(False)
            End If
        End If
    End Sub

#Region " Message System Events "

    'Private Function EncBytes(ByVal yBytes() As Byte) As Byte()
    '       'Dim lLen As Int32 = UBound(yBytes)
    '       'Dim lKey As Int32
    '       'Dim lOffset As Int32
    '       'Dim X As Int32
    '       'Dim lChrCode As Int32
    '       'Dim lMod As Int32

    '       'Dim yFinal(lLen + 1) As Byte

    '       'lKey = CInt(Math.Floor((GetNxtRnd() * 51)))
    '       'Rnd(-1)
    '       'Call Randomize(777 + lKey)

    '       'yFinal(0) = CByte(lKey)
    '       'For X = 0 To lLen
    '       '	lOffset = X - lLen
    '       '	'now, found out what we got here..
    '       '	lChrCode = yBytes(X)
    '       '	lMod = CInt(Math.Floor((GetNxtRnd() * 5) + 1))
    '       '	lChrCode = lChrCode + lMod
    '       '	If lChrCode > 255 Then lChrCode = lChrCode - 256
    '       '	yFinal(X + 1) = CByte(lChrCode)
    '       'Next X
    '       Dim lLen As Int32 = yBytes.GetUpperBound(0)
    '       Dim yFinal(lLen) As Byte
    '       For X As Int32 = 0 To lLen
    '           If yBytes(X) <> 0 Then
    '               Dim lTmp As Int32 = yBytes(X) + 1
    '               If lTmp > 255 Then lTmp = 0
    '               yFinal(X) = CByte(lTmp)
    '           Else
    '               yFinal(X) = 0
    '           End If
    '       Next X


    '	EncBytes = yFinal

    'End Function

    Private Sub moMsgSys_ConnectedToServer(ByVal yConnType As eyConnType) Handles moMsgSys.ConnectedToServer
        AddTextLine(-1, "\cf1 {\i Connected to server... validating version...}", ChatMessageType.eSysAdminMessage)

        If moMsgSys.ValidateVersion(eyConnType.PrimaryServer) = False Then         'pass 1 for the operator
            moMsgSys.DisconnectAll()
            moMsgSys = Nothing

            Dim sw As Stopwatch = Stopwatch.StartNew
            While sw.ElapsedMilliseconds < 10000
                Application.DoEvents()
                Threading.Thread.Sleep(1)
            End While
            sw.Stop()
            sw = Nothing
            Me.Close()
            Application.Exit()
            Return
        End If

    End Sub

    Private msOverrideIP As String = Nothing '= ""
    Private Sub moMsgSys_ReceivedEnvironmentsDomain(ByVal sIPAddress As String, ByVal iPort As Short) Handles moMsgSys.ReceivedEnvironmentsDomain

        If msOverrideIP Is Nothing Then
            Dim oINI As New InitFile()
            msOverrideIP = oINI.GetString("CONNECTION", "OverrideIP", "")
            oINI = Nothing
        End If
        If msOverrideIP <> "" Then sIPAddress = msOverrideIP '"epica.servegame.org"

        'If goCurrentEnvir Is Nothing = False Then
        '    'ok, got the domain ip and port
        '    If goCurrentEnvir.DomainServerPort = 0 Then
        '        'not connnected period... let's connect
        '        bNeedToConnect = True
        '    ElseIf goCurrentEnvir.DomainServerIP = sIPAddress AndAlso goCurrentEnvir.DomainServerPort = iPort Then
        '        'we're already connected to the right domain... so we just need to notify the domain that we're ready
        '        bNeedToConnect = False
        '        'moMsgSys_ConnectedToRegion()
        '        moMsgSys_ConnectedToServer(eyConnType.RegionServer)
        '    Else
        '        'we're connected, but not to the right domain... so disconnect and reconnect to the right one
        '        moMsgSys.DisconnectRegion()
        '        bNeedToConnect = True
        '    End If

        '    If bNeedToConnect Then
        '        GFXEngine.bRenderInProgress = True
        '        If moMsgSys.ConnectToDomainServer(sIPAddress, iPort) = False Then
        '            Me.Close()
        '        End If
        '    End If

        '    'ensure our values are set correctly
        '    goCurrentEnvir.DomainServerIP = sIPAddress
        '    goCurrentEnvir.DomainServerPort = iPort
        'End If
    End Sub

    Private mbNeedConnectToPrimary As Boolean = False
    Private Sub moMsgSys_ReceivedPlayerCurrentEnvironment(ByVal lPlayerID As Integer, ByVal lID As Integer, ByVal iTypeID As Short, ByVal lGalaxyID As Int32, ByVal lSystemID As Int32, ByVal lPlanetID As Int32, ByVal bInCurrDomain As Boolean) Handles moMsgSys.ReceivedPlayerCurrentEnvironment
        'right now, should always be the case... but it might be interesting to allow a player to look up another player's
        '  current environment...
        Dim sPrevIP As String = ""
        Dim X As Int32

        'AddTextLine("\cf1 {\i Current Environment: " & lID & ", " & iTypeID & "}")

        'Disconnect from Region, Disconnect from Primary, Connect to Operator
        If bInCurrDomain = False Then

            'Set these values here so we can use them soon...
            mlRequestPrimaryID = lID
            miRequestPrimaryTypeID = iTypeID

            moMsgSys.DisconnectAll()
            moMsgSys.ResetClosing()

            For X = 0 To 15
                Application.DoEvents()
                Threading.Thread.Sleep(10)
            Next

            mbNeedConnectToPrimary = True

            Return
        End If

        'Ok, for now, attempt to join testers
        Dim yData(46) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yData, lPos) : lPos += 2
        lPos += 4 'leave room for playerid
        yData(lPos) = eyChatRoomCommandType.JoinChannel : lPos += 1
        System.Text.ASCIIEncoding.ASCII.GetBytes("GENERAL").CopyTo(yData, lPos) : lPos += 20
        moMsgSys.SendToPrimary(yData)
        AddTextLine(-1, "\cf1 {\i Joining General}", ChatMessageType.eSysAdminMessage)

        If glPlayerID < 8 Then
            Dim oThread As New Threading.Thread(AddressOf CheckConnection)
            oThread.Start()
        End If

        If lPlayerID = glPlayerID Then
            'If goCurrentEnvir Is Nothing = False Then
            '    sPrevIP = goCurrentEnvir.DomainServerIP
            '    iPrevPort = goCurrentEnvir.DomainServerPort
            '    bFirstTime = False
            '    goCurrentEnvir.DisposeMe()
            'Else
            '    bFirstTime = True
            'End If
            'goCurrentEnvir = Nothing
            'If goUILib Is Nothing = False Then
            '    'goUILib.RemoveWindow("frmColonyStats")
            '    goUILib.ClearSelection()
            'End If
            'goCurrentEnvir = New BaseEnvironment()
            'With goCurrentEnvir
            '    .ObjectID = lID
            '    .ObjTypeID = iTypeID
            '    .DomainServerIP = sPrevIP
            '    .DomainServerPort = iPrevPort
            '    .oGeoObject = Nothing
            'End With

            'If goPFXEngine32 Is Nothing = False Then goPFXEngine32.ClearAllEmitters()
            'If goMissileMgr Is Nothing = False Then goMissileMgr.KillAll()
            'If goWpnMgr Is Nothing = False Then goWpnMgr.CleanAll()
            'If goSound Is Nothing = False Then goSound.KillAllSounds()

            'If bFirstTime = True Then

            '    If lGalaxyID <> -1 AndAlso goGalaxy.ObjectID = lGalaxyID Then
            '        If lSystemID <> -1 Then
            '            For X = 0 To goGalaxy.mlSystemUB
            '                If goGalaxy.moSystems(X).ObjectID = lSystemID Then
            '                    goGalaxy.CurrentSystemIdx = X
            '                    Exit For
            '                End If
            '            Next X

            '            'ensure we have the system's details
            '            If goGalaxy.CurrentSystemIdx <> -1 Then
            '                If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB = -1 Then
            '                    moMsgSys.RequestSystemDetails(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID)
            '                End If
            '            End If

            '            If lPlanetID <> -1 Then
            '                For X = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB
            '                    If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(X).ObjectID = lPlanetID Then
            '                        goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx = X
            '                        Exit For
            '                    End If
            '                Next X
            '            Else
            '                goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx = -1
            '            End If
            '        Else
            '            goGalaxy.CurrentSystemIdx = -1
            '        End If
            '    Else
            '        'TODO: what do we do here?
            '    End If

            '    If lPlanetID <> -1 Then
            '        glCurrentEnvirView = CurrentView.ePlanetMapView
            '    ElseIf lSystemID <> -1 Then
            '        glCurrentEnvirView = CurrentView.eSystemMapView1

            '        With goCamera
            '            .mlCameraAtX = 0 : .mlCameraX = 0
            '            .mlCameraAtY = 0 : .mlCameraY = 1000
            '            .mlCameraAtZ = 0 : .mlCameraZ = -1000
            '        End With

            '    Else
            '        glCurrentEnvirView = CurrentView.eGalaxyMapView
            '        With goCamera
            '            .mlCameraAtX = 0 : .mlCameraX = 0
            '            .mlCameraAtY = 0 : .mlCameraY = 1000
            '            .mlCameraAtZ = 0 : .mlCameraZ = -1000
            '        End With
            '    End If

            '    'force our startup screen to be removed
            '    moEngine.ForceCleanupLoginBackground()
            'End If

            'If goGalaxy.CurrentSystemIdx <> -1 Then
            '    If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx <> -1 Then
            '        goCurrentEnvir.oGeoObject = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx)
            '    Else
            '        goCurrentEnvir.oGeoObject = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
            '    End If
            '    goCurrentEnvir.SetExtents()
            'End If

            'If lID >= 500000000 AndAlso iTypeID = ObjectType.ePlanet Then
            '    goCurrentEnvir.oGeoObject = Planet.GetTutorialPlanet(moEngine.GetDevice)
            '    goCurrentEnvir.SetExtents()
            'End If

            If lID = 0 OrElse iTypeID = 0 Then
                moMsgSys_ProcessedBurstEnvironment()
            Else
                'ok,if we have received this, then we need to send a request for that environment's domain
                ReDim yData(7)
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestEnvironmentDomain).CopyTo(yData, 0)
                'goCurrentEnvir.GetGUIDAsString.CopyTo(yData, 2)
                System.BitConverter.GetBytes(lID).CopyTo(yData, 2)
                System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 6)
                moMsgSys.SendToPrimary(yData)
            End If
        End If


        'NOTE: This doesn't work because gyPlayertitle is never set, we'll have to force the server to send the value down
        If gyPlayerTitle = 6 OrElse glPlayerID = 1 OrElse glPlayerID = 2 OrElse glPlayerID = 6 OrElse glPlayerID = 7 Then
            ReDim yData(46)
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yData, lPos) : lPos += 2
            lPos += 4 'leave room for playerid
            yData(lPos) = eyChatRoomCommandType.JoinChannel : lPos += 1
            System.Text.ASCIIEncoding.ASCII.GetBytes("EMPEROR").CopyTo(yData, lPos) : lPos += 20
            moMsgSys.SendToPrimary(yData)
            AddTextLine(-1, "\cf1 {\i Joining Emperor}", ChatMessageType.eSysAdminMessage)
        End If

    End Sub

    '	Private Sub moMsgSys_ConnectedToPrimary() Handles moMsgSys.ConnectedToPrimary
    '		'        Me.Text = "Epica - Connected to Primary"
    '		Debug.Write("Connected to Primary" & vbCrLf)
    '#If TEST_HARNESS = 1 Then
    '        oSB.AppendLine("Connected to Primary")
    '#End If
    '	End Sub

    '	Private Sub moMsgSys_ConnectedToRegion() Handles moMsgSys.ConnectedToRegion
    '		Dim yData() As Byte

    '		'Me.Text = "Epica - Connected to Region"
    '		Debug.Write("Connected to Region" & vbCrLf)
    '#If TEST_HARNESS = 1 Then
    '        oSB.AppendLine("Connected to Region")
    '#End If
    '		'send our Player ID
    '		ReDim yData(89)
    '		Dim lPos As Int32 = 0
    '		System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginRequest).CopyTo(yData, lPos) : lPos += 2
    '		System.BitConverter.GetBytes(glActualPlayerID).CopyTo(yData, lPos) : lPos += 4
    '		System.Text.ASCIIEncoding.ASCII.GetBytes(gsUserName).CopyTo(yData, lPos) : lPos += 20
    '		System.Text.ASCIIEncoding.ASCII.GetBytes(gsPassword).CopyTo(yData, lPos) : lPos += 20
    '		System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, lPos) : lPos += 4
    '		System.Text.ASCIIEncoding.ASCII.GetBytes(gsAliasUserName).CopyTo(yData, lPos) : lPos += 20
    '		System.Text.ASCIIEncoding.ASCII.GetBytes(gsAliasPassword).CopyTo(yData, lPos) : lPos += 20
    '		moMsgSys.SendToRegion(yData)

    '		ReDim yData(9)
    '		'Now, need to request the environment
    '		If goCurrentEnvir Is Nothing = False Then
    '#If TEST_HARNESS = 1 Then
    '            oSB.AppendLine("BurstEnvironmentRequest")
    '            '            mbChangingEnvirs = False
    '            '            mbEnvirReady = True
    '            '#Else
    '#End If
    '			'request the environment objects
    '			System.BitConverter.GetBytes(GlobalMessageCode.eBurstEnvironmentRequest).CopyTo(yData, 0)
    '			goCurrentEnvir.GetGUIDAsString.CopyTo(yData, 2)
    '			moMsgSys.SendToRegion(yData)
    '		End If

    '	End Sub
    Private mbFirstLoad As Boolean = True
    Private Sub moMsgSys_ProcessedBurstEnvironment() Handles moMsgSys.ProcessedBurstEnvironment
        'this is an indication that the goCurrentEnvir object is fully populated...

    End Sub

    Private Sub moMsgSys_ReceivedGalaxyAndSystems(ByVal yData() As Byte) Handles moMsgSys.ReceivedGalaxyAndSystems
        'Ok, time to light up the ovens... this message contains the Galaxy Objects and System Objects
        AddTextLine(-1, "\cf1 {\i Requesting Player Details...}", ChatMessageType.eSysAdminMessage)
        moMsgSys.RequestPlayerDetails()

    End Sub

    Private Sub moMsgSys_ReceivedStarTypes(ByVal yData() As Byte) Handles moMsgSys.ReceivedStarTypes


    End Sub

    '	Private Sub moMsgSys_ReceivedSystemDetails(ByVal yData() As Byte) Handles moMsgSys.ReceivedSystemDetails
    '		Dim lPos As Int32 = 2	'msg id
    '		Dim oTmpPlnt As Planet
    '		Dim lID As Int32
    '		Dim iTypeID As Int16
    '		Dim sName As String
    '		Dim ySizeID As Byte
    '		Dim yMapTypeID As Byte

    '		Dim lInnerRadius As Int32
    '		Dim lOuterRadius As Int32
    '		Dim lRingDiffuse As Int32

    '		'ok, got our details, this is nothing more than a list of Planet objects... we would only get this message
    '		'  for our galaxy's current system
    '		If goGalaxy.CurrentSystemIdx <> -1 Then
    '			While lPos < yData.Length - 1


    '				'Object ID
    '				lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '				'ObjTypeID
    '				iTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

    '				'Name
    '				sName = GetStringFromBytes(yData, lPos, 20)
    '				lPos += 20

    '				'Planet type id
    '				yMapTypeID = yData(lPos) : lPos += 1
    '				'Size ID
    '				ySizeID = yData(lPos) : lPos += 1

    '				'Now, create our planet
    '#If TEST_HARNESS = 0 Then
    '				oTmpPlnt = New Planet(moEngine.GetDevice, lID, ySizeID, yMapTypeID)
    '#Else
    '				                oTmpPlnt = New Planet(nothing, lID, ySizeID, yMapTypeID)
    '#End If


    '				With oTmpPlnt
    '					.ObjTypeID = iTypeID
    '					.PlanetName = sName
    '					'ok, now the other attributes
    '					.PlanetRadius = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '					.ParentSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx) : lPos += 4
    '					.LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '					.LocY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '					.LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '					.Vegetation = yData(lPos) : lPos += 1
    '					.Atmosphere = yData(lPos) : lPos += 1
    '					.Hydrosphere = yData(lPos) : lPos += 1
    '					.Gravity = yData(lPos) : lPos += 1
    '					.SurfaceTemperature = yData(lPos) : lPos += 1
    '					.RotationDelay = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

    '					'Now, check for our ring...
    '					lInnerRadius = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '					lOuterRadius = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '					lRingDiffuse = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

    '					.SetRingProps(lInnerRadius, lOuterRadius, lRingDiffuse)

    '					'finally, the axis angle
    '					.AxisAngle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

    '					.RotateAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '					.SetLastUpdate(timeGetTime)
    '				End With
    '				goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).AddPlanet(oTmpPlnt)
    '				oTmpPlnt = Nothing
    '			End While
    '		End If

    '#If TEST_HARNESS = 1 Then
    '				        oSB.AppendLine("Received System Details")
    '#End If

    '	End Sub

    Private Sub moMsgSys_ReceivedSystemDetails(ByVal yData() As Byte) Handles moMsgSys.ReceivedSystemDetails

    End Sub

    Private Sub moMsgSys_ServerShutdown() Handles moMsgSys.ServerShutdown


        'disconnect everyone...
        moMsgSys.DisconnectAll()


        MsgBox("The server has stopped responding. This can be a result of your ISP" & vbCrLf & _
               "dropping the connection or the game servers have gone offline. Restart" & vbCrLf & _
               "the game and check the News on the updater for more information.", MsgBoxStyle.OkOnly, "Connection Severed")
        Me.Close()
        Application.Exit()
    End Sub

    Private Sub moMsgSys_ValidateVersionResponse(ByVal yConnType As eyConnType, ByVal bValue As Boolean) Handles moMsgSys.ValidateVersionResponse
        If bValue = False Then
            moMsgSys.DisconnectAll()
            moMsgSys = Nothing

            AddTextLine(-1, "\cf7 {\b Version incompatible with server!}", ChatMessageType.eSysAdminMessage)

            Dim sw As Stopwatch = Stopwatch.StartNew
            While sw.ElapsedMilliseconds < 10000
                Application.DoEvents()
                'Threading.Thread.Sleep(0)
                Threading.Thread.Sleep(1)
            End While
            sw.Stop()
            sw = Nothing
            Me.Close()
            Application.Exit()
            Return
        Else
            AddTextLine(-1, "\cf1 {\i Processing Login Request...}", ChatMessageType.eSysAdminMessage)
            moMsgSys.RequestLogin(gsUserName, gsPassword, False, "", "", yConnType)
        End If
    End Sub

    Private Sub moMsgSys_LoginResponse(ByVal yConnType As eyConnType, ByVal lVal As Integer) Handles moMsgSys.LoginResponse
        If lVal < 0 Then
            Select Case lVal
                Case LoginResponseCodes.eAccountBanned
                    AddTextLine(-1, "\cf7 {\b Login Failed! Account Banned. Please contact support.}", ChatMessageType.eSysAdminMessage)
                Case LoginResponseCodes.eAccountSuspended
                    'goUILib.SetToolTip("Login Failed: Account Suspended. Please Contact Support.", -1, lStatusTop)
                    AddTextLine(-1, "\cf7 {\bLogin Failed! Account Suspended. Please contact support.}", ChatMessageType.eSysAdminMessage)
                Case LoginResponseCodes.eInvalidPassword, LoginResponseCodes.eInvalidUserName
                    'goUILib.SetToolTip("Login Failed: Invalid Username/Password Combination.", -1, lStatusTop)
                    AddTextLine(-1, "\cf7 {\bLogin Failed! Invalid Username/Password.}", ChatMessageType.eSysAdminMessage)
                Case LoginResponseCodes.eLoginAttemptLockout
                    'goUILib.SetToolTip("Login Failed: Maximum Attempts Exceeded!", -1, lStatusTop)
                    AddTextLine(-1, "\cf7 {\bLogin Failed! Maximum attempts exceeded!}", ChatMessageType.eSysAdminMessage)
                Case LoginResponseCodes.eAccountInactive
                    'goUILib.SetToolTip("This account is inactive." & vbCrLf & "You must reactivate the account in order to login.", -1, lStatusTop)
                    AddTextLine(-1, "\cf7 {\bThis account is inactive. }", ChatMessageType.eSysAdminMessage)
                Case LoginResponseCodes.eAccountSetup
                    AddTextLine(-1, "\cf7 {\bThis account needs to be setup.}", ChatMessageType.eSysAdminMessage)
                    Return
                Case LoginResponseCodes.eAccountInUse
                    AddTextLine(-1, "\cf7 {\bAccount is in use. Please try again later.}", ChatMessageType.eSysAdminMessage)
                Case LoginResponseCodes.ePlayerIsDying
                    AddTextLine(-1, "\cf7 {\bThis account is still in the process of dying.}", ChatMessageType.eSysAdminMessage)
                Case Else
                    'goUILib.SetToolTip("Login Failed, Please try again later.", -1, lStatusTop)
                    AddTextLine(-1, "\cf7 {\bLogin failed, please try again later}", ChatMessageType.eSysAdminMessage)
            End Select
            're-enable our button

            moMsgSys.DisconnectAll()
            moMsgSys = Nothing
        Else
            glPlayerID = lVal

            Select Case yConnType
                Case eyConnType.ChatServer  '???
                Case eyConnType.OperatorServer
                    If glPlayerID = glActualPlayerID Then gbAliased = False
                    'waiting on server specifics
                    AddTextLine(-1, "\cf1 {\i Waiting on server specifics...}", ChatMessageType.eSysAdminMessage)
                Case eyConnType.PrimaryServer
                    AddTextLine(-1, "\cf1 {\i Login Successful, getting player details...}", ChatMessageType.eSysAdminMessage)
                    Dim oINI As New InitFile()
                    Dim sTemp As String = oINI.GetString("SIGNON", "LastUserName", "")
                    oINI.WriteString("SIGNON", "LastUserName", gsUserName)
                    oINI = Nothing

                    AddTextLine(-1, "\cf1 {\b Welcome to Beyond Protocol!}", ChatMessageType.eSysAdminMessage)
                    Dim yData(5) As Byte
                    AddTextLine(-1, "\cf1 {\i Entering Environment...}", ChatMessageType.eSysAdminMessage)
                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerEnvironment).CopyTo(yData, 0)
                    System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, 2)
                    moMsgSys.SendToPrimary(yData)
            End Select
        End If
    End Sub

    'Request Primary data here...
    Private mlRequestPrimaryID As Int32 = -1
    Private miRequestPrimaryTypeID As Int16 = -1
    Private Sub moMsgSys_ReceivedPrimaryServerData() Handles moMsgSys.ReceivedPrimaryServerData


        mlRequestPrimaryID = -1
        miRequestPrimaryTypeID = -1
    End Sub

    Private Sub moMsgSys_ServerDisconnected(ByVal yType As eyConnType) Handles moMsgSys.ServerDisconnected
        If yType = eyConnType.OperatorServer Then

            '7) Connect to Primary Server
            If moMsgSys.ConnectToPrimary() = False Then
                AddTextLine(-1, "\cf7 {\b Could not connect to server.}", ChatMessageType.eSysAdminMessage)
                Dim sw As Stopwatch = Stopwatch.StartNew
                While sw.ElapsedMilliseconds < 5000
                    Application.DoEvents()
                    'Threading.Thread.Sleep(0)
                    Threading.Thread.Sleep(1)
                End While
                sw.Stop()
                sw = Nothing
                Me.Close()
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub moMsgSys_ReceiverAllPlayerDetails() Handles moMsgSys.ReceivedAllPlayerDetails
        'Ok, server knows who we are and I have my data, so let's continue
        Dim yData(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerEnvironment).CopyTo(yData, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, 2)
        moMsgSys.SendToPrimary(yData)
    End Sub

#End Region

    Private Sub btnSend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSend.Click
        Dim sNewVal As String = Trim$(txtNew.Text)
        txtNew.Text = ""

        If sNewVal = "" Then Return

        If sNewVal.ToLower.StartsWith("/clientver") = True Then
            Dim lVer As Int32 = gl_BP_CLIENT_VERSION
            If frmMain.sCommandArg = "SA" Then lVer = gl_SA_CLIENT_VERSION
            AddTextLine(-1, "\cf1 {\i Client Version: " & lVer.ToString & "}", ChatMessageType.eSysAdminMessage)
            Return
        End If
        If sNewVal = "__" Then
            AddTextLine(-1, "\cf7 {\b BREAK}", ChatMessageType.eSysAdminMessage)
            Return
        End If
        If sNewVal.ToLower.StartsWith("/join ") = True Then
            Dim sRoom As String = ""
            Dim sPW As String = ""
            sNewVal = sNewVal.Substring(6)
            If sNewVal.Contains(",") = True Then
                Dim lIdx As Int32 = sNewVal.IndexOf(",")
                sPW = sNewVal.Substring(lIdx + 1)
                sRoom = sNewVal.Substring(0, lIdx - 1)
            Else : sRoom = sNewVal
            End If
            Dim yData(46) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yData, lPos) : lPos += 2
            lPos += 4 'leave room for playerid
            yData(lPos) = eyChatRoomCommandType.JoinChannel : lPos += 1
            System.Text.ASCIIEncoding.ASCII.GetBytes(sRoom).CopyTo(yData, lPos) : lPos += 20
            System.Text.ASCIIEncoding.ASCII.GetBytes(sPW).CopyTo(yData, lPos) : lPos += 20
            moMsgSys.SendToPrimary(yData)
            Return
        End If

        Dim yMsg() As Byte

        sNewVal = sNewVal.Trim

        If sNewVal.StartsWith("/") = False Then
            'Append the message only if it does not have a slash. A slash indicates the player wants to do something else
            If moCurrentTab.sMessagePrefix Is Nothing = False AndAlso moCurrentTab.sMessagePrefix <> "" Then
                sNewVal = moCurrentTab.sMessagePrefix & " " & sNewVal
            End If
        End If

        'Prepare our message
        ReDim yMsg(sNewVal.Length + 5)
        System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(sNewVal.Length).CopyTo(yMsg, 2)
        System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(yMsg, 6)

        moMsgSys.SendToPrimary(yMsg)

        txtNew.Focus()

    End Sub

    Private Sub txtNew_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNew.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnSend_Click(Nothing, Nothing)
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Process.GetCurrentProcess.Kill()
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If moMsgSys Is Nothing = False Then moMsgSys.DisconnectAll()
        moMsgSys = Nothing
        SaveSettings()
    End Sub

    Private Sub frmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        bShown = True
    End Sub

    Private Function SpawnNewTabPage(ByVal sName As String) As TabPage

        Dim lTbIdx As Int32 = Me.tbMain.TabPages.Count

        Dim rtb As New System.Windows.Forms.RichTextBox()
        With rtb
            .Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                        Or System.Windows.Forms.AnchorStyles.Left) _
                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            .Location = New System.Drawing.Point(3, 3)
            .Name = "rtb_" & lTbIdx & "_" & sName
            .ReadOnly = True
            .Size = New System.Drawing.Size(948, 564)
            .BackColor = TextBoxBackColor
            .TabIndex = lTbIdx
            .Text = ""
        End With

        Dim oTab As New TabPage(sName)
        With oTab
            .Name = "tp" & sName
            .Padding = New System.Windows.Forms.Padding(3)
            .Size = New System.Drawing.Size(954, 570)
            .TabIndex = lTbIdx
            .Text = sName
            .UseVisualStyleBackColor = True
            .Controls.Add(rtb)
        End With

        Return oTab

    End Function

    Public Sub LoadChatTabs()
        Dim oINI As InitFile = Nothing

        LoadSettings()

        Try
            tbMain.TabPages.Clear()
            moCurrentTab = Nothing
            mlTabUB = -1

            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            Dim sFile As String = sPath & gsUserName & ".cht"

            oINI = New InitFile(sFile)

            Dim lSeqNbr As Int32 = 0
            Dim sHdr As String = "ChatTab" & lSeqNbr
            Dim sName As String = oINI.GetString(sHdr, "TabName", "Chat")
            While sName <> ""
                mlTabUB += 1
                ReDim Preserve moTabs(mlTabUB)
                moTabs(mlTabUB) = New ChatTab()
                With moTabs(mlTabUB)
                    .sTabName = sName
                    .sChannel = oINI.GetString(sHdr, "Channel", "")
                    .lSequenceNumber = lSeqNbr

                    If mlTabUB = 0 Then
                        Dim lTemp As ChatTab.ChatFilter
                        lTemp = ChatTab.ChatFilter.eAllianceMessages Or ChatTab.ChatFilter.eChannelMessages Or ChatTab.ChatFilter.eLocalMessages Or ChatTab.ChatFilter.ePMs Or ChatTab.ChatFilter.eSenateMessages Or ChatTab.ChatFilter.eSysAdminMessages Or ChatTab.ChatFilter.eAliasChatMessage
                        .lFilter = CInt(Val(oINI.GetString(sHdr, "Filter", CInt(lTemp).ToString)))
                    Else : .lFilter = CInt(Val(oINI.GetString(sHdr, "Filter", "0")))
                    End If

                    .sMessagePrefix = oINI.GetString(sHdr, "MsgPrefix", "")

                    .oTab = SpawnNewTabPage(sName)
                    .oTab.Tag = moTabs(mlTabUB)
                End With
                tbMain.TabPages.Add(moTabs(mlTabUB).oTab)

                lSeqNbr += 1
                sHdr = "ChatTab" & lSeqNbr
                sName = oINI.GetString(sHdr, "TabName", "")

                If mlTabUB = 0 Then
                    moCurrentTab = moTabs(mlTabUB)
                End If
            End While

            If mlTabUB = -1 Then
                mlTabUB += 1
                ReDim Preserve moTabs(mlTabUB)
                moTabs(mlTabUB) = New ChatTab()
                With moTabs(mlTabUB)
                    .sTabName = "Chat"
                    .sChannel = ""
                    .lSequenceNumber = 0
                    .lFilter = 0
                    .sMessagePrefix = "/General"
                    .oTab = SpawnNewTabPage(.sTabName)
                    .oTab.Tag = moTabs(mlTabUB)
                End With
                tbMain.TabPages.Add(moTabs(mlTabUB).oTab)
                moCurrentTab = moTabs(mlTabUB)
            End If


            sHdr = "IgnoreList"
            Dim lIgnoreIdx As Int32 = 0
            sName = oINI.GetString(sHdr, "Ignore_" & lIgnoreIdx, "")
            While sName <> ""
                mlIgnoreListUB += 1
                ReDim Preserve msIgnoreList(mlIgnoreListUB)
                msIgnoreList(mlIgnoreListUB) = sName

                lIgnoreIdx += 1
                sName = oINI.GetString(sHdr, "Ignore_" & lIgnoreIdx, "")
            End While

        Catch
        End Try
        oINI = Nothing 
    End Sub

    Private Sub btnOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOptions.Click
        Dim ofrm As New frmChatTabProps()
        'ofrm.SetTab(moTabs(tbMain.SelectedIndex))
        ofrm.SetTab(moCurrentTab)
        ofrm.SetReturn(Me)
        ofrm.Show(Me)
    End Sub

    Private Sub btnDeleteTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteTab.Click
        If tbMain.TabPages.Count = 1 Then
            MsgBox("You cannot delete the last tab!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Cannot Delete Tab")
            Return
        End If
        If tbMain.SelectedTab Is Nothing Then Return
        If MsgBox("Are you sure you wish to delete the current tab?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To mlTabUB
                If Object.Equals(moTabs(X).oTab, tbMain.SelectedTab) = True Then
                    lIdx = X
                    Exit For
                End If
            Next X
            If lIdx <> -1 Then
                For X As Int32 = lIdx To mlTabUB - 1
                    moTabs(X) = moTabs(X + 1)
                Next X
                mlTabUB -= 1
                ReDim Preserve moTabs(mlTabUB)
            End If
            tbMain.TabPages.Remove(tbMain.SelectedTab)
            ForceSaveTabs()
        End If
    End Sub

    Private Sub btnAddTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddTab.Click
        mlTabUB += 1
        ReDim Preserve moTabs(mlTabUB)
        moTabs(mlTabUB) = New ChatTab()
        With moTabs(mlTabUB)
            .sTabName = "New Tab"
            .sChannel = ""
            .lSequenceNumber = mlTabUB
            .lFilter = 0
            .sMessagePrefix = ""
            .oTab = SpawnNewTabPage(.sTabName)
            .oTab.Tag = moTabs(mlTabUB)
        End With
        tbMain.TabPages.Add(moTabs(mlTabUB).oTab)

        moCurrentTab = moTabs(mlTabUB)
        btnOptions_Click(Nothing, Nothing)
    End Sub

    Private Sub btnChannels_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChannels.Click
        If gfrmChannels Is Nothing = False Then
            gfrmChannels.Close()
            gfrmChannels = Nothing
        End If
        gfrmChannels = New frmChannels()
        gfrmChannels.Show(Me)
        gfrmChannels.RequestRoomList()
    End Sub

    Private Sub tbMain_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbMain.SelectedIndexChanged
        If tbMain Is Nothing = False AndAlso tbMain.SelectedTab Is Nothing = False Then
            moCurrentTab = CType(tbMain.SelectedTab.Tag, ChatTab)
        End If
    End Sub

    Public Sub SendMsgToPrimary(ByVal yMsg() As Byte)
        moMsgSys.SendToPrimary(yMsg)
    End Sub

    Private Sub txtNew_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNew.TextChanged
        Dim sNewVal As String = txtNew.Text
        If sNewVal.ToLower.StartsWith("/r ") = True Then
            If moCurrentTab Is Nothing = False Then
                Dim bVal As Boolean = False
                Dim sRes As String = moCurrentTab.ProcessSlashR(bVal, sNewVal)
                If bVal = True Then
                    txtNew.Text = sRes
                    txtNew.SelectionStart = txtNew.Text.Length
                    txtNew.SelectionLength = 0
                Else
                    AddTextLine(-1, "\cf1 {\i No one has sent you any messages yet.}", ChatMessageType.eSysAdminMessage)
                    txtNew.Text = ""
                End If
            End If
        End If
    End Sub

    Private Sub btnColors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColors.Click
        If gfrmColors Is Nothing = False Then
            gfrmColors.Close()
            gfrmColors = Nothing
        End If
        gfrmColors = New frmColors()
        gfrmColors.DoInitialColors()
        gfrmColors.ShowDialog(Me)
        If gfrmColors.bDoSaveColors = True Then
            AddTextLine(-1, "\cf" & (CInt(ChatMessageType.eSysAdminMessage) + 1).ToString & " Colors saved!", ChatMessageType.eSysAdminMessage)
            For X As Int32 = 0 To mlTabUB
                If moTabs(X) Is Nothing = False Then
                    Dim oRTB As RichTextBox = CType(moTabs(X).oTab.Controls(0), RichTextBox)
                    If oRTB Is Nothing = False Then
                        oRTB.BackColor = TextBoxBackColor
                    End If
                End If
            Next X
        End If
        gfrmColors.Close()
    End Sub

    Private Sub CheckConnection()
        Threading.Thread.Sleep(5000)
        If Now.Subtract(dtLastMsg).TotalSeconds > 10 Then
            AddTextLine(-1, "\cf2 SERVER NOT RESPONDING!", ChatMessageType.eSysAdminMessage)
        End If
        Dim oThread As New Threading.Thread(AddressOf CheckConnection)
        oThread.Start()
    End Sub

    Private Sub ReConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReConnect.Click
        frmLogin.Visible = True
        frmLogin.DoEnableControls(True)
    End Sub

    Public Sub ForceSaveTabs()
        If mlTabUB = -1 Then Exit Sub
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= gsUserName & ".cht"
        If Exists(sFile) = True Then Kill(sFile)

        For X As Int32 = 0 To mlTabUB
            moTabs(X).SaveTab()
        Next X

        Dim oINI As InitFile = New InitFile(sFile)
        Dim lIdx As Int32 = 0
        For X As Int32 = 0 To mlIgnoreListUB
            If msIgnoreList(X) <> "" Then
                oINI.WriteString("IgnoreList", "Ignore_" & lIdx, msIgnoreList(X))
                lIdx += 1
            End If
        Next X
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim sName As String
        If HasAliasedRights(AliasingRights.eViewTreasury) = True Then
            Dim sCF As String = glCashFlow.ToString("#,##0")
            If glCashFlow > 0 Then sCF = "+" & sCF
            sName = "Credits: " & glCredits.ToString("#,##0") & " / " & sCF
        Else : sName = "Credits: Lack Alias Rights"
        End If
        If Credits.Text <> sName Then Credits.Text = sName

        'sName = ""
        'If HasAliasedRights(AliasingRights.eViewTreasury) = True Then
        '    sName = "Warpoints: " & glWarpoints.ToString("#,##0") & " / " & glCurrentWPUpkeepCost.ToString("#,##0")
        'Else : sName = "Credits: Lack Alias Rights"
        'End If
        'If Warpoints.Text <> sName Then Warpoints.Text = sName


    End Sub
End Class
