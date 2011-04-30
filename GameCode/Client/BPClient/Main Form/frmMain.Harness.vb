Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmMain

    Private mbConnectingToServer As Boolean = False
    Private mbNeedToLoadTacticalData As Boolean = False
    Public Shared IgnoreResizeEvents As Boolean = False

    Private Sub frmMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        GFXEngine.gbPaused = False

        'Windows Vista - 7 do not continue to pass generic keyboard events back to the BP window and alt-tab away; alt-tab back; you get the stuck radar range rings

        If mbAltKeyDown = True Then
            mbAltKeyDown = False
            Try
                If goUILib Is Nothing = False Then goUILib.mfDistFromSelection = Single.MinValue
                moEngine.bRenderSelectedRanges = mbAltKeyDown
            Catch
            End Try
        End If
    End Sub

    Private Sub frmMain_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Try

            If muSettings.Windowed = True Then
                'stop camera scroll
                If goCamera Is Nothing = False AndAlso (goCamera.bScrollUp = True OrElse goCamera.bScrollDown = True OrElse goCamera.bScrollLeft = True OrElse goCamera.bScrollRight = True) Then
                    With goCamera
                        If .bScrollUp = True Then .bScrollUp = False
                        If .bScrollDown = True Then .bScrollDown = False
                        If .bScrollLeft = True Then .bScrollLeft = False
                        If .bScrollRight = True Then .bScrollRight = False
                    End With
                End If
            End If
        Catch
        End Try

        'If muSettings.Windowed = False Then
        ''moEngine.DoToggleFullScreen()
        'GFXEngine.gbDeviceLost = True
        'End If
    End Sub

    Private Sub frmMain_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseLeave
        Try

            If muSettings.Windowed = True Then
                'stop camera scroll
                If goCamera Is Nothing = False AndAlso (goCamera.bScrollUp = True OrElse goCamera.bScrollDown = True OrElse goCamera.bScrollLeft = True OrElse goCamera.bScrollRight = True) Then
                    With goCamera
                        If .bScrollUp = True Then .bScrollUp = False
                        If .bScrollDown = True Then .bScrollDown = False
                        If .bScrollLeft = True Then .bScrollLeft = False
                        If .bScrollRight = True Then .bScrollRight = False
                    End With
                End If
            End If
        Catch
        End Try
    End Sub

    Private mlPrevSizeX As Int32
    Private mlPrevSizeZ As Int32
    Private Sub frmMain_ResizeBegin(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeBegin
        If IgnoreResizeEvents = True Then Return
        GFXEngine.gbPaused = True
        'GFXEngine.gbDeviceLost = True
        mlPrevSizeX = Me.Width
        mlPrevSizeZ = Me.Height
    End Sub
    Private Sub frmMain_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        If IgnoreResizeEvents = True Then Return
        If Me.WindowState = FormWindowState.Minimized Then Return
        If Me.Width <> mlPrevSizeX OrElse Me.Height <> mlPrevSizeZ Then GFXEngine.gbDeviceLost = True
        GFXEngine.gbPaused = False
    End Sub

    'Private Function EncBytes(ByVal yBytes() As Byte) As Byte()
    '	Dim lLen As Int32 = UBound(yBytes)
    '	Dim lKey As Int32
    '	Dim lOffset As Int32
    '	Dim X As Int32 
    '	Dim lChrCode As Int32
    '	Dim lMod As Int32

    '	Dim yFinal(lLen + 1) As Byte

    '	lKey = CInt(Math.Floor((GetNxtRnd() * 51)))
    '	Rnd(-1)
    '	Call Randomize(777 + lKey)

    '	yFinal(0) = CByte(lKey)
    '	For X = 0 To lLen
    '		lOffset = X - lLen
    '		'now, found out what we got here..
    '		lChrCode = yBytes(X)
    '		lMod = CInt(Math.Floor((GetNxtRnd() * 5) + 1))
    '		lChrCode = lChrCode + lMod
    '		If lChrCode > 255 Then lChrCode = lChrCode - 256
    '		yFinal(X + 1) = CByte(lChrCode)
    '	Next X

    '	EncBytes = yFinal

    'End Function

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim X As Int32
        Form.CheckForIllegalCrossThreadCalls = False

        'Check for an update to the updater...
        If UpdaterUpdate() = True Then
            MsgBox("The updater has been updated and will now be relaunched.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Updates")
            Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sFile.EndsWith("\") = False Then sFile &= "\"
            sFile &= "UpdaterClient.exe"
            If Exists(sFile) Then
                Shell(sFile, AppWinStyle.NormalFocus, False, -1)
            End If
            End
        End If

        If LaunchedFromUpdater() = False Then
            Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sFile.EndsWith("\") = False Then sFile &= "\"
            sFile &= "UpdaterClient.exe"
            If Exists(sFile) = True Then
                Shell("""" & sFile & """", AppWinStyle.NormalFocus, False, -1)
            Else
                MsgBox("Please launch this program from UpdaterClient.exe.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Updates")
            End If
            End
        End If

        gfrmMain = Me
    End Sub

    'Private Sub Login_StartLogin(ByVal sUserName As String, ByVal sPassword As String, ByVal bAlias As Boolean, ByVal sAliasUsername As String, ByVal sAliasPassword As String, ByRef frmLogin As frmLoginDlg)
    '	'this sub handles the login form's Login Event
    '	Dim lRes As Int32
    '	Dim yData() As Byte
    '	Dim lStatusTop As Int32 = frmLogin.Top + frmLogin.Height + 5
    '	Dim cmdLogin As UIButton = Nothing
    '	Dim cmdExit As UIButton = Nothing

    '	Dim sUserNameProper As String = sUserName

    '	sUserName = UCase$(sUserName)
    '	sPassword = UCase$(sPassword)
    '	sAliasUsername = sAliasUsername.ToUpper()
    '	sAliasPassword = sAliasPassword.ToUpper()

    '	'Dim yEncryptedPassword() As Byte = EncBytes(System.Text.ASCIIEncoding.ASCII.GetBytes(sPassword))
    '	'Dim yEncryptedAliasPW() As Byte = EncBytes(System.Text.ASCIIEncoding.ASCII.GetBytes(sAliasPassword))

    '	If sUserName = "DEBUG" Then
    '		If sPassword = "LOGIN_OFF" Then
    '			LoginScreen.mfFadeoutAlpha = 255.0F
    '			Return
    '		End If
    '	End If

    '	If goUILib Is Nothing = False Then goUILib.RetainTooltip = True

    '	'find the Login Windows Login and Exit control's
    '	For lRes = 0 To frmLogin.ChildrenUB
    '		If frmLogin.moChildren(lRes).ControlName = "cmdLogin" Then
    '			cmdLogin = CType(frmLogin.moChildren(lRes), UIButton)
    '		ElseIf frmLogin.moChildren(lRes).ControlName = "cmdExit" Then
    '			cmdExit = CType(frmLogin.moChildren(lRes), UIButton)
    '		End If
    '	Next lRes
    '	'disable it
    '	cmdLogin.Enabled = False

    '	'Application.DoEvents()

    '	'Ok, let's log in... check if we are connected
    '	If moMsgSys Is Nothing = False Then
    '		moMsgSys.DisconnectAll()
    '		moMsgSys = Nothing
    '		If goUILib Is Nothing = False Then goUILib.SetMsgSystem(Nothing)
    '	End If

    '	'Now, disable our exit button
    '	cmdExit.Enabled = False

    '	goUILib.SetToolTip("Connecting to Server...", -1, lStatusTop)
    '	mlLastKeepAliveSend = timeGetTime
    '	moMsgSys = New MsgSystem()
    '	If goUILib Is Nothing = False Then goUILib.SetMsgSystem(moMsgSys)

    '	'NEW LOGIN PROCEDURE
    '	'1) Connect to Operator Server
    '	If moMsgSys.ConnectToOperator() = False Then
    '		goUILib.AddNotification("Could not connect to server.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		goUILib.AddNotification("The server is either down or you", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		goUILib.AddNotification("are not connected to the internet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		Dim sw As Stopwatch = Stopwatch.StartNew
    '		While sw.ElapsedMilliseconds < 5000
    '			Application.DoEvents()
    '			'Threading.Thread.Sleep(0)
    '			Threading.Thread.Sleep(1)
    '		End While
    '		sw.Stop()
    '		sw = Nothing
    '		Me.Close()
    '		Application.Exit()
    '		Exit Sub
    '	End If

    '	'2) Operator Validates Version
    '	goUILib.SetToolTip("Validating Version...", -1, lStatusTop)
    '	If moMsgSys.ValidateVersion(0) = False Then		   'pass 0 for the operator
    '		moMsgSys.DisconnectAll()
    '		moMsgSys = Nothing
    '		If goUILib Is Nothing = False Then goUILib.SetMsgSystem(Nothing)
    '		cmdExit.Enabled = True

    '		goUILib.AddNotification("Version incompatible with server!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		goUILib.AddNotification("Please launch the game through the updater.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		Dim sw As Stopwatch = Stopwatch.StartNew
    '		While sw.ElapsedMilliseconds < 10000
    '			Application.DoEvents()
    '			'Threading.Thread.Sleep(0)
    '			Threading.Thread.Sleep(1)
    '		End While
    '		sw.Stop()
    '		sw = Nothing
    '		Me.Close()
    '		Application.Exit()
    '		Return
    '	End If

    '	'3) Operator Process Login
    '	goUILib.SetToolTip("Processing Login Request...", -1, lStatusTop)
    '	lRes = moMsgSys.RequestLogin(sUserName, sPassword, bAlias, sAliasUsername, sAliasPassword, 0)		'indicates from OPERATOR

    '	If lRes < 0 Then
    '		Select Case lRes
    '			Case LoginResponseCodes.eAccountBanned
    '				'goUILib.SetToolTip("Login Failed: Account Banned. Please Contact Support.", -1, lStatusTop)
    '				goUILib.AddNotification("Login Failed! Account Banned. Please contact support.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eAccountSuspended
    '				'goUILib.SetToolTip("Login Failed: Account Suspended. Please Contact Support.", -1, lStatusTop)
    '				goUILib.AddNotification("Login Failed! Account Suspended. Please contact support.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eInvalidPassword, LoginResponseCodes.eInvalidUserName
    '				'goUILib.SetToolTip("Login Failed: Invalid Username/Password Combination.", -1, lStatusTop)
    '				goUILib.AddNotification("Login Failed! Invalid Username/Password.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eLoginAttemptLockout
    '				'goUILib.SetToolTip("Login Failed: Maximum Attempts Exceeded!", -1, lStatusTop)
    '				goUILib.AddNotification("Login Failed! Maximum attempts exceeded!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eAccountInactive
    '				'goUILib.SetToolTip("This account is inactive." & vbCrLf & "You must reactivate the account in order to login.", -1, lStatusTop)
    '				goUILib.AddNotification("This account is inactive. You must reactivate", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '				goUILib.AddNotification("the account in order to login.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eAccountSetup
    '				goUILib.AddNotification("This account needs to be setup.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '				Dim ofrm As frmPlayerSetup = New frmPlayerSetup(goUILib, sUserNameProper, sPassword)
    '				ofrm.Visible = True
    '				ofrm.sUserName = sUserName
    '				ofrm.sPassword = sPassword
    '				Return
    '			Case Else
    '				'goUILib.SetToolTip("Login Failed, Please try again later.", -1, lStatusTop)
    '				goUILib.AddNotification("Login failed, please try again later", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		End Select
    '		're-enable our button
    '		cmdExit.Enabled = True
    '		cmdLogin.Enabled = True
    '		moMsgSys.DisconnectAll()
    '		moMsgSys = Nothing
    '		If goUILib Is Nothing = False Then
    '			goUILib.RetainTooltip = False
    '			goUILib.SetToolTip(False)
    '		End If

    '		Exit Sub
    '	Else
    '		glPlayerID = lRes
    '	End If

    '	gsUserName = sUserName
    '	gsPassword = sPassword
    '	gsAliasPassword = sAliasPassword
    '	gsAliasUserName = sAliasUsername
    '	If glPlayerID = glActualPlayerID Then gbAliased = False

    '	'5) Operator sends Connection info for Primary 
    '	goUILib.SetToolTip("Waiting on server specifics...", -1, lStatusTop)
    '	While moMsgSys.mbHavePrimaryConnInfo = False
    '		Threading.Thread.Sleep(10)
    '		'Application.DoEvents()
    '		'Threading.Thread.Sleep(0)
    '		'Threading.Thread.Sleep(1)
    '	End While

    '	'6) Disconnect from Operator Server
    '	goUILib.SetToolTip("Specifics received, finishing login process...", -1, lStatusTop)
    '	moMsgSys.DisconnectOperator()

    '	'7) Connect to Primary Server
    '	If moMsgSys.ConnectToPrimary() = False Then
    '		goUILib.AddNotification("Could not connect to server.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		goUILib.AddNotification("The server is either down or you", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		goUILib.AddNotification("are not connected to the internet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		Dim sw As Stopwatch = Stopwatch.StartNew
    '		While sw.ElapsedMilliseconds < 5000
    '			Application.DoEvents()
    '			'Threading.Thread.Sleep(0)
    '			Threading.Thread.Sleep(1)
    '		End While
    '		sw.Stop()
    '		sw = Nothing
    '		Me.Close()
    '		Application.Exit()
    '		Exit Sub
    '	End If

    '	'8) Primary Validates Version
    '	'goUILib.SetToolTip("Validating Version...", -1, lStatusTop)
    '	If moMsgSys.ValidateVersion(1) = False Then		   'pass 1 for the operator
    '		moMsgSys.DisconnectAll()
    '		moMsgSys = Nothing
    '		If goUILib Is Nothing = False Then goUILib.SetMsgSystem(Nothing)
    '		cmdExit.Enabled = True

    '		goUILib.AddNotification("Version incompatible with server!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		goUILib.AddNotification("Please launch the game through the updater.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		Dim sw As Stopwatch = Stopwatch.StartNew
    '		While sw.ElapsedMilliseconds < 10000
    '			Application.DoEvents()
    '			'Threading.Thread.Sleep(0)
    '			Threading.Thread.Sleep(1)
    '		End While
    '		sw.Stop()
    '		sw = Nothing
    '		Me.Close()
    '		Application.Exit()
    '		Return
    '	End If

    '	'9) Primary Processes Login
    '	lRes = moMsgSys.RequestLogin(sUserName, sPassword, bAlias, sAliasUsername, sAliasPassword, 1)		'indicates from PRIMARY
    '	If lRes < 0 Then
    '		Select Case lRes
    '			Case LoginResponseCodes.eAccountBanned
    '				'goUILib.SetToolTip("Login Failed: Account Banned. Please Contact Support.", -1, lStatusTop)
    '				goUILib.AddNotification("Login Failed! Account Banned. Please contact support.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eAccountSuspended
    '				'goUILib.SetToolTip("Login Failed: Account Suspended. Please Contact Support.", -1, lStatusTop)
    '				goUILib.AddNotification("Login Failed! Account Suspended. Please contact support.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eInvalidPassword, LoginResponseCodes.eInvalidUserName
    '				'goUILib.SetToolTip("Login Failed: Invalid Username/Password Combination.", -1, lStatusTop)
    '				goUILib.AddNotification("Login Failed! Invalid Username/Password.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eLoginAttemptLockout
    '				'goUILib.SetToolTip("Login Failed: Maximum Attempts Exceeded!", -1, lStatusTop)
    '				goUILib.AddNotification("Login Failed! Maximum attempts exceeded!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eAccountInactive
    '				'goUILib.SetToolTip("This account is inactive." & vbCrLf & "You must reactivate the account in order to login.", -1, lStatusTop)
    '				goUILib.AddNotification("This account is inactive. You must reactivate", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '				goUILib.AddNotification("the account in order to login.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case LoginResponseCodes.eAccountSetup
    '				goUILib.AddNotification("This account needs to be setup.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '				Dim ofrm As frmPlayerSetup = New frmPlayerSetup(goUILib, sUserNameProper, sPassword)
    '				ofrm.Visible = True
    '				ofrm.sUserName = sUserName
    '				ofrm.sPassword = sPassword
    '				Return
    '			Case LoginResponseCodes.eAccountInUse
    '				goUILib.AddNotification("Someone is already logged in using that account.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '			Case Else
    '				'goUILib.SetToolTip("Login Failed, Please try again later.", -1, lStatusTop)
    '				goUILib.AddNotification("Login failed, please try again later", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		End Select
    '		're-enable our button
    '		cmdExit.Enabled = True
    '		cmdLogin.Enabled = True
    '		moMsgSys.DisconnectAll()
    '		moMsgSys = Nothing
    '		If goUILib Is Nothing = False Then
    '			goUILib.RetainTooltip = False
    '			goUILib.SetToolTip(False)
    '		End If

    '		Exit Sub
    '	End If

    '	goUILib.SetToolTip("Login Successful, getting player details...", -1, lStatusTop)

    '	Dim oINI As New InitFile()
    '	Dim sTemp As String = oINI.GetString("SIGNON", "LastUserName", "")
    '	If sTemp.ToUpper <> sUserNameProper.ToUpper Then
    '		oINI.WriteString("SIGNON", "LastUserName", sUserNameProper)
    '	End If
    '	oINI = Nothing

    '	goUILib.AddNotification("Welcome to Beyond Protocol!", muSettings.InterfaceBorderColor, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

    '	'Ok, so our login is good... let's do this, first remove our window (need to release our handlers)
    '	RemoveHandler frmLogin.StartLogin, AddressOf Login_StartLogin
    '	RemoveHandler frmLogin.ExitProgram, AddressOf Login_ExitProgram
    '	goUILib.RemoveWindow(frmLogin.ControlName)

    '	'If we're here, then we need to clean up our background for the Login Screen, we do this by setting the 
    '	'  startup dialog time
    '	moEngine.swStartup.Reset()
    '	moEngine.swStartup = Nothing

    '	'Now... request our stuff...
    '	goUILib.SetToolTip("Requesting Universe Data...", -1, lStatusTop)

    '	'Application.DoEvents()
    '	'Threading.Thread.Sleep(10)

    '	'moMsgSys.RequestStarTypes()
    '	LoadStarTypes(moEngine.GetDevice)
    '	moMsgSys.RequestGalaxyAndSystems()

    '	'TODO: Remarked this out for now... was causing issues.
    '	'moMsgSys.RequestSkillList()
    '	'moMsgSys.RequestGoalList()

    '	'Now, request our player details
    '	goUILib.SetToolTip("Requesting Player Details...", -1, lStatusTop)
    '	moMsgSys.RequestPlayerDetails()

    '	'Ok, server knows who we are and I have my data, so let's continue
    '	ReDim yData(5)
    '	goUILib.SetToolTip("Login Completed Entering Universe...", -1, lStatusTop)
    '	System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerEnvironment).CopyTo(yData, 0)
    '	System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, 2)
    '	moMsgSys.SendToPrimary(yData)

    '	mbChangingEnvirs = True

    '	'Oh, and finally, set our music play list
    '	If goSound Is Nothing = False Then
    '		goSound.PlayListType = 2 'TODO: 0 = login, 1 = lull, 2 = excited
    '	End If

    '	'Show our Chat window...
    '	Dim oTmpCht As frmChat
    '	oTmpCht = New frmChat(goUILib)
    '	goUILib.AddWindow(oTmpCht)
    '	oTmpCht = Nothing

    '	'Ok, for now, attempt to join testers
    '	Dim sNewVal As String = "/join Testers"
    '	ReDim yData(sNewVal.Length + 1)
    '	System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(yData, 2)
    '	System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yData, 0)
    '	moMsgSys.SendToPrimary(yData)

    '	If goUILib Is Nothing = False Then goUILib.RetainTooltip = False
    'End Sub

    Public Sub Login_StartLogin(ByVal sUserName As String, ByVal sPassword As String, ByVal bAlias As Boolean, ByVal sAliasUsername As String, ByVal sAliasPassword As String, ByRef frmLogin As frmLoginDlg)
        'this sub handles the login form's Login Event
        Dim lRes As Int32
        Dim lStatusTop As Int32
        If frmLogin Is Nothing = False Then
            lStatusTop = frmLogin.Top + frmLogin.Height + 5
        Else
            lStatusTop = Me.Height \ 2
        End If
        Dim cmdLogin As UIButton = Nothing
        Dim cmdExit As UIButton = Nothing

        mbEnvirReady = False

        gsUserNameProper = sUserName
        sUserName = UCase$(sUserName)
        sPassword = UCase$(sPassword)
        sAliasUsername = sAliasUsername.ToUpper()
        sAliasPassword = sAliasPassword.ToUpper()
        gsUserName = sUserName
        gsPassword = sPassword
        gsAliasPassword = sAliasPassword
        gsAliasUserName = sAliasUsername
        'If glPlayerID = glActualPlayerID Then gbAliased = False
        gbAliased = bAlias

        'Dim yEncryptedPassword() As Byte = EncBytes(System.Text.ASCIIEncoding.ASCII.GetBytes(sPassword))
        'Dim yEncryptedAliasPW() As Byte = EncBytes(System.Text.ASCIIEncoding.ASCII.GetBytes(sAliasPassword))

        If sUserName = "DEBUG" Then
            If sPassword = "LOGIN_OFF" Then
                LoginScreen.mfFadeoutAlpha = 255.0F
                Return
            End If
        End If

        If goUILib Is Nothing = False Then goUILib.RetainTooltip = True

        'find the Login Windows Login and Exit control's
        If frmLogin Is Nothing = False Then
            For lRes = 0 To frmLogin.ChildrenUB
                If frmLogin.moChildren(lRes).ControlName = "cmdLogin" Then
                    cmdLogin = CType(frmLogin.moChildren(lRes), UIButton)
                ElseIf frmLogin.moChildren(lRes).ControlName = "cmdExit" Then
                    cmdExit = CType(frmLogin.moChildren(lRes), UIButton)
                End If
            Next lRes
            'disable it
            cmdLogin.Enabled = False
        End If

        'Application.DoEvents()

        'Ok, let's log in... check if we are connected
        If moMsgSys Is Nothing = False Then
            moMsgSys.DisconnectAll()
            moMsgSys = Nothing
            If goUILib Is Nothing = False Then goUILib.SetMsgSystem(Nothing)
        End If

        'Now, disable our exit button
        If cmdExit Is Nothing = False Then cmdExit.Enabled = False

        goUILib.SetToolTip("Connecting to Server...", -1, lStatusTop)
        mlLastKeepAliveSend = timeGetTime
        moMsgSys = New MsgSystem()
        If goUILib Is Nothing = False Then goUILib.SetMsgSystem(moMsgSys)

        'NEW LOGIN PROCEDURE
        '1) Connect to Operator Server
        mlConnectToOperatorTimer = timeGetTime
        mbConnectingToServer = True
        If moMsgSys.ConnectToOperator() = False Then
            goUILib.AddNotification("Could not connect to server.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            goUILib.AddNotification("The server is either down or you", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            goUILib.AddNotification("are not connected to the internet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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
            Exit Sub
        End If

        GFXEngine.bRenderInProgress = True

    End Sub
    Private mlConnectToOperatorTimer As Int32 = Int32.MinValue

    Private mbInTimer1Tick As Boolean = False
    'Private mlLastAutoFire As Int32 = 0
    'Private moRand As Random
    Public Shared bTutRelog As Boolean = False
    Dim mbRequestClaimables As Boolean = False

    Private mbRecordIt As Boolean = False
    Private moFS As IO.FileStream = Nothing
    Private moWrite As IO.StreamWriter = Nothing
    Private mlLastWriteCycle As Int32
    Private moUnit As BaseEntity = Nothing

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Dim X As Int32
        'Dim sFinal As String

        If bTutRelog = True Then
            bTutRelog = False
            glCurrentEnvirView = CurrentView.eStartupLogin
            goUILib.RemoveAllWindows()
            goUILib.GetMsgSys.DisconnectAll()
            glPlayerIntelUB = -1
            glPlayerTechKnowledgeUB = -1
            glItemIntelUB = -1
            glEntityDefUB = -1
            glMineralPropertyUB = -1
            glMineralUB = -1
            mbFirstLoad = True

            goCurrentPlayer = Nothing
            goCurrentEnvir = Nothing

            Application.DoEvents()
            bForceRequestDetails = True
            Login_StartLogin(gsUserName, gsPassword, False, "", "", Nothing)
            Return
        End If

        mbInTimer1Tick = True

        If mbRecordIt = True Then
            If moFS Is Nothing = True Then moFS = New IO.FileStream("C:\test.txt", IO.FileMode.Create)
            If moWrite Is Nothing Then
                moWrite = New IO.StreamWriter(moFS)
                moWrite.AutoFlush = True
            End If

            If glCurrentCycle <> mlLastWriteCycle Then
                Dim lCycleDiff As Int32 = glCurrentCycle - mlLastWriteCycle

                If moUnit Is Nothing Then
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) > 0 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                            moUnit = goCurrentEnvir.oEntity(X)
                            Exit For
                        End If
                    Next X
                End If
                
                moWrite.WriteLine(lCycleDiff.ToString() & "|" & gfCurrentCyclePrecise.ToString() & "|" & moUnit.LocX.ToString() & "|" & moUnit.LocZ.ToString() & "|" & moUnit.TotalVelocity.ToString)
                mlLastWriteCycle = glCurrentCycle
            End If
        Else
            If moWrite Is Nothing = False Then
                moWrite.Close()
                moWrite.Dispose()
                moWrite = Nothing
            End If
            If moFS Is Nothing = False Then
                moFS.Close()
                moFS.Dispose()
                moFS = Nothing
            End If
        End If

        'If glCurrentCycle Mod 4 = 0 AndAlso mlLastAutoFire <> glCurrentCycle Then
        '    mlLastAutoFire = glCurrentCycle
        '    Dim lIdxs() As Int32 = Nothing
        '    Dim lIdxUB As Int32 = -1
        '    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
        '        If goCurrentEnvir.lEntityIdx(X) > -1 Then
        '            Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
        '            If oEntity.bCulled = False AndAlso oEntity.ObjTypeID = ObjectType.eUnit Then
        '                lIdxUB += 1
        '                ReDim Preserve lIdxs(lIdxUB)
        '                lIdxs(lIdxUB) = X
        '                If lIdxUB = 1 Then Exit For
        '            End If
        '        End If
        '    Next X
        '    Dim oRand As New Random
        '    For X As Int32 = 0 To lIdxUB
        '        For Y As Int32 = 0 To lIdxUB
        '            If X <> Y Then
        '                Dim yType As Byte = CByte(oRand.Next(0, 4))
        '                goWpnMgr.AddNewEffect(goCurrentEnvir.oEntity(lIdxs(X)), goCurrentEnvir.oEntity(lIdxs(Y)), yType, True, False)
        '            End If
        '        Next Y
        '        'Exit For
        '    Next X
        'End If

        'If goSound Is Nothing = False AndAlso glCurrentEnvirView = CurrentView.eStartupLogin Then
        '    If moRand Is Nothing Then moRand = New Random()

        '    Dim lRoll As Int32 = moRand.Next(0, 100)

        '    If lRoll < 40 Then
        '        Dim yWpnType As Int32 = moRand.Next(0, WeaponType.eMetallicProjectile_Copper + 1)
        '        Dim sSoundFile As String
        '        Dim ySFXType As SoundMgr.SoundUsage
        '        Select Case yWpnType
        '            Case WeaponType.eSolidGreenBeam, WeaponType.eSolidPurpleBeam, WeaponType.eSolidRedBeam
        '                sSoundFile = "Beam2a.wav"
        '                'sSoundFile = "Energy Weapons\Beam2a.wav"
        '                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
        '            Case WeaponType.eSolidTealBeam, WeaponType.eSolidYellowBeam
        '                sSoundFile = "Beam2b.wav"
        '                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
        '            Case WeaponType.eFlickerGreenBeam, WeaponType.eFlickerPurpleBeam, WeaponType.eFlickerRedBeam, _
        '              WeaponType.eFlickerTealBeam, WeaponType.eFlickerYellowBeam
        '                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
        '                sSoundFile = "Beam1" & Chr(lsub) & ".wav"
        '                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
        '            Case WeaponType.eMetallicProjectile_Bronze, WeaponType.eMetallicProjectile_Copper, _
        '              WeaponType.eMetallicProjectile_Gold, WeaponType.eMetallicProjectile_Lead, WeaponType.eMetallicProjectile_Silver
        '                Dim lTemp As Int32 = moRand.Next(1, 5)
        '                sSoundFile = "LargeProj" & lTemp & ".wav"
        '                ySFXType = SoundMgr.SoundUsage.eWeaponsFireProjectile
        '            Case Else
        '                Dim lTemp As Int32 = CInt(yWpnType) + 1
        '                Dim lsub As Int32 = Asc("a") + moRand.Next(0, 3)
        '                sSoundFile = "Laser" & lTemp & Chr(lsub) & ".wav"
        '                ySFXType = SoundMgr.SoundUsage.eWeaponsFireEnergy
        '        End Select

        '        Dim vec As Vector3 = New Vector3(Rnd() * 1000 - 500, Rnd() * 1000 - 500, Rnd() * 1000 - 500)
        '        vec.X += goCamera.mlCameraX
        '        vec.Y += goCamera.mlCameraY
        '        vec.Z += goCamera.mlCameraZ
        '        goSound.StartSound(sSoundFile, False, ySFXType, vec, New Vector3(0, 0, 0))
        '    ElseIf lRoll < 80 Then
        '        Dim vec As Vector3 = New Vector3(Rnd() * 1000 - 500, Rnd() * 1000 - 500, Rnd() * 1000 - 500)
        '        vec.X += goCamera.mlCameraX
        '        vec.Y += goCamera.mlCameraY
        '        vec.Z += goCamera.mlCameraZ
        '        goSound.StartSound("Explosions\HullHit1.wav", True, SoundMgr.SoundUsage.eUnitSounds, vec, New Vector3(0, 0, 0))
        '        vec = New Vector3(Rnd() * 1000 - 500, Rnd() * 1000 - 500, Rnd() * 1000 - 500)
        '        vec.X += goCamera.mlCameraX
        '        vec.Y += goCamera.mlCameraY
        '        vec.Z += goCamera.mlCameraZ
        '        goSound.StartSound("Explosions\HullHit2.wav", True, SoundMgr.SoundUsage.eUnitSounds, vec, New Vector3(0, 0, 0))
        '    ElseIf lRoll < 90 Then
        '        Dim lIdx As Int32 = moRand.Next(1, 9)
        '        goSound.StartSound("RC" & lIdx & ".wav", False, SoundMgr.SoundUsage.eRadioChatter, New Vector3(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ), New Vector3(0, 0, 0))
        '    Else
        '        Dim vec As Vector3 = New Vector3(Rnd() * 1000 - 500, Rnd() * 1000 - 500, Rnd() * 1000 - 500)
        '        vec.X += goCamera.mlCameraX
        '        vec.Y += goCamera.mlCameraY
        '        vec.Z += goCamera.mlCameraZ
        '        goSound.StartSound("Unit Sounds\MediumShipRoar4.wav", True, SoundMgr.SoundUsage.eUnitSounds, vec, New Vector3(0, 0, 0))
        '        vec = New Vector3(Rnd() * 1000 - 500, Rnd() * 1000 - 500, Rnd() * 1000 - 500)
        '        vec.X += goCamera.mlCameraX
        '        vec.Y += goCamera.mlCameraY
        '        vec.Z += goCamera.mlCameraZ
        '        goSound.StartSound("Unit Sounds\LargeshipRoar1.wav", True, SoundMgr.SoundUsage.eUnitSounds, vec, New Vector3(0, 0, 0))
        '    End If


        'End If

        'While mbMainThread = True
        Static xlLastCycleUpdate As Int32 = timeGetTime     'timegettime is called only once
        Static xlFPSTrack As Int32 = timeGetTime
        Static xlFPS As Int32

        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.lTutorialStep > 316 AndAlso goCurrentPlayer.yPlayerPhase = 1 Then
            If glCurrentCycle - NewTutorialManager.lLastYouMayLeaveAlert > 18000 Then
                goUILib.AddNotification("You may go to live by pressing Escape at any time.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                NewTutorialManager.lLastYouMayLeaveAlert = glCurrentCycle
            End If
        End If

        If moMsgSys Is Nothing = False Then
            If mbNeedConnectToPrimary = True Then
                mbNeedConnectToPrimary = False
                If moMsgSys.ConnectToPrimary() = False Then
                    goUILib.AddNotification("Unable to establish connection, please try again later.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            End If

            'If mbConnectingToServer = True AndAlso mlConnectToOperatorTimer <> Int32.MinValue AndAlso timeGetTime - mlConnectToOperatorTimer > 30000 Then
            '    mlConnectToOperatorTimer = timeGetTime
            '    If moMsgSys.GiveUpAndConnectToBackupOperator() = False Then
            '        goUILib.AddNotification("Unable to establish connection, please try again later.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '        mlConnectToOperatorTimer = Int32.MinValue
            '    End If
            'End If

            If timeGetTime - mlLastKeepAliveSend > 5000 Then
                If moMsgSys.SendKeepAliveMsgs() = False Then
                    Timer1.Enabled = False
                    mbInTimer1Tick = False
                    moMsgSys_ServerShutdown()
                    Return
                End If
                mlLastKeepAliveSend = timeGetTime
            End If
        End If

        If mbEnvirReady = False AndAlso glCurrentEnvirView < CurrentView.eFullScreenInterface Then
            ' we still want to render our logo
            GFXEngine.bRenderInProgress = True
            moEngine.RenderOnlyLogo()
            mbInTimer1Tick = False
            Return
        ElseIf glCurrentEnvirView < CurrentView.eFullScreenInterface Then
            If GFXEngine.bRenderInProgress = True Then
                mbDoMineralPropRequests = True
            End If
            GFXEngine.bRenderInProgress = False
        End If

        If glCurrentEnvirView = CurrentView.eStartupLogin Then Timer1.Enabled = False
        'If glCurrentEnvirView = CurrentView.eStartupDSELogo Then
        '    If moEngine.mswDSELogoTimer Is Nothing = False Then
        '        If moEngine.mswDSELogoTimer.ElapsedMilliseconds > 10000 Then
        '            moEngine.ClearVideo()
        '            Return
        '        End If
        '    End If
        'End If

        If mbEnvirReady = True AndAlso mbNeedToLoadTacticalData = True Then
            mbNeedToLoadTacticalData = False
            If goCurrentEnvir Is Nothing = False Then goCurrentEnvir.LoadEnvironmentTacticalData()
        End If

        If glCurrentCycle - mlLastUIEvent > 54000 AndAlso NewTutorialManager.TutorialOn = False Then
            'If glPlayerID <> 221 AndAlso glPlayerID <> 131 AndAlso glPlayerID <> 2 AndAlso glPlayerID <> 6 AndAlso glPlayerID <> 1 AndAlso glPlayerID <> 2067 AndAlso glPlayerID <> 3253 AndAlso glPlayerID <> 2076 AndAlso glPlayerID <> 3510 Then
            If isAdmin() = False Then
                'forcefully disconnect
                Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
                If sFile.EndsWith("\") = False Then sFile &= "\"
                sFile &= "UpdaterClient.exe"
                If Exists(sFile) Then
                    Shell(sFile, AppWinStyle.NormalFocus, False, -1)
                End If

                mbfrmConfirmHandled = True
                mbInTimer1Tick = False
                Me.Timer1.Enabled = False
                Me.Close()
                Return
            Else
                mlLastUIEvent = glCurrentCycle
            End If
        End If

        If timeGetTime - xlLastCycleUpdate > 30 Then
            'glCurrentCycle += CInt(Math.Ceiling((timeGetTime - xlLastCycleUpdate) / 30))
            '...
            'timegettime - xllastcycleupdate  - number of milliseconds between last update and now
            ' / 30 = number of cycles (float form) lapsed...
            'we do a math.ceiling on it... which always causes it to raise by 1... so if it was 30.1 then we would increment by 2
            ' instead of math.ceiling, let's just assume CINT will do it for us...
            gfCurrentCyclePrecise += ((timeGetTime - xlLastCycleUpdate) / 30.0F)
            glCurrentCycle = CInt(Math.Floor(gfCurrentCyclePrecise))
            If goUILib Is Nothing = False Then goUILib.ReprocessInputs()
            xlLastCycleUpdate = timeGetTime
        End If

        If Me.WindowState <> FormWindowState.Minimized AndAlso GFXEngine.gbDeviceLost = False AndAlso GFXEngine.gbPaused = False Then
            If gbMonitorPerformance = True Then gsw_Camera.Start()
            goCamera.ScrollCamera(glCurrentEnvirView)

            CheckMineralCacheHover()

            If glCurrentEnvirView = CurrentView.eSystemView Then
                If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                    '40k is from the max entity clip plane settable by the client
                    '125 would be half of (10,000,000) / 40000 which is 250
                    Dim lTmpSectX As Int32 = (goCamera.mlCameraX \ 40000) + 125
                    Dim lTmpSectZ As Int32 = (goCamera.mlCameraZ \ 40000) + 125

                    'now, clamp to 0 - 250
                    If lTmpSectX < 0 Then lTmpSectX = 0
                    If lTmpSectX > 250 Then lTmpSectX = 250
                    If lTmpSectZ < 0 Then lTmpSectZ = 0
                    If lTmpSectZ > 250 Then lTmpSectZ = 250
                    If lTmpSectX <> glCameraSectorX OrElse lTmpSectZ <> glCameraSectorZ Then
                        glCameraSectorX = lTmpSectX
                        glCameraSectorZ = lTmpSectZ
                        Dim yMsg(3) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eCameraPosUpdate).CopyTo(yMsg, 0)
                        yMsg(2) = CByte(lTmpSectX)
                        yMsg(3) = CByte(lTmpSectZ)
                        moMsgSys.SendToRegion(yMsg)
                    End If
                End If
            End If

            goCamera.SetupMatrices(GFXEngine.moDevice, glCurrentEnvirView)
            If gbMonitorPerformance = True Then gsw_Camera.Stop()
        End If

        If goCurrentEnvir Is Nothing = False Then
            If gbMonitorPerformance = True Then gsw_Movement.Start()
            Try
                HandleMovement()
                goCurrentEnvir.UpdateEventAmbience()
            Catch
            End Try
            If gbMonitorPerformance = True Then gsw_Movement.Stop()

            If Me.WindowState <> FormWindowState.Minimized AndAlso GFXEngine.gbDeviceLost = False AndAlso GFXEngine.gbPaused = False Then
                If goCurrentPlayer Is Nothing = False Then
                    If goCurrentPlayer.lCelebrationPeriodEnd > glCurrentCycle AndAlso Rnd() * 1000 < 15 Then
                        'are we in a space or planet envir?
                        If goFireworks Is Nothing Then goFireworks = New FireworksMgr()
                        Dim yType As Byte = CByte(Rnd() * 2)
                        If glCurrentEnvirView = CurrentView.ePlanetView Then
                            If Planet.fDayTimeRatio < 0.3F OrElse Planet.fDayTimeRatio > 0.708333F Then
                                'ok, find the command center
                                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eCommandCenterSpecial Then
                                        With goCurrentEnvir.oEntity(X)
                                            goFireworks.AddEmitter(New Vector3(.LocX, .LocY, .LocZ), yType, CInt(.LocY + 3000))
                                        End With
                                        Exit For
                                    End If
                                Next X
                            End If
                        End If
                    End If
                End If
                For X As Int32 = 0 To glPlayerIntelUB
                    If glPlayerIntelIdx(X) <> -1 AndAlso goPlayerIntel(X).lCelebrationEnds > glCurrentCycle AndAlso Rnd() * 1000 < 15 Then
                        'are we in a space or planet envir?
                        If goFireworks Is Nothing Then goFireworks = New FireworksMgr()
                        Dim yType As Byte = CByte(Rnd() * 2)
                        Dim lPlayerID As Int32 = glPlayerIntelIdx(X)
                        If glCurrentEnvirView = CurrentView.ePlanetView Then
                            If Planet.fDayTimeRatio < 0.3F OrElse Planet.fDayTimeRatio > 0.708333F Then
                                'ok, find the command center
                                For Y As Int32 = 0 To goCurrentEnvir.lEntityUB
                                    If goCurrentEnvir.lEntityIdx(Y) <> -1 AndAlso goCurrentEnvir.oEntity(Y).OwnerID = lPlayerID AndAlso goCurrentEnvir.oEntity(Y).yProductionType = ProductionType.eCommandCenterSpecial Then
                                        With goCurrentEnvir.oEntity(Y)
                                            goFireworks.AddEmitter(New Vector3(.LocX, .LocY, .LocZ), yType, CInt(.LocY + 3000))
                                        End With
                                        Exit For
                                    End If
                                Next Y
                            End If
                        End If
                    End If
                Next X
            End If
        End If

        If Me.WindowState <> FormWindowState.Minimized Then
            ''Ok, before drawing the scene, let's set up our fow textures
            'If goCurrentEnvir Is Nothing = False AndAlso glCurrentEnvirView = CurrentView.ePlanetView Then
            '	If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
            '		Dim oGUID As Base_GUID = CType(goCurrentEnvir.oGeoObject, Base_GUID)
            '		If oGUID Is Nothing = False AndAlso oGUID.ObjTypeID = ObjectType.ePlanet Then
            '			Dim oPlanet As Planet = CType(goCurrentEnvir.oGeoObject, Planet)
            '			If oPlanet Is Nothing = False Then
            '				oPlanet.UpdateTerrainFOW()
            '			End If
            '		End If
            '	End If
            'End If
            'If Me.Focused = True Then glModelsRendered = moEngine.DrawScene()
            Try
                glModelsRendered = moEngine.DrawScene()
            Catch ex As Exception
                LogCrashEvent(ex, False, True, "DrawScene (" & moEngine.sDrawSceneLocation & ")", False, moMsgSys)
                mbCriticalFailure = True
            End Try

            If mbCriticalFailure = True Then
                mbCriticalFailure = False
                Timer1.Enabled = False

                muSettings.SaveSettings()
                If moEngine.RecreateEverything(Me, moMsgSys, True) = False Then
                    'force another activity report through
                    Dim lAvgFPS As Int32 = 65
                    If mlLastFrameRateMonitorCheck > 0 Then lAvgFPS = mlRunningFPSMonitor \ mlLastFrameRateMonitorCheck
                    Dim yStatus() As Byte = BPMetrics.MetricMgr.GetActivityReport(lAvgFPS)
                    If yStatus Is Nothing = False AndAlso goCurrentPlayer Is Nothing = False Then
                        Try
                            moMsgSys.SendToPrimary(yStatus)
                        Catch
                        End Try
                    End If
                    mlLastStatusUpdateToPrimary = glCurrentCycle
                    'end of the force

                    If moMsgSys Is Nothing = False Then moMsgSys.DisconnectAll()
                    If goResMgr Is Nothing = False Then goResMgr.ClearAllResources()
                    If moEngine Is Nothing = False Then moEngine.ForcefulCleanup()
                    moEngine = Nothing
                    goResMgr = Nothing

                    MsgBox("You will need to restart your client in order for the settings to take effect.", MsgBoxStyle.OkOnly, "Settings Changed")

                    mbfrmConfirmHandled = True
                    Me.Close()

                    Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                    If sPath.EndsWith("\") = False Then sPath &= "\"
                    sPath &= "UpdaterClient.exe"
                    If Exists(sPath) = True Then Shell(sPath, AppWinStyle.NormalFocus, False, -1)
                    Application.Exit()
                    End
                End If

                Timer1.Enabled = True

                Return
            End If

            If mbDoMineralPropRequests = True AndAlso glCurrentCycle - mlLastMineralPropRequest > 10 Then
                mbDoMineralPropRequests = False
                mlLastMineralPropRequest = glCurrentCycle
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) <> -1 Then
                        If goMinerals(X).bRequestedProps = False AndAlso goMinerals(X).bDiscovered = True Then
                            goMinerals(X).lLastMsgUpdate = 0
                            goMinerals(X).bRequestedProps = True
                            Dim yMsg(5) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eRequestMineral).CopyTo(yMsg, 0)
                            System.BitConverter.GetBytes(goMinerals(X).ObjectID).CopyTo(yMsg, 2)

                            Try
                                moMsgSys.SendToPrimary(yMsg)
                            Catch
                                goMinerals(X).bRequestedProps = False
                            End Try

                            mbDoMineralPropRequests = True
                            Exit For
                        End If
                    End If
                Next X

                If mbDoMineralPropRequests = False Then
                    For X As Int32 = 0 To glMineralUB
                        If glMineralIdx(X) <> -1 Then
                            If goMinerals(X).HasNonZeroProperty() = False AndAlso goMinerals(X).bDiscovered = True Then
                                mbDoMineralPropRequests = True
                                goMinerals(X).bRequestedProps = False
                            End If
                        End If
                    Next X
                End If
            End If

            If mbRequestClaimables = True Then
                'NOTE: This should only be called one time...
                mbRequestClaimables = False
                Dim yMsg(11) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eClaimItem).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(-1S).CopyTo(yMsg, 6)
                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 8)
                If moMsgSys Is Nothing = False Then moMsgSys.SendToPrimary(yMsg)
            End If

            If moMsgSys Is Nothing = False Then moMsgSys.CheckDXDiagProcess()
        End If

        xlFPS += 1
        If timeGetTime - xlFPSTrack > 1000 Then

            mlRunningFPSMonitor += xlFPS
            mlLastFrameRateMonitorCheck += 1

            If mlLastFrameRateMonitorCheck > 9 Then

                Dim lAvgFPS As Int32 = mlRunningFPSMonitor \ mlLastFrameRateMonitorCheck
                mlRunningFPSMonitor = 0
                mlLastFrameRateMonitorCheck = 0
                If goPFXEngine32 Is Nothing = False Then
                    If lAvgFPS < 30 Then
                        goPFXEngine32.IncreaseLimitOnParticles(lAvgFPS)
                    ElseIf lAvgFPS > 50 Then
                        goPFXEngine32.DecreaseLimitOnParticles()
                    End If
                End If
            End If


            xlFPS = 0
            xlFPSTrack = timeGetTime

        End If

        If glCurrentCycle - mlLastStatusUpdateToPrimary > 300 AndAlso goUILib Is Nothing = False AndAlso moMsgSys Is Nothing = False Then
            Dim lAvgFPS As Int32 = 65
            If mlLastFrameRateMonitorCheck > 0 Then lAvgFPS = mlRunningFPSMonitor \ mlLastFrameRateMonitorCheck
            Dim yStatus() As Byte = BPMetrics.MetricMgr.GetActivityReport(lAvgFPS)
            If yStatus Is Nothing = False AndAlso goCurrentPlayer Is Nothing = False Then
                Try
                    moMsgSys.SendToPrimary(yStatus)
                Catch
                End Try
                'If goUILib Is Nothing = False Then goUILib.AddNotification("Activity Report Sent", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
            mlLastStatusUpdateToPrimary = glCurrentCycle

            If mbCheckedErrorReport = False Then CheckErrorReport()
        End If
        ProcessGetCacheEntries()

        If goGalaxy Is Nothing = False Then goGalaxy.ExecuteJumpToPlanet()

        If bRestartWithUpdater = True Then
            Timer1.Enabled = False
            If moEngine Is Nothing = False Then moEngine.RecreateEverything(Me, moMsgSys, False)
            If moMsgSys Is Nothing = False Then moMsgSys.DisconnectAll()
            If goResMgr Is Nothing = False Then goResMgr.ClearAllResources()
            moEngine = Nothing
            goResMgr = Nothing
            mbInTimer1Tick = False
            mbfrmConfirmHandled = True
            Me.Close()

            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            sPath &= "UpdaterClient.exe"
            If Exists(sPath) = True Then Shell(sPath, AppWinStyle.NormalFocus, False, -1)
            Application.Exit()
            End
        End If

        If Timer1.Enabled = False AndAlso mbFormClosing = False AndAlso mbInFormClosing = False Then Timer1.Enabled = True
        'Threading.Thread.Sleep(1) 

        'End While
        mbInTimer1Tick = False
    End Sub

    Private mbCheckedErrorReport As Boolean = False
    Private Sub CheckErrorReport()
        Dim sDirectory As String = AppDomain.CurrentDomain.BaseDirectory
        If sDirectory.EndsWith("\") = True Then sDirectory = sDirectory.Substring(0, sDirectory.Length - 1)
        Dim colResults As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(sDirectory, FileIO.SearchOption.SearchAllSubDirectories, "*.txt")

        'Dim blResult As Int64 = 0

        mbCheckedErrorReport = True
        For Each sValue As String In colResults
            Dim sFileName As String = sValue.Replace(sDirectory, "").Replace("\", "")
            If sFileName.ToUpper.StartsWith("CRASHLOG") = True Then
                mbCheckedErrorReport = False

                Dim oFS As New IO.FileStream(sValue, IO.FileMode.Open)
                Dim oRead As New IO.StreamReader(oFS)
                Dim sVal As String = oRead.ReadToEnd
                oRead.Close()
                oRead.Dispose()
                oFS.Close()
                oFS.Dispose()

                If sVal.Length > 25000 Then sVal = sVal.Substring(0, 25000)
                Dim yMsg(5 + sVal.Length) As Byte

                System.BitConverter.GetBytes(GlobalMessageCode.eExceptionReport).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(sVal.Length).CopyTo(yMsg, 2)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sVal).CopyTo(yMsg, 6)
                Try
                    moMsgSys.SendToPrimary(yMsg)
                    Kill(sValue)
                Catch
                End Try
            End If
        Next

    End Sub

    Private mlLastStatusUpdateToPrimary As Int32 = 0

    Private mlLastFrameRateMonitorCheck As Int32 = 0
    Private mlRunningFPSMonitor As Int32 = 0


    Public mbDoMineralPropRequests As Boolean = False
    Public mlLastMineralPropRequest As Int32 = 0

    Private mbCriticalFailure As Boolean = False
    Private Sub CriticalFailure() Handles moEngine.CriticalFailure
        mbCriticalFailure = True
    End Sub

    Private Sub CleanupOldFiles()
        'Return
        Try
            Dim sPath As String
            sPath = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"

            Dim colResults As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(sPath & "Audio\SoundFX\Ambience", FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            For Each sValue As String In colResults
                If sValue.ToUpper.EndsWith("PAK") = False Then
                    Kill(sValue)
                End If
            Next

            colResults = My.Computer.FileSystem.GetFiles(sPath & "Audio\SoundFX\Energy Weapons", FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            For Each svalue As String In colResults
                If svalue.ToUpper.EndsWith("PAK") = False Then
                    Kill(svalue)
                End If
            Next

            colResults = My.Computer.FileSystem.GetFiles(sPath & "Audio\SoundFX\Missiles", FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            For Each svalue As String In colResults
                If svalue.ToUpper.EndsWith("PAK") = False Then
                    Kill(svalue)
                End If
            Next

            colResults = My.Computer.FileSystem.GetFiles(sPath & "Audio\SoundFX\Projectile Weapons", FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            For Each svalue As String In colResults
                If svalue.ToUpper.EndsWith("PAK") = False AndAlso svalue.ToUpper.EndsWith("FASTREPEATSMALL.WAV") = False Then
                    Kill(svalue)
                End If
            Next

            colResults = My.Computer.FileSystem.GetFiles(sPath & "Audio\SoundFX\RadioChatter", FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            For Each svalue As String In colResults
                If svalue.ToUpper.EndsWith("PAK") = False Then
                    Kill(svalue)
                End If
            Next

            colResults = My.Computer.FileSystem.GetFiles(sPath & "Audio\SoundFX\Weather", FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            For Each svalue As String In colResults
                If svalue.ToUpper.EndsWith("PAK") = False Then
                    Kill(svalue)
                End If
            Next
            'sPath &= "Audio\"
            '= sPath & "SoundFX\"
            '= sPath & "Music\"


        Catch
        End Try
    End Sub

    Private Sub frmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        'Dim oINI As New InitFile
        'Dim bShowVidPlayerBox As Boolean = CInt(oINI.GetString("SETTINGS", "ShowVidPlayerBox", "1")) <> 0
        'If bShowVidPlayerBox = True Then
        '    MsgBox("We are attempting to track down a problem with our video player. If you" & vbCrLf & _
        '           "are unable to proceed beyond the Alienware video, please get on live chat" & vbCrLf & _
        '           "and let us know. Live chat is accessible from the updater client's news section.", MsgBoxStyle.OkOnly, "Video Testing")
        '    oINI.WriteString("SETTINGS", "ShowVidPlayerBox", "0")
        'End If
        'oINI = Nothing
        'Dim oT As New TerrainClass(Nothing, 168)
        'oT.MapType = 4
        'oT.CellSpacing = gl_MEDIUM_PLANET_CELL_SPACING
        'oT.GenerateTerrain(168)
        'oT.SaveToFile(168)
        CleanupOldFiles()

        Dim oINI As New InitFile()
        Dim lWinLeft As Int32 = CInt(Val(oINI.GetString("WindowMode", "Left", "0")))
        Dim lWinTop As Int32 = CInt(Val(oINI.GetString("WindowMode", "Top", "0")))
        Dim lWinWidth As Int32 = CInt(Val(oINI.GetString("WindowMode", "Width", Manager.Adapters.Default.CurrentDisplayMode.Width.ToString)))
        Dim lWinHeight As Int32 = CInt(Val(oINI.GetString("WindowMode", "Height", Manager.Adapters.Default.CurrentDisplayMode.Height.ToString)))

        Me.Width = lWinWidth
        Me.Height = lWinHeight
        Me.Left = lWinLeft
        Me.Top = lWinTop

        oINI.WriteString("WindowMode", "Left", lWinLeft.ToString)
        oINI.WriteString("WindowMode", "Top", lWinTop.ToString)
        oINI.WriteString("WindowMode", "Width", lWinWidth.ToString)
        oINI.WriteString("WindowMode", "Height", lWinHeight.ToString)

        'TODO: remove this hack and get cache entries working
        HACK_SetCacheEntries()

        'Control Groups
        'goControlGroups = New ControlGroups()

        'Create our Graphics engine...
        goCamera = New Camera()
        moEngine = New GFXEngine()
        If moEngine.InitD3D(Me) = False Then
            MsgBox("Unable to initialize Direct3D!", MsgBoxStyle.OkOnly, "Error")
            End
        Else
            'create our resource manager (used for loading meshes and textures and ensuring that
            '  only one instance of a mesh or texture is loaded in order to conserve resources)
            goResMgr = New GFXResourceManager()

            'create the sound effects/music engine
            If muSettings.AudioOn = True Then
                goSound = New SoundMgr(Me)
                If muSettings.AudioOn = False Then goSound = Nothing
                If goSound Is Nothing = False AndAlso goSound.MusicOn = True Then
                    goSound.InitializeMusicPlayer()
                End If
            End If

            'create the User Interface engine
            goUILib = New UILib(GFXEngine.moDevice)
            'Create the Model Defs engine
            goModelDefs = New ModelDefs()

            LoadAgentData()

            'Initialize it now so we can get a faster response
            moEngine.InitializeLoginScreen()

            With goCamera
                .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                .mlCameraX = 0 : .mlCameraY = 0 : .mlCameraZ = -10000
            End With

            Dim oTmpWin As frmLoginDlg
            oTmpWin = New frmLoginDlg(goUILib)
            AddHandler oTmpWin.StartLogin, AddressOf Login_StartLogin
            AddHandler oTmpWin.ExitProgram, AddressOf Login_ExitProgram
            goUILib.AddWindow(oTmpWin)
            oTmpWin = Nothing

            glCurrentEnvirView = CurrentView.eStartupDSELogo

            Timer1.Enabled = True
            'moMainThread = New Threading.Thread(AddressOf Timer1_Tick)
            'moMainThread.Start()
        End If
    End Sub

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 5
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(13, 28)
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(992, 757)
        Me.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Beyond Protocol"
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        Me.ResumeLayout(False)

    End Sub

    'Private Sub InitializeComponent()
    '    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
    '    Me.SuspendLayout()
    '    '
    '    'frmMain
    '    '
    '    Me.ClientSize = New System.Drawing.Size(292, 273)
    '    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    '    Me.Name = "frmMain"
    '    Me.ResumeLayout(False)

    'End Sub

#End Region

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
        Select Case yConnType
            Case eyConnType.OperatorServer
                mlConnectToOperatorTimer = Int32.MinValue
                goUILib.SetToolTip("Validating Version...", -1, frmLoginDlg.lStatusTop)
                If moMsgSys.ValidateVersion(eyConnType.OperatorServer) = False Then           'pass 0 for the operator
                    moMsgSys.DisconnectAll()
                    moMsgSys = Nothing
                    If goUILib Is Nothing = False Then goUILib.SetMsgSystem(Nothing)

                    goUILib.AddNotification("Version incompatible with server!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    goUILib.AddNotification("Please launch the game through the updater.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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
                End If
            Case eyConnType.PrimaryServer
                If moMsgSys.ValidateVersion(eyConnType.PrimaryServer) = False Then         'pass 1 for the operator
                    moMsgSys.DisconnectAll()
                    moMsgSys = Nothing
                    If goUILib Is Nothing = False Then goUILib.SetMsgSystem(Nothing)

                    goUILib.AddNotification("Version incompatible with server!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    goUILib.AddNotification("Please launch the game through the updater.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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
            Case eyConnType.RegionServer
                Dim yData() As Byte

                'Me.Text = "Epica - Connected to Region"
                Debug.Write("Connected to Region" & vbCrLf)

                'send our Player ID
                ReDim yData(91) '91
                Dim lPos As Int32 = 0

                Dim oEnc As New StrEncDec()

                Dim yEncPassword() As Byte = oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(gsPassword))
                Dim yEncAPassword() As Byte = oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(gsAliasPassword))

                System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginRequest).CopyTo(yData, lPos) : lPos += 2
                System.BitConverter.GetBytes(glActualPlayerID).CopyTo(yData, lPos) : lPos += 4
                oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(gsUserName)).CopyTo(yData, lPos) : lPos += 20
                yEncPassword.CopyTo(yData, lPos) : lPos += 21
                System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, lPos) : lPos += 4
                oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(gsAliasUserName)).CopyTo(yData, lPos) : lPos += 20
                yEncAPassword.CopyTo(yData, lPos) : lPos += 21
                moMsgSys.SendToRegion(yData)

                ReDim yData(9)
                'Now, need to request the environment
                If goCurrentEnvir Is Nothing = False Then
                    'request the environment objects
                    System.BitConverter.GetBytes(GlobalMessageCode.eBurstEnvironmentRequest).CopyTo(yData, 0)
                    goCurrentEnvir.GetGUIDAsString.CopyTo(yData, 2)
                    moMsgSys.SendToRegion(yData)
                End If
            Case eyConnType.ChatServer
                '???
        End Select
    End Sub

    Private msOverrideIP As String = Nothing '= ""
    Private Sub moMsgSys_ReceivedEnvironmentsDomain(ByVal sIPAddress As String, ByVal iPort As Short) Handles moMsgSys.ReceivedEnvironmentsDomain
        Dim bNeedToConnect As Boolean

        If msOverrideIP Is Nothing Then
            Dim oINI As New InitFile()
            msOverrideIP = oINI.GetString("CONNECTION", "OverrideIP", "")
            oINI = Nothing
        End If
        If msOverrideIP <> "" Then sIPAddress = msOverrideIP '"epica.servegame.org"

        If goCurrentEnvir Is Nothing = False Then
            'ok, got the domain ip and port
            If goCurrentEnvir.DomainServerPort = 0 Then
                'not connnected period... let's connect
                bNeedToConnect = True
            ElseIf goCurrentEnvir.DomainServerIP = sIPAddress AndAlso goCurrentEnvir.DomainServerPort = iPort Then
                'we're already connected to the right domain... so we just need to notify the domain that we're ready
                bNeedToConnect = False
                'moMsgSys_ConnectedToRegion()
                moMsgSys_ConnectedToServer(eyConnType.RegionServer)
            Else
                'we're connected, but not to the right domain... so disconnect and reconnect to the right one
                moMsgSys.DisconnectRegion()
                bNeedToConnect = True
            End If

            If bNeedToConnect Then
                GFXEngine.bRenderInProgress = True
                If moMsgSys.ConnectToDomainServer(sIPAddress, iPort) = False Then
                    Me.Close()
                End If
            End If

            'ensure our values are set correctly
            goCurrentEnvir.DomainServerIP = sIPAddress
            goCurrentEnvir.DomainServerPort = iPort
        End If

        Debug.Write("Processed Environments Domain..." & vbCrLf)
    End Sub

    Private mbNeedConnectToPrimary As Boolean = False
    Private Sub moMsgSys_ReceivedPlayerCurrentEnvironment(ByVal lPlayerID As Integer, ByVal lID As Integer, ByVal iTypeID As Short, ByVal lGalaxyID As Int32, ByVal lSystemID As Int32, ByVal lPlanetID As Int32, ByVal bInCurrDomain As Boolean) Handles moMsgSys.ReceivedPlayerCurrentEnvironment
        'right now, should always be the case... but it might be interesting to allow a player to look up another player's
        '  current environment...
        Dim sPrevIP As String = ""
        Dim iPrevPort As Int16
        Dim X As Int32
        Dim bFirstTime As Boolean

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

        If lPlayerID = glPlayerID Then
            If goCurrentEnvir Is Nothing = False Then
                sPrevIP = goCurrentEnvir.DomainServerIP
                iPrevPort = goCurrentEnvir.DomainServerPort
                bFirstTime = False
                goCurrentEnvir.DisposeMe()
            Else
                bFirstTime = True
            End If
            goCurrentEnvir = Nothing
            If goUILib Is Nothing = False Then
                'goUILib.RemoveWindow("frmColonyStats")
                goUILib.ClearSelection()
            End If
            goCurrentEnvir = New BaseEnvironment()
            With goCurrentEnvir
                .ObjectID = lID
                .ObjTypeID = iTypeID
                .DomainServerIP = sPrevIP
                .DomainServerPort = iPrevPort
                .oGeoObject = Nothing
            End With

            If goPFXEngine32 Is Nothing = False Then goPFXEngine32.ClearAllEmitters()
            If goMissileMgr Is Nothing = False Then goMissileMgr.KillAll()
            If goWpnMgr Is Nothing = False Then goWpnMgr.CleanAll()
            If goSound Is Nothing = False Then goSound.KillAllSounds()

            If bFirstTime = True Then

                If lGalaxyID <> -1 AndAlso goGalaxy.ObjectID = lGalaxyID Then
                    If lSystemID <> -1 Then
                        For X = 0 To goGalaxy.mlSystemUB
                            If goGalaxy.moSystems(X).ObjectID = lSystemID Then
                                goGalaxy.CurrentSystemIdx = X
                                Exit For
                            End If
                        Next X

                        'ensure we have the system's details
                        If goGalaxy.CurrentSystemIdx <> -1 Then
                            If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB = -1 Then
                                moMsgSys.RequestSystemDetails(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID)
                            End If
                        End If

                        If lPlanetID <> -1 Then
                            For X = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB
                                If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(X).ObjectID = lPlanetID Then
                                    goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx = X
                                    Exit For
                                End If
                            Next X
                        Else
                            goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx = -1
                        End If
                    Else
                        goGalaxy.CurrentSystemIdx = -1
                    End If
                Else
                    'TODO: what do we do here?
                End If

                If lPlanetID <> -1 Then
                    glCurrentEnvirView = CurrentView.ePlanetMapView
                ElseIf lSystemID <> -1 Then
                    glCurrentEnvirView = CurrentView.eSystemMapView1

                    goGalaxy.GalaxySelectionIdx = -1
                    With goCamera
                        .mlCameraAtX = 0 : .mlCameraX = 0
                        .mlCameraAtY = 0 : .mlCameraY = 1000
                        .mlCameraAtZ = 0 : .mlCameraZ = -1000
                    End With

                Else
                    glCurrentEnvirView = CurrentView.eGalaxyMapView
                    With goCamera
                        .mlCameraAtX = 0 : .mlCameraX = 0
                        .mlCameraAtY = 0 : .mlCameraY = 1000
                        .mlCameraAtZ = 0 : .mlCameraZ = -1000
                    End With
                End If

                'force our startup screen to be removed
                moEngine.ForceCleanupLoginBackground()
            End If

            If goGalaxy.CurrentSystemIdx <> -1 Then
                If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx <> -1 Then
                    goCurrentEnvir.oGeoObject = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx)
                Else
                    goCurrentEnvir.oGeoObject = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                End If
                goCurrentEnvir.SetExtents()
            End If

            If lID >= 500000000 AndAlso iTypeID = ObjectType.ePlanet Then
                goCurrentEnvir.oGeoObject = Planet.GetTutorialPlanet()
                goCurrentEnvir.SetExtents()
            End If

            If lID = 0 OrElse iTypeID = 0 Then
                moMsgSys_ProcessedBurstEnvironment()
            Else
                'ok,if we have received this, then we need to send a request for that environment's domain
                Dim yData(7) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestEnvironmentDomain).CopyTo(yData, 0)
                goCurrentEnvir.GetGUIDAsString.CopyTo(yData, 2)
                moMsgSys.SendToPrimary(yData)
            End If

            If goUILib Is Nothing = False Then
                Dim ofrm As frmChat = CType(goUILib.GetWindow("frmChat"), frmChat)
                If ofrm Is Nothing = False Then ofrm.AddADYK()
            End If
        End If

        Debug.Write("Processed Player Current Environment..." & vbCrLf)
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
        Debug.Write("Processed Burst Environment Message..." & vbCrLf)

        'Finally, tell our environment to set primary ambience sound fx
        goCurrentEnvir.SetPrimaryAmbience()

        'Ensure that any tooltips are invis
        goUILib.SetToolTip(False)

        mbIgnoreNextMouseUp = False

        mbChangingEnvirs = False
        mbEnvirReady = True

        If goUILib Is Nothing = False Then
            If goUILib.lUISelectState = UILib.eSelectState.eNotification_ClickTo Then
                goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                Dim oTmpWin As frmNotification = CType(goUILib.GetWindow("frmNotification"), frmNotification)
                If oTmpWin Is Nothing = False AndAlso oTmpWin.Visible = True Then
                    oTmpWin.FinishClickToEvent()
                Else
                    Dim oTmpNoteHist As frmNoteHistory = CType(goUILib.GetWindow("frmNoteHistory"), frmNoteHistory)
                    If oTmpNoteHist Is Nothing = False AndAlso oTmpNoteHist.Visible = True Then
                        oTmpNoteHist.FinishClickToEvent()
                    End If
                    oTmpNoteHist = Nothing
                End If
                oTmpWin = Nothing
            ElseIf goUILib.lUISelectState = UILib.eSelectState.eEmailAttachmentJumpTo Then
                goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                PlayerComm.FinishJumpToEvent()
            ElseIf goUILib.lUISelectState = UILib.eSelectState.eBattlegroupJumpTo Then
                goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                Dim ofrm As frmFleet = CType(goUILib.GetWindow("frmFleet"), frmFleet)
                If ofrm Is Nothing = False Then ofrm.FinishClickToEvent()
                ofrm = Nothing
            ElseIf goUILib.lUISelectState = UILib.eSelectState.eJumpToEntity Then
                goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                If oEnvir Is Nothing = False Then
                    oEnvir.JumpToEntity(goUILib.lJumpToExtended1, goUILib.lJumpToExtended2)
                End If
            End If
            goUILib.AddNotification("Link Established!", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

            Dim ofrmQB As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
            If ofrmQB Is Nothing Then ofrmQB = New frmQuickBar(goUILib)
            ofrmQB = Nothing

            'Create a new Envirdisplay, because it is a singleton, it will work pretty much on its own...
            Dim ofrmED As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
            If ofrmED Is Nothing Then ofrmED = New frmEnvirDisplay(goUILib)
            ofrmED.Visible = True
            ofrmED = Nothing           'remove our local pointer, but WindowList should contain it still
        End If

        'If goTutorial Is Nothing Then goTutorial = New TutorialManager()
        'If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
        '	If goTutorial.ContinueTutorial = False Then
        '		If goTutorial.EventTriggered(TutorialManager.TutorialTriggerType.FirstTimePlayer) = False Then
        '			goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.FirstTimePlayer)
        '		End If
        '	End If
        'End If

        If NewTutorialManager.LoadScript() = False Then
            goUILib.AddNotification("Script Could not be loaded!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            If goCurrentPlayer Is Nothing = False AndAlso (goCurrentPlayer.yPlayerPhase = 0 OrElse (mbFirstLoad = True AndAlso goCurrentPlayer.yPlayerPhase = 1)) Then
                mbFirstLoad = False
                If goCurrentPlayer.lTutorialStep < 301 Then
                    glCurrentEnvirView = CurrentView.ePlanetView
                    goCamera.mlCameraAtX = 0
                    goCamera.mlCameraAtY = 0
                    goCamera.mlCameraAtZ = 0
                    goCamera.mlCameraX = 0
                    goCamera.mlCameraY = 5000
                    goCamera.mlCameraZ = 1000
                Else 'If goCurrentPlayer.lTutorialStep < 302 Then
                    If goCurrentEnvir Is Nothing = False Then
                        If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                            glCurrentEnvirView = CurrentView.ePlanetMapView
                            goCamera.mlCameraAtX = 0
                            goCamera.mlCameraAtY = 0
                            goCamera.mlCameraAtZ = 0
                            goCamera.mlCameraX = 0
                            goCamera.mlCameraY = 5000
                            goCamera.mlCameraZ = 1000
                        End If
                    End If

                End If

                NewTutorialManager.FindAndExecuteStepID(goCurrentPlayer.lTutorialStep)
            End If
        End If

        ''Now, let's place all Y values
        'If goCurrentEnvir Is Nothing = False Then
        '	Try
        '		For X As Int32 = 0 To goCurrentEnvir.lEntityUB
        '			If goCurrentEnvir.lEntityIdx(X) <> -1 Then
        '				goCurrentEnvir.oEntity(X).GetWorldMatrix()
        '				goCurrentEnvir.oEntity(X).LocY = goCurrentEnvir.oEntity(X).mlTargetY
        '			End If
        '		Next X
        '	Catch
        '	End Try
        'End If

        If goCurrentEnvir Is Nothing = False Then
            mbNeedToLoadTacticalData = True
            Dim yMsg(7) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eGetEnvirConstructionObjects).CopyTo(yMsg, 0)
            goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, 2)
            moMsgSys.SendToPrimary(yMsg)

            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eEnvironmentChanged, goCurrentEnvir.ObjectID, goCurrentEnvir.ObjTypeID, -1, "")
            End If
        End If

        GFXEngine.bRenderInProgress = False

    End Sub

    Private Sub moMsgSys_ReceivedGalaxyAndSystems(ByVal yData() As Byte) Handles moMsgSys.ReceivedGalaxyAndSystems
        'Ok, time to light up the ovens... this message contains the Galaxy Objects and System Objects
        Dim lPos As Int32 = 2   'the position of our cursor, inc. by 2 for the msg ID
        Dim lObjID As Int32
        Dim iTypeID As Int16

        Dim lTemp As Int32

        Dim X As Int32

        'Ok set out galaxy to new
        goGalaxy = New Galaxy()

        Try
            While lPos < yData.Length - 1
                'whether it is a galaxy or a system, the first 4 bytes is the ID
                lObjID = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
                'then the type id
                iTypeID = System.BitConverter.ToInt16(yData, lPos)
                lPos += 2

                'Now, we can determine whether it is a galaxy or system
                If iTypeID = ObjectType.eGalaxy Then
                    'galaxy... its our galaxy definition... we will only load one galaxy at a time...
                    With goGalaxy
                        .ObjectID = lObjID
                        .ObjTypeID = iTypeID

                        .GalaxyName = GetStringFromBytes(yData, lPos, 20)
                        lPos += 20

                        '.GalaxyName = System.Text.ASCIIEncoding.ASCII.GetString(yTempName)
                        lPos += 4
                    End With
                ElseIf iTypeID = ObjectType.eSolarSystem Then
                    'solar system
                    goGalaxy.mlSystemUB += 1
                    ReDim Preserve goGalaxy.moSystems(goGalaxy.mlSystemUB)
                    goGalaxy.moSystems(goGalaxy.mlSystemUB) = New SolarSystem()
                    With goGalaxy.moSystems(goGalaxy.mlSystemUB)
                        .ObjectID = lObjID
                        .ObjTypeID = iTypeID

                        .SystemName = GetStringFromBytes(yData, lPos, 20)
                        lPos += 20

                        lTemp = System.BitConverter.ToInt32(yData, lPos)
                        'Systems are contained within Galaxies, but since we don't really care about multiple galaxies
                        '  at this time, then we'll just pass this up for now                    
                        lPos += 4

                        .LocX = System.BitConverter.ToInt32(yData, lPos) * 2
                        lPos += 4
                        .LocY = System.BitConverter.ToInt32(yData, lPos) * 2
                        lPos += 4
                        .LocZ = System.BitConverter.ToInt32(yData, lPos) * 2
                        lPos += 4


                        lTemp = yData(lPos)
                        lPos += 1
                        'NOTE: We should have already called AND received our star types by now
                        For X = 0 To glStarTypeUB
                            If goStarTypes(X).StarTypeID = lTemp Then
                                .StarType1Idx = X
                                Exit For
                            End If
                        Next X

                        lTemp = yData(lPos)
                        lPos += 1
                        'NOTE: We should have already called AND received our star types by now
                        For X = 0 To glStarTypeUB
                            If goStarTypes(X).StarTypeID = lTemp Then
                                .StarType2Idx = X
                                Exit For
                            End If
                        Next X

                        lTemp = yData(lPos)
                        lPos += 1
                        'NOTE: We should have already called AND received our star types by now
                        For X = 0 To glStarTypeUB
                            If goStarTypes(X).StarTypeID = lTemp Then
                                .StarType3Idx = X
                                Exit For
                            End If
                        Next X

                        .SystemType = yData(lPos) : lPos += 1
                        If gl_CLIENT_VERSION >= 311 Then
                            .FleetJumpPointX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            .FleetJumpPointZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        End If
                    End With
                Else
                    'This was bad, so we're just gonna slink away and forget it ever happened
                    'Application.Exit()
                End If
            End While
        Catch
        End Try

        'Now, request our player details
        goUILib.SetToolTip("Requesting Player Details...", -1, frmLoginDlg.lStatusTop)
        moMsgSys.RequestPlayerDetails()

    End Sub

    Private Sub moMsgSys_ReceivedStarTypes(ByVal yData() As Byte) Handles moMsgSys.ReceivedStarTypes
        'first two bytes is the message ID
        Dim lPos As Int32 = 2
        Dim yStarMapRectIdx As Byte

        glStarTypeUB = -1
        ReDim goStarTypes(-1)

        While (lPos + 78) < yData.Length - 1
            glStarTypeUB += 1
            ReDim Preserve goStarTypes(glStarTypeUB)

            goStarTypes(glStarTypeUB) = New StarType()

            With goStarTypes(glStarTypeUB)
                .StarTypeID = yData(lPos)
                lPos += 1

                'Star Type Name (20)
                .StarTypeName = GetStringFromBytes(yData, lPos, 20)
                lPos += 20

                .StarAttributes = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4

                'Star Texture (20)
                .StarTexture = GetStringFromBytes(yData, lPos, 20)
                lPos += 20

                .HeatIndex = yData(lPos)
                lPos += 1

                .MatDiffuse = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
                .MatEmissive = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4

                '.StarRarity = yData(lPos)
                lPos += 1
                .LightRange = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
                .LightDiffuse = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
                .LightAmbient = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
                .LightSpecular = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
                .LightAtt0 = System.BitConverter.ToSingle(yData, lPos)
                lPos += 4
                .LightAtt1 = System.BitConverter.ToSingle(yData, lPos)
                lPos += 4
                .LightAtt2 = System.BitConverter.ToSingle(yData, lPos)
                lPos += 4
                yStarMapRectIdx = yData(lPos)       'TODO: define what this is....
                lPos += 1
                .StarRadius = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
            End With
        End While

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
        Dim lPos As Int32 = 2       'for msgcode
        Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oSystem As SolarSystem = Nothing

        Static xlReTryCnt As Int32 = 0

        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X) Is Nothing = False AndAlso goGalaxy.moSystems(X).ObjectID = lSystemID Then
                oSystem = goGalaxy.moSystems(X)
                Exit For
            End If
        Next X
        If oSystem Is Nothing Then
            xlReTryCnt += 1
            If xlReTryCnt > 5 Then
                If goUILib Is Nothing = False Then goUILib.AddNotification("Critical error in receiving system details: " & lSystemID, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            Dim yReRequest(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestSystemDetails).CopyTo(yReRequest, 0)
            System.BitConverter.GetBytes(lSystemID).CopyTo(yReRequest, 2)
            moMsgSys.bDoNotClearBusyStatus = True
            moMsgSys.SendToPrimary(yReRequest)
            Return
        End If

        Try
            Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            For X As Int32 = 0 To lCnt - 1
                'Object ID
                Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                'ObjTypeID
                Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                'Name
                Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                'Planet type id
                Dim yMapTypeID As Byte = yData(lPos) : lPos += 1
                'Size ID
                Dim ySizeID As Byte = yData(lPos) : lPos += 1

                'Now, create our planet
                Dim oTmpPlnt As New Planet(lID, ySizeID, yMapTypeID)
                With oTmpPlnt
                    .ObjTypeID = iTypeID
                    .PlanetName = sName
                    'ok, now the other attributes
                    .PlanetRadius = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .ParentSystem = oSystem : lPos += 4
                    .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .LocY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Vegetation = yData(lPos) : lPos += 1
                    .Atmosphere = yData(lPos) : lPos += 1
                    .Hydrosphere = yData(lPos) : lPos += 1
                    .Gravity = yData(lPos) : lPos += 1
                    .SurfaceTemperature = yData(lPos) : lPos += 1
                    .RotationDelay = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                    'Now, check for our ring...
                    Dim lInnerRadius As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim lOuterRadius As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim lRingDiffuse As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .SetRingProps(lInnerRadius, lOuterRadius, lRingDiffuse)

                    'finally, the axis angle
                    .AxisAngle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    .RotateAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                    .SetLastUpdate(timeGetTime)
                End With

                oSystem.AddPlanet(oTmpPlnt)
                oTmpPlnt = Nothing
            Next X
        Catch
            xlReTryCnt += 1
            If xlReTryCnt > 5 Then
                If goUILib Is Nothing = False Then goUILib.AddNotification("Critical error in receiving system details: " & lSystemID, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            Dim yReRequest(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestSystemDetails).CopyTo(yReRequest, 0)
            System.BitConverter.GetBytes(lSystemID).CopyTo(yReRequest, 2)
            moMsgSys.bDoNotClearBusyStatus = True
            moMsgSys.SendToPrimary(yReRequest)
        End Try
        If goGalaxy Is Nothing = False Then goGalaxy.JumpToPlanet(lSystemID)
        Dim ofrmUnitGoto As frmUnitGoto = CType(goUILib.GetWindow("frmUnitGoto"), frmUnitGoto)
        If ofrmUnitGoto Is Nothing = False AndAlso ofrmUnitGoto.Visible = True Then ofrmUnitGoto.FillList()
    End Sub

    Private Sub moMsgSys_ServerShutdown() Handles moMsgSys.ServerShutdown
        'server is telling us that the game is shutting down...
        mbEnvirReady = False
        mbChangingEnvirs = False

        Dim lLoopCnt As Int32 = 0
        While gb_InHandleMovement
            Threading.Thread.Sleep(10)
            lLoopCnt += 1
            If lLoopCnt > 1000 Then Exit While
        End While

        'Clear our stuff
        goCurrentEnvir = Nothing
        If goUILib Is Nothing = False Then
            goUILib.ClearSelection()
            goUILib.RemoveAllWindows()
            goUILib.oDevice.RenderState.FogEnable = False
        End If
        If goPFXEngine32 Is Nothing = False Then goPFXEngine32.ClearAllEmitters()
        With goCamera
            .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
            .mlCameraX = 0 : .mlCameraY = 0 : .mlCameraZ = -10000
        End With

        'disconnect everyone...
        moMsgSys.DisconnectAll()

        Dim oTmpWin As frmLoginDlg
        oTmpWin = New frmLoginDlg(goUILib)
        AddHandler oTmpWin.StartLogin, AddressOf Login_StartLogin
        AddHandler oTmpWin.ExitProgram, AddressOf Login_ExitProgram
        goUILib.AddWindow(oTmpWin)
        oTmpWin = Nothing

        glCurrentEnvirView = CurrentView.eStartupDSELogo

        mbfrmConfirmHandled = True

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
            If goUILib Is Nothing = False Then goUILib.SetMsgSystem(Nothing)

            goUILib.AddNotification("Version incompatible with server!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            goUILib.AddNotification("Please launch the game through the updater.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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
            If goUILib Is Nothing = False Then
                goUILib.SetToolTip("Processing Login Request...", -1, frmLoginDlg.lStatusTop)
            End If
            moMsgSys.RequestLogin(gsUserName, gsPassword, gbAliased, gsAliasUserName, gsAliasPassword, yConnType)
        End If
    End Sub

    Private Sub moMsgSys_LoginResponse(ByVal yConnType As eyConnType, ByVal lVal As Integer) Handles moMsgSys.LoginResponse
        If lVal < 0 Then
            Select Case lVal
                Case LoginResponseCodes.eAccountBanned
                    goUILib.AddNotification("Login Failed! Account Banned. Please contact support.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case LoginResponseCodes.eAccountSuspended
                    'goUILib.SetToolTip("Login Failed: Account Suspended. Please Contact Support.", -1, lStatusTop)
                    goUILib.AddNotification("Login Failed! Account Suspended. Please contact support.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case LoginResponseCodes.eInvalidPassword, LoginResponseCodes.eInvalidUserName
                    'goUILib.SetToolTip("Login Failed: Invalid Username/Password Combination.", -1, lStatusTop)
                    goUILib.AddNotification("Login Failed! Invalid Username/Password.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case LoginResponseCodes.eLoginAttemptLockout
                    'goUILib.SetToolTip("Login Failed: Maximum Attempts Exceeded!", -1, lStatusTop)
                    goUILib.AddNotification("Login Failed! Maximum attempts exceeded!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case LoginResponseCodes.eMondelisInactive
                    goUILib.AddNotification("This account is inactive. Contact Mondelis.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case LoginResponseCodes.eAccountInactive
                    'goUILib.SetToolTip("This account is inactive." & vbCrLf & "You must reactivate the account in order to login.", -1, lStatusTop)
                    goUILib.AddNotification("This account is inactive. You must have an active subscription", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    goUILib.AddNotification("for this account in order to login.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case LoginResponseCodes.eAccountSetup
                    goUILib.AddNotification("This account needs to be setup.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Dim ofrm As frmPlayerSetup = New frmPlayerSetup(goUILib, gsUserNameProper, gsPassword)
                    ofrm.Visible = True
                    ofrm.sUserName = gsUserName
                    ofrm.sPassword = gsPassword
                    Return
                Case LoginResponseCodes.eAccountInUse
                    goUILib.AddNotification("Account is in use. Please try again later.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case LoginResponseCodes.ePlayerIsDying
                    goUILib.AddNotification("This account is still in the process of dying.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    goUILib.AddNotification("Try again a few seconds.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    goUILib.AddNotification("If the problem persists for longer than five minutes, contact support.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case Else
                    'goUILib.SetToolTip("Login Failed, Please try again later.", -1, lStatusTop)
                    goUILib.AddNotification("Login failed, please try again later", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End Select
            're-enable our button

            moMsgSys.DisconnectAll()
            moMsgSys = Nothing
            If goUILib Is Nothing = False Then
                Dim frmLogin As frmLoginDlg = CType(goUILib.GetWindow("frmLoginDlg"), frmLoginDlg)
                If frmLogin Is Nothing = False Then
                    frmLogin.cmdExit.Enabled = True
                    frmLogin.cmdLogin.Enabled = True
                End If

                goUILib.RetainTooltip = False
                goUILib.SetToolTip(False)
            End If
            GFXEngine.bRenderInProgress = False
        Else
            glPlayerID = lVal

            Select Case yConnType
                Case eyConnType.ChatServer  '???
                Case eyConnType.OperatorServer
                    If glPlayerID = glActualPlayerID Then gbAliased = False
                    'waiting on server specifics
                    If goUILib Is Nothing = False Then goUILib.SetToolTip("Waiting on server specifics...", -1, frmLoginDlg.lStatusTop)
                Case eyConnType.PrimaryServer
                    If goUILib Is Nothing = False Then goUILib.SetToolTip("Login Successful, getting player details...", -1, frmLoginDlg.lStatusTop)
                    Dim oINI As New InitFile()
                    Dim sTemp As String = oINI.GetString("SIGNON", "LastUserName", "")
                    If sTemp.ToUpper <> gsUserNameProper.ToUpper Then
                        oINI.WriteString("SIGNON", "LastUserName", gsUserNameProper)
                    End If
                    oINI = Nothing
                    If goUILib Is Nothing = False Then goUILib.AddNotification("Welcome to Beyond Protocol!", muSettings.InterfaceBorderColor, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If gbAliased = False Then
                        Me.Text = "Beyond Protocol - " & gsUserNameProper
                    Else
                        Me.Text = "Beyond Protocol - " & gsUserNameProper & " (" & gsAliasUserName & ")"
                    End If

                    'Ok, so our login is good... let's do this, first remove our window (need to release our handlers)
                    goUILib.RemoveWindow("frmLoginDlg")

                    'If we're here, then we need to clean up our background for the Login Screen, we do this by setting the 
                    '  startup dialog time
                    If moEngine.swStartup Is Nothing = False OrElse bForceRequestDetails = True Then
                        bForceRequestDetails = False
                        If moEngine.swStartup Is Nothing = False Then moEngine.swStartup.Reset()
                        moEngine.swStartup = Nothing

                        'Now... request our stuff...
                        goUILib.SetToolTip("Requesting Universe Data...", -1, frmLoginDlg.lStatusTop)

                        LoadStarTypes()
                        moMsgSys.RequestGalaxyAndSystems()
                    Else
                        'Ok, server knows who we are and I have my data, so let's continue
                        Dim yData(5) As Byte
                        goUILib.SetToolTip("Login Completed Entering Universe...", -1, frmLoginDlg.lStatusTop)
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerEnvironment).CopyTo(yData, 0)
                        System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, 2)
                        moMsgSys.SendToPrimary(yData)

                        If goUILib Is Nothing = False Then goUILib.RetainTooltip = False

                    End If


            End Select
        End If
    End Sub
    Public bForceRequestDetails As Boolean = False

    'Request Primary data here...
    Private mlRequestPrimaryID As Int32 = -1
    Private miRequestPrimaryTypeID As Int16 = -1
    Private Sub moMsgSys_ReceivedPrimaryServerData() Handles moMsgSys.ReceivedPrimaryServerData

        If mlRequestPrimaryID <> -1 AndAlso miRequestPrimaryTypeID <> -1 Then
            Dim yMsg(11) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePathfindingConnectionInfo).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(mlRequestPrimaryID).CopyTo(yMsg, 6)
            System.BitConverter.GetBytes(miRequestPrimaryTypeID).CopyTo(yMsg, 10)
            moMsgSys.SendToOperator(yMsg)
        Else
            moMsgSys.DisconnectOperator()
        End If
        mlRequestPrimaryID = -1
        miRequestPrimaryTypeID = -1
    End Sub

    Private Sub moMsgSys_ServerDisconnected(ByVal yType As eyConnType) Handles moMsgSys.ServerDisconnected
        If yType = eyConnType.OperatorServer Then

            '7) Connect to Primary Server
            If moMsgSys.ConnectToPrimary() = False Then
                goUILib.AddNotification("Could not connect to server.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                goUILib.AddNotification("The server is either down or you", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                goUILib.AddNotification("are not connected to the internet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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
        'set us up for the environment change
        mbChangingEnvirs = True

        'Oh, and finally, set our music play list
        If goSound Is Nothing = False Then
            goSound.PlayListType = 2 'TODO: 0 = login, 1 = lull, 2 = excited
        End If

        'Show our Chat window...
        Dim oTmpCht As frmChat
        oTmpCht = New frmChat(goUILib)
        goUILib.AddWindow(oTmpCht)
        oTmpCht = Nothing

        If muSettings.ShowConnectionStatus = True Then
            Dim ofrmCS As frmConnectionStatus = CType(goUILib.GetWindow("frmConnectionStatus"), frmConnectionStatus)
            If ofrmCS Is Nothing Then ofrmCS = New frmConnectionStatus(goUILib)
            ofrmCS.Visible = True
            ofrmCS = Nothing
        End If

        'Ok, for now, attempt to join testers
        Dim yData(46) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yData, lPos) : lPos += 2
        lPos += 4 'leave room for playerid
        yData(lPos) = eyChatRoomCommandType.JoinChannel : lPos += 1

        Dim sChannel As String = "GENERAL"
        If goCurrentPlayer Is Nothing = False Then
            If goCurrentPlayer.yPlayerPhase = 255 Then sChannel = "GENERAL"
        End If
        System.Text.ASCIIEncoding.ASCII.GetBytes(sChannel).CopyTo(yData, lPos) : lPos += 20
        moMsgSys.SendToPrimary(yData)

        If goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Emperor AndAlso gbAliased = False Then
            ReDim yData(46)
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yData, lPos) : lPos += 2
            lPos += 4 'leave room for playerid
            yData(lPos) = eyChatRoomCommandType.JoinChannel : lPos += 1
            System.Text.ASCIIEncoding.ASCII.GetBytes("EMPEROR").CopyTo(yData, lPos) : lPos += 20
            'System.Text.ASCIIEncoding.ASCII.GetBytes("").CopyTo(yData, lPos) : lPos += 20
            moMsgSys.SendToPrimary(yData)
        End If

        'request the senate status report
        ReDim yData(5)
        lPos = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eSenateStatusReport).CopyTo(yData, lPos) : lPos += 2
        System.BitConverter.GetBytes(0I).CopyTo(yData, lPos) : lPos += 4
        moMsgSys.SendToPrimary(yData)


        'Ok, server knows who we are and I have my data, so let's continue
        ReDim yData(5)
        goUILib.SetToolTip("Login Completed Entering Universe...", -1, frmLoginDlg.lStatusTop)
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerEnvironment).CopyTo(yData, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, 2)
        moMsgSys.SendToPrimary(yData)

        If goUILib Is Nothing = False Then goUILib.RetainTooltip = False

        mbRequestClaimables = True
    End Sub

#End Region

    Public Sub CancelLogin()
        moMsgSys.DisconnectAll()
        moMsgSys = Nothing
        If goUILib Is Nothing = False Then
            Dim frmLogin As frmLoginDlg = CType(goUILib.GetWindow("frmLoginDlg"), frmLoginDlg)
            If frmLogin Is Nothing = False Then
                frmLogin.cmdExit.Enabled = True
                frmLogin.cmdLogin.Enabled = True
            End If

            goUILib.RetainTooltip = False
            goUILib.SetToolTip(False)
        End If
        GFXEngine.bRenderInProgress = False
    End Sub
End Class
