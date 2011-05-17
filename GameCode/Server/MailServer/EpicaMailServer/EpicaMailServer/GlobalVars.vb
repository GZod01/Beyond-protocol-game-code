Public Module GlobalVars
    Public Const gl_SECONDS_TO_LIVE As Int32 = 28800        '8 hrs

    Public gb_IS_TEST_SERVER As Boolean = True

    Public Enum PlayerAlertType As Byte
        eUnderAttack = 0
        eEngagedEnemy = 1
        eUnitLost = 2
        eFacilityLost = 3
        eColonyLost = 4
        eBuyOrderAccepted = 5
    End Enum

    Public gsOutHostName As String = "smtpserver.mydomain.com"
    Public gsInHostName As String = "pop3server.mydomain.com"
    Public gsEmailUserName As String = "MyEmailAddress" 'before the @
    Public gsEmailPassword As String = "MyEmailPassword"
    Public gsOutEmailFrom As String = "fleetcommander@mydomain.com"
    Public gsReplyToDomain As String = "@mydomain.com" 'This is appended to outgoing automated emails

    Public goOutQueue As MailQueue 

    Public goInboundEmailMgr As InboundEmailMgr

    Public gbRunning As Boolean = True

    Public goMailMsgs() As MailObject
    Public gyMailMsgUsed() As Byte
    Public glMailMsgUB As Int32 = -1

    Public goPlayer() As Player
    Public glPlayerIdx() As Int32
    Public glPlayerUB As Int32 = -1

    Public gfrmMain As frmMain

    Public goMsgSys As MsgSystem

    Public glBoxOperatorID As Int32
    Public gsOperatorIP As String
    Public glOperatorPort As Int32
    Public gsBackupOperatorIP As String
    Public glBackupOperatorPort As Int32

    Public gbDisabled As Boolean = gb_IS_TEST_SERVER

    Public Sub MainLoop()
        goInboundEmailMgr = New InboundEmailMgr()
        With goInboundEmailMgr
            .sEmailDir = AppDomain.CurrentDomain.BaseDirectory
            If .sEmailDir.EndsWith("\") = False Then .sEmailDir &= "\"
            '.sEmailDir &= "Inbound\"
            .sAttachDir = .sEmailDir
        End With
        'Initialize message system
        goMsgSys = New MsgSystem()
        If goMsgSys.ConnectToOperator(gsOperatorIP, glOperatorPort) = False Then
            LogEvent("Unable to connect to operator. Exiting")
            Return
        End If

        ChatRoom.CreateStandardRooms()

        'Initialize our out and in queue
        goOutQueue = New MailQueue()

        'Begin accepting our primary servers
        goMsgSys.AcceptingPrimarys = True

        'Indicate that all is well
        LogEvent("Server Started...")
        goMsgSys.SendReadyStateToOperator()

        Dim oInSW As Stopwatch = Stopwatch.StartNew()

        While gbRunning = True
            Try
                If gbDisabled = False Then
                    'Check our outbound queue for an item to send
                    Dim lNextIdx As Int32 = goOutQueue.GetNextItem()
                    If lNextIdx <> -1 Then
                        Dim oItem As MailObject = goMailMsgs(lNextIdx)
                        If oItem Is Nothing = False Then
                            If oItem.SendMailMsg() = False Then
                                'TODO: Present error or something...
                            End If
                        End If
                        oItem = Nothing
                    End If

                    'Go and get our inbound emails
                    If oInSW.ElapsedMilliseconds > 1000 Then
                        Dim lResult As Int32 = goInboundEmailMgr.RetreiveEmails()
                        'read one inbound email (if there is one)
                        goInboundEmailMgr.ParseEmailFile()
                        oInSW.Stop()
                        oInSW.Reset()
                        oInSW.Start()
                    End If
                End If

                'Now, clean up any messages that need to removed
                'Dim lCurrentTime As Int32 = CInt(Val(Now.ToString("MMddHHmmss")))
                Dim dtCurrent As Date = Now
                For X As Int32 = 0 To glMailMsgUB
                    If gyMailMsgUsed(X) <> 0 AndAlso goMailMsgs(X).dtTimeStamp <> Date.MinValue Then '  goMailMsgs(X).lTimeStamp <> Int32.MaxValue Then
                        If dtCurrent.Subtract(goMailMsgs(X).dtTimeStamp).TotalSeconds > gl_SECONDS_TO_LIVE Then
                            gyMailMsgUsed(X) = 0
                            goMailMsgs(X) = Nothing
                        End If
                        'lTemp = lCurrentTime - goMailMsgs(X).lTimeStamp
                        'If lTemp > gl_SECONDS_TO_LIVE Then
                        '	gyMailMsgUsed(X) = 0
                        '	goMailMsgs(X) = Nothing
                        'End If
                    End If
                Next X

                'check our operator
                If goMsgSys.CheckOperatorConnection() = False Then
                    goMsgSys.OperatorFailure()
                End If

                'Cause us to sleep for 10ms to give other threads a chance to do whatever
                Threading.Thread.Sleep(10)
                Application.DoEvents()
            Catch ex As Exception
                LogEvent("Main Loop Error: " & ex.Message)
            End Try
        End While

        gfrmMain.AddEventLine("Shutting Down...")
        'Now, close the application, first, no longer accepting new connections
        goMsgSys.AcceptingPrimarys = False
        'disconnect everyone
        goMsgSys.DisconnectAll()
        'kill the queues
        goOutQueue = Nothing
        glMailMsgUB = -1
        Erase goMailMsgs
        Erase gyMailMsgUsed
        glPlayerUB = -1
        Erase goPlayer
        Erase glPlayerIdx

        gfrmMain.Close()

        End
    End Sub

    Public Function AddNewOutboundMailMsg(ByVal sFrom As String, ByVal sTo As String, ByVal sSubject As String, ByVal sBody As String, ByVal lPC_ID As Int32, ByVal lPlayerID As Int32, ByVal iBaseAlertTypeID As Int16) As Int32
        For X As Int32 = 0 To glMailMsgUB
            If gyMailMsgUsed(X) = 0 Then
                goMailMsgs(X) = New MailObject(sFrom, sTo, sSubject, sBody, lPC_ID, lPlayerID, iBaseAlertTypeID)
                gyMailMsgUsed(X) = 255
                Return X
            End If
        Next X
        ReDim Preserve goMailMsgs(glMailMsgUB + 1)
        ReDim Preserve gyMailMsgUsed(glMailMsgUB + 1)
        goMailMsgs(glMailMsgUB + 1) = New MailObject(sFrom, sTo, sSubject, sBody, lPC_ID, lPlayerID, iBaseAlertTypeID)
        gyMailMsgUsed(glMailMsgUB + 1) = 255
        glMailMsgUB += 1
        Return glMailMsgUB
    End Function

    Public Function AddNewInboundMailMsg() As Int32
        'TODO: Define me
        Return -1
    End Function

    Public Function GetPlayer(ByVal lID As Int32) As Player
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) = lID Then Return goPlayer(X)
        Next X
        Return Nothing
    End Function

#Region "  Logging  "
    Public Sub LogEvent(ByVal sEventText As String)
        sEventText = Now.ToString & ": " & sEventText
        gfrmMain.AddEventLine(sEventText)

        'Any text file logging needs to be done here
        LogToFile(sEventText)
    End Sub

    Private oFS As IO.FileStream
    Private oWriter As IO.StreamWriter
    Private lVals(-1) As Int32       'just a placeholder for something to synclock on
    Private msPreviousFileName As String
    Private mlPreviousFileNum As Int32 = 0
    Private mbForceCloseLog As Boolean = False

    Private Sub LogToFile(ByVal sMsg As String)
        'need to log FromPlayerID and the Msg
        Const ml_MAX_LOG_FILE_SIZE As Int32 = 5000000

        SyncLock lVals
            If oFS Is Nothing Then
                Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                If sPath.EndsWith("\") = False Then sPath &= "\"
                mlPreviousFileNum = 0
                Dim sNewFileName As String = "Events_" & Now.ToString("MMddyyyyHHmmss")
                oFS = New IO.FileStream(sPath & sNewFileName & "_" & mlPreviousFileNum & ".log", IO.FileMode.Create)
                msPreviousFileName = sNewFileName
                oWriter = New IO.StreamWriter(oFS)
                oWriter.AutoFlush = True
            End If

            oWriter.WriteLine(sMsg)

            If oFS.Length > ml_MAX_LOG_FILE_SIZE OrElse mbForceCloseLog = True Then
                oWriter.Close()
                oFS.Close()
                oWriter.Dispose()
                oFS.Dispose()
                Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                If sPath.EndsWith("\") = False Then sPath &= "\"
                Dim sNewFileName As String = "Events_" & Now.ToString("MMddyyyyHHmmss")
                If msPreviousFileName = sNewFileName Then
                    mlPreviousFileNum += 1
                Else
                    mlPreviousFileNum = 0
                End If
                oFS = New IO.FileStream(sPath & sNewFileName & "_" & mlPreviousFileNum & ".log", IO.FileMode.Create)
                msPreviousFileName = sNewFileName
                oWriter = New IO.StreamWriter(oFS)
            End If
        End SyncLock
    End Sub
#End Region

    Public Function GenerateReplyToAddress(ByVal lPC_ID As Int32) As String
        Dim lNow As Int32 = CInt(Val(Now.ToString("HHmmssDDMM")))

        Dim sBase36 As String = ConvertBase10(lNow, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        Dim sValue As String = ConvertBase10(lPC_ID, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ")

        Dim sFinalAddress As String = sValue & sBase36
        If sFinalAddress.Length > 10 Then sFinalAddress = sFinalAddress.Substring(0, 10)
        sFinalAddress &= gsReplyToDomain
        Return sFinalAddress
    End Function

    Private Function ConvertBase10(ByVal lValue As Int32, ByVal sNewBaseDigits As String) As String
        Dim sResult As String = ""
        Dim lTemp As Int32
        Dim lI As Int32
        Dim lLastI As Int32
        Dim lBaseSize As Integer

        lBaseSize = Len(sNewBaseDigits)
        While lValue <> 0
            lTemp = lValue
            lI = 0
            While lTemp >= lBaseSize
                lI += 1
                lTemp \= lBaseSize
            End While
            If lI <> lLastI - 1 AndAlso lLastI <> 0 Then sResult &= StrDup(lLastI - lI - 1, sNewBaseDigits.Substring(0, 1))
            lTemp = lTemp
            sResult &= sNewBaseDigits.Substring(lTemp, 1)
            lValue -= CInt(lTemp * (lBaseSize ^ lI))
            lLastI = lI
        End While
        sResult &= StrDup(lI, (sNewBaseDigits.Substring(0, 1))) 'get the zero digits at the end of the number
        Return sResult
    End Function
    Public Function StringToBytes(ByVal sVal As String) As Byte()
        Return System.Text.ASCIIEncoding.ASCII.GetBytes(sVal)
    End Function


    Public Function GetDateAsNumber(ByVal dtDate As Date) As Int32
        Return CInt(Val(dtDate.ToString("yyMMddHHmm")))
    End Function

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

    Public Function GetDateFromNumberAsString(ByVal lVal As Int32) As String
        If lVal = 0 Then Return "an unknown date"

        Dim sVal As String = lVal.ToString

        'Work from right to left
        Dim lUB As Int32 = sVal.Length ' - 1

        'Ok, the bare minimum for this to work is 8
        If lUB < 8 Then Return "an unknown date"

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

        Return sMon & "/" & sDay & "/20" & sYr & " at " & sHr & ":" & sMin
    End Function

    Public Function GetDurationFromSeconds(ByVal lSeconds As Int32, ByVal bDHMS As Boolean) As String
        Dim lMinutes As Int32 = lSeconds \ 60 : lSeconds -= (lMinutes * 60)
        Dim lHours As Int32 = lMinutes \ 60 : lMinutes -= (lHours * 60)
        Dim lDays As Int32 = lHours \ 24 : lHours -= (lDays * 24)
        If bDHMS = True Then
            Dim sResult As String = ""
            If lDays <> 0 Then sResult &= lDays.ToString & "days "
            If lHours <> 0 Then sResult &= lHours.ToString & "hours "
            If lMinutes <> 0 Then sResult &= lMinutes.ToString & "minutes "
            If lSeconds <> 0 OrElse sResult = "" Then sResult &= lSeconds.ToString & "seconds "
            Return sResult.Trim()
        Else
            Return lDays.ToString("00") & ":" & lHours.ToString("00") & ":" & lMinutes.ToString("00") & ":" & lSeconds.ToString("00")
        End If
    End Function

    Public Function GetPlayerTitle(ByVal pyTitle As Byte, ByVal pbIsMale As Boolean) As String
        Dim sTemp As String = ""
        Select Case pyTitle
            Case PlayerRank.Baron
                If pbIsMale = True Then sTemp = "Baron" Else sTemp = "Baroness"
            Case PlayerRank.Duke
                If pbIsMale = True Then sTemp = "Duke" Else sTemp = "Lady"
            Case PlayerRank.Emperor
                If pbIsMale = True Then sTemp = "Emperor" Else sTemp = "Empress"
            Case PlayerRank.Governor
                If pbIsMale = True Then sTemp = "Governor" Else sTemp = "Governess"
            Case PlayerRank.King
                If pbIsMale = True Then sTemp = "King" Else sTemp = "Queen"
            Case PlayerRank.Magistrate
                sTemp = "Magistrate"
            Case PlayerRank.Overseer
                sTemp = "Overseer"
        End Select

        Return sTemp '& sName
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

#Region "  GNS DB Connection  "
    Private moGNS_CN As Odbc.OdbcConnection

    Public ReadOnly Property goGNS_CN() As Odbc.OdbcConnection
        Get
            Try
                If moGNS_CN Is Nothing = True OrElse moGNS_CN.State = ConnectionState.Broken OrElse moGNS_CN.State = ConnectionState.Closed OrElse moGNS_CN.State = ConnectionState.Connecting Then
                    Dim sConnStr As String = "DSN=WebDB"
                    'sConnStr &= AppDomain.CurrentDomain.BaseDirectory
                    'If sConnStr.EndsWith("\") = False Then sConnStr &= "\"
                    'sConnStr &= "GNS.udl"
                    moGNS_CN = New Odbc.OdbcConnection(sConnStr)
                    moGNS_CN.Open()
                End If
            Catch ex As Exception
                LogEvent("goGNS_CN: " & ex.Message)
            End Try

            Return moGNS_CN
        End Get
    End Property

    Public Sub CloseConn()
        If moGNS_CN Is Nothing = False Then moGNS_CN.Close()
        moGNS_CN = Nothing
    End Sub
#End Region

#Region "  Name Generation  "
    Private msMaleFirst() As String
    Private msFemaleFirst() As String
    Private msLastName() As String
    Private mbNameGenerationLoaded As Boolean = False

    Public Sub GenerateName(ByVal bMale As Boolean, ByRef sFirstName As String, ByRef sLastName As String)
        If mbNameGenerationLoaded = False Then LoadNameGeneration()

        If bMale = True Then
            sFirstName = msMaleFirst(CInt(Int(Rnd() * msMaleFirst.Length)))
        Else
            sFirstName = msFemaleFirst(CInt(Int(Rnd() * msFemaleFirst.Length)))
        End If
        sLastName = msLastName(CInt(Int(Rnd() * msLastName.Length)))
    End Sub

    Public Function GetFullGeneratedName(ByVal bMale As Boolean) As String
        Dim sFirstName As String = ""
        Dim sLastName As String = ""
        GenerateName(bMale, sFirstName, sLastName)
        Return sFirstName & " " & sLastName
    End Function

    Private Sub LoadNameGeneration()
        mbNameGenerationLoaded = True

        Dim oFile As IO.FileStream
        Dim oReader As IO.StreamReader
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"

        oFile = New IO.FileStream(sPath & "Male.txt", IO.FileMode.Open)
        oReader = New IO.StreamReader(oFile)

        ReDim msMaleFirst(-1)
        While oReader.EndOfStream = False
            ReDim Preserve msMaleFirst(msMaleFirst.GetUpperBound(0) + 1)
            msMaleFirst(msMaleFirst.GetUpperBound(0)) = oReader.ReadLine
        End While
        oReader.Dispose()
        oReader = Nothing
        oFile.Dispose()
        oFile = Nothing

        oFile = New IO.FileStream(sPath & "Female.txt", IO.FileMode.Open)
        oReader = New IO.StreamReader(oFile)

        ReDim msFemaleFirst(-1)
        While oReader.EndOfStream = False
            ReDim Preserve msFemaleFirst(msFemaleFirst.GetUpperBound(0) + 1)
            msFemaleFirst(msFemaleFirst.GetUpperBound(0)) = oReader.ReadLine
        End While
        oReader.Dispose()
        oReader = Nothing
        oFile.Dispose()
        oFile = Nothing

        oFile = New IO.FileStream(sPath & "LastName.txt", IO.FileMode.Open)
        oReader = New IO.StreamReader(oFile)

        ReDim msLastName(-1)
        While oReader.EndOfStream = False
            ReDim Preserve msLastName(msLastName.GetUpperBound(0) + 1)
            msLastName(msLastName.GetUpperBound(0)) = oReader.ReadLine
        End While
        oReader.Dispose()
        oReader = Nothing
        oFile.Dispose()
        oFile = Nothing
    End Sub

#End Region
End Module