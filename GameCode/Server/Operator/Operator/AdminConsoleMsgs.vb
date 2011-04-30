Public Class AdminConsoleMsg
    Public Shared Function HandleAdminLogin(ByVal yData() As Byte) As Int32
        Dim lPos As Int32 = 2 ' msgcode
        Dim sUserName As String = GetStringFromBytes(yData, lPos, 20).ToUpper : lPos += 20
        Dim sPassword As String = GetStringFromBytes(yData, lPos, 20).ToUpper : lPos += 20

        Dim lResult As Int32 = -1
        'And now lets hardwire they are valid
        Dim sSQL As String = "SELECT * FROM tblPlayer where playerusername = '" & sUserName & "' and playerpassword = '" & sPassword & "'"
        Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
        Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader(CommandBehavior.Default)
        If oData.Read = True Then
            lResult = CInt(oData("PlayerID"))
        End If
        LogEvent(LogEventType.Informational, "Admin User (" & sUserName & ") has logged in with Admin ID " & lResult)
        oData.Close()
        oData = Nothing
        oComm = Nothing

        Return lResult

    End Function

    Public Shared Function HandleRequestDxDiag(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim lPos As Int32 = 2 ' msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim yResponse() As Byte = Nothing

        'And now lets get the DXDiag
        LogEvent(LogEventType.Informational, "DXDiag requested for player " & lPlayerID & " by Admin " & lAdminID)
        Dim sSQL As String = "SELECT * FROM tblPlayerDxDiag where playerid = " & lPlayerID
        Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
        Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader(CommandBehavior.Default)
        If oData.Read = True Then
            Dim sResult As String = CStr(oData("dxdiag"))
            If sResult.Length > Int16.MaxValue Then sResult = sResult.Substring(0, 30000)
            ReDim yResponse(9 + sResult.Length)
            lPos = 0
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestDxDiag).CopyTo(yResponse, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResponse, lPos) : lPos += 4
            System.BitConverter.GetBytes(sResult.Length).CopyTo(yResponse, lPos) : lPos += 4
            StringToBytes(sResult).CopyTo(yResponse, lPos) : lPos += sResult.Length

        End If
        oData.Close()
        oData = Nothing
        oComm = Nothing

        Return yResponse


    End Function

    Public Shared Function HandleRequestPlayerDetails(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()

        Dim lPos As Int32 = 2
        Dim sUserName As String = GetStringFromBytes(yData, lPos, 20).ToUpper : lPos += 20
        Dim lPlayerID As Int32 = -1

        Dim sFirstName As String = ""
        Dim sLastName As String = ""
        Dim sAddress As String = ""
        Dim sCity As String = ""
        Dim sState As String = ""
        Dim sZipCode As String = ""
        Dim sEmailAddress As String = ""
        Dim sNetworkIP As String = ""
        Dim sWebPW As String = ""

        Dim yResponse() As Byte = Nothing

        Dim sSQL As String = "SELECT * FROM nuke_users where "
        Dim needsOR As Boolean = False
        If sUserName.Length > 0 Then
            sSQL = sSQL & "username = '" & sUserName & "' "
        End If

        Dim oComm As Odbc.OdbcCommand = New Odbc.OdbcCommand(sSQL, goWebCN)
        Dim oData As Odbc.OdbcDataReader = oComm.ExecuteReader(CommandBehavior.Default)
        Dim counter As Int32 = 0
        Dim userlist() As Byte = Nothing

        If oData.Read = True Then
            sFirstName = CStr(oData("FirstName"))
            sLastName = CStr(oData("LastName"))
            sAddress = CStr(oData("StreetAddress"))
            sCity = CStr(oData("City"))
            sState = CStr(oData("StateProv"))
            sZipCode = CStr(oData("ZipCode"))
            sEmailAddress = CStr(oData("user_email"))
            sWebPW = CStr(oData("user_password"))
        End If

        sSQL = "Select * from tblPlayer where playerusername = '" & sUserName & "'"
        Dim oCommGame As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
        Dim oDataGame As OleDb.OleDbDataReader = oCommGame.ExecuteReader(CommandBehavior.Default)
        Dim iAccountStatus As Int32 = 0
        Dim iLastLogin As Int32 = 0
        Dim iTotalPlayTime As Int32 = 0


        If oDataGame.Read = True Then
            lPlayerID = CInt(oDataGame("PlayerID"))
        End If
        oDataGame.Close()
        oDataGame = Nothing
        oCommGame = Nothing

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then
            'there is no player with that ID?  PROBLEM
            LogEvent(LogEventType.Warning, "Admin User " & lAdminID & " Requesting Player Details of a non player: " & sFirstName & " " & sLastName & "(" & lPlayerID & ")")
            Return Nothing
        Else

            ReDim yResponse(670)
            lPos = 0
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestPlayerDetails).CopyTo(yResponse, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResponse, lPos) : lPos += 4
            StringToBytes(sUserName).CopyTo(yResponse, lPos) : lPos += 20
            StringToBytes(sFirstName).CopyTo(yResponse, lPos) : lPos += 100
            StringToBytes(sLastName).CopyTo(yResponse, lPos) : lPos += 100
            StringToBytes(sAddress).CopyTo(yResponse, lPos) : lPos += 40
            StringToBytes(sCity).CopyTo(yResponse, lPos) : lPos += 20
            StringToBytes(sState).CopyTo(yResponse, lPos) : lPos += 20
            StringToBytes(sZipCode).CopyTo(yResponse, lPos) : lPos += 10
            StringToBytes(sEmailAddress).CopyTo(yResponse, lPos) : lPos += 255
            System.BitConverter.GetBytes(1233331).CopyTo(yResponse, lPos) : lPos += 4
            System.BitConverter.GetBytes(oPlayer.TotalPlayTime).CopyTo(yResponse, lPos) : lPos += 4
            System.BitConverter.GetBytes(GetDateAsNumber(oPlayer.LastLogin)).CopyTo(yResponse, lPos) : lPos += 4
            'TODO: HARDCODE
            System.BitConverter.GetBytes(10222).CopyTo(yResponse, lPos) : lPos += 4 'HomePrimary
            System.BitConverter.GetBytes(10277722).CopyTo(yResponse, lPos) : lPos += 4 'LastPrimary
            StringToBytes("LastRegionHardCode").CopyTo(yResponse, lPos) : lPos += 20 'LastEnvironment
            System.BitConverter.GetBytes(10222).CopyTo(yResponse, lPos) : lPos += 4 'LastPrimaryRequest

            oPlayer.PlayerPassword.CopyTo(yResponse, lPos) : lPos += 20
            StringToBytes(sWebPW).CopyTo(yResponse, lPos) : lPos += 20

            Dim subs() As Byte = Nothing

            Dim subcounter As Int32 = 0

            sSQL = "Select * from fc_subscriptions where email = '" & sEmailAddress & "' order by startdate desc"

            Dim oCommSuite As Odbc.OdbcCommand = New Odbc.OdbcCommand(sSQL, goSuiteCN)
            Dim oDataSuite As Odbc.OdbcDataReader = oCommSuite.ExecuteReader(CommandBehavior.Default)
            Dim sublPos As Int32 = 0
            While oDataSuite.Read = True
                subcounter += 1
                ReDim Preserve subs(subcounter * 74)
                StringToBytes(CStr(oDataSuite("SubID"))).CopyTo(subs, sublPos) : sublPos += 32
                System.BitConverter.GetBytes(GetDateAsNumber(CDate(oDataSuite("StartDate")))).CopyTo(subs, sublPos) : sublPos += 4
                System.BitConverter.GetBytes(GetDateAsNumber(CDate(oDataSuite("EndDate")))).CopyTo(subs, sublPos) : sublPos += 4
                System.BitConverter.GetBytes(CInt(oDataSuite("TotalRecur"))).CopyTo(subs, sublPos) : sublPos += 4
                StringToBytes(CStr(oDataSuite("Interval"))).CopyTo(subs, sublPos) : sublPos += 30
            End While
            If subs Is Nothing = False Then
                ReDim Preserve yResponse(lPos + 4 + subs.Length)
                System.BitConverter.GetBytes(subcounter).CopyTo(yResponse, lPos) : lPos += 4 ' Number of subscription entries
                subs.CopyTo(yResponse, lPos) : lPos += subs.Length
            Else
                ReDim Preserve yResponse(lPos + 4)
                System.BitConverter.GetBytes(subcounter).CopyTo(yResponse, lPos) : lPos += 4 ' Number of subscription entries
            End If

            oDataSuite.Close()
            oDataSuite = Nothing
            oCommSuite = Nothing

        End If

        'aliases 
        If oPlayer.lAliasUB > -1 Then
            Dim aliascount As Integer = 0
            Dim aliases() As Byte = Nothing

            Dim aliaslPos As Int32 = 0
            For x As Int32 = 0 To oPlayer.lAliasUB
                If oPlayer.lAliasIdx(x) > -1 Then
                    aliascount += 1
                    ReDim Preserve aliases(aliascount * 28)
                    oPlayer.oAliases(x).PlayerUserName.CopyTo(aliases, aliaslPos) : aliaslPos += 20
                    System.BitConverter.GetBytes(GetDateAsNumber(oPlayer.oAliases(x).LastLogin)).CopyTo(aliases, aliaslPos) : aliaslPos += 4
                    System.BitConverter.GetBytes(oPlayer.oAliases(x).ObjectID).CopyTo(aliases, aliaslPos) : aliaslPos += 4
                End If
            Next x
            ReDim Preserve yResponse(lPos + 4 + aliases.Length)
            System.BitConverter.GetBytes(aliascount).CopyTo(yResponse, lPos) : lPos += 4
            aliases.CopyTo(yResponse, lPos) : lPos += aliases.Length


        End If




        LogEvent(LogEventType.Informational, "Admin User " & lAdminID & " Requesting Player Details: " & sFirstName & " " & sLastName)


        oData.Close()
        oData = Nothing
        oComm = Nothing

        Return yResponse
    End Function

    Public Shared Function HandleAliasPermissions(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sUserName As String = GetStringFromBytes(yData, lPos, 20).ToUpper : lPos += 20

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

        Dim yResponse() As Byte = Nothing
        If oPlayer Is Nothing Then
            'there is no player with that ID?  PROBLEM
            LogEvent(LogEventType.Warning, "Admin User " & lAdminID & " Requesting Alias Details of a non player: " & "(" & lPlayerID & ")")
        Else
            For X As Int32 = 0 To oPlayer.lAliasUB
                If oPlayer.lAliasIdx(X) > -1 Then
                    If oPlayer.oAliases(X).sPlayerName = sUserName.ToUpper Then
                        'System.BitConverter.GetBytes( .uAliasLogin(x).lRights()

                    End If
                End If
            Next X
        End If
        Return yResponse
    End Function

    Public Shared Function HandleSearchForPlayer(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()

        Dim lPos As Int32 = 2
        Dim sUserName As String = GetStringFromBytes(yData, lPos, 20).ToUpper : lPos += 20
        Dim sEmailAddress As String = GetStringFromBytes(yData, lPos, 255).ToUpper : lPos += 255
        Dim sFirstName As String = GetStringFromBytes(yData, lPos, 100).ToUpper : lPos += 100
        Dim sLastName As String = GetStringFromBytes(yData, lPos, 100).ToUpper : lPos += 100

        Dim yResponse() As Byte = Nothing

        Dim sSQL As String = "SELECT * FROM nuke_users where "
        Dim needsOR As Boolean = False
        If sUserName.Length > 0 Then
            sSQL = sSQL & "username like '" & sUserName & "%' "
        End If
        If sLastName.length > 0 Then
            If needsOR Then
                sSQL = sSQL & " or "
            End If
            sSQL = sSQL & " (firstname  like '" & sFirstName & "%' and lastname like '" & sLastName & "%' )"
        End If
        If sEmailAddress.length > 0 Then
            If needsOR Then
                sSQL = sSQL & " or "
            End If
            sSQL = sSQL & " user_email like '" & sEmailAddress & "%' "
        End If

        Dim oComm As Odbc.OdbcCommand = New Odbc.OdbcCommand(sSQL, goWebCN)
        Dim oData As Odbc.OdbcDataReader = oComm.ExecuteReader(CommandBehavior.Default)
        Dim counter As Int32 = 0
        Dim userlist() As Byte = Nothing

        While oData.Read = True
            Dim respUser As String = CStr(oData("username"))
            If userlist Is Nothing Then ReDim userlist(-1)
            ReDim Preserve userlist(userlist.GetUpperBound(0) + 20)
            StringToBytes(respUser).CopyTo(userlist, counter * 20) : counter += 1
        End While

        LogEvent(LogEventType.Informational, "Admin User Search ")

        ReDim yResponse(5 + userlist.Length)

        System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.SearchForPlayer).CopyTo(yResponse, 0)

        System.BitConverter.GetBytes(counter).CopyTo(yResponse, 2)
        userlist.CopyTo(yResponse, 6)


        oData.Close()
        oData = Nothing
        oComm = Nothing

        Return yResponse
    End Function

    Public Shared Function HandleGetPlayerColonies(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()

        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)

        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") requesting Colonies for player " & lPlayerID)

        Try
            Dim sSQL As String = "select tblColony.*, (Select PlanetID FROm tblPlanet WHERE PlanetID = tblColony.ParentID And tblColony.ParentTypeID = 3"
            sSQL &= " UNION select tblStructure.ParentID FROM tblStructure WHERE tblColony.ParentTypeID = 10 AND tblColony.ParentID = tblStructure.StructureID) As ActualParentID"
            sSQL &= " FROM tblColony Where OwnerID = " & lPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)

            oData = oComm.ExecuteReader()

            Dim lID(-1) As Int32
            Dim sName(-1) As String
            Dim yTax(-1) As Byte
            Dim sParentName(-1) As String
            Dim lUB As Int32 = -1

            While oData.Read = True
                'ok, get the stuff we need.
                lUB += 1
                ReDim Preserve lID(lUB)
                ReDim Preserve sName(lUB)
                ReDim Preserve yTax(lUB)
                ReDim Preserve sParentName(lUB)
                lID(lUB) = CInt(oData("ColonyID"))
                sName(lUB) = CStr(oData("ColonyName"))
                yTax(lUB) = CByte(oData("TaxRate"))
                sParentName(lUB) = "Unknown"

                Dim iParentTypeID As Int16 = CShort(oData("ParentTypeID"))
                Dim lActualParentID As Int32 = CInt(oData("ActualParentID"))
                If iParentTypeID = ObjectType.ePlanet Then
                    Dim oPlanet As Planet = GetEpicaPlanet(lActualParentID)
                    If oPlanet Is Nothing = False Then sParentName(lUB) = BytesToString(oPlanet.PlanetName)
                Else
                    Dim oSystem As SolarSystem = GetEpicaSystem(lActualParentID)
                    If oSystem Is Nothing = False Then sParentName(lUB) = BytesToString(oSystem.SystemName)
                End If
            End While

            'Now, configure our response
            ReDim yResp(9 + (45 * (lUB + 1)))
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestColoniesForPlayer).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lUB + 1).CopyTo(yResp, lPos) : lPos += 4
            For X As Int32 = 0 To lUB
                System.BitConverter.GetBytes(lID(X)).CopyTo(yResp, lPos) : lPos += 4
                StringToBytes(sName(X)).CopyTo(yResp, lPos) : lPos += 20
                yResp(lPos) = yTax(X) : lPos += 1
                StringToBytes(sParentName(X)).CopyTo(yResp, lPos) : lPos += 20
            Next X

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleGetPlayerColonies: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp
    End Function

    Private Structure AgentSkill
        Public AgentID As Int32
        Public SkillID As Int32
        Public SkillValue As Int16
    End Structure
    Public Shared Function HandleRequestAgentsForPlayer(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)

        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") requesting agents for player " & lPlayerID)

        Try
            Dim sSQL As String = "SELECT * FROM tblAgentSkill WHERE AgentID IN (Select AgentID FROM tblAgent WHERE OwnerID = " & lPlayerID & ")"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            Dim uSkills() As AgentSkill = Nothing
            Dim lUB As Int32 = -1

            While oData.Read = True
                lUB += 1
                ReDim Preserve uSkills(lUB)
                With uSkills(lUB)
                    .AgentID = CInt(oData("AgentID"))
                    .SkillID = CInt(oData("SkillID"))
                    .SkillValue = CShort(oData("SkillValue"))
                End With
            End While
            oData.Close()
            oComm.Dispose()

            Dim yTemp() As Byte
            Dim lTempUB As Int32 = 9

            ReDim yTemp(lTempUB)

            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestAgentsForPlayer).CopyTo(yTemp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yTemp, lPos) : lPos += 4
            Dim lCntPos As Int32 = lPos
            lPos += 4
            Dim lCnt As Int32 = 0

            sSQL = "SELECT * FROM tblAgent WHERE OwnerID = " & lPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                Dim lSkillCnt As Int32 = 0
                Dim lAgentID As Int32 = CInt(oData("AgentID"))
                For X As Int32 = 0 To lUB
                    If uSkills(X).AgentID = lAgentID Then
                        lSkillCnt += 1
                    End If
                Next X
                lCnt += 1

                lTempUB += ((lSkillCnt * 6) + 121)
                ReDim Preserve yTemp(lTempUB)
                System.BitConverter.GetBytes(lAgentID).CopyTo(yTemp, lPos) : lPos += 4
                StringToBytes(CStr(oData("AgentName"))).CopyTo(yTemp, lPos) : lPos += 30
                yTemp(lPos) = CByte(oData("Infiltration")) : lPos += 1
                yTemp(lPos) = CByte(oData("Dagger")) : lPos += 1
                yTemp(lPos) = CByte(oData("Resourcefulness")) : lPos += 1
                yTemp(lPos) = CByte(oData("Luck")) : lPos += 1
                yTemp(lPos) = CByte(oData("Loyalty")) : lPos += 1
                yTemp(lPos) = CByte(oData("InfiltrationLevel")) : lPos += 1
                yTemp(lPos) = CByte(oData("InfiltrationType")) : lPos += 1
                Dim lTargetID As Int32 = CInt(oData("TargetID"))
                Dim iTargetTypeID As Int16 = CShort(oData("TargetTypeID"))
                Dim sTarget As String = "Unassigned"
                If iTargetTypeID = ObjectType.ePlayer Then
                    Dim oPlayer As Player = GetEpicaPlayer(lTargetID)
                    If oPlayer Is Nothing = False Then sTarget = oPlayer.sPlayerNameProper
                ElseIf iTargetTypeID = ObjectType.ePlanet Then
                    Dim oPlanet As Planet = GetEpicaPlanet(lTargetID)
                    If oPlanet Is Nothing = False Then sTarget = BytesToString(oPlanet.PlanetName)
                ElseIf iTargetTypeID = ObjectType.eSolarSystem Then
                    Dim oSystem As SolarSystem = GetEpicaSystem(lTargetID)
                    If oSystem Is Nothing = False Then sTarget = BytesToString(oSystem.SystemName)
                End If
                System.BitConverter.GetBytes(lTargetID).CopyTo(yTemp, lPos) : lPos += 4
                System.BitConverter.GetBytes(iTargetTypeID).CopyTo(yTemp, lPos) : lPos += 2
                StringToBytes(sTarget).CopyTo(yTemp, lPos) : lPos += 20
                System.BitConverter.GetBytes(CInt(oData("MissionID"))).CopyTo(yTemp, lPos) : lPos += 4
                yTemp(lPos) = CByte(oData("IsMale")) : lPos += 1
                System.BitConverter.GetBytes(CInt(oData("RecruitedOn"))).CopyTo(yTemp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("UpfrontCost"))).CopyTo(yTemp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("MaintCost"))).CopyTo(yTemp, lPos) : lPos += 4

                yTemp(lPos) = CByte(oData("Health")) : lPos += 1
                System.BitConverter.GetBytes(CInt(oData("AgentStatus"))).CopyTo(yTemp, lPos) : lPos += 4

                Dim lCapturedByID As Int32 = CInt(oData("CapturedBy"))
                System.BitConverter.GetBytes(lCapturedByID).CopyTo(yTemp, lPos) : lPos += 4
                If lCapturedByID > -1 Then
                    Dim oCapturedBy As Player = GetEpicaPlayer(lCapturedByID)
                    If oCapturedBy Is Nothing = False Then
                        oCapturedBy.PlayerName.CopyTo(yTemp, lPos)
                    End If
                End If
                lPos += 20

                System.BitConverter.GetBytes(lSkillCnt).CopyTo(yTemp, lPos) : lPos += 4
                For X As Int32 = 0 To lUB
                    If uSkills(X).AgentID = lAgentID Then
                        System.BitConverter.GetBytes(uSkills(X).SkillID).CopyTo(yTemp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(uSkills(X).SkillValue).CopyTo(yTemp, lPos) : lPos += 2
                    End If
                Next X
            End While

            System.BitConverter.GetBytes(lCnt).CopyTo(yTemp, lCntPos) ': lPos += 4

            yResp = yTemp


        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleRequestAgentsForPlayer: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp

    End Function

    Private Structure StationListing
        Public StructureID As Int32
        Public StructureName As String
        Public ColonyID As Int32
        Public ColonyName As String
        Public lStatus As Int32
        Public lLocX As Int32
        Public lLocZ As Int32
        Public lParentID As Int32
        Public iParentTypeID As Int16
        Public sParentName As String

        Public Children() As StationListing
        Public ChildrenUB As Int32
    End Structure
    Public Shared Function HandleRequestStationsForPlayer(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)


        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") requesting stations for player " & lPlayerID)

        Try
            Dim sSQL As String = "SELECT tblStructure.*, (SELECT ColonyName FROM tblColony WHERE tblColony.ColonyID = tblStructure.ColonyID) As ColonyName FROM tblStructure WHERE OwnerID = " & lPlayerID & " AND ParentTypeID = " & ObjectType.eSolarSystem
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            Dim uStations(-1) As StationListing
            Dim lStationUB As Int32 = -1

            While oData.Read = True
                lStationUB += 1
                ReDim Preserve uStations(lStationUB)
                With uStations(lStationUB)
                    ReDim .Children(-1)
                    .ChildrenUB = -1

                    If oData("ColonyName") Is DBNull.Value = False Then .ColonyName = CStr(oData("ColonyName")) Else .ColonyName = "No Colony"
                    .ColonyID = CInt(oData("ColonyID"))
                    .iParentTypeID = CShort(oData("ParentTypeID"))
                    .lLocX = CInt(oData("LocX"))
                    .lLocZ = CInt(oData("LocY"))
                    .lParentID = CInt(oData("ParentID"))
                    .lStatus = CInt(oData("CurrentStatus"))
                    .sParentName = "Unknown"

                    Dim oSystem As SolarSystem = GetEpicaSystem(.lParentID)
                    If oSystem Is Nothing = False Then
                        .sParentName = BytesToString(oSystem.SystemName)
                    End If

                    .StructureID = CInt(oData("StructureID"))
                    .StructureName = CStr(oData("FacilityName"))
                End With
            End While
            oData.Close()
            oComm.Dispose()

            Dim lChildren As Int32 = 0

            sSQL = "SELECT * FROM tblStructure WHERE OwnerID = " & lPlayerID & " AND ParentTypeID = " & ObjectType.eFacility
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                Dim lParentID As Int32 = CInt(oData("parentID"))
                For X As Int32 = 0 To lStationUB
                    If uStations(X).StructureID = lParentID Then
                        uStations(X).ChildrenUB += 1

                        lChildren += 1

                        ReDim Preserve uStations(X).Children(uStations(X).ChildrenUB)
                        With uStations(X).Children(uStations(X).ChildrenUB)
                            .lStatus = CInt(oData("CurrentStatus"))
                            .StructureID = CInt(oData("StructureID"))
                            .StructureName = CStr(oData("FacilityName"))
                        End With
                    End If
                Next X
            End While

            'Now, create our msg
            Dim lPos As Int32 = 0
            ReDim yResp(9 + ((lStationUB + 1) * 90) + (lChildren * 28)) '9 + ((lStationUB + 1) * 90) + (lChildren * 28)
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestStationsForPlayer).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lStationUB + 1).CopyTo(yResp, lPos) : lPos += 4
            For X As Int32 = 0 To lStationUB
                With uStations(X)
                    System.BitConverter.GetBytes(.StructureID).CopyTo(yResp, lPos) : lPos += 4
                    StringToBytes(.StructureName).CopyTo(yResp, lPos) : lPos += 20
                    System.BitConverter.GetBytes(.ColonyID).CopyTo(yResp, lPos) : lPos += 4
                    StringToBytes(.ColonyName).CopyTo(yResp, lPos) : lPos += 20
                    System.BitConverter.GetBytes(.lParentID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.iParentTypeID).CopyTo(yResp, lPos) : lPos += 2
                    StringToBytes(.sParentName).CopyTo(yResp, lPos) : lPos += 20
                    System.BitConverter.GetBytes(.lLocX).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lLocZ).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lStatus).CopyTo(yResp, lPos) : lPos += 4

                    System.BitConverter.GetBytes(.ChildrenUB + 1).CopyTo(yResp, lPos) : lPos += 4

                    For Y As Int32 = 0 To .ChildrenUB
                        System.BitConverter.GetBytes(.Children(Y).StructureID).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(.Children(Y).StructureName).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(.Children(Y).lStatus).CopyTo(yResp, lPos) : lPos += 4
                    Next Y
                End With
            Next X

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleRequestStationsForPlayer: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp

    End Function

    Public Shared Function HandleGetPlayerSpecialTechs(ByVal yData() As Byte, ByVal lAdminID As Int32, ByRef oSocket As NetSock) As Byte()
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)


        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") requesting player special techs for player " & lPlayerID)

        Try
            Dim sSQL As String = "SELECT tblSpecialTech.TechID, tblspecialtech.TechName, tblspecialtech.MaxLinkChanceAttempts, tblspecialtech.BenefitsDesc, tblspecialtech.NewAttrValue, " & _
              "tblplayerspecialtech.LinkAttempts, tblplayerspecialtech.bArchived, tblplayerspecialtech.SuccessfulLink, tblplayerspecialtech.CreditResearchAttempts, " & _
              "tblplayerspecialtech.ResearchAttempts from tblplayerspecialtech LEFT OUTER JOIN tblSpecialTech ON tblPlayerSpecialTech.SpecialTechID = tblSpecialTech.TechID " & _
              "WHERE tblPlayerSpecialTech.PlayerID = " & lPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            Dim lTempUB As Int32 = 9
            Dim yTemp(lTempUB) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestSpecialTechsForPlayer).CopyTo(yTemp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yTemp, lPos) : lPos += 4
            Dim lCntPos As Int32 = lPos
            Dim lCnt As Int32 = 0
            lPos += 4

            While oData.Read = True

                If lTempUB > 30000 Then
                    System.BitConverter.GetBytes(lCnt).CopyTo(yTemp, lCntPos)
                    oSocket.SendData(yTemp)
                    lTempUB = 9
                    ReDim yTemp(lTempUB)
                    System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestSpecialTechsForPlayer).CopyTo(yTemp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lPlayerID).CopyTo(yTemp, lPos) : lPos += 4
                    lCntPos = lPos
                    lPos += 4
                    lCnt = 0
                End If

                lCnt += 1

                Dim sDesc As String = CStr(oData("BenefitsDesc"))

                lTempUB += (79 + sDesc.Length)
                ReDim Preserve yTemp(lTempUB)

                System.BitConverter.GetBytes(CInt(oData("TechID"))).CopyTo(yTemp, lPos) : lPos += 4
                Dim sName As String = CStr(oData("TechName"))
                StringToBytes(sName).CopyTo(yTemp, lPos) : lPos += 50
                System.BitConverter.GetBytes(CInt(oData("MaxLinkChanceAttempts"))).CopyTo(yTemp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("NewAttrValue"))).CopyTo(yTemp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("LinkAttempts"))).CopyTo(yTemp, lPos) : lPos += 4

                Dim yVal As Byte
                If CByte(oData("bArchived")) <> 0 Then yVal = 1
                If CByte(oData("SuccessfulLink")) <> 0 Then yVal = CByte(yVal Or 128)
                yTemp(lPos) = yVal : lPos += 1

                System.BitConverter.GetBytes(CInt(oData("CreditResearchAttempts"))).CopyTo(yTemp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("ResearchAttempts"))).CopyTo(yTemp, lPos) : lPos += 4

                System.BitConverter.GetBytes(sDesc.Length).CopyTo(yTemp, lPos) : lPos += 4
                StringToBytes(sDesc).CopyTo(yTemp, lPos) : lPos += sDesc.Length

            End While

            System.BitConverter.GetBytes(lCnt).CopyTo(yTemp, lCntPos)
            yResp = yTemp

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleGetPlayerSpecialTechs: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp
    End Function

    Public Shared Function HandleRequestComponentsForPlayer(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        Try
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)

            LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") requesting component design list for player " & lPlayerID)

            Dim sSQL As String = "select alloyid As TechID, 6 As TechTypeID, alloyname As TechName, resphase, ErrorReasonCode, MajorDesignFlaw, bArchived FROM tblAlloy WHERE OwnerID = " & lPlayerID
            sSQL &= " UNION Select ArmorID, 7, armorname, resphase, ErrorReasonCode, MajorDesignFlaw, bArchived FROM tblArmor WHERE OwnerID = " & lPlayerID
            sSQL &= " UNION Select EngineID, 9, enginename, resphase, ErrorReasonCode, MajorDesignFlaw, bArchived FROM tblEngine WHERE OwnerID = " & lPlayerID
            sSQL &= " UNION Select HullID, 15, hullname, resphase, 0, 0, bArchived FROM tblHull WHERE OwnerID = " & lPlayerID
            sSQL &= " UNION SELECT PrototypeID, 23, prototypename, resphase, 0, 0, bArchived FROM tblPrototype WHERE OwnerID = " & lPlayerID
            sSQL &= " UNION SELECT RadarID, 24, RadarName, resphase, ErrorReasonCode, MajorDesignFlaw, bArchived FROM tblRadar where ownerid = " & lPlayerID
            sSQL &= " UNION SELECT ShieldID, 27, shieldname, resphase, ErrorReasonCode, MajorDesignFlaw, bArchived FROM tblshield where ownerid = " & lPlayerID
            sSQL &= " UNION Select WeaponID, 34, weaponname, resphase, ErrorReasonCode, MajorDesignFlaw, bArchived FROM tblweapon where ownerid = " & lPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            'We are only returning a list of items
            '  ID (4)
            '  TypeID (2)
            '  Name (20)
            '  ResearchPhase (4)
            '  ErrorCode (1)
            '  MajorDesignFlaw (1)
            '  Archived (1)

            Dim lTempUB As Int32 = 9
            Dim yTemp(lTempUB) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestComponentsForPlayer).CopyTo(yTemp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yTemp, lPos) : lPos += 4
            Dim lCntPos As Int32 = lPos
            Dim lCnt As Int32 = 0
            lPos += 4

            While oData.Read() = True
                lCnt += 1

                lTempUB += 33
                ReDim Preserve yTemp(lTempUB)

                System.BitConverter.GetBytes(CInt(oData("TechID"))).CopyTo(yTemp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CShort(oData("TechTypeID"))).CopyTo(yTemp, lPos) : lPos += 2
                StringToBytes(CStr(oData("TechName"))).CopyTo(yTemp, lPos) : lPos += 20
                System.BitConverter.GetBytes(CInt(oData("ResPhase"))).CopyTo(yTemp, lPos) : lPos += 4
                yTemp(lPos) = CByte(oData("ErrorReasonCode")) : lPos += 1
                yTemp(lPos) = CByte(oData("MajorDesignFlaw")) : lPos += 1
                yTemp(lPos) = CByte(oData("bArchived")) : lPos += 1
            End While
            System.BitConverter.GetBytes(lCnt).CopyTo(yTemp, lCntPos)

            yResp = yTemp

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleRequestComponentsForPlayer: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp
    End Function
    Public Shared Function HandleRequestComponentDetails(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        Try
            Dim lTechID As Int32 = System.BitConverter.ToInt32(yData, 2)
            Dim iTechTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

            Dim sSQL As String
            Dim sFields() As String = Nothing
            Select Case iTechTypeID
                Case ObjectType.eAlloyTech
                    sSQL = "select a.*, m1.mineralname as mineral1name, m2.mineralname as mineral2name, m3.mineralname as mineral3name, m4.mineralname as mineral4name"
                    sSQL &= ", mp1.MineralPropertyName As MinPropName1, mp2.MineralPropertyName As MinPropName2, mp3.MineralPropertyName As MinPropName3 "
                    sSQL &= " from tblalloy a left outer join tblmineral as m1 on m1.mineralid = a.mineral1id left outer join tblmineral as m2 on m2.mineralid = a.mineral2id"
                    sSQL &= " left outer join tblmineral as m3 on m3.mineralid = a.mineral3id left outer join tblmineral as m4 on m4.mineralid = a.mineral4id "
                    sSQL &= " left outer join tblMineralProperty As mp1 ON mp1.MineralPropertyID = a.MineralProperty1ID LEFT OUTER JOIN tblMineralProperty As mp2 "
                    sSQL &= " ON mp2.MineralPropertyID = a.MineralProperty2ID LEFT OUTER JOIN tblMineralProperty As mp3 ON Mp3.MineralPropertyID = a.MineralProperty3ID "
                    sSQL &= " WHERE AlloyID = " & lTechID

                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader()
                    If oData.Read() = True Then
                        ReDim yResp(209)
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestComponentDetails).CopyTo(yResp, lPos) : lPos += 2

                        System.BitConverter.GetBytes(lTechID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(iTechTypeID).CopyTo(yResp, lPos) : lPos += 2
                        StringToBytes(CStr(oData("AlloyName"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("Mineral1ID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("Mineral1Name") Is DBNull.Value = False Then StringToBytes(CStr(oData("Mineral1Name"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("Mineral2ID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("Mineral2Name") Is DBNull.Value = False Then StringToBytes(CStr(oData("Mineral2Name"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("Mineral3ID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("Mineral3Name") Is DBNull.Value = False Then StringToBytes(CStr(oData("Mineral3Name"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("Mineral4ID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("Mineral4Name") Is DBNull.Value = False Then StringToBytes(CStr(oData("Mineral4Name"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("MineralProperty1ID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("MinPropName1") Is DBNull.Value = False Then StringToBytes(CStr(oData("MinPropName1"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("MineralProperty2ID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("MinPropName2") Is DBNull.Value = False Then StringToBytes(CStr(oData("MinPropName2"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("MineralProperty3ID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("MinPropName3") Is DBNull.Value = False Then StringToBytes(CStr(oData("MinPropName3"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("Property1Value"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("Property2Value"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("Property3Value"))).CopyTo(yResp, lPos) : lPos += 4
                        yResp(lPos) = CByte(oData("PopIntel")) : lPos += 1
                        yResp(lPos) = CByte(oData("ResearchLevel")) : lPos += 1
                    End If

                Case ObjectType.eArmorTech
                    'sSQL = "SELECT * FROM tblArmor WHERE ArmorID = " & lTechID
                    sSQL = "select a.*, m1.mineralname as mineral1name, m2.mineralname as mineral2name, m3.mineralname as mineral3name from tblArmor a "
                    sSQL &= " left outer join tblmineral as m1 on m1.mineralid = a.OuterLayerMineralID "
                    sSQL &= " left outer join tblmineral as m2 on m2.mineralid = a.MiddleLayerMineralID "
                    sSQL &= " left outer join tblmineral as m3 on m3.mineralid = a.InnerLayerMineralID "
                    sSQL &= " WHERE ArmorID = " & lTechID

                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader()
                    If oData.Read() = True Then
                        ReDim yResp(170)
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestComponentDetails).CopyTo(yResp, lPos) : lPos += 2

                        System.BitConverter.GetBytes(lTechID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(iTechTypeID).CopyTo(yResp, lPos) : lPos += 2
                        StringToBytes(CStr(oData("ArmorName"))).CopyTo(yResp, lPos) : lPos += 20
                        yResp(lPos) = CByte(oData("PiercingResist")) : lPos += 1
                        yResp(lPos) = CByte(oData("ImpactResist")) : lPos += 1
                        yResp(lPos) = CByte(oData("BeamResist")) : lPos += 1
                        yResp(lPos) = CByte(oData("ECMResist")) : lPos += 1
                        yResp(lPos) = CByte(oData("FlameResist")) : lPos += 1
                        yResp(lPos) = CByte(oData("ChemicalResist")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("HitPoints"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("HullUsagePerPlate"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("OuterLayerMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("mineral1name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("MiddleLayerMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("mineral2name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("InnerLayerMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("mineral3name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CSng(oData("RandomSeed"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("Integrity"))).CopyTo(yResp, lPos) : lPos += 4

                        yResp(lPos) = CByte(oData("PopIntel")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedHull"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedPower"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedColonist"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedEnlisted"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedOfficer"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin1"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin2"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin3"))).CopyTo(yResp, lPos) : lPos += 4
                    End If
  
                Case ObjectType.eEngineTech
                    'sSQL = "SELECT * FROM tblEngine WHERE EngineID = " & lTechID
                    sSQL = "select e.*, m1.mineralname as mineral1name, m2.mineralname as mineral2name, m3.mineralname as mineral3name,"
                    sSQL &= " m4.mineralname as mineral4name, m5.mineralname as mineral5name, m6.mineralname as mineral6name from tblEngine e "
                    sSQL &= " left outer join tblmineral as m1 on m1.mineralid = e.StructuralBodyMineralID "
                    sSQL &= " left outer join tblmineral as m2 on m2.mineralid = e.StructuralFrameMineralID "
                    sSQL &= " left outer join tblmineral as m3 on m3.mineralid = e.StructuralMeldMineralID "
                    sSQL &= " left outer join tblmineral as m4 on m4.mineralid = e.DriveBodyMineralID "
                    sSQL &= " left outer join tblmineral as m5 on m5.mineralid = e.DriveFrameMineralID "
                    sSQL &= " left outer join tblmineral as m6 on m6.mineralid = e.DriveMeldMineralID "
                    sSQL &= " WHERE EngineID = " & lTechID

                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader()
                    If oData.Read() = True Then
                        ReDim yResp(354)
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestComponentDetails).CopyTo(yResp, lPos) : lPos += 2

                        System.BitConverter.GetBytes(lTechID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(iTechTypeID).CopyTo(yResp, lPos) : lPos += 2
                        StringToBytes(CStr(oData("EngineName"))).CopyTo(yResp, lPos) : lPos += 20
                        yResp(lPos) = CByte(oData("HullTypeID")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("Thrust"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("PowerProd"))).CopyTo(yResp, lPos) : lPos += 4
                        yResp(lPos) = CByte(oData("Maneuver")) : lPos += 1
                        yResp(lPos) = CByte(oData("Speed")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("StructuralBodyMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral1Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("StructuralFrameMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral2Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("StructuralMeldMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral3Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("DriveBodyMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral4Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("DriveFrameMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral5Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("DriveMeldMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral6Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CSng(oData("RandomSeed"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("HullRequired"))).CopyTo(yResp, lPos) : lPos += 4

                        yResp(lPos) = CByte(oData("PopIntel")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedHull"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedPower"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedColonist"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedEnlisted"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedOfficer"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin1"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin2"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin3"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin4"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin5"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin6"))).CopyTo(yResp, lPos) : lPos += 4

                    End If
                Case ObjectType.eHullTech
                    sSQL = "SELECT * FROM tblHull WHERE HullID = " & lTechID
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader()
                    If oData.Read() = True Then
                        ReDim yResp(37)
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestComponentDetails).CopyTo(yResp, lPos) : lPos += 2

                        System.BitConverter.GetBytes(lTechID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(iTechTypeID).CopyTo(yResp, lPos) : lPos += 2
                        StringToBytes(CStr(oData("HullName"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CShort(oData("ModelID"))).CopyTo(yResp, lPos) : lPos += 2
                        System.BitConverter.GetBytes(CInt(oData("HullSize"))).CopyTo(yResp, lPos) : lPos += 4
                        yResp(lPos) = CByte(oData("TypeID")) : lPos += 1
                        yResp(lPos) = CByte(oData("SubTypeID")) : lPos += 1
                        yResp(lPos) = CByte(oData("ChassisType")) : lPos += 1
                        yResp(lPos) = CByte(oData("PopIntel")) : lPos += 1
                    End If

                Case ObjectType.ePrototype
                    sSQL = "SELECT *, (Select EngineName FROM tblEngine WHERE tblEngine.EngineID = tblPrototype.EngineID) As EngineName, "
                    sSQL &= " (SELECT ArmorName FROM tblArmor WHERE tblArmor.ArmorID = tblPrototype.ArmorID) As ArmorName, "
                    sSQL &= " (SELECT HullName FROM tblHull WHERE TblHull.HullID = tblPrototype.HullID) AS HullName, "
                    sSQL &= " (SELECT RadarName FROM tblRadar WHERE tblRadar.RadarID = tblPrototype.RadarID) AS RadarName, "
                    sSQL &= " (SELECT ShieldName FROM tblShield WHERE tblShield.ShieldID = tblPrototype.ShieldID) AS ShieldName"
                    sSQL &= " FROM tblPrototype WHERE PrototypeID = " & lTechID
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader()
                    If oData.Read = True Then
                        ReDim yResp(176)
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestComponentDetails).CopyTo(yResp, lPos) : lPos += 2

                        System.BitConverter.GetBytes(lTechID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(iTechTypeID).CopyTo(yResp, lPos) : lPos += 2
                        StringToBytes(CStr(oData("PrototypeName"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("EngineID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("EngineName") Is DBNull.Value = False Then StringToBytes(CStr(oData("EngineName"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("ArmorID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("ArmorName") Is DBNull.Value = False Then StringToBytes(CStr(oData("ArmorName"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("HullID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("HullName") Is DBNull.Value = False Then StringToBytes(CStr(oData("HullName"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("RadarID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("RadarName") Is DBNull.Value = False Then StringToBytes(CStr(oData("RadarName"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("ShieldID"))).CopyTo(yResp, lPos) : lPos += 4
                        If oData("ShieldName") Is DBNull.Value = False Then StringToBytes(CStr(oData("ShieldName"))).CopyTo(yResp, lPos)
                        lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("ForeArmorUnits"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("LeftArmorUnits"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("AftArmorUnits"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("RightArmorUnits"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("MaxCrew"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("MinCrew"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("ProductionHull"))).CopyTo(yResp, lPos) : lPos += 4
                        yResp(lPos) = CByte(oData("PopIntel")) : lPos += 1
                    End If
                Case ObjectType.eRadarTech
                    'sSQL = "SELECT * FROM tblRadar WHERE RadarID = " & lTechID

                    sSQL = "select r.*, m1.mineralname as mineral1name,  m2.mineralname as mineral2name, m3.mineralname as mineral3name,"
                    sSQL &= " m4.mineralname as mineral4name from tblRadar r left outer join tblmineral as m1 on m1.mineralid = r.CasingMineralID"
                    sSQL &= " left outer join tblmineral as m2 on m2.mineralid = r.CollectionMineralID "
                    sSQL &= " left outer join tblmineral as m3 on m3.mineralid = r.EmitterMineralID "
                    sSQL &= " left outer join tblmineral as m4 on m4.mineralid = r.DetectionMineralID WHERE RadarID = " & lTechID

                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader()
                    If oData.Read = True Then
                        ReDim yResp(194)
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestComponentDetails).CopyTo(yResp, lPos) : lPos += 2

                        System.BitConverter.GetBytes(lTechID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(iTechTypeID).CopyTo(yResp, lPos) : lPos += 2
                        StringToBytes(CStr(oData("RadarName"))).CopyTo(yResp, lPos) : lPos += 20
                        yResp(lPos) = CByte(oData("RadarType")) : lPos += 1
                        yResp(lPos) = CByte(oData("WeaponAcc")) : lPos += 1
                        yResp(lPos) = CByte(oData("ScanResolution")) : lPos += 1
                        yResp(lPos) = CByte(oData("OptimumRange")) : lPos += 1
                        yResp(lPos) = CByte(oData("MaximumRange")) : lPos += 1
                        yResp(lPos) = CByte(oData("DisruptionResist")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("CasingMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral1Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("CollectionMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral2Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("EmitterMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral3Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("DetectionMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral4Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("PowerRequired"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("HullRequired"))).CopyTo(yResp, lPos) : lPos += 4
                        yResp(lPos) = CByte(oData("JamImmunity")) : lPos += 1
                        yResp(lPos) = CByte(oData("JamStrength")) : lPos += 1
                        yResp(lPos) = CByte(oData("JamTargets")) : lPos += 1
                        yResp(lPos) = CByte(oData("JamEffect")) : lPos += 1

                        yResp(lPos) = CByte(oData("PopIntel")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedHull"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedPower"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedColonist"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedEnlisted"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedOfficer"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin1"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin2"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin3"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin4"))).CopyTo(yResp, lPos) : lPos += 4
                    End If

                Case ObjectType.eShieldTech
                    'sSQL = "SELECT * FROM tblShield WHERE ShieldID = " & lTechID

                    sSQL = "select s.*, m1.mineralname as mineral1name, m2.mineralname as mineral2name, m3.mineralname as mineral3name "
                    sSQL &= " from tblshield s left outer join tblmineral as m1 on m1.mineralid = s.CasingMineralID"
                    sSQL &= " left outer join tblmineral as m2 on m2.mineralid = s.CoilMineralID "
                    sSQL &= " left outer join tblmineral as m3 on m3.mineralid = s.AcceleratorMineralID WHERE ShieldID = " & lTechID

                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader()
                    If oData.Read = True Then
                        ReDim yResp(173)
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestComponentDetails).CopyTo(yResp, lPos) : lPos += 2

                        System.BitConverter.GetBytes(lTechID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(iTechTypeID).CopyTo(yResp, lPos) : lPos += 2
                        StringToBytes(CStr(oData("ShieldName"))).CopyTo(yResp, lPos) : lPos += 20
                        yResp(lPos) = CByte(oData("HullTypeID")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("MaxHitPoints"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("RechargeRate"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("RechargeFreq"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("ProjectionHullSize"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("CasingMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral1Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("CoilMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral2Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("AcceleratorMineralID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral3Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("PowerRequired"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("HullRequired"))).CopyTo(yResp, lPos) : lPos += 4

                        yResp(lPos) = CByte(oData("PopIntel")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedHull"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedPower"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedColonist"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedEnlisted"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedOfficer"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin1"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin2"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin3"))).CopyTo(yResp, lPos) : lPos += 4 
                    End If

                Case ObjectType.eWeaponTech
                    sSQL = "SELECT w.*, m1.MineralName As Mineral1Name, m2.MineralName As Mineral2Name, m3.MineralName As Mineral3Name, "
                    sSQL &= "m4.MineralName As Mineral4Name, m5.MineralName As Mineral5Name FROM tblWeapon w LEFT OUTER JOIN tblMineral m1 "
                    sSQL &= " on m1.MineralID = w.Mineral1ID LEFT OUTER JOIN tblMineral m2 on m2.MineralID = w.Mineral2ID LEFT OUTER JOIN "
                    sSQL &= " tblMineral m3 on m3.MineralID = w.Mineral3ID LEFT OUTER JOIN tblMineral m4 on m4.MineralID = w.Mineral4ID"
                    sSQL &= " LEFT OUTER JOIN tblMineral m5 on m5.MineralID = w.Mineral5ID WHERE WeaponID = " & lTechID
                    ' FROM tblWeapon WHERE WeaponID = " & lTechID
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader()
                    If oData.Read = True Then
                        ReDim yResp(249)
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestComponentDetails).CopyTo(yResp, lPos) : lPos += 2

                        System.BitConverter.GetBytes(lTechID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(iTechTypeID).CopyTo(yResp, lPos) : lPos += 2
                        StringToBytes(CStr(oData("WeaponName"))).CopyTo(yResp, lPos) : lPos += 20
                        yResp(lPos) = CByte(oData("HullTypeID")) : lPos += 1
                        yResp(lPos) = CByte(oData("WeaponClassType")) : lPos += 1
                        yResp(lPos) = CByte(oData("WeaponTypeID")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("PowerRequired"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("HullRequired"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CShort(oData("ROF"))).CopyTo(yResp, lPos) : lPos += 2
                        System.BitConverter.GetBytes(CInt(oData("MinDmg"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("MaxDmg"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("OptimumRange"))).CopyTo(yResp, lPos) : lPos += 4
                        yResp(lPos) = CByte(oData("PierceRating")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("Mineral1ID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral1Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("Mineral2ID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral2Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("Mineral3ID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral3Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("Mineral4ID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral4Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("Mineral5ID"))).CopyTo(yResp, lPos) : lPos += 4
                        StringToBytes(CStr(oData("Mineral5Name"))).CopyTo(yResp, lPos) : lPos += 20
                        System.BitConverter.GetBytes(CInt(oData("ShotHullSize"))).CopyTo(yResp, lPos) : lPos += 4
                        yResp(lPos) = CByte(oData("MaxSpeed")) : lPos += 1
                        yResp(lPos) = CByte(oData("Maneuver")) : lPos += 1
                        System.BitConverter.GetBytes(CShort(oData("FlightTime"))).CopyTo(yResp, lPos) : lPos += 2
                        yResp(lPos) = CByte(oData("Accuracy")) : lPos += 1
                        yResp(lPos) = CByte(oData("PayloadType")) : lPos += 1
                        yResp(lPos) = CByte(oData("ExplosionRadius")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("StructureHP"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("ProjectileHullSize"))).CopyTo(yResp, lPos) : lPos += 4

                        yResp(lPos) = CByte(oData("PopIntel")) : lPos += 1
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedHull"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedPower"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedProdTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResCost"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedResTime"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedColonist"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedEnlisted"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedOfficer"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin1"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin2"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin3"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin4"))).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CInt(oData("SpecifiedMin5"))).CopyTo(yResp, lPos) : lPos += 4
                    End If


                Case Else
                    LogEvent(LogEventType.CriticalError, "Admin.HandleRequestComponentDetails Admin " & lAdminID & " requesting invalid type.")
                    Return Nothing
            End Select

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleRequestComponentDetails: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp

    End Function

    Public Shared Function HandleRequestDiplomacyContacts(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()

        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        Try
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
            LogEvent(LogEventType.Informational, "Admin " & lAdminID & " requesting diplomacy contacts for " & lPlayerID)

            Dim sSQL As String = "SELECT pr1.player2id, pr1.reltypeid, p.PlayerName, pr2.reltypeid As TheirScore "
            sSQL &= " from tblPlayerToPlayerRel pr1 LEFT OUTER JOIN tblPlayerToPlayerRel pr2 ON pr1.Player2ID = pr2.player1id AND pr1.player1id = pr2.player2id"
            sSQL &= " LEFT OUTER JOIN tblPlayer p ON pr1.player2id = p.playerid WHERE pr1.player1id = " & lPlayerID

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            Dim lPos As Int32 = 0
            Dim lTempUB As Int32 = 9
            ReDim yResp(lTempUB)
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestDiplomacyContacts).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
            Dim lCntPos As Int32 = lPos
            Dim lCnt As Int32 = 0
            lPos += 4

            While oData.Read() = True
                lCnt += 1
                lTempUB += 26
                ReDim Preserve yResp(lTempUB)

                System.BitConverter.GetBytes(CInt(oData("Player2ID"))).CopyTo(yResp, lPos) : lPos += 4
                yResp(lPos) = CByte(oData("RelTypeID")) : lPos += 1
                StringToBytes(CStr(oData("PlayerName"))).CopyTo(yResp, lPos) : lPos += 20
                yResp(lPos) = CByte(oData("TheirScore")) : lPos += 1
            End While
            System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lCntPos)

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleRequestDiplomacyContacts: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp

    End Function

    Public Shared Function HandleRequestEntitiesByEnvironment(ByVal yData() As Byte, ByVal lAdminID As Int32, ByRef oSocket As NetSock) As Byte()

        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        Try
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
            Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 6)
            Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 10)
            Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)

            LogEvent(LogEventType.Informational, "Admin " & lAdminID & " Request player entities for " & lPlayerID & ". (" & lEnvirID & ", " & iEnvirTypeID & ", " & iObjTypeID & ")")

            Dim sSQL As String

            If iObjTypeID = ObjectType.eUnit Then
                sSQL = "SELECT u.unitid as ObjID, ud.unitdefid as ObjDefID, ud.defname As ObjDefName, u.UnitName As ObjName, u.LocX, u.LocY, u.LocAngle, u.Structure_HP, "
                sSQL &= "ud.Structure_MaxHP, u.Shield_HP, ud.Shield_MaxHP, u.Q1_HP, u.Q2_HP, u.Q3_HP, u.Q4_HP, ud.Q1_MaxHP, ud.Q2_MaxHP, ud.Q3_MaxHP, "
                sSQL &= "ud.Q4_MaxHP, u.ExpLevel, u.CurrentStatus, ud.ProductionTypeID, ud.RequiredProductionTypeID, u.CombatTactics, u.TargetingTactics, "
                sSQL &= "u.Hangar_Cap, ud.Hangar_Cap As MaxHangarCap, u.Cargo_Cap, ud.Cargo_Cap As MaxCargoCap, ud.MaxSpeed, ud.Maneuver, ud.ChassisType, "
                sSQL &= "ud.OptRadarRange, ud.MaxRadarRange, ud.Weapon_acc, ud.ScanResolution, ud.DisruptionResistance, ud.PiercingResist, ud.ImpactResist, "
                sSQL &= "ud.BeamResist, ud.ECMResist, ud.FlameResist, ud.ChemicalResist"
                sSQL &= " FROM TblUnit u LEFT OUTER JOIN tblUnitDef ud ON u.UnitDefID = ud.UnitDefID WHERE u.ParentID = "
                sSQL &= lEnvirID & " AND u.ParentTypeID = " & iEnvirTypeID & " AND u.OwnerID = " & lPlayerID & " ORDER BY u.UnitName"
            ElseIf iObjTypeID = ObjectType.eFacility Then
                sSQL = "SELECT  s.StructureID As ObjID, sd.FacilityDefID As ObjDefID, sd.FacilityDefName As ObjDefName, s.FacilityName As ObjName, "
                sSQL &= "s.LocX, s.LocY, s.LocAngle, s.Structure_HP, sd.Structure_MaxHP, s.ExpLevel, s.CurrentStatus, sd.ProductionTypeID, sd.RequiredProductionTypeID, "
                sSQL &= "s.CombatTactics, s.TargetingTactics, s.shield_hp, sd.shield_maxhp, s.q1_hp, s.q2_hp, s.q3_hp, s.q4_hp, sd.q1_maxhp, sd.q2_maxhp, sd.q3_maxhp, sd.q4_maxhp, s.Hangar_Cap, sd.Hangar_Cap AS MaxHangarCap, s.Cargo_Cap, sd.Cargo_Cap as MaxCargoCap, sd.MaxSpeed, "
                sSQL &= "sd.Maneuver, sd.ChassisType, sd.OptRadarRange, sd.MaxRadarRange, sd.Weapon_Acc, sd.ScanResolution, sd.DisruptionResistance, sd.PiercingResist, "
                sSQL &= "sd.ImpactResist, sd.BeamResist, sd.ECMResist, sd.FlameResist, sd.ChemicalResist from tblStructure s LEFT OUTER JOIN tblStructureDef sd ON "
                sSQL &= "s.FacilityDefID = sd.FacilityDefID WHERE ParentID = "
                sSQL &= lEnvirID & " AND ParentTypeID = " & iEnvirTypeID & " AND s.OwnerID = " & lPlayerID & " ORDER BY FacilityName"
            Else : Return Nothing
            End If

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            Dim lTempUB As Int32 = 17
            ReDim yResp(lTempUB)

            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestUnitsByEnvironment).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lEnvirID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, lPos) : lPos += 2

            Dim lCntPos As Int32 = lPos
            lPos += 4
            Dim lCnt As Int32 = 0

            While oData.Read = True
                If lTempUB > 30000 Then
                    System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lCntPos)
                    oSocket.SendData(yResp)

                    lTempUB = 17
                    ReDim yResp(lTempUB)
                    lPos = 0

                    System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestUnitsByEnvironment).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(lEnvirID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, lPos) : lPos += 2
                    lCntPos = lPos
                    lCnt = 0
                End If

                lCnt += 1

                lTempUB += 147
                ReDim Preserve yResp(lTempUB)

                System.BitConverter.GetBytes(CInt(oData("ObjID"))).CopyTo(yResp, lPos) : lPos += 4
                StringToBytes(CStr(oData("ObjDefName"))).CopyTo(yResp, lPos) : lPos += 20
                System.BitConverter.GetBytes(CInt(oData("ObjDefID"))).CopyTo(yResp, lPos) : lPos += 4
                StringToBytes(CStr(oData("ObjName"))).CopyTo(yResp, lPos) : lPos += 20
                System.BitConverter.GetBytes(CInt(oData("LocX"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("LocY"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CShort(oData("LocAngle"))).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(CInt(oData("Structure_HP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Structure_MaxHP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Shield_HP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Shield_MaxHP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Q1_HP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Q2_HP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Q3_HP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Q4_HP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Q1_MaxHP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Q2_MaxHP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Q3_MaxHP"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Q4_MaxHP"))).CopyTo(yResp, lPos) : lPos += 4
                yResp(lPos) = CByte(oData("ExpLevel")) : lPos += 1
                System.BitConverter.GetBytes(CInt(oData("CurrentStatus"))).CopyTo(yResp, lPos) : lPos += 4
                yResp(lPos) = CByte(oData("ProductionTypeID")) : lPos += 1
                yResp(lPos) = CByte(oData("RequiredProductionTypeID")) : lPos += 1
                System.BitConverter.GetBytes(CShort(oData("CombatTactics"))).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(CShort(oData("TargetingTactics"))).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(CInt(oData("Hangar_Cap"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("MaxHangarCap"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Cargo_Cap"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("MaxCargoCap"))).CopyTo(yResp, lPos) : lPos += 4
                yResp(lPos) = CByte(oData("MaxSpeed")) : lPos += 1
                yResp(lPos) = CByte(oData("Maneuver")) : lPos += 1
                yResp(lPos) = CByte(oData("ChassisType")) : lPos += 1
                yResp(lPos) = CByte(oData("OptRadarRange")) : lPos += 1
                yResp(lPos) = CByte(oData("MaxRadarRange")) : lPos += 1
                yResp(lPos) = CByte(oData("Weapon_Acc")) : lPos += 1
                yResp(lPos) = CByte(oData("ScanResolution")) : lPos += 1
                yResp(lPos) = CByte(oData("DisruptionResistance")) : lPos += 1
                yResp(lPos) = CByte(oData("PiercingResist")) : lPos += 1
                yResp(lPos) = CByte(oData("ImpactResist")) : lPos += 1
                yResp(lPos) = CByte(oData("BeamResist")) : lPos += 1
                yResp(lPos) = CByte(oData("ECMResist")) : lPos += 1
                yResp(lPos) = CByte(oData("FlameResist")) : lPos += 1
                yResp(lPos) = CByte(oData("ChemicalResist")) : lPos += 1
            End While
            System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lCntPos)

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleRequestEntitiesByEnvironment: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp
    End Function

    Public Shared Function HandleRequestEntityContents(ByVal yData() As Byte, ByVal lAdminID As Int32, ByRef oSocket As NetSock) As Byte()
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        Try
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
            Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

            LogEvent(LogEventType.Informational, "Admin " & lAdminID & " requesting Entity Cargo for " & lObjID & ", " & iObjTypeID)

            'Cargo and Hangar... can be units, components, minerals

            Dim sSQL As String = "SELECT tblUnit.UnitID, tblUnit.UnitName, tblUnitDef.HullSize, tblUnitDef.UnitDefID FROM tblUnit LEFT OUTER JOIN tblUnitDef ON tblUnit.UnitDefID = tblUnitDef.UnitDefID WHERE ParentID = " & lObjID & " AND ParentTypeID = " & iObjTypeID

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            Dim lPos As Int32 = 0
            Dim lTempUB As Int32 = 11
            ReDim yResp(lTempUB)
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestUnitCargo).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lObjID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, lPos) : lPos += 2

            Dim lCntPos As Int32 = lPos
            Dim lCnt As Int32 = 0
            lPos += 4

            While oData.Read() = True
                If lTempUB > 30000 Then
                    System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lCntPos)
                    oSocket.SendData(yResp)
                    lTempUB = 11
                    ReDim yResp(lTempUB)
                    lPos = 0
                    System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestUnitCargo).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lObjID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, lPos) : lPos += 2
                    lCntPos = lPos
                    lPos += 4
                End If

                lCnt += 1
                lTempUB += 34
                ReDim Preserve yResp(lTempUB)

                System.BitConverter.GetBytes(CInt(oData("UnitID"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(ObjectType.eUnit).CopyTo(yResp, lPos) : lPos += 2
                StringToBytes(CStr(oData("UnitName"))).CopyTo(yResp, lPos) : lPos += 20
                System.BitConverter.GetBytes(CInt(oData("HullSize"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("UnitDefID"))).CopyTo(yResp, lPos) : lPos += 4
            End While

            sSQL = "SELECT mc.CacheID, mc.MineralID, mc.Quantity, mc.Concentration, m.MineralName FROM tblmineralCache mc Left outer join tblMineral m ON mc.MineralID = m.MineralID WHERE ParentID = " & lObjID & " AND ParentTypeID = " & iObjTypeID
            oData.Close()
            oComm.Dispose()

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            While oData.Read = True

                If lTempUB > 30000 Then
                    System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lCntPos)
                    oSocket.SendData(yResp)
                    lTempUB = 11
                    ReDim yResp(lTempUB)
                    lPos = 0
                    System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestUnitCargo).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lObjID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, lPos) : lPos += 2
                    lCntPos = lPos
                    lPos += 4
                End If

                lCnt += 1
                lTempUB += 38
                ReDim Preserve yResp(lTempUB)

                System.BitConverter.GetBytes(CInt(oData("CacheID"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(ObjectType.eMineralCache).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(CInt(oData("Quantity"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Concentration"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("MineralID"))).CopyTo(yResp, lPos) : lPos += 4
                StringToBytes(CStr(oData("MineralName"))).CopyTo(yResp, lPos) : lPos += 20
            End While

            oData.Close()
            oComm.Dispose()

            sSQL = "select cc.ComponentCacheID, cc.ComponentID, cc.ComponentTypeID, cc.ComponentOwnerID,  cc.Quantity,"
            sSQL &= " (Select ArmorName FROM tblArmor WHERE ArmorID = cc.ComponentID) As ComponentName "
            sSQL &= " from tblcomponentcache cc WHERE ComponentTypeID = " & ObjectType.eArmorTech & " AND cc.ParentID = " & lObjID
            sSQL &= " AND cc.ParentTypeID = " & iObjTypeID & " UNION "
            sSQL &= " select cc.ComponentCacheID, cc.ComponentID, cc.ComponentTypeID, cc.ComponentOwnerID,  cc.Quantity,"
            sSQL &= "(Select EngineName FROM tblEngine WHERE EngineID = cc.ComponentID) As ComponentName"
            sSQL &= " from tblcomponentcache cc WHERE ComponentTypeID = " & ObjectType.eEngineTech & " AND cc.ParentID = " & lObjID
            sSQL &= " AND cc.ParentTypeID = " & iObjTypeID & " UNION "
            sSQL &= "select cc.ComponentCacheID, cc.ComponentID, cc.ComponentTypeID, cc.ComponentOwnerID,  cc.Quantity,"
            sSQL &= " (Select RadarName FROM tblRadar WHERE RadarID = cc.ComponentID) As ComponentName"
            sSQL &= " from tblcomponentcache cc WHERE ComponentTypeID = " & ObjectType.eRadarTech & " AND cc.ParentID = " & lObjID
            sSQL &= " AND cc.ParentTypeID = " & iObjTypeID & " UNION "
            sSQL &= " select cc.ComponentCacheID, cc.ComponentID, cc.ComponentTypeID, cc.ComponentOwnerID,  cc.Quantity,"
            sSQL &= " (Select ShieldName FROM tblShield WHERE ShieldID = cc.ComponentID) As ComponentName"
            sSQL &= " from tblcomponentcache cc WHERE ComponentTypeID = " & ObjectType.eShieldTech & " AND cc.ParentID = " & lObjID
            sSQL &= " AND cc.ParentTypeID = " & iObjTypeID & " UNION "
            sSQL &= " select cc.ComponentCacheID, cc.ComponentID, cc.ComponentTypeID, cc.ComponentOwnerID,  cc.Quantity,"
            sSQL &= " (Select WeaponName FROM tblWeapon WHERE WeaponID = cc.ComponentID) As ComponentName"
            sSQL &= " from tblcomponentcache cc WHERE ComponentTypeID = " & ObjectType.eWeaponTech & " AND cc.ParentID = " & lObjID
            sSQL &= " AND cc.ParentTypeID = " & iObjTypeID

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                If lTempUB > 30000 Then
                    System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lCntPos)
                    oSocket.SendData(yResp)
                    lTempUB = 11
                    ReDim yResp(lTempUB)
                    lPos = 0
                    System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestUnitCargo).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lObjID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, lPos) : lPos += 2
                    lCntPos = lPos
                    lPos += 4
                End If

                lCnt += 1
                lTempUB += 40
                ReDim Preserve yResp(lTempUB)

                System.BitConverter.GetBytes(CInt(oData("ComponentCacheID"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(ObjectType.eComponentCache).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(CInt(oData("ComponentID"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CShort(oData("ComponentTypeID"))).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(CInt(oData("ComponentOwnerID"))).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CInt(oData("Quantity"))).CopyTo(yResp, lPos) : lPos += 4
                StringToBytes(CStr(oData("ComponentName"))).CopyTo(yResp, lPos) : lPos += 20
            End While
            System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lCntPos)

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleRequestEntityCargo: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp
    End Function

    Private Structure RouteListing
        Public LocX As Int32
        Public LocZ As Int32
        Public LoadItemID As Int32
        Public LoadItemTypeID As Int16
        Public DestID As Int32
        Public DestTypeID As Int16
        Public OrderNum As Byte
        Public DestName As String
        Public PickupName As String
    End Structure

    Public Shared Function HandleRequestRouteForUnit(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        Try
            Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, 2)
            LogEvent(LogEventType.Informational, "Admin " & lAdminID & " requesting Unit Routes for " & lUnitID)

            Dim sSQL As String = "SELECT * FROM tblRouteItem WHERE UnitID = " & lUnitID & " ORDER BY OrderNum"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            Dim uItems(-1) As RouteListing
            Dim lUB As Int32 = -1

            Dim sFacList As String = ""
            Dim sUnitList As String = ""
            Dim sPlanetList As String = ""
            Dim sSystemList As String = ""
            Dim sMineralList As String = ""

            While oData.Read = True
                lUB += 1
                ReDim Preserve uItems(lUB)
                With uItems(lUB)
                    .DestID = CInt(oData("DestID"))
                    .DestTypeID = CShort(oData("DestTypeID"))
                    .LoadItemID = CInt(oData("LoadItemID"))
                    .LoadItemTypeID = CShort(oData("LoadItemTypeID"))
                    .LocX = CInt(oData("LocX"))
                    .LocZ = CInt(oData("LocZ"))
                    .OrderNum = CByte(oData("OrderNum"))
                    .PickupName = "Unknown"
                    .DestName = "Unknown"

                    If .LoadItemTypeID < 0 Then
                        .PickupName = "Any/All"
                    ElseIf .LoadItemTypeID = ObjectType.eMineral Then
                        If sMineralList <> "" Then sMineralList &= ", "
                        sMineralList &= .LoadItemID.ToString
                    Else : .PickupName = "Unknown"
                    End If

                    If .DestTypeID = ObjectType.eUnit Then
                        If sUnitList <> "" Then sUnitList &= ", "
                        sUnitList &= .DestID
                    ElseIf .DestTypeID = ObjectType.eFacility Then
                        If sFacList <> "" Then sFacList &= ", "
                        sFacList &= .DestID
                    ElseIf .DestTypeID = ObjectType.eSolarSystem Then
                        If sSystemList <> "" Then sSystemList &= ", "
                        sSystemList &= .DestID
                    ElseIf .DestTypeID = ObjectType.ePlanet Then
                        If sPlanetList <> "" Then sPlanetList &= ", "
                        sPlanetList &= .DestID
                    End If
                End With
            End While

            oData.Close()
            oComm.Dispose()

            If sFacList <> "" Then
                sSQL = "SELECT StructureID, FacilityName FROM tblStructure WHERE StructureID IN (" & sFacList & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader()
                While oData.Read = True
                    Dim lID As Int32 = CInt(oData("StructureID"))
                    Dim sName As String = CStr(oData("FacilityName"))
                    For X As Int32 = 0 To lUB
                        If uItems(X).DestID = lID AndAlso uItems(X).DestTypeID = ObjectType.eFacility Then
                            uItems(X).DestName = sName
                        End If
                    Next X
                End While
                oData.Close()
                oComm.Dispose()
            End If
            If sUnitList <> "" Then
                sSQL = "SELECT UnitID, UnitName FROM tblUnit WHERE UnitID IN (" & sUnitList & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader()
                While oData.Read = True
                    Dim lID As Int32 = CInt(oData("UnitID"))
                    Dim sName As String = CStr(oData("UnitName"))
                    For X As Int32 = 0 To lUB
                        If uItems(X).DestID = lID AndAlso uItems(X).DestTypeID = ObjectType.eUnit Then
                            uItems(X).DestName = sName
                        End If
                    Next X
                End While
                oData.Close()
                oComm.Dispose()
            End If
            If sSystemList <> "" Then
                sSQL = "SELECT SystemName, SystemID FROM tblSolarSystem WHERE SystemID IN (" & sSystemList & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader()
                While oData.Read = True
                    Dim lID As Int32 = CInt(oData("SystemID"))
                    Dim sName As String = CStr(oData("SystemName"))
                    For X As Int32 = 0 To lUB
                        If uItems(X).DestID = lID AndAlso uItems(X).DestTypeID = ObjectType.eSolarSystem Then
                            uItems(X).DestName = sName
                        End If
                    Next X
                End While
                oData.Close()
                oComm.Dispose()
            End If
            If sPlanetList <> "" Then
                sSQL = "SELECT PlanetName, PlanetID FROM TblPlanet WHERE PlanetID IN (" & sPlanetList & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader()
                While oData.Read = True
                    Dim lID As Int32 = CInt(oData("PlanetID"))
                    Dim sName As String = CStr(oData("PlanetName"))
                    For X As Int32 = 0 To lUB
                        If uItems(X).DestID = lID AndAlso uItems(X).DestTypeID = ObjectType.ePlanet Then
                            uItems(X).DestName = sName
                        End If
                    Next X
                End While
                oData.Close()
                oComm.Dispose()
            End If
            If sMineralList <> "" Then
                sSQL = "SELECT MineralName, MineralID FROM tblMineral WHERE MineralID IN (" & sMineralList & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader()
                While oData.Read = True
                    Dim lID As Int32 = CInt(oData("MineralID"))
                    Dim sName As String = CStr(oData("MineralName"))
                    For X As Int32 = 0 To lUB
                        If uItems(X).LoadItemID = lID AndAlso uItems(X).LoadItemTypeID = ObjectType.eMineral Then
                            uItems(X).PickupName = sName
                        End If
                    Next X
                End While
                oData.Close()
                oComm.Dispose()
            End If

            Dim lPos As Int32 = 0
            ReDim yResp(((lUB + 1) * 61) + 9)
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.RequestRouteForUnit).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lUnitID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lUB + 1).CopyTo(yResp, lPos) : lPos += 4
            For X As Int32 = 0 To lUB
                With uItems(X)
                    yResp(lPos) = .OrderNum : lPos += 1
                    System.BitConverter.GetBytes(.DestID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.DestTypeID).CopyTo(yResp, lPos) : lPos += 2
                    StringToBytes(.DestName).CopyTo(yResp, lPos) : lPos += 20
                    System.BitConverter.GetBytes(.LoadItemID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.LoadItemTypeID).CopyTo(yResp, lPos) : lPos += 2
                    StringToBytes(.PickupName).CopyTo(yResp, lPos) : lPos += 20
                    System.BitConverter.GetBytes(.LocX).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.LocZ).CopyTo(yResp, lPos) : lPos += 4
                End With
            Next X

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleRequestDiplomacyContacts: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp
    End Function

    Private Shared Function GetIDFromSQL(ByVal sSQL As String, ByVal sIDField As String) As Int32
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim lID As Int32 = -1

        Try
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            If oData Is Nothing = False Then
                If oData.Read = True Then
                    lID = CInt(oData(sIDField))
                End If
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "GetIDFromSQL: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try
        Return lID
    End Function

    Private Shared Sub GetParentFromSQL(ByVal sSQL As String, ByVal sIDField As String, ByVal sTypeIDField As String, ByRef lParentID As Int32, ByVal iParentTypeID As Int16)
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing

        Try
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            If oData Is Nothing = False Then
                If oData.Read = True Then
                    lParentID = CInt(oData(sIDField))
                    iParentTypeID = CShort(oData(sTypeIDField))
                End If
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "GetParentFromSQL: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try
    End Sub

    Private Shared Function GetParentSystem(ByVal lID As Int32, ByVal iTypeID As Int16, ByRef lParentID As Int32, ByRef iParentTypeID As Int16) As Boolean
        Select Case iTypeID
            Case ObjectType.eColony
                GetParentFromSQL("SELECT ParentID, ParentTypeID FROM tblColony WHERE ColonyID = " & lID, "ParentID", "ParentTypeID", lParentID, iParentTypeID)
            Case ObjectType.eComponentCache
                GetParentFromSQL("SELECT ParentID, ParentTypeID FROM tblComponentCache WHERE ComponentCacheID = " & lID, "ParentID", "ParentTypeID", lParentID, iParentTypeID)
            Case ObjectType.eFacility
                GetParentFromSQL("SELECT ParentID, ParentTypeID FROM tblStructure WHERE StructureID = " & lID, "ParentID", "ParentTypeID", lParentID, iParentTypeID)
            Case ObjectType.eGalaxy
                Return False
            Case ObjectType.eGuild
            Case ObjectType.eMineralCache
                GetParentFromSQL("SELECT ParentID, ParentTypeID FROM tblMineralCache WHERE CacheID = " & lID, "ParentID", "ParentTypeID", lParentID, iParentTypeID)
            Case ObjectType.ePlanet
                Dim oPlanet As Planet = GetEpicaPlanet(lID)
                If oPlanet Is Nothing = False Then
                    iParentTypeID = ObjectType.eSolarSystem
                    lParentID = oPlanet.ParentSystem.ObjectID
                End If
            Case ObjectType.eSolarSystem
                lParentID = lID
                iParentTypeID = iTypeID
            Case ObjectType.eTrade
            Case ObjectType.eUnit
                GetParentFromSQL("SELECT ParentID, ParentTypeID FROM tblUnit WHERE UnitID = " & lID, "ParentID", "ParentTypeID", lParentID, iParentTypeID)
            Case ObjectType.eUnitGroup
            Case Else
                Return False
        End Select

        If iParentTypeID = -1 Then Return False
        If iParentTypeID = ObjectType.eSolarSystem Then
            Return True
        Else
            Dim lNewID As Int32 = lParentID
            Dim iNewTypeID As Int16 = iParentTypeID
            Return GetParentSystem(lNewID, iNewTypeID, lParentID, iParentTypeID)
        End If
    End Function

    Public Shared Function GetOwnerPrimary(ByVal lID As Int32, ByVal iTypeID As Int16) As ServerObject
        'now, determine what we are changing
        Dim lPlayerID As Int32 = -1
        Dim lPrimaryIdx As Int32 = -1

        Select Case iTypeID
            Case ObjectType.eAgent
                lPlayerID = GetIDFromSQL("SELECT OwnerID FROM tblAgent WHERE AgentID = " & lID, "OwnerID")
            Case ObjectType.eAlloyTech
                lPlayerID = GetIDFromSQL("SELECT OwnerID FROM tblAlloy WHERE AlloyID = " & lID, "OwnerID")
            Case ObjectType.eArmorTech
                lPlayerID = GetIDFromSQL("SELECT OwnerID FROM tblArmor WHERE ArmorID = " & lID, "OwnerID")
            Case ObjectType.eEngineTech
                lPlayerID = GetIDFromSQL("SELECT OwnerID FROM tblEngine WHERE EngineID = " & lID, "OwnerID")
            Case ObjectType.eHullTech
                lPlayerID = GetIDFromSQL("SELECT OwnerID FROM tblHull WHERE HullID = " & lID, "OwnerID")
            Case ObjectType.ePlayer
                lPlayerID = lID
            Case ObjectType.ePlayerComm
                lPlayerID = GetIDFromSQL("SELECT PlayerID FROM tblPlayerComm WHERE PC_ID = " & lID, "OwnerID")
            Case ObjectType.ePrototype
                lPlayerID = GetIDFromSQL("SELECT OwnerID FROM tblPrototype WHERE PrototypeID = " & lID, "OwnerID")
            Case ObjectType.eRadarTech
                lPlayerID = GetIDFromSQL("SELECT OwnerID FROM tblRadar WHERE RadarID = " & lID, "OwnerID")
            Case ObjectType.eShieldTech
                lPlayerID = GetIDFromSQL("SELECT OwnerID FROM tblShield WHERE ShieldID = " & lID, "OwnerID")
            Case ObjectType.eWeaponTech
                lPlayerID = GetIDFromSQL("SELECT OwnerID FROM tblWeapon WHERE WeaponID = " & lID, "OwnerID")

            Case ObjectType.eColony, ObjectType.eComponentCache, ObjectType.eFacility, ObjectType.eMineralCache, ObjectType.ePlanet, ObjectType.eUnit
                Dim lParentID As Int32 = -1
                Dim iParentTypeID As Int16 = -1
                If GetParentSystem(lID, iTypeID, lParentID, iParentTypeID) = False Then Return Nothing
                If iParentTypeID = ObjectType.eSolarSystem Then
                    Dim oSystem As SolarSystem = GetEpicaSystem(lID)
                    If oSystem Is Nothing = False AndAlso oSystem.oPrimaryServer Is Nothing = False AndAlso oSystem.oPrimaryServer.oSocket Is Nothing = False Then
                        lPrimaryIdx = oSystem.oPrimaryServer.oSocket.SocketIndex
                    End If
                End If
            Case ObjectType.eGalaxy
            Case ObjectType.eGuild
            Case ObjectType.eNebula
            Case ObjectType.eSolarSystem
            Case ObjectType.eTrade
            Case ObjectType.eUnitGroup
            Case ObjectType.eWormhole
        End Select

        If lPrimaryIdx = -1 Then
            If lPlayerID < 1 Then Return Nothing
            Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
            If oPlayer Is Nothing = False Then
                lPrimaryIdx = oPlayer.lOwnerPrimaryIdx
            End If
        End If

        If lPrimaryIdx = -1 Then Return Nothing

        Return goMsgSys.oServerObject(lPrimaryIdx)

    End Function

    Public Shared Function HandleSetEntityName(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim yResp() As Byte = Nothing

        Try
            Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, 2)
            Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
            Dim sName As String = GetStringFromBytes(yData, 8, 20)

            LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") changing entity name (" & lEntityID & ", " & iEntityTypeID & ") to " & sName)

            Dim oPrimary As ServerObject = GetOwnerPrimary(lEntityID, iEntityTypeID)
            If oPrimary Is Nothing = False Then
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityName).CopyTo(yData, 0)
                oPrimary.oSocket.SendData(yData)
            Else
                ReDim yResp(5)
                System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.SetUnitName).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(-2I).CopyTo(yResp, 2)
            End If

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.SetEntityName: " & ex.Message)
            ReDim yResp(5)
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.SetUnitName).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(-1I).CopyTo(yResp, 2)
        End Try

        Return yResp
    End Function

    Public Shared Function HandleRemoveEntity(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim yResp() As Byte = Nothing

        Try
            Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, 2)
            Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

            LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") Removing entity (" & lEntityID & ", " & iEntityTypeID & ")")

            'First, validate it is a proper object type
            Select Case iEntityTypeID
                Case ObjectType.eAgent, ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eColony, ObjectType.eEngineTech, _
                    ObjectType.eFacility, ObjectType.eFacilityDef, ObjectType.eGuild, ObjectType.eHullTech, ObjectType.ePlayerComm, _
                    ObjectType.ePlayerIntel, ObjectType.ePlayerTechKnowledge, ObjectType.ePrototype, ObjectType.eRadarTech, _
                    ObjectType.eSenateLaw, ObjectType.eShieldTech, ObjectType.eUnit, ObjectType.eUnitDef, ObjectType.eWeaponTech
                Case Else
                    ReDim yResp(4)
                    System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.SetUnitName).CopyTo(yResp, 0)
                    System.BitConverter.GetBytes(-3I).CopyTo(yResp, 2)
                    Return yResp
            End Select

            Dim oPrimary As ServerObject = GetOwnerPrimary(lEntityID, iEntityTypeID)
            If oPrimary Is Nothing = False Then
                System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yData, 0)
                oPrimary.oSocket.SendData(yData)
            Else
                ReDim yResp(5)
                System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.SetUnitName).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(-2I).CopyTo(yResp, 2)
            End If

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.SetEntityName: " & ex.Message)
            ReDim yResp(5)
            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.SetUnitName).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(-1I).CopyTo(yResp, 2)
        End Try

        Return yResp
    End Function

    Public Shared Sub HandleClearColonyProductionQueues(ByVal yData() As Byte, ByVal lAdminID As Int32)

        Try
            Dim lColonyID As Int32 = System.BitConverter.ToInt32(yData, 2)

            LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") clearing colony production queues for colony " & lColonyID)

            Dim oPrimary As ServerObject = GetOwnerPrimary(lColonyID, ObjectType.eColony)
            If oPrimary Is Nothing = False Then
                System.BitConverter.GetBytes(GlobalMessageCode.eColonyLowResources).CopyTo(yData, 0)
                oPrimary.oSocket.SendData(yData)
            End If

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.ClearColonyProductionQueues: " & ex.Message)
        End Try

    End Sub

    Public Shared Sub HandleChangeColonyTaxRate(ByVal yData() As Byte, ByVal lAdminID As Int32)
        Try
            Dim lColonyID As Int32 = System.BitConverter.ToInt32(yData, 2)
            Dim yTax As Byte = yData(6)
            LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") change Colony Tax Rate to " & yTax & " for colony " & lColonyID)

            Dim oPrimary As ServerObject = GetOwnerPrimary(lColonyID, ObjectType.eColony)
            If oPrimary Is Nothing = False Then
                System.BitConverter.GetBytes(GlobalMessageCode.eSetColonyTaxRate).CopyTo(yData, 0)
                oPrimary.oSocket.SendData(yData)
            End If

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.ChangeColonyTaxRate: " & ex.Message)
        End Try
    End Sub

    'new as of 09/10/08
    Public Shared Function HandleGetPlayerFactions(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim yResp() As Byte = Nothing

        Try
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
            LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") Getting Faction List for playerID " & lPlayerID)

 
            Dim sSQL As String = "SELECT Slot1ID, Slot2ID, Slot3ID, Slot4ID, Slot5ID, Faction1ID, Faction2ID, Faction3ID FROM tblPlayer WHERE PlayerID = " & lPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            While oData.Read = True
                '
            End While

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleGetPlayerFactions: " & ex.Message)
            yResp = Nothing
        Finally
            If oData Is Nothing = False Then oData.Close()
            If oComm Is Nothing = False Then oComm.Dispose()
            oData = Nothing
            oComm = Nothing
        End Try

        Return yResp
    End Function

    Public Shared Function HandleSetPlayerCredits(ByVal yData() As Byte, ByVal lAdminID As Int32) As Byte()
        Dim yResp(6) As Byte
        Try
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
            Dim blCredits As Int64 = System.BitConverter.ToInt64(yData, 6)
            LogEvent(LogEventType.Informational, "Admin (" & lAdminID & ") setting player " & lPlayerID & " credits to " & blCredits)

            System.BitConverter.GetBytes(MsgSystem.AdminMessageCode.SetPlayerCredits).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, 2)

            Dim yMsg(13) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerCredits).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(blCredits).CopyTo(yMsg, lPos) : lPos += 8
            goMsgSys.SendToPrimaryServers(yMsg)

            yResp(6) = 255
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Admin.HandleSetPlayerCredits: " & ex.Message)
            yResp(6) = 0
        End Try
        Return yResp
    End Function
End Class
