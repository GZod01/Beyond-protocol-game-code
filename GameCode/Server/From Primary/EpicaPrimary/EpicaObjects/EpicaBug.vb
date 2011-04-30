Public Class EpicaBug
    Public lBugID As Int32
    Public yCategory As Byte
    Public ySubCat As Byte
    Public yOccurs As Byte
    Public yPriority As Byte
    Public yStatus As Byte
    Public sProblemDesc As String = String.Empty
    Public sStepsToProduce As String = String.Empty
    Public sDevNotes As String = String.Empty

    Public lCreatedBy As Int32
    Public lAssignedTo As Int32

    Private mbDataChanged As Boolean = True
    Private mySendString() As Byte
    Public Sub DataChanged()
        mbDataChanged = True
    End Sub

    Public Function GetObjAsString() As Byte()
        If mbDataChanged = True Then
            Dim lPos As Int32

            ReDim mySendString(42 + sProblemDesc.Length + sStepsToProduce.Length + sDevNotes.Length)

            System.BitConverter.GetBytes(lBugID).CopyTo(mySendString, 0)
            mySendString(4) = yCategory
            mySendString(5) = ySubCat
            mySendString(6) = yOccurs
            mySendString(7) = yPriority
            mySendString(8) = yStatus

            lPos = 9
            System.BitConverter.GetBytes(lCreatedBy).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(lAssignedTo).CopyTo(mySendString, lPos) : lPos += 4

            'Username... (playername)
            Dim oPlayer As Player = GetEpicaPlayer(lCreatedBy)
            If oPlayer Is Nothing = False Then
                System.Text.ASCIIEncoding.ASCII.GetBytes(oPlayer.sPlayerNameProper).CopyTo(mySendString, lPos)
            End If
            lPos += 20

            'Copy the problem desc
            System.BitConverter.GetBytes(CShort(sProblemDesc.Length)).CopyTo(mySendString, lPos) : lPos += 2
            Array.Copy(System.Text.ASCIIEncoding.ASCII.GetBytes(sProblemDesc), 0, mySendString, lPos, sProblemDesc.Length)
            lPos += sProblemDesc.Length

            'Copy the Steps to Produce
            System.BitConverter.GetBytes(CShort(sStepsToProduce.Length)).CopyTo(mySendString, lPos) : lPos += 2
            Array.Copy(System.Text.ASCIIEncoding.ASCII.GetBytes(sStepsToProduce), 0, mySendString, lPos, sStepsToProduce.Length)
            lPos += sStepsToProduce.Length

            'Copy the Dev Notes
            System.BitConverter.GetBytes(CShort(sDevNotes.Length)).CopyTo(mySendString, lPos) : lPos += 2
            Array.Copy(System.Text.ASCIIEncoding.ASCII.GetBytes(sDevNotes), 0, mySendString, lPos, sDevNotes.Length)
            mbDataChanged = False
        End If
        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand
        Dim bResult As Boolean = False

        Try
            If lBugID < 1 Then
                'Insert
				sSQL = "INSERT INTO tblBug (ProblemDesc, StepsToProduce, UserID, DevNotes, CategoryID, occursID, PriorityID, StatusID, SubCatID, AssignedTo, DateCreated, DateModified)" & _
				  " VALUES ('" & MakeDBStr(sProblemDesc) & "', '" & MakeDBStr(sStepsToProduce) & "', " & lCreatedBy & ", '" & MakeDBStr(sDevNotes) & "', " & _
				  yCategory & ", " & yOccurs & ", " & yPriority & ", " & yStatus & ", " & ySubCat & "," & lAssignedTo & "," & GetDateAsNumber(Now) & ", " & GetDateAsNumber(Now) & ")"
            Else
                'update
				sSQL = "UPDATE tblBug SET ProblemDesc = '" & MakeDBStr(sProblemDesc) & "', StepsToProduce = '" & MakeDBStr(sStepsToProduce) & _
				  "', UserID = " & lCreatedBy & ", DevNotes = '" & MakeDBStr(sDevNotes) & "', CategoryID = " & yCategory & ", OccursID = " & _
				  yOccurs & ", PriorityID = " & yPriority & ", StatusID = " & yStatus & ", SubCatID = " & ySubCat & ", AssignedTo = " & lAssignedTo & _
				  ", DateModified = " & GetDateAsNumber(Now) & " WHERE BugID = " & lBugID
            End If

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If lBugID < 1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(BugID) FROM tblBug WHERE UserID = " & lCreatedBy & " AND CategoryID = " & yCategory
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    lBugID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object BUG " & lBugID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bresult
    End Function

    Public Shared Function GetSubCategoryText(ByVal yCatCode As Byte, ByVal ySubCatCode As Byte) As String
        Select Case yCatCode
            Case 0
                Select Case ySubCatCode
                    Case 0 : Return "Client-Side Crash"
                    Case 1 : Return "Server-Side Crash"
                    Case Else : Return "Other Crash"
                End Select
            Case 1
                Select Case ySubCatCode
                    Case 0 : Return "Gameplay - Exploit"
                    Case 1 : Return "Gameplay - Game Logic Bug"
                    Case 2 : Return "Gameplay - Geography Bug"
                    Case 3 : Return "Gameplay - Unexpected Result"
                    Case 4 : Return "Gameplay - User Interface"
                    Case Else : Return "Gameplay - Other"
                End Select
            Case 2
                Select Case ySubCatCode
                    Case 0 : Return "Graphical - Geography"
                    Case 1 : Return "Graphical - Models"
                    Case 2 : Return "Graphical - Particle FX"
                    Case 3 : Return "Graphical - Textures"
                    Case 4 : Return "Graphical - User Interface"
                    Case Else : Return "Graphical - Other"
                End Select
            Case 3
                Select Case ySubCatCode
                    Case 0 : Return "Performance - Low Frame Rate"
                    Case 1 : Return "Performance - Memory Leak"
                    Case 2 : Return "Performance - Perceived Lag"
                    Case 3 : Return "Performance - Game Stutter"
                    Case Else : Return "Performance - Other"
                End Select
            Case 4
                Return "Other Miscellaneous"
            Case 5
                Select Case ySubCatCode
                    Case 0 : Return "Login - Login Credentials"
                    Case 1 : Return "Login - Connections Issue"
                    Case Else : Return "Login - Other"
                End Select
            Case 6
                Select Case ySubCatCode
                    Case 0 : Return "Suggestions - Gameplay-Related"
                    Case 1 : Return "Suggestions - User Interface"
                    Case 2 : Return "Suggestions - Hotkeys"
                    Case Else : Return "Suggestions - Other"
                End Select
            Case 7
                Select Case ySubCatCode
                    Case 0 : Return "Combat"
                    Case 1 : Return "Movement"
                    Case 2 : Return "User Interface"
                    Case 3 : Return "Performance"
                    Case 4 : Return "Production"
                    Case 5 : Return "Tech Builder"
                    Case 6 : Return "Special Techs"
                    Case 7 : Return "Mining"
                    Case 8 : Return "Trading"
                    Case 9 : Return "Battlegroups"
                    Case 10 : Return "In-Game Email"
                    Case 11 : Return "Diplomacy"
                    Case 12 : Return "Colony/Budget"
                    Case 13 : Return "Chat"
                    Case 14 : Return "Repair"
                    Case 15 : Return "Space stations"
                    Case 16 : Return "Guilds/Alliances"
                    Case 17 : Return "Espionage"
                    Case 18 : Return "Sound"
                    Case Else : Return "Other - be specific"
                End Select
            Case Else
                Return "Unknown"
        End Select
    End Function

    Public Shared Function GetOccurrenceText(ByVal yOccurs As Byte) As String
        Select Case yOccurs
            Case 0 : Return "Easily Reproduceable"
            Case 1 : Return "Intermittent"
            Case Else : Return "Rarely"
        End Select
    End Function

    Public Shared Function GetPriorityText(ByVal yPriority As Byte) As String
        Select Case yPriority
            Case 0 : Return "Extremely High"
            Case 1 : Return "High"
            Case 2 : Return "Normal"
            Case 3 : Return "Low"
            Case Else : Return "Extremely Low"
        End Select
    End Function

    Public Shared Function GetStatusText(ByVal yStatus As Byte) As String
        Select Case yStatus
            Case 0 : Return "New"
            Case 1 : Return "Open"
            Case 2 : Return "In Progress"
            Case 3 : Return "Unreleased"
            Case 4 : Return "Released"
            Case 5 : Return "On Hold"
            Case Else : Return "Waiting"
        End Select
    End Function

    Public Shared Function GetCategoryText(ByVal yCategory As Byte) As String
        Select Case yCategory
            Case 0
                Return "Crash"
            Case 1
                Return "Gameplay"
            Case 2
                Return "Graphical"
            Case 3
                Return "Performance"
            Case 4
                Return "Other"
            Case 5
                Return "Login"
            Case 6
                Return "Suggestions"
            Case 7
                Return "Test Case"
            Case Else
                Return "Unknown"
        End Select
    End Function

	Public Function NewSaveObject() As Boolean
		Dim oPlayer As Player = GetEpicaPlayer(lCreatedBy)
		Dim sRaceName As String = ""
		If oPlayer Is Nothing = False Then
			sRaceName = BytesToString(oPlayer.RaceName)
		End If
		sRaceName = sRaceName.Replace(" ", "_")

		Dim dtBase As Date = CDate("January 1, 1970 00:00:00")
		Dim lNow As Int32 = CInt(Now.ToUniversalTime.Subtract(dtBase).TotalSeconds)

		Dim lTestCatID As Int32 = 100 + (CInt(yCategory) * 10) + CInt(ySubCat)
		Dim lNewPriority As Int32 = CInt(yPriority) + 1

		Dim lUserID As Int32 = -1
		Dim oData As Odbc.OdbcDataReader = GetWebsiteData("SELECT `userid` from `users` where `username` = '" & MakeDBStr(sRaceName) & "'")
		If oData.Read = True Then
			lUserID = CInt(oData(0))
		End If
		oData.Close()
		oData = Nothing

		Dim sSQL As String = "insert into `issues` (`gid`, `opened_by`, `opened`, `modified`, `summary`, `problem`, `status`, `category`, " & _
		  "`product`, `severity`) VALUES (4, " & _
		  lUserID & ", " & lNow.ToString & ", " & lNow.ToString & _
		  ", '" & MakeDBStr(Mid$(sProblemDesc.Replace(vbCrLf, " "), 1, 60).PadRight(63, " "c)) & "', '" & MakeDBStr(sProblemDesc) & _
		  "', 1, " & lTestCatID & ", 2, " & lNewPriority & ")"

		'DO INSERT HERE
		If ExecToWebsite(sSQL, False) = False Then
			LogEvent(LogEventType.CriticalError, "SaveBugFailed: " & sSQL)
		End If

		'Get our issue id
		sSQL = "SELECT MAX(Issueid) FROM `issues` WHERE gid = 4"
		oData = GetWebsiteData(sSQL)
		Dim lIssueID As Int32 = -1
		If oData.Read = True Then
			lIssueID = CInt(oData(0))
		End If
		oData.Close()
		oData = Nothing

		sSQL = "INSERT INTO `issue_groups` (`issueid`, `gid`, `opened`, `first_response`) VALUES (" & lIssueID & ", 4, " & lNow.ToString & ", 0)"
		If ExecToWebsite(sSQL, False) = False Then
			LogEvent(LogEventType.CriticalError, "SaveBugFailed: " & sSQL)
		End If

		If sStepsToProduce Is Nothing = False AndAlso sStepsToProduce <> "" Then
			'ok, now insert our event...
			sSQL = "INSERT INTO `events` (`issueid`, `status`, `userid`, `performed_on`, `action`) VALUES (" & lIssueID & ", 1, " & lUserID & _
			 ", " & lNow.ToString & ", '" & MakeDBStr(sStepsToProduce) & "')"

			'DO INSERT HERE
			If ExecToWebsite(sSQL, False) = False Then
				LogEvent(LogEventType.CriticalError, "SaveBugFailed: " & sSQL)
			End If
		End If

		Return True
	End Function

	Private Shared moBugList() As EpicaBug
	Private Shared mlBugListUB As Int32 = -1
	Private Shared mlLastBugListUpdate As Int32
	Private Const ml_BUG_LIST_UPDATE_DELAY As Int32 = 9000		'5 minutes
	Public Shared Function GetBugListResponse() As Byte()
		'k, first,	check if we need to update our bug list
		If glCurrentCycle - mlLastBugListUpdate > ml_BUG_LIST_UPDATE_DELAY Then
			mlLastBugListUpdate = glCurrentCycle

			'Now, do our update...
			Dim oData As Odbc.OdbcDataReader = GetWebsiteData("SELECT * FROM `issues` WHERE `status` IN (11, 23) and gid = 4 ORDER BY `opened`")
			Dim lNewUB As Int32 = -1
			Dim oNewBug() As EpicaBug = Nothing
			Dim lArrayUB As Int32 = -1
			While oData.Read
				Try
					lNewUB += 1
					If lNewUB > lArrayUB Then
						lArrayUB += 10
						ReDim Preserve oNewBug(lArrayUB)
					End If
					oNewBug(lNewUB) = New EpicaBug
					With oNewBug(lNewUB)

						If oData("problem") Is DBNull.Value = False Then .sProblemDesc = CStr(oData("problem")) Else .sProblemDesc = ""
						.yStatus = CByte(oData("Status"))
						Dim lTempCat As Int32 = 0
						If oData("Category") Is DBNull.Value = False Then lTempCat = CInt(oData("Category"))
						If lTempCat > 100 Then lTempCat -= 100
						.ySubCat = CByte(lTempCat Mod 10)
						lTempCat -= CInt(.ySubCat)
						.yCategory = CByte(lTempCat \ 10)
						Dim lSeverity As Int32 = 0
						If oData("severity") Is DBNull.Value = False Then lSeverity = CInt(oData("severity"))
						.yPriority = CByte(Math.Min(Math.Max(0, lSeverity - 1), 255))
						.sStepsToProduce = ""
					End With
				Catch ex As Exception
					LogEvent(LogEventType.CriticalError, "LoadBugList: " & ex.Message)
				End Try
			End While
			oData.Close()
			oData = Nothing

			mlBugListUB = lNewUB
			moBugList = oNewBug
		End If


		Dim yCache(0) As Byte
		Dim lPos As Int32
		Dim yTemp() As Byte
		lPos = 0
		For X As Int32 = 0 To mlBugListUB
			yTemp = goMsgSys.GetAddObjectMessage(moBugList(X), GlobalMessageCode.eBugList)
			ReDim Preserve yCache(yCache.Length + yTemp.Length + 2)
			System.BitConverter.GetBytes(CShort(yTemp.Length)).CopyTo(yCache, lPos)
			lPos += 2
			yTemp.CopyTo(yCache, lPos)
			lPos += yTemp.Length
		Next X
		Return yCache
	End Function
	Private Shared Function GetWebsiteData(ByVal sSQL As String) As Odbc.OdbcDataReader
		Try
			Dim oComm As Odbc.OdbcCommand = New Odbc.OdbcCommand(sSQL, GetWebConnection)
			Dim oReader As Odbc.OdbcDataReader = oComm.ExecuteReader()
			Return oReader
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "GetWebsiteData: " & ex.Message)
		End Try
		Return Nothing
	End Function
	Private Shared Function ExecToWebsite(ByVal sSQL As String, ByVal bZeroRecsOK As Boolean) As Boolean
		Dim bRes As Boolean = False
		Dim oComm As Odbc.OdbcCommand = Nothing
		Try
			oComm = New Odbc.OdbcCommand(sSQL, GetWebConnection)
			Dim lResult As Int32 = oComm.ExecuteNonQuery()
			If lResult = 0 AndAlso bZeroRecsOK = False Then Return False
			bRes = True
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "ExecToWebsite: " & ex.Message)
		Finally
			If oComm Is Nothing = False Then oComm.Dispose()
			oComm = Nothing
		End Try
		Return bRes
	End Function


End Class
