Public Enum eInfiltrationType As Byte
    eGeneralInfiltration = 0
    eMilitaryInfiltration = 1
    eFederalInfiltration = 2
    eColonialInfiltration = 3
    eProductionInfiltration = 4
    eResearchInfiltration = 5
    eTradeInfiltration = 6
    ePowerCenterInfiltration = 7
    eMiningInfiltration = 8
    eSolarSystemInfiltration = 9
    eCapitalShipInfiltration = 10
    ePlanetInfiltration = 11
    eCombatUnitInfiltration = 12
    eAgencyInfiltration = 13
    eCorporationInfiltration = 14

    eLastInfiltrationType = 15
End Enum
Public Enum AgentStatus As Int32
	NormalStatus = 0
	UsedAsCoverAgent = 1
	HasBeenCaptured = 2
	IsDead = 4
	NewRecruit = 8
	CounterAgent = 16
	ReturningHome = 32
	OnAMission = 64
	Infiltrating = 128
	IsInfiltrated = 256
	Dismissed = 512
End Enum
Public Enum eReportFreq As Byte
	OncePerHalfHour = 0
	OncePerHour = 1
	OncePerTwoHours = 2
	OncePerSixHours = 3
	OncePerTwelveHours = 4
	OncePerDay = 5
	OncePerTwoDays = 6
	OncePerWeek = 7
End Enum
Public Enum eSpilledData As Byte
	NoDataSpilled = 0
	AgentName = 1
	InfiltrationSpecifics = 2
	OwnerName = 4
	CurrentMission = 8
	MissionTarget = 16
	Accomplices = 32
End Enum

Public Class Agent
	Inherits Epica_GUID

	Public ServerIndex As Int32 = -1	'index in the global array
	Public AgentName(29) As Byte
    Public Infiltration As Byte
    Public Dagger As Byte
    Public Resourcefulness As Byte
    Public Luck As Byte
    Public Loyalty As Byte
    Public Suspicion As Byte

    Public InfiltrationLevel As Byte
	Public InfiltrationType As eInfiltrationType = 0
	Public yReportFreq As Byte
	Public lArrivalCycles As Int32 = -1
	Public lArrivalStart As Int32 = -1 

	Public lUpfrontCost As Int32
	Public lMaintCost As Int32

	Public lTargetID As Int32 = -1
	Public iTargetTypeID As Int16 = -1

	Public oOwner As Player
	Public oMission As PlayerMission = Nothing

	Public bMale As Boolean

	Public Skills() As Skill
	Public SkillProf() As Byte
	Public SkillUB As Int32 = -1

	Public dtRecruited As Date

	'Used for when the agent is captured
    Public lCapturedBy As Int32 = -1
    Public lCapturedOn As Int32 = -1
    Public lPrisonTestCycles As Int32 = -1
	Public yHealth As Byte = 100
	Public lInterrogatorID As Int32 = -1
	Public ySpilledData As eSpilledData = eSpilledData.NoDataSpilled
	Public yInterrogationState As Byte = 0

	'Mission Specifics....
	Public lAgentStatus As AgentStatus = AgentStatus.NormalStatus

	Private mlPreviousDeinfiltrationAttempt As Int32
	Private Const ml_DEINFILTRATION_DELAY As Int32 = 90		'once per 3 seconds

	'Private mySendString() As Byte

	Public Function GetAgentStatusMsg() As Byte()
		Dim yMsg(15) As Byte
		'Hardcoded for speed
		System.BitConverter.GetBytes(GlobalMessageCode.eGetAgentStatus).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(ObjectID).CopyTo(yMsg, 2)
		System.BitConverter.GetBytes(lAgentStatus).CopyTo(yMsg, 6)
		yMsg(10) = Suspicion
		yMsg(11) = InfiltrationLevel

		If lArrivalCycles <> -1 Then
			Dim lTemp As Int32 = lArrivalCycles - (glCurrentCycle - lArrivalStart)
			System.BitConverter.GetBytes(lTemp).CopyTo(yMsg, 12)
		Else : System.BitConverter.GetBytes(lArrivalCycles).CopyTo(yMsg, 12)
		End If
		Return yMsg
	End Function

	Public Function GetObjAsString() As Byte()
		'here we will return the entire object as a string
		'If mbStringReady = False Then
		Dim lPos As Int32 = 0

		'ReDim mySendString(81 + ((SkillUB + 1) * 5))  '0 to 46 = 47 bytes
		Dim mySendString(81) As Byte '+ ((SkillUB + 1) * 5)) As Byte

		GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
		AgentName.CopyTo(mySendString, lPos) : lPos += 30
		mySendString(lPos) = Infiltration : lPos += 1
		mySendString(lPos) = Dagger : lPos += 1
		mySendString(lPos) = Resourcefulness : lPos += 1
		mySendString(lPos) = Suspicion : lPos += 1
		mySendString(lPos) = InfiltrationLevel : lPos += 1
		mySendString(lPos) = InfiltrationType : lPos += 1

		Dim lTemp As Int32 = 0
		If (lAgentStatus And AgentStatus.NewRecruit) = 0 Then
			lTemp = CInt(Now.Subtract(dtRecruited).TotalSeconds)
		End If
		System.BitConverter.GetBytes(lTemp).CopyTo(mySendString, lPos) : lPos += 4

		System.BitConverter.GetBytes(lTargetID).CopyTo(mySendString, lPos) : lPos += 4
		System.BitConverter.GetBytes(iTargetTypeID).CopyTo(mySendString, lPos) : lPos += 2

		System.BitConverter.GetBytes(oOwner.ObjectID).CopyTo(mySendString, lPos) : lPos += 4

		If oMission Is Nothing Then
			System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, lPos)
		Else : System.BitConverter.GetBytes(oMission.PM_ID).CopyTo(mySendString, lPos)
		End If
		lPos += 4

		If bMale = True Then mySendString(lPos) = 1 Else mySendString(lPos) = 0
		lPos += 1

		System.BitConverter.GetBytes(lUpfrontCost).CopyTo(mySendString, lPos) : lPos += 4
		System.BitConverter.GetBytes(lMaintCost).CopyTo(mySendString, lPos) : lPos += 4
		mySendString(lPos) = yReportFreq : lPos += 1
		System.BitConverter.GetBytes(lAgentStatus).CopyTo(mySendString, lPos) : lPos += 4

		If lArrivalCycles <> -1 Then
			lTemp = lArrivalCycles - (glCurrentCycle - lArrivalStart)
			System.BitConverter.GetBytes(lTemp).CopyTo(mySendString, lPos) : lPos += 4
		Else : System.BitConverter.GetBytes(lArrivalCycles).CopyTo(mySendString, lPos) : lPos += 4
		End If


		'System.BitConverter.GetBytes(SkillUB + 1).CopyTo(mySendString, lPos) : lPos += 4

		'For X As Int32 = 0 To SkillUB
		'	System.BitConverter.GetBytes(Skills(X).ObjectID).CopyTo(mySendString, lPos) : lPos += 4
		'	mySendString(lPos) = SkillProf(X) : lPos += 1
		'Next X

		'mbStringReady = True
		'End If
		Return mySendString
	End Function

	Public Function GetCapturedAgentMsg() As Byte()
		Dim yMsg(58) As Byte
        Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eCaptureKillAgent).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4

		yMsg(lPos) = 0 : lPos += 1		'0 to indicate new captured agent...

		If (Me.ySpilledData And eSpilledData.AgentName) <> 0 Then
			Me.AgentName.CopyTo(yMsg, lPos) : lPos += 30
		Else : StringToBytes("Unknown").CopyTo(yMsg, lPos) : lPos += 30
		End If
		If (Me.ySpilledData And eSpilledData.OwnerName) <> 0 Then
			System.BitConverter.GetBytes(Me.oOwner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
		Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
		End If
		yMsg(lPos) = yHealth : lPos += 1
		If (Me.ySpilledData And eSpilledData.InfiltrationSpecifics) <> 0 Then
			yMsg(lPos) = Me.InfiltrationLevel : lPos += 1
			yMsg(lPos) = Me.InfiltrationType : lPos += 1
		Else
			yMsg(lPos) = 0 : lPos += 1
			yMsg(lPos) = 255 : lPos += 1
		End If

		yMsg(lPos) = yInterrogationState : lPos += 1
		System.BitConverter.GetBytes(lInterrogatorID).CopyTo(yMsg, lPos) : lPos += 4

		Dim lID As Int32 = -1
		If (Me.ySpilledData And eSpilledData.CurrentMission) <> 0 Then
			If Me.oMission Is Nothing = False Then
				lID = Me.oMission.oMission.ObjectID
			Else : lID = 0
			End If
		End If
		System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4

		lID = -1
		Dim iID As Int16 = -1
		If (Me.ySpilledData And eSpilledData.MissionTarget) <> 0 Then
			If Me.oMission Is Nothing = False Then
				lID = oMission.lTargetID
				iID = oMission.iTargetTypeID
			Else
				lID = 0
				iID = 0
			End If
		End If
		System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(iID).CopyTo(yMsg, lPos) : lPos += 2

		Return yMsg
	End Function

	Public Function GetNonOwnMsg() As Byte()
		Dim lPos As Int32 = 0

		Dim yResp(51 + ((SkillUB + 1) * 5)) As Byte	'+ ((SkillUB + 1) * 5)) As Byte

		GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
		AgentName.CopyTo(yResp, lPos) : lPos += 30
		yResp(lPos) = Infiltration : lPos += 1
		yResp(lPos) = Dagger : lPos += 1
		yResp(lPos) = Resourcefulness : lPos += 1
		yResp(lPos) = Suspicion : lPos += 1

		System.BitConverter.GetBytes(lUpfrontCost).CopyTo(yResp, lPos) : lPos += 4
		System.BitConverter.GetBytes(lMaintCost).CopyTo(yResp, lPos) : lPos += 4

		System.BitConverter.GetBytes(SkillUB + 1I).CopyTo(yResp, lPos) : lPos += 4
		For X As Int32 = 0 To SkillUB
			System.BitConverter.GetBytes(Skills(X).ObjectID).CopyTo(yResp, lPos) : lPos += 4
			yResp(lPos) = SkillProf(X) : lPos += 1
		Next X

		Return yResp
	End Function

	Public Sub AddSkill(ByVal oSkill As Skill, ByVal yProficiency As Byte)

        For X As Int32 = 0 To SkillUB
            If Skills(X).ObjectID = oSkill.ObjectID Then
                SkillProf(X) = yProficiency
                Return
            End If
        Next X

        SkillUB += 1
		ReDim Preserve Skills(SkillUB)
		ReDim Preserve SkillProf(SkillUB)
		Skills(SkillUB) = oSkill
		SkillProf(SkillUB) = yProficiency
	End Sub

	Public Function GetSaveObjectText() As String
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim X As Int32
		Dim oSB As New System.Text.StringBuilder

		If ObjectID < 1 Then
			If SaveObject() = False Then Return ""
		End If

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
		If lArrivalCycles > 0 Then
			lArrivalCycles -= (glCurrentCycle - lArrivalStart)
			If lArrivalCycles < 0 Then lArrivalCycles = 1
		End If
        If lPrisonTestCycles > 0 Then
            lPrisonTestCycles -= (glCurrentCycle - lCapturedOn)
            If lPrisonTestCycles < 0 Then lPrisonTestCycles = 1
        End If

		Try
			If ObjectID = -1 Then
				'INSERT
                sSQL = "INSERT INTO tblAgent (AgentName, Infiltration, Dagger, Resourcefulness, Luck, Loyalty, UpfrontCost, MaintCost, " & _
                  "Suspicion, InfiltrationLevel, InfiltrationType, TargetID, TargetTypeID, ArrivalCycles, OwnerID, MissionID, RecruitedOn, " & _
                  "CapturedBy, Health, InterrogatorID, ySpilledData, AgentStatus, IsMale, ReportFreq, CapturedOn, PrisonTestCycles) VALUES ('" & _
                  MakeDBStr(BytesToString(AgentName)) & "', " & Infiltration & ", " & Dagger & ", " & Resourcefulness & ", " & Luck & ", " & _
                  Loyalty & ", " & lUpfrontCost & ", " & lMaintCost & ", " & Suspicion & ", " & InfiltrationLevel & ", " & _
                  CByte(InfiltrationType) & ", " & lTargetID & ", " & iTargetTypeID & ", " & lArrivalCycles & ", "

				If oOwner Is Nothing = False Then
					sSQL = sSQL & oOwner.ObjectID & ", "
				Else : sSQL = sSQL & "-1, "
				End If
				If oMission Is Nothing = False Then
					sSQL = sSQL & oMission.PM_ID '& ")"
				Else : sSQL = sSQL & "-1" ')"
				End If
				sSQL &= ", " & GetDateAsNumber(dtRecruited) & ", " & lCapturedBy & ", " & yHealth & ", " & lInterrogatorID & ", " & _
				  CInt(ySpilledData) & ", " & CInt(lAgentStatus) & ", "
                If bMale = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", " & yReportFreq
                sSQL &= ", " & lCapturedOn & ", " & lPrisonTestCycles & ")"

			Else
				'UPDATE
                sSQL = "UPDATE tblAgent SET AgentName = '" & MakeDBStr(BytesToString(AgentName)) & "', Infiltration = " & Infiltration & _
                  ", Dagger = " & Dagger & ", Resourcefulness = " & Resourcefulness & ", Luck = " & Luck & _
                  ", Loyalty = " & Loyalty & ", Suspicion = " & Suspicion & ", InfiltrationLevel = " & InfiltrationLevel & _
                  ", InfiltrationType = " & CByte(InfiltrationType) & ", TargetID = " & lTargetID & ", TargetTypeID = " & iTargetTypeID & _
                  ", UpfrontCost = " & lUpfrontCost & ", MaintCost = " & lMaintCost & ", ArrivalCycles = " & lArrivalCycles & _
                  ", CapturedBy = " & lCapturedBy & ", Health = " & yHealth & ", InterrogatorID = " & lInterrogatorID & ", ySpilledData = " & _
                  CInt(ySpilledData) & ", AgentStatus = " & CInt(lAgentStatus) & ", ReportFreq = " & yReportFreq & ", IsMale = "
                If bMale = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", CapturedOn=" & lCapturedOn & ", PrisonTestCycles=" & lPrisonTestCycles

				If oOwner Is Nothing = False Then
					sSQL = sSQL & ", OwnerID = " & oOwner.ObjectID
				Else : sSQL = sSQL & ", OwnerID = -1"
				End If
				If oMission Is Nothing = False Then
					sSQL = sSQL & ", MissionID = " & oMission.PM_ID
				Else : sSQL = sSQL & ", MissionID = -1"
				End If
				sSQL = sSQL & ", RecruitedOn = " & GetDateAsNumber(dtRecruited) & " WHERE AgentID = " & ObjectID
			End If
			oSB.AppendLine(sSQL)

			'first, delete all entries from agentskill
			sSQL = "DELETE FROM tblAgentSkill WHERE AgentID = " & ObjectID
			oSB.AppendLine(sSQL)

			'Now, add the skills to the database
			For X = 0 To SkillUB
				sSQL = "INSERT INTO tblAgentSkill (AgentID, SkillID, SkillValue) VALUES (" & ObjectID & ", " & _
				  Skills(X).ObjectID & ", " & SkillProf(X) & ")"
				oSB.AppendLine(sSQL)
			Next X

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		End Try
		Return oSB.ToString
	End Function

	Public Function SaveObject() As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand
		Dim X As Int32

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
        Try
            If lArrivalCycles > 0 Then
                lArrivalCycles -= (glCurrentCycle - lArrivalStart)
                If lArrivalCycles < 0 Then lArrivalCycles = 1
            End If

            If lPrisonTestCycles > 0 Then
                lPrisonTestCycles -= (glCurrentCycle - lCapturedOn)
                If lPrisonTestCycles < 0 Then lPrisonTestCycles = 1
            End If
        Catch
        End Try

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblAgent (AgentName, Infiltration, Dagger, Resourcefulness, Luck, Loyalty, UpfrontCost, MaintCost, " & _
                  "Suspicion, InfiltrationLevel, InfiltrationType, TargetID, TargetTypeID, ArrivalCycles, OwnerID, MissionID, RecruitedOn, " & _
                  "CapturedBy, Health, InterrogatorID, ySpilledData, AgentStatus, IsMale, ReportFreq, CapturedOn, PrisonTestCycles) VALUES ('" & _
                  MakeDBStr(BytesToString(AgentName)) & "', " & Infiltration & ", " & Dagger & ", " & Resourcefulness & ", " & Luck & ", " & _
                  Loyalty & ", " & lUpfrontCost & ", " & lMaintCost & ", " & Suspicion & ", " & InfiltrationLevel & ", " & _
                  CByte(InfiltrationType) & ", " & lTargetID & ", " & iTargetTypeID & ", " & lArrivalCycles & ", "

                If oOwner Is Nothing = False Then
                    sSQL = sSQL & oOwner.ObjectID & ", "
                Else : sSQL = sSQL & "-1, "
                End If
                If oMission Is Nothing = False Then
                    sSQL = sSQL & oMission.PM_ID '& ")"
                Else : sSQL = sSQL & "-1" ')"
                End If
                sSQL &= ", " & GetDateAsNumber(dtRecruited) & ", " & lCapturedBy & ", " & yHealth & ", " & lInterrogatorID & ", " & _
                  CInt(ySpilledData) & ", " & CInt(lAgentStatus) & ", "
                If bMale = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", " & yReportFreq & ", " & lCapturedOn & ", " & lPrisonTestCycles & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblAgent SET AgentName = '" & MakeDBStr(BytesToString(AgentName)) & "', Infiltration = " & Infiltration & _
                  ", Dagger = " & Dagger & ", Resourcefulness = " & Resourcefulness & ", Luck = " & Luck & _
                  ", Loyalty = " & Loyalty & ", Suspicion = " & Suspicion & ", InfiltrationLevel = " & InfiltrationLevel & _
                  ", InfiltrationType = " & CByte(InfiltrationType) & ", TargetID = " & lTargetID & ", TargetTypeID = " & iTargetTypeID & _
                  ", UpfrontCost = " & lUpfrontCost & ", MaintCost = " & lMaintCost & ", ArrivalCycles = " & lArrivalCycles & _
                  ", CapturedBy = " & lCapturedBy & ", Health = " & yHealth & ", InterrogatorID = " & lInterrogatorID & ", ySpilledData = " & _
                  CInt(ySpilledData) & ", AgentStatus = " & CInt(lAgentStatus) & ", ReportFreq = " & yReportFreq & ", IsMale = "
                If bMale = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", CapturedOn=" & lCapturedOn & ", PrisonTestCycles=" & lPrisonTestCycles

                If oOwner Is Nothing = False Then
                    sSQL = sSQL & ", OwnerID = " & oOwner.ObjectID
                Else : sSQL = sSQL & ", OwnerID = -1"
                End If
                If oMission Is Nothing = False Then
                    sSQL = sSQL & ", MissionID = " & oMission.PM_ID
                Else : sSQL = sSQL & ", MissionID = -1"
                End If
                sSQL = sSQL & ", RecruitedOn = " & GetDateAsNumber(dtRecruited) & " WHERE AgentID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(AgentID) FROM tblAgent WHERE AgentName = '" & MakeDBStr(BytesToString(AgentName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            'first, delete all entries from agentskill
            sSQL = "DELETE FROM tblAgentSkill WHERE AgentID = " & ObjectID
            oComm = Nothing
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm = Nothing

            'Now, add the skills to the database
            For X = 0 To SkillUB
                sSQL = "INSERT INTO tblAgentSkill (AgentID, SkillID, SkillValue) VALUES (" & ObjectID & ", " & _
                  Skills(X).ObjectID & ", " & SkillProf(X) & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm = Nothing
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Public Function GetSkillValue(ByVal lSkillID As Int32, ByVal bExactMatch As Boolean, ByVal lAlternativeValue As Int32) As Int32
		For X As Int32 = 0 To SkillUB
			If Skills(X).ObjectID = lSkillID Then Return SkillProf(X)
		Next X

		'ok, if we are here, check bExactMatch
		If bExactMatch = False Then
			'ok, Naturally Talented may work
			Dim lTmp As Int32 = GetSkillValue(lSkillHardcodes.eNaturallyTalented, True, 0)
			If lAlternativeValue > lTmp Then Return lAlternativeValue Else Return lTmp
		End If
		Return lAlternativeValue
	End Function

    'Public Function AttemptInfiltration() As Byte
    '	If lTargetID < 1 Then Return 0

    '	If iTargetTypeID = ObjectType.ePlayer Then

    '		If lTargetID = Me.oOwner.ObjectID Then
    '			Me.InfiltrationLevel = 50
    '			Return 255
    '		End If

    '		Dim oTarget As Player = GetEpicaPlayer(lTargetID)

    '		Dim lCntrAgents As Int32 = 0
    '		Dim lCntrAgentInf As Int32 = 0
    '		'An Agent is infiltrated if the agent succeeds on the initial infiltration test. 
    '		'This test occurs on the cycle that the agent begins infiltrating. 
    '		'The Agent succeeds to infiltrate the target empire based on a 1D100 roll versus the to hit number of: 
    '		'   100 - Empire_Resist + (Agent_Infiltration - Agent_Suspicion) - (Sum_Of_Counter_Agent_Inf/Counter_Agents)
    '		With oTarget.oSecurity
    '			For X As Int32 = 0 To .lCounterAgentUB
    '				If .lCounterAgentIdx(X) <> -1 AndAlso .oCounterAgents(X).InfiltrationType = eInfiltrationType.eGeneralInfiltration Then
    '					lCntrAgents += 1
    '					lCntrAgentInf += .oCounterAgents(X).InfiltrationLevel
    '					lCntrAgentInf += .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)
    '				End If
    '			Next X

    '			Dim lVal As Int32 = 100
    '			'- .lInfiltrationResistance(eInfiltrationType.eGeneralInfiltration) 
    '			If .lInfiltrationResistance Is Nothing = False Then
    '				If .lInfiltrationResistance.GetUpperBound(0) >= eInfiltrationType.eGeneralInfiltration Then lVal -= .lInfiltrationResistance(eInfiltrationType.eGeneralInfiltration)
    '			End If
    '			'+ (Me.Infiltration - Me.Suspicion) - CInt(lCntrAgentInf / lCntrAgents)
    '               lVal += (CInt(Me.Infiltration) - CInt(Me.Suspicion))
    '			If lCntrAgents > 0 Then
    '				lVal -= CInt(lCntrAgentInf / lCntrAgents)
    '			End If

    '			Dim lTempValue As Int32 = Me.GetSkillValue(lSkillHardcodes.eDisguises, True, 0)
    '			If lTempValue <> 0 Then
    '				If CInt(Rnd() * 100) < lTempValue Then
    '					'when disguises succeeds, 20 modifier
    '					lVal += 20
    '				End If
    '			End If
    '			lTempValue = Me.GetSkillValue(lSkillHardcodes.eEtiquetteGeneral, True, 0)
    '			If lTempValue <> 0 Then
    '				If CInt(Rnd() * 100) < lTempValue Then
    '					lVal += 10
    '				End If
    '			End If
    '			lTempValue = Me.GetSkillValue(lSkillHardcodes.eLegitimateCover, True, 0)
    '			If lTempValue <> 0 Then
    '				If CInt(Rnd() * 100) < lTempValue Then
    '					lVal += 15
    '				End If
    '			End If
    '			lVal += Me.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)

    '			'Now, roll 1D100 and get under lVal
    '			Dim lRoll As Int32 = CInt(Rnd() * 100)
    '			If lRoll < lVal Then
    '				'If the agent infiltrates successfully, the agent's infiltration level is set to 50.
    '				Me.InfiltrationLevel = 50
    '				Return 255
    '			Else
    '				'If the agent does not infiltrate successfully, then they have to make a test versus the hidden Luck ability. 
    '				lRoll = CInt(Rnd() * 100)

    '				'If the Luck roll works, 
    '				If lRoll < Me.Luck Then
    '					'then the agent returns home. 
    '					lTargetID = -1
    '					iTargetTypeID = -1
    '					Return 1
    '				Else
    '					'If the Luck roll fails, the agent is captured by the enemy empire for interrogation. 
    '					oTarget.oSecurity.CaptureAgent(Me)
    '					Return 2
    '				End If
    '			End If
    '		End With
    '	Else
    '		'There are planets and systems, these automatically infiltrate
    '		Me.InfiltrationLevel = 50
    '		Return 255
    '	End If

    'End Function

    Public Function AttemptInfiltration() As Byte
        If lTargetID < 1 Then Return 0

        If iTargetTypeID = ObjectType.ePlayer Then

            If lTargetID = Me.oOwner.ObjectID Then
                Me.InfiltrationLevel = 50
                Return 255
            End If

            Dim oTarget As Player = GetEpicaPlayer(lTargetID)

            Dim lCntrAgents As Int32 = 0
            Dim lCntrAgentInf As Int32 = 0
            
            With oTarget.oSecurity
                Dim lVal As Int32 = (CInt(Me.Infiltration) - CInt(Me.Suspicion))
                lVal += .GetInfiltrationCounterAgentModifier(Me.InfiltrationType, Me.GetSkillValue(lSkillHardcodes.eDisguises, False, 0), Me.GetSkillValue(lSkillHardcodes.ePersuasive, True, 0), Me.GetSkillValue(lSkillHardcodes.eForger, True, 0))
                lVal += Me.GetSkillValue(lSkillHardcodes.eLegitimateCover, True, 0)
                lVal += Me.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)
                If .lInfiltrationResistance Is Nothing = False Then
                    If .lInfiltrationResistance.GetUpperBound(0) >= Me.InfiltrationType Then lVal -= .lInfiltrationResistance(Me.InfiltrationType)
                End If

                'And now, our etiquette bonus...
                Dim lEtiquette As Int32 = 0
                Select Case Me.InfiltrationType
                    Case eInfiltrationType.eColonialInfiltration
                        lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteColony, True, 0)
                    Case eInfiltrationType.eFederalInfiltration
                        lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteFederal, True, 0)
                    Case eInfiltrationType.eGeneralInfiltration
                        lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteGeneral, True, 0)
                    Case eInfiltrationType.eMilitaryInfiltration
                        lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteMilitary, True, 0)
                    Case eInfiltrationType.eProductionInfiltration, eInfiltrationType.eCorporationInfiltration, eInfiltrationType.eMiningInfiltration, eInfiltrationType.ePowerCenterInfiltration
                        lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteWorker, True, 0)
                End Select
                If lEtiquette > 0 AndAlso lVal < 100 Then
                    '= (MAX(0, lEtiquette- (RAND()*100) )/lEtiquette)*(100-lVal)
                    Dim fTemp As Single = Math.Max(0, lEtiquette - ((Rnd() * 100) / lEtiquette))
                    'fTemp *= (100 - lVal)
                    fTemp = (fTemp / 100.0F) * lVal
                    lVal = CInt(lVal + fTemp)
                End If

                'Now, roll 1D100 and get under lVal
                Dim lRoll As Int32 = CInt(Rnd() * 100)
                If lRoll < lVal Then
                    'If the agent infiltrates successfully, the agent's infiltration level is set to 50.
                    Me.InfiltrationLevel = 50
                    Return 255
                Else
                    'If the agent does not infiltrate successfully, then they have to make a test versus the hidden Luck ability. 
                    lRoll = CInt(Rnd() * 100)

                    'If the Luck roll works, 
                    If lRoll < Me.Luck Then
                        'then the agent returns home. 
                        lTargetID = -1
                        iTargetTypeID = -1
                        Return 1
                    Else
                        'If the Luck roll fails, the agent is captured by the enemy empire for interrogation. 
                        oTarget.oSecurity.CaptureAgent(Me)
                        Return 2
                    End If
                End If
            End With
        Else
            'There are planets and systems, these automatically infiltrate
            Me.InfiltrationLevel = 50
            Return 255
        End If

    End Function

    'Public Sub ReprocessInfiltrateTest()
    '	'100 - (Empire_Resist + Area_Resist) + (Agent_Infiltration - Agent_Suspicion) + (Agent_Infiltration_Level - 50) - 
    '	'       (Sum_Of_Specific_Counter_Agents_Inf / Specific_Counter_Agents)

    '	If lTargetID < 1 OrElse iTargetTypeID <> ObjectType.ePlayer Then Return

    '	Dim oTarget As Player = GetEpicaPlayer(lTargetID)

    '	If Me.InfiltrationLevel > 99 Then
    '		If Me.Suspicion = 0 AndAlso Me.InfiltrationLevel < 200 Then
    '			If CInt(Rnd() * 100) < Me.Resourcefulness Then
    '                   Dim lTemp As Int32 = CInt(Me.InfiltrationLevel) + 1
    '				If lTemp > 199 Then lTemp = 200
    '				Me.InfiltrationLevel = CByte(lTemp)
    '			End If
    '		End If
    '		Return
    '	End If

    '	Dim lVal As Int32 = CInt(Me.Infiltration) - CInt(Me.Suspicion)
    '	lVal += (CInt(Me.InfiltrationLevel) - 20)
    '	If oTarget.oSecurity.lInfiltrationResistance Is Nothing = False AndAlso oTarget.oSecurity.lInfiltrationResistance.GetUpperBound(0) >= Me.InfiltrationType Then
    '           lVal -= (oTarget.oSecurity.lInfiltrationResistance(Me.InfiltrationType))
    '	End If

    '	Dim lCntrAgents As Int32 = 0
    '	Dim lCntrAgentInf As Int32 = 0
    '	With oTarget.oSecurity
    '		For X As Int32 = 0 To .lCounterAgentUB
    '			If .lCounterAgentIdx(X) <> -1 AndAlso .oCounterAgents(X).InfiltrationType = Me.InfiltrationType Then '  eInfiltrationType.eGeneralInfiltration Then
    '				lCntrAgents += 1
    '				lCntrAgentInf += CInt(.oCounterAgents(X).InfiltrationLevel) + 10
    '				lCntrAgentInf += .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)
    '			End If
    '		Next X
    '	End With
    '	If lCntrAgents > 0 Then lVal -= (lCntrAgentInf \ lCntrAgents)
    '	lVal = 100 - lVal

    '	Dim lTempValue As Int32 = Me.GetSkillValue(lSkillHardcodes.eDisguises, True, 0)
    '	If lTempValue <> 0 Then
    '		If CInt(Rnd() * 100) < lTempValue Then
    '			'when disguises succeeds, 20 modifier
    '			lVal += 20
    '		End If
    '	End If
    '	lTempValue = 0
    '	Select Case Me.InfiltrationType
    '		Case eInfiltrationType.eColonialInfiltration
    '			lTempValue = Me.GetSkillValue(lSkillHardcodes.eLaw, True, 0)
    '			If lTempValue <> 0 Then If CInt(Rnd() * 100) < lTempValue Then lVal += 5
    '			lTempValue = 0
    '			lTempValue = Me.GetSkillValue(lSkillHardcodes.eEtiquetteColony, True, 0)
    '		Case eInfiltrationType.eFederalInfiltration
    '			lTempValue = Me.GetSkillValue(lSkillHardcodes.eEtiquetteFederal, True, 0)
    '		Case eInfiltrationType.eMilitaryInfiltration
    '			lTempValue = Me.GetSkillValue(lSkillHardcodes.eEtiquetteMilitary, True, 0)
    '		Case eInfiltrationType.eProductionInfiltration, eInfiltrationType.eResearchInfiltration
    '			lTempValue = Me.GetSkillValue(lSkillHardcodes.eEtiquetteWorker, True, 0)
    '	End Select
    '	If lTempValue <> 0 Then If CInt(Rnd() * 100) < lTempValue Then lVal += 10
    '	lVal += Me.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)


    '	'lVal is the target number
    '	If CInt(Rnd() * 100) < lVal Then
    '		'If the agent succeeds, the agent's Infiltration level is bumped up 10 points.
    '           Dim lTemp As Int32 = CInt(Me.InfiltrationLevel) + 10
    '		If lTemp > 255 Then lTemp = 255
    '		Me.InfiltrationLevel = CByte(lTemp)
    '	Else
    '		If Me.lTargetID = Me.oOwner.ObjectID AndAlso Me.iTargetTypeID = ObjectType.ePlayer Then
    '			'do nothing
    '		Else
    '			'If the agent fails, then the agent rolls against the hidden Luck attribute. 
    '			If CInt(Rnd() * 100) < Me.Luck Then
    '				'If the agent is successful against the luck attribute, then the agent receives 5 Suspicion Points. 
    '				Dim lTemp As Int32 = CInt(Me.Suspicion) + 5
    '				If lTemp > 255 Then lTemp = 255
    '				Me.Suspicion = CByte(lTemp)
    '			Else
    '				'Otherwise, the agent receives 15 Suspicion Points. 
    '				Dim lTemp As Int32 = CInt(Me.Suspicion) + 15
    '				If lTemp > 255 Then lTemp = 255
    '				Me.Suspicion = CByte(lTemp)
    '			End If
    '			goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.ReduceSuspicion, Me, Nothing, glCurrentCycle + Agent.ml_DEINFILTRATE_TIME)
    '			'If Me.Suspicion > 200 Then
    '			'	oTarget.oSecurity.CaptureAgent(Me)
    '			'End If
    '		End If
    '	End If

    'End Sub

    Public Sub ReprocessInfiltrateTest()
        '100 - (Empire_Resist + Area_Resist) + (Agent_Infiltration - Agent_Suspicion) + (Agent_Infiltration_Level - 50) - 
        '       (Sum_Of_Specific_Counter_Agents_Inf / Specific_Counter_Agents)

        If lTargetID < 1 OrElse iTargetTypeID <> ObjectType.ePlayer Then Return

        Dim oTarget As Player = GetEpicaPlayer(lTargetID)

        If Me.InfiltrationLevel > 99 Then
            If Me.Suspicion = 0 AndAlso Me.InfiltrationLevel < 200 Then
                If CInt(Rnd() * 100) < Me.Resourcefulness Then
                    Dim lTemp As Int32 = CInt(Me.InfiltrationLevel) + 1
                    If lTemp > 199 Then lTemp = 200
                    Me.InfiltrationLevel = CByte(lTemp)
                End If
            End If
            Return
        End If

        Dim lVal As Int32 = CInt(Me.Resourcefulness) + (CInt(Me.Infiltration) - CInt(Me.Suspicion))
        lVal += CInt(Me.InfiltrationLevel) - 20

        With oTarget.oSecurity
            If .lInfiltrationResistance Is Nothing = False Then
                If .lInfiltrationResistance.GetUpperBound(0) >= Me.InfiltrationType Then lVal -= .lInfiltrationResistance(Me.InfiltrationType)
            End If

            lVal += .GetInfiltrationCounterAgentModifier(Me.InfiltrationType, Me.GetSkillValue(lSkillHardcodes.eDisguises, False, 0), Me.GetSkillValue(lSkillHardcodes.ePersuasive, True, 0), Me.GetSkillValue(lSkillHardcodes.eForger, True, 0))
        End With
        Dim lEtiquette As Int32 = 0
        Select Case Me.InfiltrationType
            Case eInfiltrationType.eColonialInfiltration
                lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteColony, True, 0)
            Case eInfiltrationType.eFederalInfiltration
                lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteFederal, True, 0)
            Case eInfiltrationType.eGeneralInfiltration
                lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteGeneral, True, 0)
            Case eInfiltrationType.eMilitaryInfiltration
                lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteMilitary, True, 0)
            Case eInfiltrationType.eProductionInfiltration, eInfiltrationType.eCorporationInfiltration, eInfiltrationType.eMiningInfiltration, eInfiltrationType.ePowerCenterInfiltration
                lEtiquette = Me.GetSkillValue(lSkillHardcodes.eEtiquetteWorker, True, 0)
        End Select
        If Rnd() * 100 < lEtiquette Then
            lVal += 10
        End If
        lVal += Me.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)

        'lVal is the target number
        If CInt(Rnd() * 100) < lVal Then
            'If the agent succeeds, the agent's Infiltration level is bumped up 10 points.
            Dim lTemp As Int32 = CInt(Me.InfiltrationLevel) + 10
            If lTemp > 255 Then lTemp = 255
            Me.InfiltrationLevel = CByte(lTemp)
        Else
            If Me.lTargetID = Me.oOwner.ObjectID AndAlso Me.iTargetTypeID = ObjectType.ePlayer Then
                'do nothing
            Else
                'If the agent fails, then the agent rolls against the hidden Luck attribute. 
                If CInt(Rnd() * 100) < Me.Luck Then
                    'If the agent is successful against the luck attribute, then the agent receives 5 Suspicion Points. 
                    Dim lTemp As Int32 = CInt(Me.Suspicion) + 5
                    If lTemp > 255 Then lTemp = 255
                    Me.Suspicion = CByte(lTemp)
                Else
                    'Otherwise, the agent receives 15 Suspicion Points. 
                    Dim lTemp As Int32 = CInt(Me.Suspicion) + 15
                    If lTemp > 255 Then lTemp = 255
                    Me.Suspicion = CByte(lTemp)
                End If
                goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.ReduceSuspicion, Me, Nothing, glCurrentCycle + Agent.ml_DEINFILTRATE_TIME)
                'If Me.Suspicion > 200 Then
                '	oTarget.oSecurity.CaptureAgent(Me)
                'End If
            End If
        End If

    End Sub

	Public Sub ReduceSuspicion()
		'If oTarget Is Nothing = False Then
		Dim lTemp As Int32 = Me.Suspicion
		If Me.InfiltrationLevel >= 100 Then
			lTemp = CInt(Me.Suspicion) - 2
		Else
			If CInt(Rnd() * 100) < Me.Resourcefulness Then lTemp = CInt(Me.Suspicion) - 1
		End If
		If lTemp < 0 Then lTemp = 0
		Me.Suspicion = CByte(lTemp)
		'End If
    End Sub

    Public Function PrisonEscapeTest() As Boolean
        'Do an escape test.  If failed, hit health.  If pass, they escape.

        Dim yRoll As Byte = CByte(Rnd() * 100)
        Dim oCaptor As Player = GetEpicaPlayer(Me.lCapturedBy)
        If oCaptor Is Nothing = True Then Exit Function

        Dim sCaptorStr As String = vbCrLf & vbCrLf & "They were being jailed by " & oCaptor.sPlayerNameProper & "."
        Dim oPC As PlayerComm = Nothing
        Dim sPlayerMsg As String = ""


        Dim yToHit As Int32 = 5 'base 5%
        Dim lTmp As Int32

        'Add up to 5% for Escape Artist
        lTmp = Me.GetSkillValue(lSkillHardcodes.eEscapeArtist, False, 0) \ 20
        yToHit = yToHit + lTmp

        'Add up to 5% for Dagger
        lTmp = Me.Dagger \ 20
        yToHit = yToHit + lTmp

        If yRoll < yToHit Then
            'Escape
            Me.SetAsFugitive(oCaptor, True)
            Dim yMsg(6) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eCaptureKillAgent).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
            yMsg(6) = 3
            oCaptor.SendPlayerMessage(yMsg, False, AliasingRights.eViewAgents)
            Return True
        Else
            Me.yHealth -= CByte(Math.Min(5, Me.yHealth))
            If Me.yHealth < 1 Then
                Me.KillMe()
                If (Me.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                    'Alert Agent Owner
                    sPlayerMsg = ""
                    sPlayerMsg = "Your agent, " & BytesToString(Me.AgentName) & ", was killed while attempting to escape captivity."
                    sPlayerMsg &= sCaptorStr
                    oPC = Me.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, sPlayerMsg, "Agent Killed", Me.oOwner.ObjectID, GetDateAsNumber(Now), False, Me.oOwner.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then
                        'Update Client Email List
                        Me.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                    'Update Client Agent List
                    Me.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                End If

                If (oCaptor.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                    'Alert Captor
                    sPlayerMsg = ""
                    If (Me.ySpilledData Or eSpilledData.AgentName) <> 0 Then
                        sPlayerMsg = "An agent, " & BytesToString(Me.AgentName) & ", was killed while attempting to escape captivity."
                    Else
                        sPlayerMsg = "An agent was killed while attempting to escape captivity."
                    End If
                    oPC = oCaptor.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, sPlayerMsg, "Captured Agent Killed", oCaptor.ObjectID, GetDateAsNumber(Now), False, oCaptor.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then
                        'Update Client Email List
                        oCaptor.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                    'Update Client Agent List
                    oCaptor.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                End If

            End If
            Return False
        End If
        Me.SaveObject()
    End Function

    Public Sub SetAsFugitive(Optional ByRef oCaptor As EpicaPrimary.Player = Nothing, Optional ByVal AlertCaptor As Boolean = True)
        Dim oPlayer As Player = GetEpicaPlayer(Me.lCapturedBy)
        If oPlayer Is Nothing = False Then
            With oPlayer.oSecurity
                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To .lCapturedAgentUB
                    If .lCapturedAgentIdx(X) = Me.ObjectID Then
                        lIdx = X
                        Exit For
                    End If
                Next
                If lIdx <> -1 Then
                    .lCapturedAgentIdx(lIdx) = -1
                    .oCapturedAgents(lIdx) = Nothing
                End If
            End With
        End If
        Me.lCapturedBy = -1
        Me.lCapturedOn = -1
        Me.lPrisonTestCycles = 0
        If (Me.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then Me.lAgentStatus = Me.lAgentStatus Xor AgentStatus.HasBeenCaptured
        If (Me.lAgentStatus And AgentStatus.IsInfiltrated) = 0 Then Me.lAgentStatus = Me.lAgentStatus Or AgentStatus.IsInfiltrated

        Me.lInterrogatorID = -1
        Me.yInterrogationState = 0
        Me.yHealth = 100        'Nero - are we sure we should raise health? since it is infil, it could cause escape > heal > cap > intero loop, maybe not heal until deinfil
        Me.Suspicion = 200      'perhaps add +heal to the reducesuspicion sub if health not 100, natural heal
        Me.InfiltrationLevel = 0
        Me.InfiltrationType = eInfiltrationType.eGeneralInfiltration

        goAgentEngine.CancelAllAgentEvents(Me.ObjectID)

        'Dim sPlayerMsg As String
        Dim sPlayerMsg As String = ""
        Dim oPC As PlayerComm = Nothing
        Dim sCaptorStr As String = vbCrLf & vbCrLf & "They were being jailed by " & oCaptor.sPlayerNameProper & "."
        sPlayerMsg = "Your agent, " & BytesToString(Me.AgentName) & ", has escaped captivity."
        sPlayerMsg &= sCaptorStr
        If (Me.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
            oPC = Me.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sPlayerMsg, "Agent Escaped", Me.oOwner.ObjectID, GetDateAsNumber(Now), False, Me.oOwner.sPlayerNameProper, Nothing)
            If oPC Is Nothing = False Then
                'Update Client Email List
                Me.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
        End If
        oPC = Nothing
        'Update Client Agent List
        Me.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)

        'Alert Captor
        If AlertCaptor = True Then
            If (Me.ySpilledData Or eSpilledData.AgentName) <> 0 Then
                sPlayerMsg = "An agent, " & BytesToString(Me.AgentName) & ", has escaped captivity. Somehow the fugitive broke through security and moved outside of the confines of the compound."
            Else
                sPlayerMsg = "An agent has escaped captivity. Somehow the fugitive broke through security and moved outside of the confines of the compound.  We never got the agent's name."
            End If
            sPlayerMsg &= vbCrLf & vbCrLf & "Authorities have been alerted and are in pursuit."
            If (oCaptor.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                oPC = oCaptor.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sPlayerMsg, "Captured Agent Escaped", oCaptor.ObjectID, GetDateAsNumber(Now), False, oCaptor.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then
                    'Update Client Email List
                    oCaptor.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If
            End If
            'Update Client Agent List
            oCaptor.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
        End If
        Me.SaveObject()
    End Sub


	''' <summary>
	''' Returns the result of Me.DaggerRoll - Other.DaggerRoll. A Negative value indicates Other wins. Positive indicates I win. This will NEVER return 0.
	''' Me is assumed to be the agent who will be negatively impacted by the dagger test the most (capturee, assassinated, etc...)
	''' </summary>
	''' <param name="oOtherAgent"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function HandleDaggerTest(ByRef oOtherAgent As Agent, ByVal yType As eDaggerTestType) As Int32
        Dim lMyVal As Int32 = CInt(Me.Dagger) + CInt(Rnd() * 100)
        Dim lOtherVal As Int32 = CInt(oOtherAgent.Dagger) + CInt(Rnd() * 100)

		Dim lTmp As Int32
		If yType = eDaggerTestType.AvoidCapture Then
			lTmp = Me.GetSkillValue(lSkillHardcodes.eEscapeArtist, False, 0)
			If lTmp <> 0 Then lMyVal += CInt(Rnd() * lTmp)
		End If
        lTmp = Me.GetSkillValue(lSkillHardcodes.eWeaponSpecialist, False, 0)
		If lTmp <> 0 Then If CInt(Rnd() * 100) < lTmp Then lMyVal += 15
        lTmp = oOtherAgent.GetSkillValue(lSkillHardcodes.eWeaponSpecialist, False, 0)
		If lTmp <> 0 Then If CInt(Rnd() * 100) < lTmp Then lOtherVal += 15
		lMyVal += Me.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)
		lOtherVal += oOtherAgent.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)

		Dim lResult As Int32 = lMyVal - lOtherVal
		If lResult = 0 Then
			'ok, the higher luck wins
			If Me.Luck > oOtherAgent.Luck Then
				lResult = 1
			ElseIf Me.Luck < oOtherAgent.Luck Then
				lResult = -1
			Else
				'k, the highest dagger wins
				If Me.Dagger > oOtherAgent.Dagger Then
					lResult = 1
				ElseIf Me.Dagger < oOtherAgent.Dagger Then
					lResult = -1
				Else
					'rnd
					If CInt(Rnd() * 100) < 50 Then
						lResult = 1
					Else : lResult = -1
					End If
				End If
			End If
		End If
		Return lResult
    End Function

    Public Enum eyCoverAgentVsCounterResult As Byte
        CoverAgentWins = 1
        CounterAgentWins = 2
        CoverAgentDies = 4
        CounterAgentDies = 8
        CoverAgentCaptured = 16
        AlarmThrown = 32
    End Enum
    Public Function HandleCoverAgentVsCounter(ByRef oCounter As Agent, ByVal lMissionMod As Int32) As eyCoverAgentVsCounterResult
        'three tests, counter vs. cover
        'DAGGER TEST: dagger + 1D100 + spygames vs. dagger + 1D100 + spygames
        'WEAPONS TEST: MAX(Dagger / 2, WeaponSkill) + 1D100 + SpyGames Vs. MAX(Dagger / 2, WeaponSkill) + 1D100 + SpyGames
        'EVASION TEST: Tracking + Perceptive + 1D100 + SpyGames + IF(Either has Disguises, Disguises, 0)
        '                   vs.
        '               Escape Artist + SpyGames + 1D100 + Disguises

        'we determine the successes of each test...
        Dim lCounterSuccess As Int32 = 0
        Dim lCoverSuccess As Int32 = 0
        Dim lCounterDamage As Int32 = 0
        Dim lCoverDamage As Int32 = 0

        Dim lCounterSpyGames As Int32 = oCounter.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0) - lMissionMod
        Dim lCoverSpyGames As Int32 = GetSkillValue(lSkillHardcodes.eSpyGames, True, 0) + GetSkillValue(lSkillHardcodes.eSecurity, True, 0)

        'First the dagger test
        Dim lCounterValue As Int32 = CInt(oCounter.Dagger) + lCounterSpyGames + CInt(Rnd() * 100)
        Dim lCoverValue As Int32 = CInt(Dagger) + lCoverSpyGames + CInt(Rnd() * 100)

        'A victor gets a point, no victor... no point
        If lCounterValue > lCoverValue Then
            lCounterSuccess += 1
            lCoverDamage += 1
        ElseIf lCoverValue > lCounterValue Then
            lCoverSuccess += 1
            lCounterDamage += 1
        Else
            lCounterDamage += 1
            lCoverDamage += 1
        End If

        'Next is the weapon test
        lCounterValue = CInt(Math.Max(CInt(oCounter.Dagger) / 2, oCounter.GetSkillValue(lSkillHardcodes.eWeaponSpecialist, True, 0)))
        lCounterValue += CInt(Rnd() * 100) + lCounterSpyGames
        lCoverValue = CInt(Math.Max(CInt(Dagger) / 2, GetSkillValue(lSkillHardcodes.eWeaponSpecialist, True, 0)))
        lCoverValue += CInt(Rnd() * 100) + lCoverSpyGames

        'A victor gets a point, no victor... no point
        If lCounterValue > lCoverValue Then
            lCounterSuccess += 1
            lCoverDamage += 2
        ElseIf lCoverValue > lCounterValue Then
            lCoverSuccess += 1
            lCounterDamage += 2
        Else
            lCoverDamage += 1
            lCounterDamage += 1
        End If

        'Finally, the evasion test
        Dim lCoverDisguises As Int32 = GetSkillValue(lSkillHardcodes.eDisguises, True, 0)
        lCounterValue = oCounter.GetSkillValue(lSkillHardcodes.eTracking, True, 0) + oCounter.GetSkillValue(lSkillHardcodes.ePerceptive, True, 0)
        lCounterValue += CInt(Rnd() * 100) + lCounterSpyGames
        If lCoverDisguises > 0 Then lCounterValue += oCounter.GetSkillValue(lSkillHardcodes.eDisguises, True, 0)
        lCoverValue = GetSkillValue(lSkillHardcodes.eEscapeArtist, True, 0) + lCoverSpyGames + CInt(Rnd() * 100) + lCoverDisguises

        'A victor gets a point, no victor... no point
        Dim bEscaped As Boolean = False
        If lCounterValue > lCoverValue Then
            lCounterSuccess += 1
        ElseIf lCoverValue > lCounterValue Then
            lCoverSuccess += 1
            lCoverDamage -= 1
            bEscaped = True
        End If

        'Ok, we'll determine the result now
        Dim yResult As eyCoverAgentVsCounterResult = 0

        If lCoverSuccess = lCounterSuccess Then
            lCoverDamage += 1
            lCounterDamage += 1
            If oCounter.Luck > Me.Luck Then
                lCounterSuccess += 1
            ElseIf oCounter.Luck < Me.Luck Then
                lCoverSuccess += 1
            ElseIf Me.Dagger > oCounter.Dagger Then
                lCoverSuccess += 1
            ElseIf Me.Dagger < oCounter.Dagger Then
                lCounterSuccess += 1
            ElseIf Rnd() * 100 < 50 Then
                lCoverSuccess += 1
            Else
                lCounterSuccess += 1
            End If
        End If

        If lCoverSuccess > lCounterSuccess Then
            yResult = yResult Or eyCoverAgentVsCounterResult.CoverAgentWins
        ElseIf lCounterSuccess > lCoverSuccess Then
            yResult = yResult Or eyCoverAgentVsCounterResult.CounterAgentWins
            If bEscaped = False Then
                yResult = yResult Or eyCoverAgentVsCounterResult.CoverAgentCaptured
            Else : yResult = yResult Or eyCoverAgentVsCounterResult.AlarmThrown
            End If
        End If

        If lCounterDamage > 0 Then
            Dim lTmpHealth As Int32 = oCounter.yHealth
            lTmpHealth -= CInt(Rnd() * (lCounterDamage * 10))
            If lTmpHealth > 100 Then lTmpHealth = 100
            If lTmpHealth < 0 Then lTmpHealth = 0

            'Ok, did our counter agent die?
            oCounter.yHealth = CByte(lTmpHealth)
            If oCounter.yHealth < 1 Then
                'yes...
                yResult = yResult Or eyCoverAgentVsCounterResult.CounterAgentDies
            End If
        End If
        If lCoverDamage > 0 Then
            Dim lTmpHealth As Int32 = yHealth
            lTmpHealth -= CInt(Rnd() * (lCoverDamage * 10))
            If lTmpHealth > 100 Then lTmpHealth = 100
            If lTmpHealth < 0 Then lTmpHealth = 0

            'Ok, did our counter agent die?
            yHealth = CByte(lTmpHealth)
            If yHealth < 1 Then
                'yes...
                yResult = yResult Or eyCoverAgentVsCounterResult.CoverAgentDies
            End If
        End If


        Return yResult
    End Function

	Public Sub KillMe()
		'Do the act of killing this agent here, notify the owner that the agent is dead
		goAgentEngine.CancelAllAgentEvents(Me.ObjectID)
		RemoveMeFromMissions()
		lAgentStatus = AgentStatus.IsDead
        oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
    End Sub

    'Public Function RunAgentCaptureTest(ByRef oCapturee As Agent) As Boolean
    '       Dim lToHit As Int32 = (100 - (CInt(oCapturee.Resourcefulness) + (((CInt(oCapturee.InfiltrationLevel) \ 2) + (CInt(oCapturee.Luck) \ 2)) \ 2)))
    '       lToHit += ((CInt(Dagger) + CInt(Resourcefulness)) \ 2)

    '	Return CInt(Rnd() * 100) < lToHit
    'End Function

	Public Const ml_DEINFILTRATE_TIME As Int32 = 9000	   '5 minutes
	Public Function AttemptDeinfiltration() As Int32

		If (Me.lAgentStatus And AgentStatus.IsDead) <> 0 Then Return -2 'agent is dead
		If (Me.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then Return -3 'agent is captured
		If (Me.lAgentStatus And AgentStatus.ReturningHome) <> 0 Then Return -4 'agent is already returning home
		If (Me.lAgentStatus And AgentStatus.OnAMission) <> 0 Then Return -5 'agent is on a mission and cannot deinfiltrate

		If glCurrentCycle - mlPreviousDeinfiltrationAttempt > ml_DEINFILTRATION_DELAY Then
			Dim lChance As Int32 = (CInt(InfiltrationLevel) - CInt(Suspicion)) + (Resourcefulness \ 3)
			If iTargetTypeID <> ObjectType.ePlayer OrElse lTargetID = oOwner.ObjectID Then lChance = 100

			If lChance < 5 Then lChance = 5
			If CInt(Rnd() * 100) < lChance Then
				'Deinfiltration succeeded! 
				Me.ReturnHome()

				'return 0 to indicate success
				Return 0
			Else
                Dim lTemp As Int32 = CInt(Me.Suspicion) + CInt(Rnd() * 5) + 5
				If lTemp > 255 Then lTemp = 255
				Me.Suspicion = CByte(lTemp)
				goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.ReduceSuspicion, Me, Nothing, glCurrentCycle + Agent.ml_DEINFILTRATE_TIME)
				Return Me.Suspicion		'the new suspicion value
			End If
		Else
			Return -1	'indicates delaying
		End If
	End Function

    Private Shared mlAgentType As Int32 = 0
	Public Shared Function GenerateNewAgent(ByRef poOwner As Player) As Agent
		Dim oAgent As New Agent()
 
		Randomize()

		With oAgent
			.bMale = (Rnd() * 100) > 50
			.AgentName = StringToBytes(GetFullGeneratedName(.bMale))
			.oOwner = poOwner

			Dim lMaxVal As Int32 = 100

            If gb_IS_TEST_SERVER = True Then

                Dim yScore As Byte
                Dim y20Score As Byte
                mlAgentType += 1
                Select Case mlAgentType
                    Case 1
                        yScore = 66
                        y20Score = 10
                    Case 2
                        yScore = 98
                        y20Score = 20
                    Case Else
                        yScore = 33
                        mlAgentType = 0
                        y20Score = 0
                End Select

                .Infiltration = yScore
                .Dagger = yScore
                .Resourcefulness = yScore
                .Loyalty = yScore
                .Luck = yScore
                .InfiltrationLevel = 0
                .InfiltrationType = 0
                .ObjTypeID = ObjectType.eAgent
                .oMission = Nothing
                .oOwner = poOwner
                .lTargetID = -1
                .iTargetTypeID = -1
                .Suspicion = 0
                .lAgentStatus = AgentStatus.NewRecruit

                For X As Int32 = 0 To glSkillUB
                    If glSkillIdx(X) > -1 Then
                        If goSkill(X).MaxVal > 90 Then
                            .AddSkill(goSkill(X), yScore)
                        Else
                            .AddSkill(goSkill(X), y20Score)
                        End If
                    End If
                Next X

            Else
                .Infiltration = CByte(Int(Rnd() * lMaxVal) + 1)

                If .Infiltration > 80 Then lMaxVal = 80
                .Dagger = CByte(Int(Rnd() * lMaxVal) + 1)

                If CInt(.Infiltration) + CInt(.Dagger) > 120 Then lMaxVal = 70
                .Resourcefulness = CByte(Int(Rnd() * lMaxVal) + 1)

                .Loyalty = CByte(Int(Rnd() * 80) + 20)
                .Luck = CByte(Int(Rnd() * 100) + 1)

                .InfiltrationLevel = 0
                .InfiltrationType = 0
                .ObjTypeID = ObjectType.eAgent
                .oMission = Nothing
                .oOwner = poOwner
                .lTargetID = -1
                .iTargetTypeID = -1
                .Suspicion = 0
                .lAgentStatus = AgentStatus.NewRecruit

                'Ok, find the first skill number...
                Dim lSkillIdx As Int32 = CInt(Int(Rnd() * (glSkillUB + 1)))
                Dim lProf As Int32
                Dim lTemp As Int32
                Dim bDone As Boolean = False
                Dim bFound As Boolean = False

                lProf = CInt(Int(Rnd() * (goSkill(lSkillIdx).MaxVal - goSkill(lSkillIdx).MinVal)) + goSkill(lSkillIdx).MinVal)
                .AddSkill(goSkill(lSkillIdx), CByte(lProf))

                While bDone = False
                    For X As Int32 = 0 To goSkill(lSkillIdx).lRelatedSkillUB
                        lTemp = CInt(Int(Rnd() * 100) + 1)
                        If lTemp < goSkill(lSkillIdx).RelatedSkills(X).lToHitNumber Then
                            'Check if the agent already has the skill
                            lTemp = goSkill(lSkillIdx).RelatedSkills(X).oToSkill.ObjectID
                            bFound = False
                            For Y As Int32 = 0 To .SkillUB
                                If .Skills(Y).ObjectID = lTemp Then
                                    bFound = True
                                End If
                            Next Y
                            If bFound = False Then
                                lProf = CInt(Int(Rnd() * (goSkill(lSkillIdx).RelatedSkills(X).oToSkill.MaxVal - goSkill(lSkillIdx).RelatedSkills(X).oToSkill.MinVal)) + goSkill(lSkillIdx).RelatedSkills(X).oToSkill.MinVal)
                                .AddSkill(goSkill(lSkillIdx).RelatedSkills(X).oToSkill, CByte(lProf))
                            End If
                        End If
                    Next X

                    lMaxVal = 100 - ((.SkillUB + 1) * 7)
                    If CInt(Int(Rnd() * 100) + 1) < lMaxVal Then
                        bFound = True

                        While bFound = True
                            lSkillIdx = CInt(Int(Rnd() * (glSkillUB + 1)))
                            lTemp = goSkill(lSkillIdx).ObjectID
                            bFound = False
                            For Y As Int32 = 0 To .SkillUB
                                If .Skills(Y).ObjectID = lTemp Then
                                    bFound = True
                                End If
                            Next Y
                            If bFound = False Then
                                lProf = CInt(Int(Rnd() * (goSkill(lSkillIdx).MaxVal - goSkill(lSkillIdx).MinVal)) + goSkill(lSkillIdx).MinVal)
                                .AddSkill(goSkill(lSkillIdx), CByte(lProf))
                            End If
                        End While
                    Else : bDone = True
                    End If
                End While
            End If

			'Ok, let's determine our costs now
			Dim lAttrScore As Int32 = CInt(.Dagger) + CInt(.Infiltration) + CInt(.Resourcefulness) + CInt(.Luck)
			Dim lSkillScore As Int32 = 0
			Dim lAbilityScore As Int32 = 0
			For X As Int32 = 0 To .SkillUB
				If .Skills(X).SkillType = 0 Then
					lSkillScore += .SkillProf(X)
				Else : lAbilityScore += .SkillProf(X)
				End If
			Next X
			lAbilityScore *= 5

			Dim lOverallScore As Int32 = lAttrScore + lSkillScore + lAbilityScore
			'now, upfront is the overall score * 5000
			.lUpfrontCost = lOverallScore * 5000
			'the maint cost is =ROUND((Overall - Loyalty)*(1.3+(RAND()*0.5)),0)
			.lMaintCost = CInt(Math.Abs(lOverallScore - CInt(.Loyalty)) * (1.3F + (Rnd() * 0.5F)))
		End With

		If oAgent.SaveObject() = False Then
			LogEvent(LogEventType.CriticalError, "GenerateNewAgent could not save the Agent!")
			Return Nothing
		Else : LogEvent(LogEventType.Informational, "New Agent Generated (" & oAgent.ObjectID & "): " & BytesToString(oAgent.AgentName))
		End If

		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To glAgentUB
			If glAgentIdx(X) = -1 Then
				lIdx = X
				Exit For
			End If
		Next X
		If lIdx = -1 Then
			lIdx = glAgentUB + 1
			ReDim Preserve glAgentIdx(lIdx)
			ReDim Preserve goAgent(lIdx)
			glAgentIdx(lIdx) = -1
			glAgentUB += 1
		End If
		oAgent.ServerIndex = lIdx
		goAgent(lIdx) = oAgent
		glAgentIdx(lIdx) = oAgent.ObjectID

		poOwner.AddAgentLookup(oAgent.ObjectID, lIdx)

		goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.RecruitDismiss, oAgent, Nothing, glCurrentCycle + 5184000)

		Return oAgent
	End Function

    Public Shared Function GenerateNewbieAgent(ByRef poOwner As Player) As Agent
        Dim oAgent As New Agent()

        With oAgent
            .bMale = True
            .AgentName = StringToBytes("Csaj Schnutic")
            .oOwner = poOwner

            .Infiltration = 85
            .Dagger = 33
            .Resourcefulness = 52
            .Loyalty = 80
            .Luck = 80

            .InfiltrationLevel = 0
            .InfiltrationType = 0
            .ObjTypeID = ObjectType.eAgent
            .oMission = Nothing
            .oOwner = poOwner
            .lTargetID = -1
            .iTargetTypeID = -1
            .Suspicion = 0
            .lAgentStatus = AgentStatus.NewRecruit

            .AddSkill(GetEpicaSkill(56), 98)
            .AddSkill(GetEpicaSkill(37), 94)

            'Ok, let's determine our costs now
            Dim lAttrScore As Int32 = CInt(.Dagger) + CInt(.Infiltration) + CInt(.Resourcefulness) + CInt(.Luck)
            Dim lSkillScore As Int32 = 0
            Dim lAbilityScore As Int32 = 0
            For X As Int32 = 0 To .SkillUB
                If .Skills(X).SkillType = 0 Then
                    lSkillScore += .SkillProf(X)
                Else : lAbilityScore += .SkillProf(X)
                End If
            Next X
            lAbilityScore *= 5

            Dim lOverallScore As Int32 = lAttrScore + lSkillScore + lAbilityScore
            'now, upfront is the overall score * 5000
            .lUpfrontCost = lOverallScore * 5000
            'the maint cost is =ROUND((Overall - Loyalty)*(1.3+(RAND()*0.5)),0)
            .lMaintCost = CInt(Math.Abs(lOverallScore - CInt(.Loyalty)) * (1.3F + (Rnd() * 0.5F)))
        End With

        If oAgent.SaveObject() = False Then
            LogEvent(LogEventType.CriticalError, "GenerateNewbieAgent could not save the Agent!")
            Return Nothing
        Else : LogEvent(LogEventType.Informational, "New Agent Generated (" & oAgent.ObjectID & "): " & BytesToString(oAgent.AgentName))
        End If

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To glAgentUB
            If glAgentIdx(X) = -1 Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            lIdx = glAgentUB + 1
            ReDim Preserve glAgentIdx(lIdx)
            ReDim Preserve goAgent(lIdx)
            glAgentIdx(lIdx) = -1
            glAgentUB += 1
        End If
        oAgent.ServerIndex = lIdx
        goAgent(lIdx) = oAgent
        glAgentIdx(lIdx) = oAgent.ObjectID

        poOwner.AddAgentLookup(oAgent.ObjectID, lIdx)

        'goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.RecruitDismiss, oAgent, Nothing, glCurrentCycle + 5184000)

        Return oAgent
    End Function

	Public Sub DeleteMe()
		'Do delete code here
		If Me.ObjectID > -1 Then
			Dim sSQL As String = "DELETE FROM tblAgentSkill WHERE AgentID = " & Me.ObjectID
			Dim oComm As New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm.Dispose()
			oComm = Nothing
			sSQL = "DELETE FROM tblAgent WHERE AgentID = " & Me.ObjectID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm.Dispose()
			oComm = Nothing

			'delete this agent from any assignments in progress, etc...
			RemoveMeFromMissions()

			'now, remove me from the server array
			If ServerIndex > -1 AndAlso glAgentIdx(ServerIndex) = ObjectID Then glAgentIdx(ServerIndex) = -1
		End If
	End Sub

	Public Sub ReturnHome()
        'RemoveMeFromMissions()
        Try
            Dim lOwnerID As Int32 = Me.oOwner.ObjectID
            For X As Int32 = 0 To glPlayerMissionUB
                If glPlayerMissionIdx(X) <> -1 AndAlso goPlayerMission(X).oPlayer.ObjectID = lOwnerID Then
                    If goPlayerMission(X).lCurrentPhase >= eMissionPhase.ePreparationTime AndAlso goPlayerMission(X).lCurrentPhase < eMissionPhase.eReinfiltrationPhase Then
                        goPlayerMission(X).RemoveAgentFromAssignments(Me.ObjectID)
                    End If
                End If
            Next X
        Catch
        End Try

		Me.lAgentStatus = AgentStatus.ReturningHome
		If Me.iTargetTypeID = ObjectType.ePlayer AndAlso Me.lTargetID = oOwner.ObjectID Then
			Me.lArrivalCycles = 0
			Me.lArrivalStart = glCurrentCycle
		Else
			Me.lArrivalCycles = ml_DEINFILTRATE_TIME
			Me.lArrivalStart = glCurrentCycle
		End If
		goAgentEngine.CancelAgentEvent(AgentEngine.EventTypeID.AgentReInfiltrate, Me.ObjectID)
		goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentDeinfiltrateComplete, Me, Nothing, glCurrentCycle + Me.lArrivalCycles)
		Me.lTargetID = -1
		Me.iTargetTypeID = -1
		Me.InfiltrationType = eInfiltrationType.eGeneralInfiltration
		Me.InfiltrationLevel = 0

		Dim yMsg(13) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(ObjectID).CopyTo(yMsg, 2)
		System.BitConverter.GetBytes(AgentStatus.ReturningHome).CopyTo(yMsg, 6)
		System.BitConverter.GetBytes(lArrivalCycles).CopyTo(yMsg, 10)
		Me.oOwner.SendPlayerMessage(yMsg, False, AliasingRights.eViewAgents)
	End Sub

	Public Sub RemoveMeFromMissions()
		Try
			Dim lOwnerID As Int32 = Me.oOwner.ObjectID
			For X As Int32 = 0 To glPlayerMissionUB
				If glPlayerMissionIdx(X) <> -1 AndAlso goPlayerMission(X).oPlayer.ObjectID = lOwnerID Then
					goPlayerMission(X).RemoveAgentFromAssignments(Me.ObjectID)
				End If
			Next X
		Catch
		End Try
	End Sub

	'returns whether the next agent check in should occur
	Public Function HandleAgentCheckIn() As Boolean
		If Me.iTargetTypeID <> ObjectType.ePlayer Then Return False
		Dim oTarget As Player = GetEpicaPlayer(Me.lTargetID)
		If oTarget Is Nothing Then Return False
        Dim oPI As PlayerIntel = Me.oOwner.GetOrAddPlayerIntel(Me.lTargetID, False)
        If oPI Is Nothing Then Return False

        'need to do a detection test
        Dim lAgentVal As Int32 = Me.Resourcefulness
        lAgentVal += Me.GetSkillValue(lSkillHardcodes.eCryptography, False, 0)
        lAgentVal += Math.Max(Math.Max(Me.GetSkillValue(lSkillHardcodes.eElectronics, True, 0), Me.GetSkillValue(lSkillHardcodes.eTinyElectronics, True, 0)), Me.GetSkillValue(lSkillHardcodes.eNaturallyTalented, True, 0))
        lAgentVal += Me.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)
        For X As Int32 = 0 To oTarget.oSecurity.lCounterAgentUB
            If oTarget.oSecurity.lCounterAgentIdx(X) > 0 Then
                Dim oCounter As Agent = oTarget.oSecurity.oCounterAgents(X)
                If oCounter Is Nothing = False Then
                    If oCounter.InfiltrationType = Me.InfiltrationType Then
                        If (oCounter.lAgentStatus And (AgentStatus.Dismissed Or AgentStatus.HasBeenCaptured Or AgentStatus.Infiltrating Or AgentStatus.IsDead Or AgentStatus.ReturningHome)) = 0 Then
                            Dim lCounterVal As Int32 = oCounter.Resourcefulness
                            lCounterVal += oCounter.GetSkillValue(lSkillHardcodes.eCryptography, False, 0)
                            lCounterVal += oCounter.GetSkillValue(lSkillHardcodes.eBattleLanguage, False, 0)
                            lCounterVal += oCounter.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)

                            Dim lTestVal As Int32 = lAgentVal + CInt(Rnd() * 100 - 50)
                            If lTestVal < lCounterVal Then
                                'ok, i was caught reporting... increase my suspicion
                                Dim lTmpSus As Int32 = Me.Suspicion
                                lTmpSus += CInt(Rnd() * 5 + 1)
                                If lTmpSus > 255 Then lTmpSus = 255
                                Me.Suspicion = CByte(lTmpSus)
                            End If
                        End If
                    End If
                End If
            End If
        Next X

        Dim yTypeID As Byte = 0
		Select Case Me.InfiltrationType
			Case eInfiltrationType.eColonialInfiltration
                oPI.PopulationScore = oTarget.PopulationScore
                oPI.PopulationUpdate = GetDateAsNumber(Now)
                yTypeID = 3
			Case eInfiltrationType.eFederalInfiltration
                oPI.DiplomacyScore = oTarget.DiplomacyScore
                oPI.DiplomacyUpdate = GetDateAsNumber(Now)
                yTypeID = 6
			Case eInfiltrationType.eMilitaryInfiltration
                oPI.MilitaryScore = oTarget.lMilitaryScore \ 50
                oPI.MilitaryUpdate = GetDateAsNumber(Now)
                yTypeID = 5
			Case eInfiltrationType.eProductionInfiltration
                oPI.ProductionScore = oTarget.ProductionScore
                oPI.ProductionUpdate = GetDateAsNumber(Now)
                yTypeID = 4
			Case eInfiltrationType.eResearchInfiltration
                oPI.TechnologyScore = oTarget.TechnologyScore
                oPI.TechnologyUpdate = GetDateAsNumber(Now)
                yTypeID = 1
			Case eInfiltrationType.eTradeInfiltration
                oPI.WealthScore = oTarget.WealthScore
                oPI.WealthUpdate = GetDateAsNumber(Now)
                yTypeID = 2
			Case Else
				Return False
		End Select
        'If yTypeID <> 0 Then goMsgSys.SendRequestGlobalPlayerScores(Me.lTargetID, yTypeID, Me.oOwner.ObjectID)

        Me.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPI, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
		Return True
	End Function

	Public Shared Function GetCycleDelayByFreq(ByVal yFreq As Byte) As Int32
		Select Case yFreq
			Case eReportFreq.OncePerDay
				Return 2592000
			Case eReportFreq.OncePerHalfHour
				Return 54000
			Case eReportFreq.OncePerHour
				Return 108000
			Case eReportFreq.OncePerSixHours
				Return 648000
			Case eReportFreq.OncePerTwelveHours
				Return 1296000
			Case eReportFreq.OncePerTwoDays
				Return 5184000
			Case eReportFreq.OncePerTwoHours
				Return 216000
			Case eReportFreq.OncePerWeek
				Return 18144000
		End Select
		Return 54000
    End Function

    Public Sub LoadFromDataReader(ByVal oData As OleDb.OleDbDataReader, ByVal lSrvrIndex As Int32)
        With Me
            .ServerIndex = lSrvrIndex
            .AgentName = StringToBytes(CStr(oData("AgentName")))
            .Dagger = CByte(oData("Dagger"))
            .Infiltration = CByte(oData("Infiltration"))
            .InfiltrationLevel = CByte(oData("InfiltrationLevel"))
            .InfiltrationType = CType(oData("InfiltrationType"), eInfiltrationType)
            .Loyalty = CByte(oData("Loyalty"))
            .Luck = CByte(oData("Luck"))
            .ObjectID = CInt(oData("AgentID"))
            .ObjTypeID = ObjectType.eAgent
            .oOwner = GetEpicaPlayer(CInt(oData("OwnerID")))
            .lTargetID = CInt(oData("TargetID"))
            .iTargetTypeID = CShort(oData("TargetTypeID"))
            .Resourcefulness = CByte(oData("Resourcefulness"))
            .Suspicion = CByte(oData("Suspicion"))
            .bMale = CBool(oData("IsMale"))
            .dtRecruited = CDate(GetDateFromNumber(CInt(oData("RecruitedOn"))))
            .lMaintCost = CInt(oData("MaintCost"))
            .lUpfrontCost = CInt(oData("UpfrontCost"))

            .lArrivalCycles = CInt(oData("ArrivalCycles"))
            .lCapturedBy = CInt(oData("CapturedBy"))
            .lCapturedOn = CInt(oData("CapturedOn"))
            .lPrisonTestCycles = CInt(oData("PrisonTestCycles"))
            If lCapturedBy > 0 And .lPrisonTestCycles <= 0 Then 'Self Healing for already captured agents
                .lPrisonTestCycles = CInt((2052000 * Rnd()) - (540000)) '5-24hr (19*rnd+5)
            End If
            .yHealth = CByte(oData("Health"))
            .lInterrogatorID = CInt(oData("InterrogatorID"))
            .ySpilledData = CType(oData("ySpilledData"), eSpilledData)
            .lAgentStatus = CType(CInt(oData("AgentStatus")), AgentStatus)
            .yReportFreq = CByte(oData("ReportFreq"))

            If .oOwner Is Nothing = False Then .oOwner.AddAgentLookup(.ObjectID, lSrvrIndex)

            If .Suspicion > 0 Then
                goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.ReduceSuspicion, goAgent(lSrvrIndex), Nothing, glCurrentCycle + Agent.ml_DEINFILTRATE_TIME + CInt(Rnd() * 1000))
            End If
            If .lPrisonTestCycles > 0 Then
                goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.PrisonEscapeTest, goAgent(lSrvrIndex), Nothing, glCurrentCycle + .lPrisonTestCycles)
            End If
        End With
    End Sub
End Class


Public Enum eDaggerTestType As Byte
    AvoidCapture = 0
    Assassination = 1
End Enum