Public Enum eMissionResult As Int32
    eUndefined = 0
    eShowQueueAndContents = 1           'show prod/res/whatever else queue and contents
    eSlowProduction = 2                 'Slow production (research/mining/etc...)
    eSabotageProduction = 3             'sabotage production
    eCorruptProduction = 4              'reduces reliability
    eDestroyFacility = 5                'destroys target facility
    eClutterCargoBay = 6                'reduces cargo bay capacity
    eDoorJamHangar = 7                  'increases docking/undocking delay
    eStealCargo = 8                     'steal cargo
    eSpecialProjectList = 9             'displays the special project technology list
    eComponentDesignList = 10           'Gets list of designed components
    eComponentResearchList = 11         'gets list of researched components
    eGetMineralList = 12                'gets the mineral list
    eAcquireTechData = 13               'gets the data of a component/prototype/special tech
    eGetTradeTreaties = 14              'gets the current trade treaties
    eGetFacilityList = 15               'gets a list of the facilities for the provided production type
    eSiphonItem = 16                    'pulls credits/minerals/production/etc... based on infiltration
    eBadPublicity = 17                  'reduces trade revenue from allies
    eGetColonyBudget = 18               'returns budget for the colony
    eAssassinateGovernor = 19           'short-term morale and production penalty
    eDecreaseTaxMorale = 20             'colony morale decreases based on Taxes (tax morale modifier)
    eDecreaseUnemploymentMorale = 21    'colony morale decreases based on unemployment (unemployment modifier)
    eDecreaseHousingMorale = 22       'colony morale decreases based on housing (housing modifier)
    eGetAlliesList = 23                 'gets the allies with the values
    eGetEnemyList = 24                  'gets the enemies with the values
    ePlantEvidence = 25                 'plants evidence
    eDiplomaticStrength = 26            'gets the diplomacy score
    eMilitaryStrength = 27              'gets the military score
    eTechnologyStrength = 28            'gets the tech score
    ePopulationScore = 29               'gets the population score
    eProductionScore = 30               'gets the production score
    eMilitaryCoup = 31                  'half units in an environment change alignment

    eFindMineral = 32                   'searches for a specific mineral within the infiltrated system
    eFindCommandCenters = 33            'searches for command centers within the infiltrated planet belonging to the target player
    eFindSpaceStations = 34             'searches for space stations within the infiltrated system belonging to the target player
    eFindProductionFac = 35             'searches for production facilities within the infiltrated planet belonging to the target player
    eFindResearchFac = 36               'searches for research facilities within the infiltrated planet belonging to the target player

    eAssassinateAgent = 37              'searches for an agent (specific or any) belonging to a target player within the infiltrated system and kills them
    eCaptureAgent = 38                  'searches for an agent (specific or any) belonging to a target player within the infiltrated system and captures them
    eSearchAndRescueAgent = 39          'searches for an agent within the infiltrated system and rescues them

    eReconPlanetMap = 40                'exposes an area of the fog of war for a period of time (temporary or permanent based on mission effect) of the infiltrated planet
    eGeologicalSurvey = 41              'Lists minerals for the infiltrated planet

    eRelayBattleReports = 42            'Provides detailed reports about combat
    eIFFSabotage = 43                   'enables friendly fire
    eDegradePay = 44                    'degrade the pay for corporation
    eTutorialFindFactory = 45

    eAgencyAnalysis = 46                'get the strength of the target counter agent network
    eGetAgentList = 47
    eDropInvulField = 48
    eLocateWormhole = 49
    eAudit = 50
    eIncreasePowerNeeds = 51
    eGetAliasList = 52
    eImpedeCurrentDevelopment = 53
    eDestroyCurrentSpecialProject = 54
End Enum

Public Class Mission
    Inherits Epica_GUID

    Public MissionName(19) As Byte
    Public MissionDesc(254) As Byte

    Public lInfiltrationType As eInfiltrationType = eInfiltrationType.eGeneralInfiltration

    Public ProgramControlID As Int16
    Public BaseEffect As Int16
    Public Modifier As Int16

    Public Goals() As Goal
    Public MethodIDs() As Int32     'stored as int32, but should follow BYTE rules
    Public GoalUB As Int32 = -1

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        If mbStringReady = False Then
            Dim lPos As Int32 = 0
            ReDim mySendString(291 + ((GoalUB + 1) * 5))
            Me.GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
            MissionName.CopyTo(mySendString, lPos) : lPos += 20
            System.BitConverter.GetBytes(Modifier).CopyTo(mySendString, lPos) : lPos += 2
            System.BitConverter.GetBytes(BaseEffect).CopyTo(mySendString, lPos) : lPos += 2
            System.BitConverter.GetBytes(ProgramControlID).CopyTo(mySendString, lPos) : lPos += 2
            mySendString(lPos) = lInfiltrationType : lPos += 1
            MissionDesc.CopyTo(mySendString, lPos) : lPos += 255

            System.BitConverter.GetBytes(GoalUB + 1).CopyTo(mySendString, lPos) : lPos += 4
            For X As Int32 = 0 To GoalUB
                mySendString(lPos) = CByte(MethodIDs(X)) : lPos += 1
                System.BitConverter.GetBytes(Goals(X).ObjectID).CopyTo(mySendString, lPos) : lPos += 4
            Next X

            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Sub AddGoal(ByRef oGoal As Goal, ByVal lMethodID As Int32)
        GoalUB += 1
        ReDim Preserve Goals(GoalUB)
        ReDim Preserve MethodIDs(GoalUB)
        Goals(GoalUB) = oGoal
        MethodIDs(GoalUB) = lMethodID
    End Sub

	'If you use this routine, add MethodID as a lookup parm
	'Public Sub RemoveGoal(ByRef oGoal As Goal)
	'    Dim lIdx As Int32 = -1
	'    For X As Int32 = 0 To GoalUB
	'        If Goals(X).ObjectID = oGoal.ObjectID Then
	'            lIdx = X
	'            Exit For
	'        End If
	'    Next X

	'    'go back thru and shift
	'    If Not (lIdx = -1 AndAlso GoalUB = 0) Then
	'        For X As Int32 = GoalUB To lIdx + 1 Step -1
	'            Goals(X) = Goals(X - 1)
	'            MethodIDs(X) = MethodIDs(X - 1)
	'        Next X
	'    End If
	'    GoalUB -= 1
	'End Sub

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand
        Dim X As Int32

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblMission (MissionName, MissionDesc, InfiltrationType, ProgramControlID, BaseEffect, Modifier) VALUES ('" & _
                  MakeDBStr(BytesToString(MissionName)) & "', '" & MakeDBStr(BytesToString(MissionDesc)) & "', " & CByte(lInfiltrationType) & _
                  ", " & ProgramControlID & ", " & BaseEffect & ", " & Modifier & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblMission SET MissionName = '" & MakeDBStr(BytesToString(MissionName)) & "', '" & _
                  MakeDBStr(BytesToString(MissionDesc)) & "', InfiltrationType = " & CByte(lInfiltrationType) & ", ProgramControlID = " & _
                  ProgramControlID & ", BaseEffect = " & BaseEffect & ", Modifier = " & Modifier & " WHERE MissionID = " & Me.ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(MissionID) FROM tblMission WHERE MissionName = '" & MakeDBStr(BytesToString(MissionName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            For X = 0 To GoalUB
                If Goals(X).SaveObject() = False Then
                    Err.Raise(-1, "Mission.Goals.SaveObject", "Unable to save Mission Goal!")
                End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function
End Class

Public Structure uMissionMethod
    Public lID As Int32
    Public yName() As Byte
End Structure
