Public Enum ePM_ErrorCode As Int32
	eUnknownError = Int32.MinValue
	eNoMissionSet = -2
	eNoErrorsFound = 1
	eInvalidMethod = -4
	eMissingSkillset = -5
	eMissingAssignment = -6
	eMissingAssignmentAgent = -7
	eAssignedAgentUnavail = -8
	eInvalidTargetSelection = -9
End Enum

Public Class PlayerMission
	Public PM_ID As Int32				'unique identifier

	Public oMission As Mission
	Public CasualtyCnt As Int32 = 0

	Public oPlayer As Player				'player who started this mission (agents belong to this player)
	Public lTargetPlayerID As Int32

	'Specific item GUID... for example, a specific factory, unit, or something could be selected to be destroyed...
    '  a specific agent could be selected to be rescued.... etc...
	Public lTargetID As Int32
	Public iTargetTypeID As Int16
	Public lTargetID2 As Int32
	Public iTargetTypeID2 As Int16

	Public Modifier As Int32
	Public bAlarmThrown As Boolean = False

	Public oMissionGoals() As PlayerMissionGoal
	Public lMethodID As Int32
    Public ySafeHouseSetting As Byte = 0        '0 = No Safe House, Non-zero is a bonus to rolls that benefit from safehouses
    Public oSafehouseMissionGoal As PlayerMissionGoal

	Public oPhases() As PlayerMissionPhase
	Public lPhaseUB As Int32 = -1

	Public bArchived As Boolean = False

	Public yCurrentPhase As eMissionPhase = eMissionPhase.eInPlanning

	Public Function GetListBoxText() As String

        '' Old List Box
        '      Dim sResult As String = oMission.sMissionName.PadRight(50, " "c)
        'If lTargetPlayerID > -1 Then
        '	sResult &= GetCacheObjectValue(lTargetPlayerID, ObjectType.ePlayer).PadRight(25, " "c)
        'Else
        '	sResult &= GetCacheObjectValue(lTargetID, iTargetTypeID).PadRight(25, " "c)
        'End If
        'sResult &= GetPhaseText(yCurrentPhase)
        'Return sResult
        '' End Old

        Dim sResult As String = oMission.sMissionName.PadRight(35, " "c)
        Dim sText As String = ""
        If lTargetPlayerID > -1 Then
            sText &= GetCacheObjectValue(lTargetPlayerID, ObjectType.ePlayer)
        End If
        If lTargetID > -1 AndAlso iTargetTypeID > -1 Then
            If iTargetTypeID = 0 AndAlso oMission Is Nothing = False AndAlso oMission.ObjectID = eMissionResult.eGetFacilityList Then
                If sText <> "" Then sText &= " - "
                Select Case lTargetID
                    Case ProductionType.eEnlisted
                        sText &= "Barracks"
                    Case ProductionType.eCommandCenterSpecial
                        sText &= "Command Center"
                    Case ProductionType.eProduction
                        sText &= "Factory"
                    Case ProductionType.eProduction
                        sText &= "Land Production"
                    Case ProductionType.eMining
                        sText &= "Mining"
                    Case ProductionType.eNavalProduction
                        sText &= "Naval Production"
                    Case ProductionType.eOfficers
                        sText &= "Officers"
                    Case ProductionType.ePowerCenter
                        sText &= "Power Center"
                    Case ProductionType.eRefining
                        sText &= "Refining"
                    Case ProductionType.eResearch
                        sText &= "Research"
                    Case ProductionType.eColonists
                        sText &= "Residence"
                    Case ProductionType.eAerialProduction
                        sText &= "Spaceport"
                    Case ProductionType.eSpaceStationSpecial
                        sText &= "Space Station"
                    Case ProductionType.eTradePost
                        sText &= "Tradepost"
                    Case ProductionType.eWareHouse
                        sText &= "Warehouse"
                    Case Else
                        sText &= "Unknown"
                End Select
            ElseIf iTargetTypeID = 0 Then
                'sText &= "Any"
            Else
                If sText <> "" Then sText &= " - "
                sText &= GetCacheObjectValue(lTargetID, iTargetTypeID)
                For X As Int32 = 0 To glItemIntelUB
                    If glItemIntelIdx(X) <> -1 AndAlso goItemIntel(X).lOtherPlayerID = lTargetPlayerID AndAlso goItemIntel(X).iItemTypeID = iTargetTypeID AndAlso goItemIntel(X).lItemID = lTargetID AndAlso goItemIntel(X).EnvirTypeID > 0 AndAlso goItemIntel(X).EnvirID > 0 Then
                        sText &= " - " & GetCacheObjectValue(goItemIntel(X).EnvirID, goItemIntel(X).EnvirTypeID)
                    End If
                Next X
            End If
        End If
        If lTargetID2 > -1 AndAlso iTargetTypeID2 > -1 Then
            If sText <> "" Then sText &= " - "
            sText &= GetCacheObjectValue(lTargetID2, iTargetTypeID2)
        End If
        sResult &= sText.PadRight(40, " "c)

        sResult &= GetPhaseText(yCurrentPhase)
        'TODO: Show the progress of running missions.  Appears to only show the first phase?
        'If yCurrentPhase <= eMissionPhase.eWaitingToExecute AndAlso yCurrentPhase >= eMissionPhase.ePreparationTime Then
        '    If oMissionGoals Is Nothing = False Then
        '        Dim sProgress As String = ""
        '        Dim iProgress As Int32 = 0
        '        Dim iTotal As Int32 = 0
        '        For x As Int32 = 0 To oMissionGoals.GetUpperBound(0)
        '            If oMissionGoals(x).oAssignments Is Nothing = False Then
        '                For y As Int32 = 0 To oMissionGoals(x).oAssignments.GetUpperBound(0)
        '                    iProgress += oMissionGoals(x).oAssignments(y).PointsProduced
        '                    iTotal += oMissionGoals(x).oAssignments(y).PointsRequired
        '                Next
        '            End If
        '        Next

        '        If iTotal > 0 Then
        '            sProgress = iProgress.ToString & "/" & iTotal.ToString
        '        End If

        '        If sProgress <> "" Then
        '            sResult &= " (" & sProgress & ")"
        '        End If
        '    End If
        'End If
        Return sResult
    End Function

	Public Shared Function GetPhaseText(ByVal lPhase As eMissionPhase) As String

		Dim lTemp As Int32 = lPhase
		Dim sSuffix As String = ""
		If (lTemp And eMissionPhase.eMissionPaused) <> 0 Then
			lTemp = lTemp Xor eMissionPhase.eMissionPaused
			sSuffix = " - Paused"
		End If

		Select Case lTemp
			Case eMissionPhase.eCancelled
				Return "Cancelled" & sSuffix
			Case eMissionPhase.eFlippingTheSwitch
				Return "Executing" & sSuffix
			Case eMissionPhase.eInPlanning
				Return "In Planning" & sSuffix
			Case eMissionPhase.eMissionOverFailure
				Return "Completed: Failure" & sSuffix
			Case eMissionPhase.eMissionOverSuccess
				Return "Completed: Success" & sSuffix
			Case eMissionPhase.ePreparationTime
				Return "Preparation" & sSuffix
			Case eMissionPhase.eReinfiltrationPhase
				Return "Reinfiltrating" & sSuffix
			Case eMissionPhase.eSettingTheStage
				Return "In Progress" & sSuffix
			Case eMissionPhase.eWaitingToExecute
				Return "Waiting to Execute" & sSuffix
		End Select
		Return "In Planning" & sSuffix
	End Function

	Public Function CreateClone() As PlayerMission
		Dim oNew As New PlayerMission()
		With oNew
			.bAlarmThrown = bAlarmThrown
			.CasualtyCnt = CasualtyCnt
			.iTargetTypeID = iTargetTypeID
			.iTargetTypeID2 = iTargetTypeID2
			.yCurrentPhase = yCurrentPhase
			.lMethodID = lMethodID
			.lPhaseUB = lPhaseUB
			.lTargetID = lTargetID
			.lTargetID2 = lTargetID2
			.lTargetPlayerID = lTargetPlayerID
			.Modifier = Modifier
			.oMission = oMission

			ReDim .oMissionGoals(oMissionGoals.GetUpperBound(0))
			For X As Int32 = 0 To oMissionGoals.GetUpperBound(0)
				.oMissionGoals(X) = oMissionGoals(X)
			Next X
			ReDim .oPhases(lPhaseUB)
			For X As Int32 = 0 To lPhaseUB
				.oPhases(X) = oPhases(X)
			Next X

			.oPlayer = oPlayer
            '.lTargetPlayerID = 2
			.PM_ID = PM_ID
			.ySafeHouseSetting = ySafeHouseSetting
		End With
		Return oNew
	End Function

	Public Function FillFromMsg(ByRef yData() As Byte) As Boolean
		Try
			Dim lPos As Int32 = 2	'for msgcode
			PM_ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            'Only initial GetAddObject sends this flag.  All the mission submitions do not, and would routinly error not displaying the proper message.
            If yData.Length > 11 Then Me.bArchived = yData(lPos) <> 0 : lPos += 1

			Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			'lTemp is either a MissionID, -1, or an Error Code (ePM_ErrorCode) which is less than -1
			If lTemp < -1 Then
				If goUILib Is Nothing = False Then
					Select Case CType(lTemp, ePM_ErrorCode)
						Case ePM_ErrorCode.eAssignedAgentUnavail
							goUILib.AddNotification("Mission plan failed: Assigned Agent Unavailable (Captured or Dead).", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Case ePM_ErrorCode.eInvalidMethod
							goUILib.AddNotification("Mission plan failed: Invalid Method Selection!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Case ePM_ErrorCode.eInvalidTargetSelection
							goUILib.AddNotification("Mission plan failed: Invalid Target Selection!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Case ePM_ErrorCode.eMissingAssignment
							goUILib.AddNotification("Mission plan failed: Missing Assignment!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Case ePM_ErrorCode.eMissingAssignmentAgent
							goUILib.AddNotification("Mission Plan Failed: Missing Agent Assignment!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Case ePM_ErrorCode.eMissingSkillset
							goUILib.AddNotification("Mission Plan Failed: Missing Skillset selection!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Case ePM_ErrorCode.eNoMissionSet
							goUILib.AddNotification("Mission Plan Failed: No mission selected!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Case Else
					End Select
				End If
				Return False
			Else
				If lTemp <> -1 Then
					For X As Int32 = 0 To glMissionUB
						If glMissionIdx(X) = lTemp Then
							Me.oMission = goMissions(X)
							ReDim oMissionGoals(oMission.GoalUB)
							For Y As Int32 = 0 To oMission.GoalUB
								oMissionGoals(Y) = New PlayerMissionGoal()
								oMissionGoals(Y).oGoal = oMission.Goals(Y)
								oMissionGoals(Y).oMission = Me
							Next Y
							Exit For
						End If
					Next X
				End If
			End If

			lTargetPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			lTargetID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			iTargetTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			lTargetID2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			iTargetTypeID2 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			ySafeHouseSetting = yData(lPos) : lPos += 1
			lMethodID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			yCurrentPhase = CType(yData(lPos), eMissionPhase) : lPos += 1
            bAlarmThrown = yData(lPos) <> 0 : lPos += 1

            'ok, first, check our safehousesetting
            If ySafeHouseSetting <> 0 Then
                If oSafehouseMissionGoal Is Nothing Then
                    oSafehouseMissionGoal = New PlayerMissionGoal()
                    oSafehouseMissionGoal.oGoal = Goal.GetSafehouseGoal()
                    oSafehouseMissionGoal.oMission = Me
                End If
                'skillsetid
                lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                oSafehouseMissionGoal.oSkillSet = oSafehouseMissionGoal.oGoal.GetSkillset(lTemp)
                'assignment1 agentid
                Dim lAgent1 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                'assignment1 skillid
                Dim lSkill1 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iPoints1 As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim yStatus1 As Byte = yData(lPos) : lPos += 1
                'assignment2 agentid
                Dim lAgent2 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                'assignment2 skillid
                Dim lSkill2 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iPoints2 As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim yStatus2 As Byte = yData(lPos) : lPos += 1
                If lAgent1 > -1 Then
                    Dim oAgent As Agent = Nothing
                    For Z As Int32 = 0 To goCurrentPlayer.AgentUB
                        If goCurrentPlayer.AgentIdx(Z) = lAgent1 Then
                            oAgent = goCurrentPlayer.Agents(Z) : Exit For
                        End If
                    Next Z
                    Dim oSkill As Skill = Nothing
                    For Z As Int32 = 0 To glSkillUB
                        If glSkillIdx(Z) = lSkill1 Then
                            oSkill = goSkills(Z) : Exit For
                        End If
                    Next Z
                    If oAgent Is Nothing = False Then
                        If oAgent Is Nothing = False AndAlso oSkill Is Nothing = False Then

                            Dim lPointsRequired As Int32 = 0
                            If oSafehouseMissionGoal.oSkillSet Is Nothing = False Then
                                For Z As Int32 = 0 To oSafehouseMissionGoal.oSkillSet.SkillUB
                                    If oSkill.ObjectID = oSafehouseMissionGoal.oSkillSet.Skills(Z).oSkill.ObjectID Then
                                        lPointsRequired = oSafehouseMissionGoal.oSkillSet.Skills(Z).PointRequirement
                                        Exit For
                                    End If
                                Next Z
                            End If

                            Dim oAA As AgentAssignment = oSafehouseMissionGoal.AddAgentAssignment(oAgent, oSkill)
                            oAA.PointsProduced = iPoints1
                            oAA.PointsRequired = lPointsRequired
                            oAA.yStatus = yStatus1
                        End If
                    End If
                End If
                If lAgent2 > -1 Then
                    Dim oAgent As Agent = Nothing
                    For Z As Int32 = 0 To goCurrentPlayer.AgentUB
                        If goCurrentPlayer.AgentIdx(Z) = lAgent2 Then
                            oAgent = goCurrentPlayer.Agents(Z) : Exit For
                        End If
                    Next Z
                    Dim oSkill As Skill = Nothing
                    For Z As Int32 = 0 To glSkillUB
                        If glSkillIdx(Z) = lSkill2 Then
                            oSkill = goSkills(Z) : Exit For
                        End If
                    Next Z

                    If oAgent Is Nothing = False Then
                        If oAgent Is Nothing = False AndAlso oSkill Is Nothing = False Then

                            Dim lPointsRequired As Int32 = 0
                            If oSafehouseMissionGoal.oSkillSet Is Nothing = False Then
                                For Z As Int32 = 0 To oSafehouseMissionGoal.oSkillSet.SkillUB
                                    If oSkill.ObjectID = oSafehouseMissionGoal.oSkillSet.Skills(Z).oSkill.ObjectID Then
                                        lPointsRequired = oSafehouseMissionGoal.oSkillSet.Skills(Z).PointRequirement
                                        Exit For
                                    End If
                                Next Z
                            End If

                            Dim oAA As AgentAssignment = oSafehouseMissionGoal.AddAgentAssignment(oAgent, oSkill)
                            oAA.PointsProduced = iPoints2
                            oAA.PointsRequired = lPointsRequired
                            oAA.yStatus = yStatus2
                        End If
                    End If
                End If 
            End If

            lPhaseUB = CInt(yData(lPos)) - 1 : lPos += 1
            ReDim oPhases(lPhaseUB)
            For X As Int32 = 0 To lPhaseUB
                oPhases(X) = New PlayerMissionPhase()
                oPhases(X).yPhase = CType(yData(lPos), eMissionPhase) : lPos += 1
                Dim lCvrAgntCnt As Int32 = CInt(yData(lPos)) : lPos += 1
                For Y As Int32 = 0 To lCvrAgntCnt - 1
                    Dim oAgent As Agent = Nothing
                    Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    For Z As Int32 = 0 To goCurrentPlayer.AgentUB
                        If goCurrentPlayer.AgentIdx(Z) = lAgentID Then
                            oAgent = goCurrentPlayer.Agents(Z)
                            Exit For
                        End If
                    Next Z
                    If oAgent Is Nothing = False Then oPhases(X).AddCoverAgent(oAgent, 0)
                Next Y
            Next X

            Dim lCnt As Int32 = CInt(yData(lPos)) : lPos += 1
            For X As Int32 = 0 To lCnt - 1
                Dim lGoalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lSSID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lAssCnt As Int32 = yData(lPos) : lPos += 1

                Dim oMG As PlayerMissionGoal = Nothing
                For Y As Int32 = 0 To oMission.GoalUB
                    If oMission.MethodIDs(Y) = lMethodID AndAlso oMission.Goals(Y).ObjectID = lGoalID Then
                        oMG = oMissionGoals(Y)
                        Exit For
                    End If
                Next Y

                Dim oSkillset As SkillSet = Nothing
                If oMG Is Nothing = False Then
                    For Y As Int32 = 0 To oMG.oGoal.SkillSetUB
                        If oMG.oGoal.SkillSets(Y).SkillSetID = lSSID Then
                            oSkillset = oMG.oGoal.SkillSets(Y)
                            Exit For
                        End If
                    Next Y
                    If oSkillset Is Nothing = False Then
                        oMG.oSkillSet = oSkillset
                    Else : Continue For ' Err.Raise(-1, "PM:Fill:", "MissionGoalSkillsetNotFound")
                    End If
                Else : Continue For ' Err.Raise(-1, "PM:Fill", "MissionGoalNotFound")
                End If

                oMG.lAssignmentUB = -1
                For Y As Int32 = 0 To lAssCnt - 1
                    Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4    '		AgentID (4)
                    Dim lSkillID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4    '		SkillID (4)
                    Dim lPointProd As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    Dim yStatus As Byte = yData(lPos) : lPos += 1
                    Dim oAgent As Agent = Nothing
                    Dim oSkill As Skill = Nothing

                    For Z As Int32 = 0 To goCurrentPlayer.AgentUB
                        If goCurrentPlayer.AgentIdx(Z) = lAgentID Then
                            oAgent = goCurrentPlayer.Agents(Z) : Exit For
                        End If
                    Next Z
                    For Z As Int32 = 0 To glSkillUB
                        If glSkillIdx(Z) = lSkillID Then
                            oSkill = goSkills(Z) : Exit For
                        End If
                    Next Z
                    If oAgent Is Nothing = False AndAlso oSkill Is Nothing = False Then

                        Dim lPointsRequired As Int32 = 0
                        If oSkillset Is Nothing = False Then
                            For Z As Int32 = 0 To oSkillset.SkillUB
                                If oSkill.ObjectID = oSkillset.Skills(Z).oSkill.ObjectID Then
                                    lPointsRequired = oSkillset.Skills(Z).PointRequirement
                                    Exit For
                                End If
                            Next Z
                        End If

                        Dim oAA As AgentAssignment = oMG.AddAgentAssignment(oAgent, oSkill)
                        oAA.PointsProduced = lPointProd
                        oAA.PointsRequired = lPointsRequired
                        oAA.yStatus = yStatus
                    End If
                Next Y

            Next X


            Return True
        Catch ex As Exception
            If goUILib Is Nothing = False Then goUILib.AddNotification("PM:Fill:" & ex.Message, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End Try
	End Function
End Class

Public Class PlayerMissionPhase
    Public oParent As PlayerMission
	Public yPhase As eMissionPhase

    Public oCoverAgents() As Agent
    Public lCoverAgentIdx() As Int32
    Public yUsedAsCoverAgent() As Byte
    Public lCoverAgentUB As Int32 = -1

    Public Sub AddCoverAgent(ByRef oAgent As Agent, ByVal yUsed As Byte)
        Dim lID As Int32 = oAgent.ObjectID
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lCoverAgentUB
            If lCoverAgentIdx(X) = lID Then
                yUsedAsCoverAgent(X) = yUsed
                Return
            End If
            If lIdx = -1 AndAlso lCoverAgentIdx(X) = -1 Then lIdx = X
        Next X
        If lIdx = -1 Then
            lCoverAgentUB += 1
            ReDim Preserve oCoverAgents(lCoverAgentUB)
            ReDim Preserve yUsedAsCoverAgent(lCoverAgentUB)
            ReDim Preserve lCoverAgentIdx(lCoverAgentUB)
            lIdx = lCoverAgentUB
        End If
        oCoverAgents(lIdx) = oAgent
        lCoverAgentIdx(lIdx) = oAgent.ObjectID
        yUsedAsCoverAgent(lIdx) = yUsed
    End Sub

    Public Sub RemoveCoverAgent(ByVal lAgentID As Int32)
        For X As Int32 = 0 To lCoverAgentUB
            If lCoverAgentIdx(X) = lAgentID Then
                lCoverAgentIdx(X) = -1
                oCoverAgents(X) = Nothing
            End If
        Next X
    End Sub
 
End Class