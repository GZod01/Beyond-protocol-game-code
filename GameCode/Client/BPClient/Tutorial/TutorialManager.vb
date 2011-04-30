Option Strict On

Public Class NewTutorialManager
	Public Enum StartupTriggerType As Integer
		eNoTrigger = 0
		eClickOnSelect = 1						'pass in what item is being selected
		eDeselectSelection = 2
		eControlEToSelectEngineer = 3
		eEngineerBuildWindowOpen = 4
		eSelectItemInBuildWindow = 5			'we will pass in what item was selected
		eBuildGhostNonBuildable = 6
		eBuildOrderSent = 7						'pass in what item is being built
		eLeftClickOnMinimap = 8
		eNotificationAlertClick = 9
		eColonyNameChanged = 10
		eTaxRateSet = 11
		eControlUSelection = 12					'passes in what was selected
		eControlSPressed = 13
		eEngagedEnemy = 14
		eIncrementPirateTurretDeath = 15		'indicates a pirate turret is destroyed
		eIncrementTankDeath = 16				'indicates a player's tank is destroyed
		eUnitBuilt = 17
		eControlISelect = 18					'passes in what was selected
		eFacilityBuildWindowOpen = 19			'passes in what facility is selected
		eMineralDepositHover = 20
		eWindowOpened = 21						'passes in the name of the window opened
		eWindowClosed = 22						'passes in the name of the window closed
		eRoutePlusClick = 23
		eRouteSelectDest = 24					'passes in what was selected (mining facility, refinery, point, etc...)
		eBeginRouteClicked = 25
		eFacilityBuilt = 26						'passes in what was built (type)
		'eResearchMineralButton = 27			'Use WindowOpen event instead
		eDiscoverMineralButton = 27
		eResearchCompleted = 28					'passes in what was researched
		eStudyMineralButton = 29
		eStudyMineralListSelected = 30			'passes in what mineral was selected
		eResearchWindowFilterSelected = 31		'passes in the new filter value
		eViewResultsClicked = 32				'passes in what is being viewed (the tech that view results is viewing)
		'eResearchScreenCancelClicked = 34		'Use WindowClosed event instead
		'eResearchComponentDesignButton = 35	'Use WindowOpen event instead
		eSubmitDesignClick = 33					'passes in what type of component was being design (objecttype)
		'eResearchButtonClicked = 37			'Use BuildOrderSent instead
		'eResearchSpecialsButtonClick = 38		'Use WindowOpen event instead
		eSpecialProjectSelected = 34			'passes in the GUID of special tech
		eIdleFacilitySelected = 35				'passes in the production type of the selection
		eHullBuilderScrollTypeSelected = 36		'passes in the TypeID
		eHullBuilderScrollSubTypeSelected = 37	'passes in the SubTypeID
		eHullBuilderScrollFrameSelected = 38	'passes in the ModelID
		eHullBuilderHullSizeChange = 39			'passes in the new hull size
		eHullBuilderMaterialSelected = 40		'passes in the new material selected
		eHullBuilderHitpointsChange = 41		'passes in the new hitpoints
		eHullBuilderSlotPaletteSelected = 42	'passes in the slot config value selected
		eHullBuilderSlotConfigChange = 43		'passes in the slot config value, slot group and the total slots defined
		eHullBuilderWeaponGroupChange = 44		'passes in the new group number

		ePrototypeNameChange = 45				'passes in the new name
		ePrototypeComponentSelected = 46		'passes in the Component GUID selected
		ePrototypeAllArmorSliderMax = 47
		ePrototypeWeaponTypeSelected = 48		'passes in the weapon type selected
		ePrototypeWeaponSelected = 49			'passes in the weapon selected
		ePrototypePlaceButtonClick = 50
		ePrototypeWeaponPlaced = 51				'passes in the group and the weapon

		eEnlistedExceeds20 = 52
		eInsufficientResources = 53
		eAvailableResourcesRefreshClick = 54

		ePirateMediumPowerKilled = 55
		eAgentAvailable = 56
		'eViewAgentClicked = 62					'Use WindowOpen event instead
		eRecruitButtonClicked = 57				'includes confirm
		'eCreateMissionClicked = 64				'Use WindowOpen event instead
		eCreateMissionMissionSelected = 58		'passes in the mission id
		eCreateMissionMethodSelected = 59		'passes in the methodid
		eCreateMissionGoalSelected = 60			'passes in the goal selected
		eCreateMissionSkillsetSelected = 61		'passes in the skillset selected
		eCreateMissionSelectAgent = 62			'passes in the agent id
		eCreateMissionSkillsetAssign = 63		'passes in the agent id and the skillsetid and the skillid
		eCreateMissionExecuteNow = 64
		eAgentInfiltrationTypeSelected = 65		'passes in the type of infiltration selected
		eAgentInfiltrationTargetSelected = 66	'passes in the target for infiltration selected
		eAgentInfiltrationSetSettings = 67
		eAgentInfiltrated = 68
		eMissionPhaseChange = 69				'passes in missionid, PhaseID
		eControlGroupAssigned = 70				'passes in the control group used
		eControlGroupSelected = 71				'passes in the control group selected
		eMissionCompleted = 72					'passes in the missionid
		eEmailSelected = 73
		eEmailWaypointJumpedTo = 74
		eIdleEngineerSelected = 75				'passes in the guid, prodtype

		eTutorialStepNextClick = 76

		eProductionCompleted = 77				'passes in the def guid and prod type

		eMoveCommandGiven = 78
		eRotateCamera = 79
		eScrollCamera = 80
		eZoomCamera = 81

		eConstructionStarted = 82

		eAdvanceDisplayPowerClick = 83

		eCreateMissionTargetSelected = 84
		ePirateBaseKilled = 85
		eBudgetItemSelected = 86
		eDeathBudgetDeposit = 87
		eViewChanged = 88
		eEnvironmentChanged = 89
        eRelationSet = 90
        eTechBuilderCostValueChange = 91
        eFirstChatMsgSent = 92
        eMiningBidAmountChange = 93
        eMiningBidQuantityChange = 94
        eMiningBidSubmitButton = 95

	End Enum
	Public Class ScriptStep
		Public StepID As Int32
		Public StepTitle As String = ""
		Public WAVFile As String = ""
		Public DisplayText As String = ""
		Public DisplayLocX As Int32 = -1
		Public DisplayLocY As Int32 = -1
		Public GroupID As Int32 = -1
		Public AlertControl As String = ""
        Public StepType As Int32 = 0

        Public lImStuckDelay As Int32 = -1

		Public PreqStepID As Int32 = -1
		Private moPreqStep As ScriptStep = Nothing
		Public ReadOnly Property oPreqStep() As ScriptStep
			Get
				If moPreqStep Is Nothing Then
					If moSteps Is Nothing = False Then
						For X As Int32 = 0 To moSteps.GetUpperBound(0)
							If moSteps(X).StepID = PreqStepID Then
								moPreqStep = moSteps(X)
								Exit For
							End If
						Next X
					End If
				End If
				Return moPreqStep
			End Get
		End Property

		Public sUIControl As String = "UIDisable:ALL" & vbCrLf & "HotkeyDisable:ALL"

		Public lRequirementFireStepID As Int32 = -1
		Public lRequirementInitializeStepID As Int32 = -1
		Public oRequirements() As Requirement

		Public oInitializeStepsReqs() As ScriptStep = Nothing

		Public bStepExecuted As Boolean = False
        Public bStepFinished As Boolean = False

        Public lBeginStepGroupNum As Int32 = -1
 
        Public Sub AddRequirement(ByRef oRequirement As Requirement)
            oRequirement.oFiresStep = Me
            If oRequirements Is Nothing Then ReDim oRequirements(-1)
            ReDim Preserve oRequirements(oRequirements.GetUpperBound(0) + 1)
            oRequirements(oRequirements.GetUpperBound(0)) = oRequirement
        End Sub

		Public Function CheckRequirements() As Boolean
			If oRequirements Is Nothing = False Then
				For X As Int32 = 0 To oRequirements.GetUpperBound(0)
					If oRequirements(X) Is Nothing = False AndAlso oRequirements(X).bRequirementMet = False Then
						Return False
					End If
				Next X
            End If
            mbNextButtonClickable = True
			Return True
		End Function

	End Class
	Public Class Requirement
		Public lTriggerType As StartupTriggerType = StartupTriggerType.eNoTrigger
		Public lExtended1 As Int32 = -1
		Public lExtended2 As Int32 = -1
		Public lExtended3 As Int32 = -1
		Public sExtendedText As String = ""
		Public lRequirementCount As Int32 = 0

		Public bRequirementMet As Boolean = False

		Public oFiresStep As ScriptStep = Nothing

		Public ReadOnly Property sDisplay() As String
			Get
				Dim sTemp As String = CType(lTriggerType, StartupTriggerType).ToString
				sTemp = sTemp.Substring(sTemp.LastIndexOf("."c) + 2)
				Return sTemp
			End Get
		End Property
	End Class
	Private Structure StoredGUIDs
		Public sVariableName As String
		Public lID As Int32
		Public iTypeID As Int16
	End Structure

	Private Shared moSteps() As ScriptStep

	Private Shared muStoredGUIDs(-1) As StoredGUIDs

    Private Shared msStoredCommands(-1) As String

	Public Shared TutorialOn As Boolean = False 

	Private Shared oRequirementsInWaiting(-1) As Requirement

    Private Shared moCurrentStep As ScriptStep = Nothing

    Public Shared mbNextButtonClickable As Boolean = False

    Public Shared bStepHasCenterScreen As Boolean = False

    Public Shared lRouteEngineerID As Int32 = -1

    Public Shared lStepGroupNumber As Int32 = 0
    Public Shared lStepGroupNumberStartedOnStepID As Int32 = 0

    Public Shared lLastYouMayLeaveAlert As Int32 = 0

    Public Shared Sub ExecuteCenterScreenCmd()
        If moCurrentStep Is Nothing = False AndAlso moCurrentStep.sUIControl Is Nothing = False Then
            Dim sCommands() As String = Split(moCurrentStep.sUIControl, vbCrLf)
            If sCommands Is Nothing = False Then
                For X As Int32 = 0 To sCommands.GetUpperBound(0)
                    Dim sCurCmd As String = sCommands(X)
                    If sCurCmd.IndexOf(":") > -1 Then
                        Dim sValPair() As String = Split(sCurCmd, ":")
                        sValPair(0) = sValPair(0).Trim.ToUpper
                        If sValPair(0) = "CENTERSCREEN" Then
                            CenterScreenCmd(sValPair(1))
                            Exit For
                        End If
                    End If
                Next X
            End If
             
        End If
    End Sub

	Public Shared Function GetTutorialStepID() As Int32
		If moCurrentStep Is Nothing = False Then Return moCurrentStep.StepID Else Return -1
	End Function

    Public Shared Sub InitializeRequirement(ByRef oRequirement As Requirement)
        If oRequirement.lTriggerType = StartupTriggerType.eTutorialStepNextClick Then mbNextButtonClickable = True
        For X As Int32 = 0 To oRequirementsInWaiting.GetUpperBound(0)
            If oRequirementsInWaiting(X) Is Nothing Then
                oRequirementsInWaiting(X) = oRequirement
                Return
            End If
        Next X
        ReDim Preserve oRequirementsInWaiting(oRequirementsInWaiting.GetUpperBound(0) + 1)
        oRequirementsInWaiting(oRequirementsInWaiting.GetUpperBound(0)) = oRequirement
    End Sub

	Private Shared mbScriptLoaded As Boolean = False
	Public Shared Function LoadScript() As Boolean
		If mbScriptLoaded = True Then Return True
		moSteps = Nothing

		'Dim oAssembly As System.Reflection.Assembly
		'oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
		'Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.TutorialScript.dat")
		Dim oStream As IO.Stream = goResMgr.UnpackAndGetDataFile("tut.pak", "TutScript.dat")
		If oStream Is Nothing Then Return False
		Dim oReader As New IO.StreamReader(oStream)

		Dim bInStep As Boolean = False
		Dim bInRequirement As Boolean = False
		Dim bInUIControl As Boolean = False
		Dim lLineNum As Int32 = 0
		Dim oCurStep As ScriptStep = Nothing
		Dim oCurReq As Requirement = Nothing

		Dim bResult As Boolean = True

		While oReader.EndOfStream = False
			Dim sLine As String = oReader.ReadLine().Trim
			lLineNum += 1
			If sLine = "" Then Continue While
			Dim sCompare As String = sLine.ToUpper

			If sCompare.ToUpper.StartsWith("BEGIN STEP") = True Then
				If bInStep = True Then
					If goUILib Is Nothing = False Then
						goUILib.AddNotification("ERROR: Invalid Begin Step at " & lLineNum & ". Already inside of a step.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					End If
					bResult = False
					Exit While
				Else
					oCurStep = New ScriptStep()
					bInStep = True
				End If
			Else
				If bInStep = False Then
					'ignore it...
				Else
					'ok, we are in a step
					If sCompare.StartsWith("BEGIN REQUIREMENT") = True Then
						'begin requirement
						If bInRequirement = True Then
							If goUILib Is Nothing = False Then
								goUILib.AddNotification("ERROR: Invalid Begin Requirement at " & lLineNum & ". Already inside of a requirement.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
							End If
							bResult = False
							Exit While
						Else
							oCurReq = New Requirement()
							bInRequirement = True
						End If
					ElseIf sCompare.StartsWith("END REQUIREMENT") = True Then
						'end requirement
						If bInRequirement = True Then
							oCurStep.AddRequirement(oCurReq)
							oCurReq = Nothing
							bInRequirement = False
						Else
							'ignore it
						End If
					ElseIf sCompare.StartsWith("END STEP") = True Then
						'end step
						If moSteps Is Nothing Then ReDim moSteps(-1)
						ReDim Preserve moSteps(moSteps.GetUpperBound(0) + 1)
						moSteps(moSteps.GetUpperBound(0)) = oCurStep

						oCurStep = Nothing
						bInStep = False
					ElseIf sCompare.StartsWith("BEGIN UICONTROL") = True Then
						If bInUIControl = True Then
							If goUILib Is Nothing = False Then
								goUILib.AddNotification("ERROR: Begin UIControl found while inside of Begin UIControl. Line: " & lLineNum, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
							End If
							bResult = False
							Exit While
						Else
							oCurStep.sUIControl = ""
							bInUIControl = True
						End If
					ElseIf sCompare.StartsWith("END UICONTROL") = True Then
						If bInUIControl = True Then
							bInUIControl = False
						Else
							'ignore it
						End If
					Else
						If bInUIControl = True Then
							If oCurStep.sUIControl <> "" Then oCurStep.sUIControl &= vbCrLf
							oCurStep.sUIControl &= sLine
							Continue While
						End If
						'Ok, likely a property value pair
						If sCompare.Contains("=") = False Then
							'ignore it
						Else
							'ok, property value pair
							Dim sProperty As String = sLine.Substring(0, sLine.IndexOf("="c))
							Dim sValue As String = sLine.Substring(sLine.IndexOf("="c) + 1)

							'Now, place the value
							If bInRequirement = True Then
								Select Case sProperty.ToUpper.Trim
									Case "EXTENDEDTEXT"
										oCurReq.sExtendedText = sValue
									Case "TRIGGERTYPE"
										oCurReq.lTriggerType = CType(CInt(Val(sValue)), StartupTriggerType)
									Case "COUNT"
										oCurReq.lRequirementCount = CInt(Val(sValue))
									Case "EXTENDED1"
										oCurReq.lExtended1 = CInt(Val(sValue))
									Case "EXTENDED2"
										oCurReq.lExtended2 = CInt(Val(sValue))
									Case "EXTENDED3"
										oCurReq.lExtended3 = CInt(Val(sValue))
									Case Else
										'ignore it
										'ofrm.AddEventLine("WARNING: Value Property not recognized for requirement on line " & lLineNum & ". " & sLine)
								End Select
							ElseIf bInStep = True Then
								Select Case sProperty.ToUpper.Trim
									Case "WAVFILE"
										oCurStep.WAVFile = sValue
									Case "STEPTYPE"
										oCurStep.StepType = CInt(Val(sValue))
									Case "STEPTITLE"
										oCurStep.StepTitle = sValue
									Case "STEPID"
										oCurStep.StepID = CInt(Val(sValue))
									Case "PREQSTEPID"
										oCurStep.PreqStepID = CInt(Val(sValue))
									Case "REQUIREMENTINITIALIZESTEPID"
										oCurStep.lRequirementInitializeStepID = CInt(Val(sValue))
									Case "REQUIREMENTFIRESTEPID"
										oCurStep.lRequirementFireStepID = CInt(Val(sValue))
									Case "GROUPID"
										oCurStep.GroupID = CInt(Val(sValue))
									Case "DISPLAYTEXT"
										oCurStep.DisplayText = sValue.Replace("[VBCRLF]", vbCrLf)
									Case "DISPLAYLOCX"
										oCurStep.DisplayLocX = CInt(Val(sValue))
									Case "DISPLAYLOCY"
										oCurStep.DisplayLocY = CInt(Val(sValue))
									Case "ALERTCONTROL"
                                        oCurStep.AlertControl = sValue
                                    Case "BEGINSTEPGROUPNUM"
                                        oCurStep.lBeginStepGroupNum = CInt(Val(sValue))
                                    Case "IMSTUCKTIMER"
                                        oCurStep.lImStuckDelay = CInt(Val(sValue))
                                    Case Else
                                        'ignore it
                                        'ofrm.AddEventLine("WARNING: Value Property not recognized for step on line " & lLineNum & ". " & sLine)
                                End Select
							Else
								If goUILib Is Nothing = False Then
									goUILib.AddNotification("ERROR: Not in Requirement or Step but parsed property value on line " & lLineNum, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
								End If
								bResult = False
								Exit While
							End If
						End If
					End If

				End If
			End If
		End While

		oReader.Close()
		oReader.Dispose()
		oStream.Close()
		oStream.Dispose()

		'Now, associate our initialize reqs to the steps
		If moSteps Is Nothing = False Then
			For X As Int32 = 0 To moSteps.GetUpperBound(0)
				If moSteps(X) Is Nothing = False Then
					If moSteps(X).lRequirementInitializeStepID > 0 Then
						'Ok, this step's requirements get initialize in step id lRequirementInitializeStepID
						Dim lInitStepID As Int32 = moSteps(X).lRequirementInitializeStepID
						For Y As Int32 = 0 To moSteps.GetUpperBound(0)
							If moSteps(Y) Is Nothing = False AndAlso moSteps(Y).StepID = lInitStepID Then
								If moSteps(Y).oInitializeStepsReqs Is Nothing Then ReDim moSteps(Y).oInitializeStepsReqs(-1)
								ReDim Preserve moSteps(Y).oInitializeStepsReqs(moSteps(Y).oInitializeStepsReqs.GetUpperBound(0) + 1)
								moSteps(Y).oInitializeStepsReqs(moSteps(Y).oInitializeStepsReqs.GetUpperBound(0)) = moSteps(X)
								Exit For
							End If
						Next Y
					End If
				End If
			Next X
		End If

		mbScriptLoaded = True

		Return bResult

	End Function

    Private Shared mbInTriggerFired As Boolean = False
	Public Shared Sub TriggerFired(ByVal lTrigger As StartupTriggerType, ByVal lExt1 As Int32, ByVal lExt2 As Int32, ByVal lExt3 As Int32, ByVal sExtText As String)
        'If lTrigger <> StartupTriggerType.eWindowClosed Then Stop
        If mbInTriggerFired = True Then Return
        mbInTriggerFired = False
		Dim bCanExecuteStep As Boolean = True
		For X As Int32 = 0 To oRequirementsInWaiting.GetUpperBound(0)
            Dim oReq As Requirement = oRequirementsInWaiting(X)
            If oReq Is Nothing = False AndAlso oReq.lTriggerType = lTrigger Then

                'goUILib.AddNotification("Trigger Fire Match: " & lTrigger.ToString, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Debug.WriteLine("Trigger Fire Match: " & lTrigger.ToString)

                If oReq.lExtended1 = -1 OrElse oReq.lExtended1 = lExt1 Then
                    If oReq.lExtended2 = -1 OrElse oReq.lExtended2 = lExt2 Then
                        If oReq.lExtended3 = -1 OrElse oReq.lExtended3 = lExt3 Then
                            If oReq.sExtendedText = "" OrElse oReq.sExtendedText.ToUpper = sExtText.ToUpper Then
                                'ok, found it
                                If lTrigger = StartupTriggerType.eBuildOrderSent Then
                                    oReq.lRequirementCount -= Math.Max(1, lExt3)
                                Else : oReq.lRequirementCount -= 1
                                End If

                                'CMC - 07/25/08 - hard-code to fix missing tank/turret bug
                                ' if we are here, we'll simply count our items and make our requirement count appropriately (requirement count = number remaining)
                                If moCurrentStep Is Nothing OrElse moCurrentStep.StepID < 55 Then
                                    If lTrigger = StartupTriggerType.eIncrementPirateTurretDeath Then
                                        oReq.lRequirementCount = GetPirateTurretCount()
                                    ElseIf lTrigger = StartupTriggerType.eIncrementTankDeath Then
                                        oReq.lRequirementCount = GetPlayerTankCount()
                                    End If
                                End If

                                If oReq.lRequirementCount < 1 Then
                                    'ok, this requirement is met...
                                    oReq.bRequirementMet = True

                                    If oReq.oFiresStep Is Nothing = False Then
                                        If oReq.oFiresStep.CheckRequirements() = True Then
                                            'any preq step id?
                                            If oReq.oFiresStep.oPreqStep Is Nothing = False Then
                                                If oReq.oFiresStep.oPreqStep.bStepExecuted = False Then Continue For
                                                'If oReq.oFiresStep.oPreqStep.bStepFinished = False Then Continue For
                                            End If

                                            'Ok, we are here, that means we can execute the fires step
                                            If bCanExecuteStep = True Then ExecuteStep(oReq.oFiresStep)
                                            bCanExecuteStep = False
                                            'Exit For
                                        End If
                                    End If

                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next X
        mbInTriggerFired = False
    End Sub

    Private Shared Function GetPirateTurretCount() As Int32
        Dim lCurUB As Int32 = -1
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return 0

        Dim lCnt As Int32 = 0

        If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If oEnvir.lEntityIdx(X) > -1 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                If oEntity Is Nothing = False Then
                    If oEntity.OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID AndAlso oEntity.bObjectDestroyed = False Then
                        If oEntity.oUnitDef Is Nothing = False Then
                            'MSC: We only care about the mesh portion of the modelid
                            If (oEntity.oUnitDef.ModelID And 255) = 16 Then
                                lCnt += 1
                            End If
                        End If
                    End If
                End If
            End If
        Next X

        Return lCnt
    End Function
    Private Shared Function GetPlayerTankCount() As Int32
        Dim lCurUB As Int32 = -1
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return 0

        Dim lCnt As Int32 = 0

        If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If oEnvir.lEntityIdx(X) > -1 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.bObjectDestroyed = False Then
                    If oEntity.OwnerID = glPlayerID AndAlso oEntity.ObjTypeID = ObjectType.eUnit AndAlso oEntity.yProductionType = 0 Then
                        lCnt += 1
                    End If
                End If
            End If
        Next X

        Return lCnt
    End Function

	Private Shared Sub ExecuteStep(ByRef oStep As ScriptStep)
        moCurrentStep = oStep

        If oStep.StepID < 3 Then
            Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sFile.EndsWith("\") = False Then sFile &= "\"
            sFile &= goCurrentPlayer.PlayerName & ".tut"
            Dim oINI As New InitFile(sFile)
            oINI.WriteString("Tutorial", "FirstLiveMsg", "0")
            oINI = Nothing
        End If

        If oStep.StepID = 317 Then
            TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.EndsTutorialInAurelium)
        End If
        If oStep.StepID < 315 Then
            TutorialOn = True
        End If

		oStep.bStepExecuted = True
		Dim oWin As frmTutorialStep = CType(goUILib.GetWindow("frmTutorialStep"), frmTutorialStep)
		If oWin Is Nothing Then oWin = New frmTutorialStep(goUILib)
		oWin.Visible = True
		oWin.SetFromStepNew(oStep)
		oWin = Nothing

		moAlertControl = Nothing

		If oStep.AlertControl <> "" Then AlertControl(oStep.AlertControl)

		'Now, clear the requirements in the list that relate to this step
		For X As Int32 = 0 To oRequirementsInWaiting.GetUpperBound(0)
			If oRequirementsInWaiting(X) Is Nothing = False Then
				If Object.Equals(oRequirementsInWaiting(X).oFiresStep, oStep) = True Then
					oRequirementsInWaiting(X) = Nothing
				End If
			End If
		Next X

        'CSAJ 8/20/08 We are on a new step and have yet to upload the requirements so mark next button unclickable if this is wrong the added requirements will fix it
        mbNextButtonClickable = False
        'If moCurrentStep.oRequirements Is Nothing = True Then mbNextButtonClickable = True
        If oStep.oInitializeStepsReqs Is Nothing = True Then
            mbNextButtonClickable = True
            If oRequirementsInWaiting Is Nothing = False Then
                For X As Int32 = 0 To oRequirementsInWaiting.GetUpperBound(0)
                    If oRequirementsInWaiting(X) Is Nothing = False AndAlso oRequirementsInWaiting(X).oFiresStep.StepID = oStep.StepID + 1 AndAlso oRequirementsInWaiting(X).lTriggerType <> StartupTriggerType.eTutorialStepNextClick Then
                        mbNextButtonClickable = False
                        Exit For
                    End If
                Next X
            End If
        Else
            mbNextButtonClickable = True
            For X As Int32 = 0 To oStep.oInitializeStepsReqs.GetUpperBound(0)
                If oStep.oInitializeStepsReqs(X).StepID = oStep.StepID + 1 Then
                    mbNextButtonClickable = False
                    Exit For
                End If
            Next X
        End If
        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then mbNextButtonClickable = True

        'Now, initialize any requirements that this step is requested to initialize
        If oStep.oInitializeStepsReqs Is Nothing = False Then
            For X As Int32 = 0 To oStep.oInitializeStepsReqs.GetUpperBound(0)
                Dim oInitStep As ScriptStep = oStep.oInitializeStepsReqs(X)
                If oInitStep Is Nothing = False AndAlso oInitStep.oRequirements Is Nothing = False Then
                    For Y As Int32 = 0 To oInitStep.oRequirements.GetUpperBound(0)
                        InitializeRequirement(oInitStep.oRequirements(Y))
                    Next Y
                End If
            Next X
        End If

        'Do our disables now
        ProcessStepDisableCommands(oStep)

        'Append stored commands to current step commands
        Dim sUIControl As String = oStep.sUIControl + GetStoredCommands()

        'Store the user interface/hotkey commands for this step
        StoreStepEnableCommands(oStep)

        'If oStep.StepID = 314 Then
        '    frmEnvirDisplay.dtEndTimer = Now.Add(New TimeSpan(4, 0, 0))
        'End If

        bStepHasCenterScreen = False

        'Finally, send our commands for locking down the user interface/hotkeys
        If goUILib Is Nothing = False AndAlso sUIControl <> "" Then
            Dim sCommands() As String = Split(sUIControl, vbCrLf)
            For X As Int32 = 0 To sCommands.GetUpperBound(0)
                Dim sCurCmd As String = sCommands(X)
                If sCurCmd.IndexOf(":") > -1 Then
                    Dim sValPair() As String = Split(sCurCmd, ":")
                    sValPair(0) = sValPair(0).Trim.ToUpper
                    If sValPair(0) = "UIENABLE" Then
                        goUILib.HandleEnableCmd(sValPair(1), True)
                    ElseIf sValPair(0) = "UIDISABLE" Then
                        goUILib.HandleDisableCmd(sValPair(1), True)
                    ElseIf sValPair(0) = "HOTKEYENABLE" Then
                        goUILib.HandleEnableCmd(sValPair(1), False)
                    ElseIf sValPair(0) = "HOTKEYDISABLE" Then
                        goUILib.HandleDisableCmd(sValPair(1), False)
                    ElseIf sValPair(0) = "CENTERSCREEN" Then
                        bStepHasCenterScreen = True
                        CenterScreenCmd(sValPair(1))
                    ElseIf sValPair(0) = "SELECTALL" Then
                        SelectAllCmd(sValPair(1))
                    ElseIf sValPair(0) = "SETZOOMLEVEL" Then
                        SetZoomLevel(CInt(Val(sValPair(1))))
                    ElseIf sValPair(0) = "CLOSEWINDOW" Then
                        Dim sTemp As String = Split(sCurCmd, ":")(1)
                        goUILib.RemoveWindow(sTemp)
                    ElseIf sValPair(0) = "STOREGUIDOFNEXTADD" Then
                        Dim sVars() As String = Split(sValPair(1), ",")
                        If sVars Is Nothing = False AndAlso sVars.GetUpperBound(0) = 1 Then
                            sStoreNextGUIDName = sVars(0).Trim.ToUpper
                            iStoreNextGUIDType = CShort(sVars(1))
                        End If
                    ElseIf sValPair(0) = "ENSURESELECTION" Then
                        EnsureSelected(sValPair(1))
                        'ElseIf sValPair(0) = "BEGINSTEPGROUPNUM" Then
                        '   lStepGroupNumber = CInt(Val(sValPair(1)))
                        '  lStepGroupNumberStartedOnStepID = oStep.StepID
                    ElseIf sValPair(0) = "ADDNOTIFICATION" Then
                        '<This is a test notification from your engineer><CLICKTO:MyEngineer>
                        If goUILib Is Nothing = False Then
                            'Ok, there are possibly 2 params... first is the text, the second is a click to command
                            Dim lIdx As Int32 = sValPair(1).IndexOf(">")
                            Dim sFirst As String = sValPair(1)
                            Dim sSecond As String = ""
                            If lIdx > -1 Then
                                sFirst = sValPair(1).Substring(0, lIdx)
                                sSecond = sValPair(1).Substring(lIdx)
                            End If
                            If sValPair.GetUpperBound(0) = 2 Then sSecond = sValPair(2)
                            sFirst = sFirst.Replace("<", "").Replace(">", "")
                            sSecond = sSecond.Replace("<", "").Replace(">", "")

                            If sSecond.ToUpper.StartsWith("CLICKTO:") = True Then sSecond = sSecond.Substring(8)

                            'Now, get our lookup object
                            Dim oObj As Base_GUID = GetLookupObject(sSecond)
                            Dim lID As Int32 = -1
                            Dim iTypeID As Int16 = -1
                            Dim lPID As Int32 = -1
                            Dim iPTypeID As Int16 = -1
                            If oObj Is Nothing = False Then
                                lID = oObj.ObjectID : iTypeID = oObj.ObjTypeID : lPID = goCurrentEnvir.ObjectID : iPTypeID = goCurrentEnvir.ObjTypeID
                            End If
                            goUILib.AddNotification(sFirst, Color.White, lPID, iPTypeID, lID, iTypeID, Int32.MinValue, Int32.MinValue)
                        End If
                    ElseIf sValPair(0).StartsWith("<SEND_TUTORIAL_PRODUCTION_FINISH") = True Then
                        sCurCmd = sCurCmd.Substring(33).Replace(">", "")
                        'ok, we should now have ID and a typeID... OR a variable name
                        Dim lID As Int32 = -1
                        Dim iTypeID As Int16 = -1
                        If sCurCmd.Contains(",") = True Then
                            Dim sVals() As String = Split(sCurCmd, ",")
                            If sVals.GetUpperBound(0) >= 1 Then
                                lID = CInt(sVals(0))
                                iTypeID = CShort(sVals(1))
                            End If
                        Else
                            FillStoredGUIDValues(sCurCmd, lID, iTypeID)
                            If lID = -1 AndAlso iTypeID = -1 Then
                                goUILib.AddNotification("Unknown token " & sCurCmd.Trim & ". Step " & oStep.StepID, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                Continue For
                            End If
                        End If

                        Dim yMsg(11) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eTutorialProdFinish).CopyTo(yMsg, 0)
                        System.BitConverter.GetBytes(lID).CopyTo(yMsg, 2)
                        System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, 6)
                        System.BitConverter.GetBytes(oStep.StepID).CopyTo(yMsg, 8)
                        goUILib.SendMsgToPrimary(yMsg)
                    End If
                ElseIf sCurCmd.ToUpper.Trim = "UIFORCENOFOCUS" Then
                    If goUILib.FocusedControl Is Nothing = False Then
                        goUILib.FocusedControl.HasFocus = False
                        Dim oTmpCtrl As UIControl = goUILib.FocusedControl
                        While oTmpCtrl.ParentControl Is Nothing = False
                            oTmpCtrl.ParentControl.HasFocus = False
                            oTmpCtrl = oTmpCtrl.ParentControl
                        End While
                        goUILib.FocusedControl = Nothing
                    End If
                ElseIf sCurCmd.ToUpper.Trim = "UPDATESTEPID" Then
                    Dim yUpdateMsg(5) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerTutorialStep).CopyTo(yUpdateMsg, 0)
                    System.BitConverter.GetBytes(oStep.StepID).CopyTo(yUpdateMsg, 2)
                    If goUILib Is Nothing = False Then goUILib.SendMsgToPrimary(yUpdateMsg)
                ElseIf sCurCmd.ToUpper.Trim = "DESELECTALL" Then
                    If goCurrentEnvir Is Nothing = False Then goCurrentEnvir.DeselectAll()
                End If
            Next X
        End If

    End Sub

	Public Shared moAlertControl As UIControl = Nothing
	Private Shared mlAlertCount As Int32
	Private Shared mlLastAlertUpdate As Int32
	Private Shared mbRepeatAlertControl As Boolean = False
	Private Shared msAlertControlName As String
	Private Shared mbAlertSounded As Boolean = False

	Public Shared Sub AlertControl(ByVal sControlName As String)
		If sControlName = "" Then Return
		msAlertControlName = sControlName
		mbAlertSounded = False

		If moAlertControl Is Nothing = False Then moAlertControl.Visible = True
		moAlertControl = Nothing

		Dim sPath() As String = Split(sControlName, ".")
		If sPath Is Nothing OrElse sPath.GetUpperBound(0) < 0 Then Return

		mbRepeatAlertControl = True

		Dim sWindow As String = sPath(0)
		Dim ofrmWindow As UIWindow = goUILib.GetWindow(sWindow)
		Dim oCurrent As UIControl = ofrmWindow

		If ofrmWindow Is Nothing Then Return

		Dim bFound As Boolean = False
		For X As Int32 = 1 To sPath.GetUpperBound(0)
			Dim lIdx As Int32 = -1
			For Y As Int32 = 0 To oCurrent.ChildrenUB
				If oCurrent.moChildren(Y).ControlName = sPath(X) Then
					lIdx = Y
					Exit For
				End If
			Next Y
			If lIdx = -1 Then Exit For

			oCurrent = oCurrent.moChildren(lIdx)
			If X = sPath.GetUpperBound(0) Then bFound = True
		Next X

		If bFound = False Then moAlertControl = ofrmWindow Else moAlertControl = oCurrent

		mlAlertCount = 7
		mlLastAlertUpdate = glCurrentCycle
		mbRepeatAlertControl = False
	End Sub

	Public Shared Sub HandleAlertControlUpdate()
		If moAlertControl Is Nothing = False Then
			If mbAlertSounded = False Then
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			End If
			mbAlertSounded = True
			mbRepeatAlertControl = False

			'If glCurrentCycle - mlLastAlertUpdate > 3 Then
			'	moAlertControl.Visible = Not moAlertControl.Visible
			'	moAlertControl.IsDirty = False
			'	moAlertControl.IsDirty = True
			'	mlAlertCount -= 1
			'	mlLastAlertUpdate = glCurrentCycle
			'	If mlAlertCount < 1 Then
			'		moAlertControl.Visible = False
			'		moAlertControl.Visible = True
			'		moAlertControl.IsDirty = False
			'		moAlertControl.IsDirty = True
			'		'moAlertControl = Nothing
			'		mlLastAlertUpdate = Int32.MaxValue
			'	End If
			'End If

		ElseIf mbRepeatAlertControl = True Then
			AlertControl(msAlertControlName)
		End If
	End Sub

	Private Shared Sub SetZoomLevel(ByVal lZoom As Int32)
		If goCamera Is Nothing = False Then
			If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
				If goCurrentEnvir.oGeoObject Is Nothing = False Then
					Dim lHt As Int32 = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraX, goCamera.mlCameraZ, True))
					lZoom += lHt + 1
				End If
			End If
			muSettings.lPlanetViewCameraY = lZoom
			goCamera.mlCameraY = lZoom
		End If
	End Sub

	Private Shared Sub EnsureSelected(ByVal sVal As String)
		Dim oObj As BaseEntity = CType(GetLookupObject(sVal), BaseEntity)
		If oObj Is Nothing = False Then
			If oObj.bSelected = True Then Return
			goCurrentEnvir.DeselectAll()
			oObj.bSelected = True
			goUILib.AddSelection(oObj.lEnvirEntityIdx)
		End If
	End Sub

	Private Shared Sub CenterScreenCmd(ByVal sVal As String)
		Dim oObj As Base_GUID = GetLookupObject(sVal)
		If oObj Is Nothing = False Then
			If goCamera Is Nothing = False Then

				Dim lX As Int32 = 0
				Dim lZ As Int32 = 0
				Dim iA As Int16 = 0
				Select Case oObj.ObjTypeID
					Case ObjectType.eUnit, ObjectType.eFacility
						With CType(oObj, BaseEntity)
							lX = CInt(.LocX) : lZ = CInt(.LocZ) : iA = CShort(.LocAngle / 10)
						End With
					Case ObjectType.eMineralCache
						With CType(oObj, MineralCache)
							lX = .LocX : lZ = .LocZ : iA = 0
						End With
				End Select
				goCamera.ZoomToPosition(lX, lZ, iA)
				muSettings.lPlanetViewCameraY = 7500
				goCamera.mlCameraY = goCamera.mlCameraAtY + 7500

				'goCamera.mlCameraX = lX
				'goCamera.mlCameraAtX = goCamera.mlCameraX
				'goCamera.mlCameraAtZ = lZ
				'goCamera.mlCameraZ = goCamera.mlCameraAtZ - 500

				'If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
				'	'goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ)) + 3000
				'	goCamera.mlCameraAtY = 0

				'	goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

				'	goCamera.mlCameraX += muSettings.lPlanetViewCameraX
				'	goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
				'End If


			End If
		End If
	End Sub

	Private Shared Sub SelectAllCmd(ByVal sVal As String)
		Select Case sVal.Trim.ToUpper
			Case "MYTANKS"
				If goCurrentEnvir Is Nothing = False Then
					goCurrentEnvir.DeselectAll()
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) <> -1 Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
							If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eUnit AndAlso oEntity.yProductionType = 0 AndAlso oEntity.OwnerID = glPlayerID Then
								oEntity.bSelected = True
								goUILib.AddSelection(X)
							End If
						End If
					Next X
				End If
		End Select
	End Sub

	Private Shared Function GetLookupObject(ByVal sVal As String) As Base_GUID
		Select Case sVal.Trim.ToUpper
			Case "MYFACTORY"
				If goCurrentEnvir Is Nothing = False Then
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) <> -1 Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
							If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.yProductionType = ProductionType.eProduction Then
								Return oEntity
							End If
						End If
					Next X
				End If
			Case "MYRESEARCHLAB"
				If goCurrentEnvir Is Nothing = False Then
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) <> -1 Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
							If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.yProductionType = ProductionType.eResearch Then
								Return oEntity
							End If
						End If
					Next X
				End If
			Case "PIRATETURRETS"
				If goCurrentEnvir Is Nothing = False Then
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) <> -1 Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
                            If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID AndAlso oEntity.oUnitDef Is Nothing = False AndAlso (oEntity.oUnitDef.ModelID And 255S) = 16 Then
                                Return oEntity
                            End If
						End If
					Next X
				End If
			Case "PIRATEMEDIUMFACILITY"
				If goCurrentEnvir Is Nothing = False Then
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) <> -1 Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
                            If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID AndAlso oEntity.oUnitDef Is Nothing = False AndAlso (oEntity.oUnitDef.ModelID And 255S) = 12 Then
                                Return oEntity
                            End If
						End If
					Next X
				End If
			Case "MYCOMMANDCENTER"
				If goCurrentEnvir Is Nothing = False Then
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) <> -1 Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
							If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.yProductionType = ProductionType.eCommandCenterSpecial Then
								Return oEntity
							End If
						End If
					Next X
				End If
			Case "ENOCHINE"
				If goCurrentEnvir Is Nothing = False Then
					Dim lCurUB As Int32 = goCurrentEnvir.lCacheUB
					If goCurrentEnvir.lCacheIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lCacheIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lCacheIdx(X) <> -1 Then
							Dim oCache As MineralCache = goCurrentEnvir.oCache(X)
							If oCache Is Nothing = False AndAlso oCache.MineralID = 157 Then
								Return oCache
							End If
						End If
					Next X
				End If
			Case "MYENGINEER"
				If goCurrentEnvir Is Nothing = False Then
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) <> -1 Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
							If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eUnit AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.yProductionType = ProductionType.eFacility Then
								Return oEntity
							End If
						End If
					Next X
				End If
			Case "MYBARRACKS"
				If goCurrentEnvir Is Nothing = False Then
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) <> -1 Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
							If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.yProductionType = ProductionType.eEnlisted Then
								Return oEntity
							End If
						End If
					Next X
				End If
			Case "MYOFFICERS"
				If goCurrentEnvir Is Nothing = False Then
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) <> -1 Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
							If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.yProductionType = ProductionType.eOfficers Then
								Return oEntity
							End If
						End If
					Next X
				End If
			Case Else
				'ok, go through all of my stored guids
				Dim sTemp As String = sVal.Trim.ToUpper
				Dim lID As Int32 = -1
				Dim iTypeID As Int16 = -1
				For X As Int32 = 0 To muStoredGUIDs.GetUpperBound(0)
					If muStoredGUIDs(X).sVariableName.Trim.ToUpper = sTemp Then
						lID = muStoredGUIDs(X).lID
						iTypeID = muStoredGUIDs(X).iTypeID
						Exit For
					End If
				Next X
				If lID <> -1 AndAlso iTypeID <> -1 Then
					Dim lCurUB As Int32 = goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx Is Nothing Then lCurUB = -1 Else lCurUB = Math.Min(lCurUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						If goCurrentEnvir.lEntityIdx(X) = lID Then
							Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
							If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iTypeID Then
								Return oEntity
							End If
						End If
					Next X
				End If

		End Select
		Return Nothing
	End Function

	Public Shared iStoreNextGUIDType As Int16 = -1
	Private Shared sStoreNextGUIDName As String = ""

	Public Shared Sub StoreGUID(ByVal lID As Int32, ByVal iTypeID As Int16)
        If iStoreNextGUIDType = iTypeID Then
            iStoreNextGUIDType = -1

            Dim uNew As StoredGUIDs
            With uNew
                .sVariableName = sStoreNextGUIDName
                .lID = lID
                .iTypeID = iTypeID
            End With

            Dim bFoundGUID As Boolean = False
            Dim lFirstIdx As Int32 = -1
            For X As Int32 = 0 To muStoredGUIDs.GetUpperBound(0)
                If muStoredGUIDs(X).sVariableName = "" Then
                    lFirstIdx = X
                ElseIf muStoredGUIDs(X).sVariableName.ToUpper = uNew.sVariableName.ToUpper Then
                    muStoredGUIDs(X) = uNew
                    bFoundGUID = True
                    Exit For
                End If
            Next X

            If bFoundGUID = False Then
                If lFirstIdx = -1 Then
                    lFirstIdx = muStoredGUIDs.GetUpperBound(0) + 1
                    ReDim Preserve muStoredGUIDs(lFirstIdx)
                End If
                muStoredGUIDs(lFirstIdx) = uNew
            End If

            ' Write GUID to player's local .TUT file
            If goCurrentPlayer Is Nothing = False Then
                Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
                If sFile.EndsWith("\") = False Then sFile &= "\"
                sFile &= goCurrentPlayer.PlayerName & ".tut"

                Dim oINI As InitFile = New InitFile(sFile)

                Dim sGUIDName As String = uNew.sVariableName
                Dim sGUIDValue As String = uNew.iTypeID.ToString() + "," + uNew.lID.ToString()
                Dim sGUIDKeys As String = oINI.GetString("Tutorial", "GUIDKeys", "")
                Dim sGUIDExistingValue As String = oINI.GetString("Tutorial", sGUIDName, "")

                oINI.WriteString("Tutorial", "RouteEngineer", lRouteEngineerID.ToString)

                'If goUILib Is Nothing = False Then goUILib.AddNotification("StoreGUID: " & sGUIDName & " = " & sGUIDValue, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

                If sGUIDExistingValue.Length = 0 Then
                    If sGUIDKeys.Length = 0 Then
                        sGUIDKeys = sGUIDName
                    Else
                        sGUIDKeys += "," + sGUIDName
                    End If
                    oINI.WriteString("Tutorial", "GUIDKeys", sGUIDKeys)
                End If
                oINI.WriteString("Tutorial", sGUIDName, sGUIDValue)
            End If
        End If
	End Sub

	Public Shared Sub FillStoredGUIDValues(ByVal sVarName As String, ByRef lID As Int32, ByRef iTypeID As Int16)
		lID = -1 : iTypeID = -1

		sVarName = sVarName.Trim.ToUpper
		If muStoredGUIDs Is Nothing = False Then
			For Y As Int32 = 0 To muStoredGUIDs.GetUpperBound(0)
				If muStoredGUIDs(Y).sVariableName.Trim.ToUpper = sVarName Then
					lID = muStoredGUIDs(Y).lID
					iTypeID = muStoredGUIDs(Y).iTypeID
					Exit For
				End If
			Next Y
		End If
	End Sub

	Public Shared Function TutorialStepFinished() As Boolean
		If moCurrentStep Is Nothing = False Then
			moCurrentStep.bStepFinished = True
		End If

		If oRequirementsInWaiting Is Nothing OrElse oRequirementsInWaiting.GetUpperBound(0) = -1 Then
			If moSteps Is Nothing = False Then
				For X As Int32 = 0 To moSteps.GetUpperBound(0)
					If moSteps(X) Is Nothing = False AndAlso moSteps(X).StepID = 1 Then
						If moSteps(X).bStepExecuted = False Then
							ExecuteStep(moSteps(X))
							Return True
						End If
						Exit For
					End If
				Next
			End If
		End If

		Dim bFired As Boolean = False
		Dim bMissingReqs As Boolean = False
		Dim lMissingReqsFor() As Int32 = Nothing
		For X As Int32 = 0 To oRequirementsInWaiting.GetUpperBound(0)
			Dim oReq As Requirement = oRequirementsInWaiting(X)
			If oReq Is Nothing Then Continue For
			If oReq.lRequirementCount < 1 Then
				'ok, this requirement is met...
				oReq.bRequirementMet = True

				If oReq.oFiresStep Is Nothing = False Then
					If oReq.oFiresStep.CheckRequirements() = True Then
						'any preq step id?
						If oReq.oFiresStep.oPreqStep Is Nothing = False Then
							If oReq.oFiresStep.oPreqStep.bStepExecuted = False OrElse oReq.oFiresStep.oPreqStep.bStepFinished = False Then
								'bMissingReqs = True
								If lMissingReqsFor Is Nothing Then ReDim lMissingReqsFor(-1)
								ReDim Preserve lMissingReqsFor(lMissingReqsFor.GetUpperBound(0) + 1)
								lMissingReqsFor(lMissingReqsFor.GetUpperBound(0)) = oReq.oFiresStep.StepID
								Continue For
							End If
						End If

						'Ok, we are here, that means we can execute the fires step
						ExecuteStep(oReq.oFiresStep)
						bFired = True
						Exit For
					Else : bMissingReqs = True
					End If
				End If
			Else
				If lMissingReqsFor Is Nothing Then ReDim lMissingReqsFor(-1)
				ReDim Preserve lMissingReqsFor(lMissingReqsFor.GetUpperBound(0) + 1)
				lMissingReqsFor(lMissingReqsFor.GetUpperBound(0)) = oReq.oFiresStep.StepID
				'bMissingReqs = True
			End If
		Next X

		If bFired = False AndAlso bMissingReqs = False Then
			If moSteps Is Nothing = False Then
				For X As Int32 = 0 To moSteps.GetUpperBound(0)
					If moSteps(X) Is Nothing = False AndAlso moSteps(X).PreqStepID = moCurrentStep.StepID Then
						If lMissingReqsFor Is Nothing = False Then
							Dim bBad As Boolean = False
							For Y As Int32 = 0 To lMissingReqsFor.GetUpperBound(0)
								If lMissingReqsFor(Y) = moSteps(X).StepID Then
									bBad = True
									Exit For
								End If
							Next Y
							If bBad = True Then Continue For
						End If
						ExecuteStep(moSteps(X))
						Return True
						Exit For
					End If
				Next X
			End If
		End If
		Return False
	End Function

	Public Shared Sub FindAndExecuteStepID(ByVal lStepID As Int32)
		If moSteps Is Nothing = False Then
			For X As Int32 = 0 To moSteps.GetUpperBound(0)
				If moSteps(X) Is Nothing = False AndAlso moSteps(X).StepID = lStepID Then
					ExecuteStep(moSteps(X))
					Exit For
				End If
            Next X

            'Now, see if we missed the event for the tanks/turrets dying
            If moCurrentStep Is Nothing = False AndAlso moCurrentStep.StepID < 55 Then
                Dim bCanExecuteStep As Boolean = True
                For X As Int32 = 0 To oRequirementsInWaiting.GetUpperBound(0)
                    Dim oReq As Requirement = oRequirementsInWaiting(X)
                    If oReq Is Nothing = True Then Continue For

                    If oReq.lTriggerType = StartupTriggerType.eIncrementPirateTurretDeath Then
                        oReq.lRequirementCount = GetPirateTurretCount()
                    ElseIf oReq.lTriggerType = StartupTriggerType.eIncrementTankDeath Then
                        oReq.lRequirementCount = GetPlayerTankCount()
                    ElseIf oReq.lTriggerType = StartupTriggerType.eEngagedEnemy Then
                        Dim lTankCnt As Int32 = GetPlayerTankCount()
                        Dim lTurretCnt As Int32 = GetPirateTurretCount()
                        If lTankCnt = 0 OrElse lTurretCnt = 0 Then oReq.lRequirementCount = 0
                    ElseIf oReq.lTriggerType = StartupTriggerType.eConstructionStarted Then
                        If oReq.lExtended1 = 5 AndAlso oReq.lExtended2 = 11 Then
                            If GetPlayerCommandCenterCnt() > 0 Then
                                oReq.lRequirementCount = 0
                            End If
                        End If
                    End If

                    If oReq.lRequirementCount < 1 Then
                        'ok, this requirement is met...
                        oReq.bRequirementMet = True

                        If oReq.oFiresStep Is Nothing = False Then
                            If oReq.oFiresStep.CheckRequirements() = True Then
                                'any preq step id?
                                If oReq.oFiresStep.oPreqStep Is Nothing = False Then
                                    If oReq.oFiresStep.oPreqStep.bStepExecuted = False Then Continue For
                                    'If oReq.oFiresStep.oPreqStep.bStepFinished = False Then Continue For
                                End If

                                'Ok, we are here, that means we can execute the fires step
                                If bCanExecuteStep = True Then ExecuteStep(oReq.oFiresStep)
                                bCanExecuteStep = False
                                'Exit For
                            End If
                        End If

                    End If
                Next X
            End If

        End If
	End Sub

    Private Shared Function GetPlayerCommandCenterCnt() As Int32
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return 0
        Dim lCurUB As Int32 = -1
        If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))

        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lCurUB
            If oEnvir.lEntityIdx(X) > -1 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.yProductionType = ProductionType.eCommandCenterSpecial Then
                    lCnt += 1
                End If
            End If
        Next X
        Return lCnt
    End Function

	Public Shared Sub SetTutorialStep(ByVal lStepID As Int32)
			' Load stored GUIDs from player's local .TUT file
		If goCurrentPlayer Is Nothing = False Then
			Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
			If sFile.EndsWith("\") = False Then sFile &= "\"
			sFile &= goCurrentPlayer.PlayerName & ".tut"

			Dim oINI As InitFile = New InitFile(sFile)

			Dim sGUIDKeys As String = oINI.GetString("Tutorial", "GUIDKeys", "")

            lRouteEngineerID = CInt(oINI.GetString("Tutorial", "RouteEngineer", "-1"))

			Dim sGUIDNames() As String = sGUIDKeys.Split(CChar(","))

			Dim sGUIDName As String
			For Each sGUIDName In sGUIDNames
				Dim sGUIDValue As String = oINI.GetString("Tutorial", sGUIDName, "")
				If sGUIDValue.Length > 0 Then
					Dim sGUIDValues() As String = sGUIDValue.Split(CChar(","))
					If sGUIDValues.Length = 2 Then
						Dim uNew As StoredGUIDs
						With uNew
							.sVariableName = sGUIDName
							.iTypeID = CShort(sGUIDValues(0))
							.lID = CInt(sGUIDValues(1))
						End With
						ReDim Preserve muStoredGUIDs(muStoredGUIDs.GetUpperBound(0) + 1)
						muStoredGUIDs(muStoredGUIDs.GetUpperBound(0)) = uNew
					End If
				End If
			Next sGUIDName
		End If
	
        If moSteps Is Nothing Then LoadScript()
		If moSteps Is Nothing = False Then
			For X As Int32 = 0 To moSteps.GetUpperBound(0)
				If moSteps(X) Is Nothing = False Then
					If moSteps(X).StepID < lStepID Then
						moSteps(X).bStepFinished = True
                        moSteps(X).bStepExecuted = True
                        StoreStepEnableCommands(moSteps(X))
                    ElseIf moSteps(X).StepID = lStepID Then
                        moCurrentStep = moSteps(X)
                    End If
                End If
            Next X

            For X As Int32 = 0 To moSteps.GetUpperBound(0)
                If moSteps(X).bStepExecuted = True Then
                    If moSteps(X).oInitializeStepsReqs Is Nothing = False Then
                        For Z As Int32 = 0 To moSteps(X).oInitializeStepsReqs.GetUpperBound(0)
                            Dim oInitStep As ScriptStep = moSteps(X).oInitializeStepsReqs(Z)
                            If oInitStep Is Nothing = False AndAlso oInitStep.oRequirements Is Nothing = False Then
                                For Y As Int32 = 0 To oInitStep.oRequirements.GetUpperBound(0)
                                    If oInitStep.StepID > lStepID Then
                                        InitializeRequirement(oInitStep.oRequirements(Y))
                                    End If
                                Next Y
                            End If
                        Next Z
                    End If
                End If
            Next X


            ''Now, do a requirement check
            'Dim bCanExecuteStep As Boolean = True
            'If moCurrentStep Is Nothing OrElse moCurrentStep.StepID < 55 Then
            '    For X As Int32 = 0 To oRequirementsInWaiting.GetUpperBound(0)
            '        Dim oReq As Requirement = oRequirementsInWaiting(X)
            '        Dim bForceFire As Boolean = False
            '        If oReq.lTriggerType = StartupTriggerType.eIncrementPirateTurretDeath Then
            '            Dim lOrigCnt As Int32 = oReq.lRequirementCount
            '            Dim lNewCnt As Int32 = GetPirateTurretCount()
            '            If lOrigCnt > 0 AndAlso lNewCnt = 0 Then
            '                oReq.bRequirementMet = True
            '                If oReq.oFiresStep Is Nothing = False AndAlso oReq.oFiresStep.StepID > moCurrentStep.StepID Then
            '                    bForceFire = True
            '                End If
            '            End If
            '        ElseIf oReq.lTriggerType = StartupTriggerType.eIncrementTankDeath Then
            '            Dim lOrigCnt As Int32 = oReq.lRequirementCount
            '            Dim lNewCnt As Int32 = GetPlayerTankCount()
            '            If lOrigCnt > 0 AndAlso lNewCnt = 0 Then
            '                oReq.bRequirementMet = True
            '                If oReq.oFiresStep Is Nothing = False AndAlso oReq.oFiresStep.StepID > moCurrentStep.StepID Then
            '                    bForceFire = True
            '                End If
            '            End If
            '        End If

            '        If bForceFire = True AndAlso oReq.oFiresStep.CheckRequirements() = True Then
            '            'any preq step id?
            '            If oReq.oFiresStep.oPreqStep Is Nothing = False Then
            '                If oReq.oFiresStep.oPreqStep.bStepExecuted = False Then Continue For
            '                'If oReq.oFiresStep.oPreqStep.bStepFinished = False Then Continue For
            '            End If

            '            'Ok, we are here, that means we can execute the fires step
            '            If bCanExecuteStep = True Then ExecuteStep(oReq.oFiresStep)
            '            bCanExecuteStep = False
            '        End If
            '    Next X
            'End If

        End If

    End Sub

    Public Shared Sub ClearSteps()
        mbScriptLoaded = False
        moSteps = Nothing
        ReDim oRequirementsInWaiting(-1)
    End Sub

    Private Shared Function GetStoredCommands() As String
        Dim sStoredCommands As String = ""
        Try
            For X As Int32 = 0 To msStoredCommands.GetUpperBound(0)
                sStoredCommands += vbCrLf + msStoredCommands(X)
            Next X
        Catch
        End Try
        GetStoredCommands = sStoredCommands
    End Function

    Private Shared Sub StoreStepEnableCommands(ByRef oStep As ScriptStep)
        If oStep.sUIControl <> "" Then

            If oStep.lBeginStepGroupNum > -1 Then
                lStepGroupNumber = oStep.lBeginStepGroupNum
                lStepGroupNumberStartedOnStepID = oStep.StepID
            End If

            Dim sCommands() As String = Split(oStep.sUIControl, vbCrLf)
            For X As Int32 = 0 To sCommands.GetUpperBound(0)
                Dim bIsEnableCommand As Boolean = False
                Dim sCurCmd As String = sCommands(X).Trim.ToUpper
                If sCurCmd.IndexOf(":") > -1 Then
                    Dim sValPair() As String = Split(sCurCmd, ":")
                    sValPair(0) = sValPair(0).Trim.ToUpper
                    If sValPair(0) = "UIENABLE" OrElse sValPair(0) = "HOTKEYENABLE" Then
                        bIsEnableCommand = True
                        'ElseIf sValPair(0) = "BEGINSTEPGROUPNUM" Then
                        '    lStepGroupNumber = CInt(Val(sValPair(1)))
                        '    lStepGroupNumberStartedOnStepID = oStep.StepID
                    End If
                End If

                If bIsEnableCommand Then
                    Dim bIsCommandFound As Boolean = False
                    For Y As Int32 = 0 To msStoredCommands.GetUpperBound(0)
                        If msStoredCommands(Y) = sCurCmd Then
                            bIsCommandFound = True
                            Exit For
                        End If
                    Next Y
                    If bIsCommandFound = False Then
                        ReDim Preserve msStoredCommands(msStoredCommands.GetUpperBound(0) + 1)
                        msStoredCommands(msStoredCommands.GetUpperBound(0)) = sCurCmd
                    End If
                Else
                    'Check if the command is not a generic disableall or something
                    'UIDisable:ALL
                    'HotkeyDisable:ALL
                    Dim sValPair() As String = Split(sCurCmd, ":")
                    sValPair(0) = sValPair(0).Trim.ToUpper
                    If sValPair(0) = "UIDISABLE" OrElse sValPair(0) = "HOTKEYDISABLE" Then
                        If sValPair.GetUpperBound(0) > 0 Then
                            sValPair(1) = sValPair(1).Trim.ToUpper
                            If sValPair(1) <> "ALL" Then
                                'ok, specific disable, loop through and remove all enable instances
                                Dim sTempCmd As String = sCurCmd.Replace("UIDISABLE", "UIENABLE").Replace("HOTKEYDISABLE", "HOTKEYENABLE")
                                RemoveStoredCommand(sTempCmd)
                            End If
                        End If
                    End If
                End If
            Next X
        End If
	End Sub

	Private Shared Sub RemoveStoredCommand(ByVal sCmd As String)
		If msStoredCommands Is Nothing Then Return

		Dim bDone As Boolean = False
		Dim lUB As Int32 = msStoredCommands.GetUpperBound(0)

		While bDone = False
			bDone = True
			For Y As Int32 = 0 To lUB
				If msStoredCommands(Y) = sCmd Then
					bDone = False

					For Z As Int32 = Y To lUB - 1
						msStoredCommands(Z) = msStoredCommands(Z + 1)
					Next Z

					lUB -= 1
					Exit For
				End If
			Next Y
		End While

		ReDim Preserve msStoredCommands(lUB)
	End Sub

	Private Shared Sub ProcessStepDisableCommands(ByRef oStep As ScriptStep)
		If oStep.sUIControl <> "" Then
			Dim sCommands() As String = Split(oStep.sUIControl, vbCrLf)
			For X As Int32 = 0 To sCommands.GetUpperBound(0)
				Dim bIsEnableCommand As Boolean = False
				Dim sCurCmd As String = sCommands(X).Trim.ToUpper
				If sCurCmd.IndexOf(":") > -1 Then
					Dim sValPair() As String = Split(sCurCmd, ":")
					sValPair(0) = sValPair(0).Trim.ToUpper
					If sValPair(0) = "UIENABLE" OrElse sValPair(0) = "HOTKEYENABLE" Then
						bIsEnableCommand = True
					End If
				End If

				If bIsEnableCommand = False Then
					'Check if the command is not a generic disableall or something
					'UIDisable:ALL
					'HotkeyDisable:ALL
					Dim sValPair() As String = Split(sCurCmd, ":")
					sValPair(0) = sValPair(0).Trim.ToUpper
					If sValPair(0) = "UIDISABLE" OrElse sValPair(0) = "HOTKEYDISABLE" Then
						If sValPair.GetUpperBound(0) > 0 Then
							sValPair(1) = sValPair(1).Trim.ToUpper
							If sValPair(1) <> "ALL" Then
								'ok, specific disable, loop through and remove all enable instances
								Dim sTempCmd As String = sCurCmd.Replace("UIDISABLE", "UIENABLE").Replace("HOTKEYDISABLE", "HOTKEYENABLE")
								RemoveStoredCommand(sTempCmd)
							End If
						End If
					End If
				End If
			Next X
		End If
    End Sub

    Public Shared Sub CheckReceivedFirstMsg()
        If goCurrentPlayer Is Nothing = False Then
            Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sFile.EndsWith("\") = False Then sFile &= "\"
            sFile &= goCurrentPlayer.PlayerName & ".tut"
            If Exists(sFile) = False Then Return

            Dim oINI As InitFile = New InitFile(sFile)

            If CInt(oINI.GetString("Tutorial", "FirstLiveMsg", "0")) = 0 Then
                'save our new setting
                oINI.WriteString("Tutorial", "FirstLiveMsg", "1")
                'show the first msg
                Dim oItem As New TOCList.TOCItem
                With oItem
                    ReDim .oItems(0)
                    .oItems(0) = New TOCList.TOCItem
                    .oItems(0).sText = "You are now in the live game. Should you need further help, ask for it using the chat interface or click the question mark (F1) button on the Quickbar." & vbCrLf & vbCrLf & _
                        "Now that you are in the live game, a good objective would be to establish a tradepost to begin trade with other players." & vbCrLf & vbCrLf & "Good luck!"
                    .oItems(0).lType = TOCList.lItemType.ePage
                    .lType = TOCList.lItemType.eTopic
                    .sText = "Welcome to the Live Game"
                End With

                Dim oFrm As frmHelpItem = CType(goUILib.GetWindow("frmHelpItem"), frmHelpItem)
                If oFrm Is Nothing Then oFrm = New frmHelpItem(goUILib)
                oFrm.SetFromTopicNode(oItem)
                oFrm.Visible = True

                'Dim oChat As frmChat = CType(goUILib.GetWindow("frmChat"), frmChat)
                'If oChat Is Nothing = False Then
                '    oChat.SetDefaultChatTabPrefix("/GENERAL ")
                'End If
                'oChat = Nothing
            End If
            oINI = Nothing
        End If
    End Sub
End Class
