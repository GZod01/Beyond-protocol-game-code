Option Strict On

Public Class TutorialManager
    'Ok, two things about the tutorial manager:
    '1) The step-by-step tutorial system which guides players through the initial steps of the tutorial
    '2) The in-game help system which is a TOC approach to listening to critical sections of the tutorial

    Public Enum TutorialTriggerType As Integer
        FirstTimePlayer = 1
        SelectedEngineer = 2
        CommandCenterSelectedInBuildWindow = 3
        CommandCenterPlaced = 4
        CommandCenterBuilt = 5
        ColonyDetailsSmallShown = 6
        ColonyDetailsFullShown = 7
        PowerGeneratorBuilt = 8
        BarracksBuilt = 9
        BarracksSelected = 10
        BarracksBuildWindow = 11
        ResidentialFacilityBuilt = 12
        OfficersFacilityBuilt = 13
        FactoryBuilt = 14
        CargoTruckBuilt = 15
        ComponentsLoadIntoFactory = 16                  'TODO: Not currently captured
        MineralCacheHover = 17
        EngineerOrderedToBuildMine = 18
        MiningFacilityBuilt = 19
        DiscoveredNewMaterial = 20

        ResearchFacilityPlaced = 21
        ResearchFacilityBuilt = 22
        ResearchWindow = 23

        MaterialResearchWindow = 24
        MatResDiscoverButton = 25
        MatResDiscoverResearchStart = 26
        MineralDiscovered = 27
        MatResStudyButton = 28

        SchoolOfSciencesResearched = 29

        FirstEngagement = 30
        AssaultPirateBase = 31

        RandomTriggerEventThatDoesntMatter = 32     'never gets fired!!!!
        BackspaceKeyPressed = 33
        HomeKeyPressed = 34

        ChangedToSolarSystemEnvir = 35
        IAmLostF7Key = 36
        FirstBudgetScreenJumpTo = 37

        'Trade window stuff
        BuyScreenSelected = 38
        SellScreenSelected = 39
        DirectTradeScreenSelected = 40
        BuyOrderScreenSelected = 41
        CreateBuyOrderSelected = 42
        TradeHistoryScreenSelected = 43
        TradeInProgressScreenSelected = 44

        EmailComposeButtonClick = 45
        EmailContactsButtonClick = 46
        EmailOptionsButtonClick = 47

        ExceedsCommandPoints = 48

        PiratesEliminated = 49

        TutorialStepEnded = Int32.MaxValue
    End Enum

    Public Enum TutorialStepType As Byte
        eNormalTutorial = 0         'the normal tutorial steps for the beginning player
        eArmorBuilder = 1           'armor builder tutorial
        eEngineBuilder = 2          'engine builder tutorial
        eHullBuilder = 3            'Hull Builder tutorail
        eMatRes = 4                 'material research tutorial
        eMinPropHlpr = 5            'mineral property helper for builders
        ePrototypeBuilder = 6       'prototype builder tut
        eRadarBuilder = 7           'radar builder tut
        eResearchMain = 8           '? on the Research Build window
        eShieldBuilder = 9          'shield builder tut
        eTradeScreen = 10
        eBattlegroup = 11           'battlegroup tutorial
        eBudget = 12                'budget tutorial
        eChat = 13                  'chat window
        eMining = 14                'Mining window
        eRepairing = 15             'Repair Tutorial
        eEmail = 16                 'Email Tutorial
        eDiplomacy = 17             'Diplomacy Window tutorial
        eRepairAUnit = 18
        eRepairAFacility = 19
        eCPExceeding = 20
        eBugMainWindow = 21
        eBugEntryWindow = 22
        eStepUnused = 255
    End Enum

    Public Structure TutorialStep
        Public lStepID As Int32
        Public sStepTitle As String     'only shown in TOC, if this is blank, then this step does not show up in the TOC
        Public sWAVFile As String
        Public lPreqStepID As Int32
        Public lPreqActionID As Int32
        Public sDisplayText As String
        Public lDisplayLocX As Int32
        Public lDisplayLocY As Int32

        Public GroupID As Int32         'for similar items that need to be grouped together

        Public bStepExecuted As Boolean

        Public lResetTriggerID As Int32

        Public sAlertControl As String

        Public yStepType As Byte
    End Structure

    Public muSteps() As TutorialStep
    Public mlStepUB As Int32 = -1

    'Monitors what triggers have passed
    Private mlTriggers() As Int32
    Private mlTriggerUB As Int32 = -1

    Public TutorialOn As Boolean = False

    Private mlCurrentStepID As Int32 = -1

    Public Shared bFirstFactoryBuilt As Boolean = False

    Public bAutoTrigger As Boolean = True

    Private myCurrentStepType As Byte = TutorialStepType.eNormalTutorial

    Public lTemporaryOnGroupID As Int32 = -1

    Public Sub HandleTutorialTrigger(ByVal lTrigger As TutorialTriggerType)

        If lTrigger <> TutorialTriggerType.TutorialStepEnded Then
            Dim bFound As Boolean = False
            For X As Int32 = 0 To mlTriggerUB
                If mlTriggers(X) = lTrigger Then
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False Then
                mlTriggerUB += 1
                ReDim Preserve mlTriggers(mlTriggerUB)
                mlTriggers(mlTriggerUB) = lTrigger
            End If
        End If

        If bAutoTrigger = False Then Return

        Dim bStepTriggerred As Boolean = False

        For X As Int32 = 0 To mlStepUB
            If muSteps(X).lPreqStepID = mlCurrentStepID Then
                If muSteps(X).bStepExecuted = False AndAlso (muSteps(X).yStepType = myCurrentStepType OrElse myCurrentStepType = 0) Then
                    'Ok, found one, has its trigger been fired? If not, this is our step
                    If ((lTrigger = TutorialTriggerType.TutorialStepEnded AndAlso muSteps(X).lPreqActionID < 1) OrElse EventTriggered(muSteps(X).lPreqActionID) = True) Then
                        Dim oWin As frmTutorialStep_Old = CType(goUILib.GetWindow("frmTutorialStep_Old"), frmTutorialStep_Old)
                        If oWin Is Nothing Then oWin = New frmTutorialStep_Old(goUILib)
                        oWin.Visible = True
                        oWin.SetFromStep(muSteps(X))
                        oWin = Nothing

                        ResetTrigger(muSteps(X).lResetTriggerID)
                        muSteps(X).bStepExecuted = True

                        bStepTriggerred = True

                        mlCurrentStepID = muSteps(X).lStepID

                        If muSteps(X).sAlertControl <> "" Then AlertControl(muSteps(X).sAlertControl)

                        Exit For
                    End If
                End If
            End If
        Next X

        'For X As Int32 = 0 To mlStepUB
        '    If muSteps(X).bStepExecuted = False Then
        '        If lTemporaryOnGroupID = -1 OrElse muSteps(X).GroupID = lTemporaryOnGroupID Then
        '            If muSteps(X).lPreqStepID > 0 Then
        '                Dim bGood As Boolean = False
        '                For Y As Int32 = 0 To mlStepUB
        '                    If muSteps(Y).lStepID = muSteps(X).lPreqStepID Then
        '                        bGood = muSteps(Y).bStepExecuted
        '                        Exit For
        '                    End If
        '                Next Y

        '                If bGood = False Then Continue For
        '            End If

        '            If muSteps(X).yStepType = myCurrentStepType AndAlso ((lTrigger = TutorialTriggerType.TutorialStepEnded AndAlso muSteps(X).lPreqActionID < 1) OrElse EventTriggered(muSteps(X).lPreqActionID) = True) Then
        '                Dim oWin As frmTutorialStep = CType(goUILib.GetWindow("frmTutorialStep"), frmTutorialStep)
        '                If oWin Is Nothing Then oWin = New frmTutorialStep(goUILib)
        '                oWin.Visible = True
        '                oWin.SetFromStep(muSteps(X))
        '                oWin = Nothing

        '                ResetTrigger(muSteps(X).lResetTriggerID)
        '                muSteps(X).bStepExecuted = True

        '                bStepTriggerred = True

        '                mlCurrentStepID = muSteps(X).lStepID

        '                If muSteps(X).sAlertControl <> "" Then AlertControl(muSteps(X).sAlertControl)

        '                Exit For
        '            End If
        '        End If
        '    End If
        'Next X

        If bStepTriggerred = False AndAlso lTemporaryOnGroupID <> -1 Then
            Dim bShutOff As Boolean = True
            For X As Int32 = 0 To mlStepUB
                If muSteps(X).GroupID = lTemporaryOnGroupID AndAlso muSteps(X).bStepExecuted = False Then
                    bShutOff = False
                    Exit For
                End If
            Next X
            If bShutOff = True Then
                Me.TutorialOn = False
                lTemporaryOnGroupID = -1
            End If
        End If

        'If bStepTriggerred = False AndAlso myCurrentStepType <> 0 AndAlso lTrigger = TutorialTriggerType.TutorialStepEnded Then
        '    myCurrentStepType = 0
        '    goUILib.RemoveWindow("frmTutorialStep")
        'End If
    End Sub

    Public Function EventTriggered(ByVal lTrigger As Int32) As Boolean
        For X As Int32 = 0 To mlTriggerUB
            If mlTriggers(X) = lTrigger Then Return True
        Next X
        Return False
    End Function

    Public Sub ResetGroupOfStep(ByVal lStepID As Int32)
        For X As Int32 = 0 To mlStepUB
            If muSteps(X).lStepID = lStepID Then
                If muSteps(X).GroupID > 0 Then
                    For Y As Int32 = 0 To mlStepUB
                        If muSteps(Y).GroupID = muSteps(X).GroupID Then
                            muSteps(Y).bStepExecuted = False
                        End If
                    Next Y
                End If
            End If
        Next X

    End Sub

    Public Sub ExecuteTutorialStep(ByVal lStepID As Int32)

        For X As Int32 = 0 To mlStepUB
            If muSteps(X).lStepID = lStepID Then

                mlCurrentStepID = lStepID

                muSteps(X).bStepExecuted = True

                Dim oWin As frmTutorialStep_Old = CType(goUILib.GetWindow("frmTutorialStep_Old"), frmTutorialStep_Old)
                If oWin Is Nothing Then oWin = New frmTutorialStep_Old(goUILib)
                oWin.Visible = True
                oWin.SetFromStep(muSteps(X))
                oWin = Nothing

                ResetTrigger(muSteps(X).lResetTriggerID)

                If muSteps(X).sAlertControl <> "" Then AlertControl(muSteps(X).sAlertControl)

                Exit For
            End If
        Next X
    End Sub

    Private Sub ResetTrigger(ByVal lTriggerID As Int32)
        If lTriggerID = -1 Then Return

        For X As Int32 = 0 To mlTriggerUB
            If mlTriggers(X) = lTriggerID Then
                For Y As Int32 = X To mlTriggerUB - 1
                    mlTriggers(Y) = mlTriggers(Y + 1)
                Next Y
                mlTriggerUB -= 1
                Exit For
            End If
        Next X
    End Sub

    Public Sub New()
        Dim sDATFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sDATFile.EndsWith("\") = False Then sDATFile &= "\"
        sDATFile &= "Audio\SoundFX\Tutorial\TutorialSteps.dat"
        Dim oINI As New InitFile(sDATFile)

        Dim lStep As Int32 = 1

        Dim bDone As Boolean = False

        While bDone = False
            Dim sHdr As String = "Step" & lStep
            Dim lID As Int32 = CInt(Val(oINI.GetString(sHdr, "StepID", "-1")))

            If lID = -1 Then
                bDone = True
            Else
                mlStepUB += 1
                ReDim Preserve muSteps(mlStepUB)

                With muSteps(mlStepUB)
                    .lStepID = lID
                    'If lID = 156 Then Stop
                    .sStepTitle = oINI.GetString(sHdr, "StepTitle", "")
                    .sWAVFile = oINI.GetString(sHdr, "WAVFile", "")
                    .lPreqStepID = CInt(Val(oINI.GetString(sHdr, "PreqStepID", "-1")))
                    .lPreqActionID = CInt(Val(oINI.GetString(sHdr, "PreqActionID", "-1")))
                    .sDisplayText = Replace$(oINI.GetString(sHdr, "DisplayText", ""), "[VBCRLF]", vbCrLf, , , CompareMethod.Text)
                    .lDisplayLocX = CInt(Val(oINI.GetString(sHdr, "DisplayLocX", "-1")))
                    .lDisplayLocY = CInt(Val(oINI.GetString(sHdr, "DisplayLocY", "-1")))
                    .GroupID = CInt(Val(oINI.GetString(sHdr, "GroupID", "0")))
                    .lResetTriggerID = CInt(Val(oINI.GetString(sHdr, "ResetTrigger", "-1")))
                    .sAlertControl = oINI.GetString(sHdr, "AlertControl", "")
                    .yStepType = CByte(Val(oINI.GetString(sHdr, "StepType", "0")))
                End With
            End If

            lStep += 1
        End While

        oINI = Nothing


        'Now, load up the player
        If goCurrentPlayer Is Nothing = False Then
            sDATFile = AppDomain.CurrentDomain.BaseDirectory
            If sDATFile.EndsWith("\") = False Then sDATFile &= "\"
            sDATFile &= goCurrentPlayer.PlayerName & ".tut"

            oINI = New InitFile(sDATFile)

            Dim sSteps As String = oINI.GetString("Tutorial", "Steps", "")
            Dim sTriggers As String = oINI.GetString("Tutorial", "Triggers", "")

            'Me.TutorialOn = Val(oINI.GetString("Tutorial", "TutorialOn", "1")) <> 0
            Me.TutorialOn = False
            TutorialManager.bFirstFactoryBuilt = TutorialManager.bFirstFactoryBuilt OrElse Val(oINI.GetString("Tutorial", "FirstFactoryBuilt", "0")) <> 0

            ParseLoadedSteps(sSteps)
            ParseLoadedTriggers(sTriggers)

            oINI = Nothing
        End If
    End Sub

    Private Sub ParseLoadedSteps(ByVal sSteps As String)
        For X As Int32 = 0 To sSteps.Length - 1
            If sSteps.Substring(X, 1) <> "0" Then
                muSteps(X).bStepExecuted = True
            End If
        Next X
    End Sub

    Private Sub ParseLoadedTriggers(ByVal sTriggers As String)
        Dim sValues() As String = Split(sTriggers, ";")

        mlTriggerUB = -1

        For X As Int32 = 0 To sValues.GetUpperBound(0)
            If sValues(X) <> "" Then
                mlTriggerUB += 1
                ReDim Preserve mlTriggers(mlTriggerUB)
                mlTriggers(mlTriggerUB) = CInt(Val(sValues(X)))
            End If
        Next X
    End Sub

    Protected Overrides Sub Finalize()
        If goCurrentPlayer Is Nothing = False Then
            Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sFile.EndsWith("\") = False Then sFile &= "\"
            sFile &= goCurrentPlayer.PlayerName & ".tut"

            Dim sSteps As String = ""
            For X As Int32 = 0 To mlStepUB
                If muSteps(X).bStepExecuted = True Then
                    sSteps &= "1"
                Else
                    sSteps &= "0"
                End If
            Next X

            Dim sTriggers As String = ""
            For X As Int32 = 0 To mlTriggerUB
                sTriggers &= mlTriggers(X).ToString & ";"
            Next X

            Dim oINI As InitFile = New InitFile(sFile)
            If Me.TutorialOn = True Then
                oINI.WriteString("Tutorial", "TutorialOn", "1")
            Else : oINI.WriteString("Tutorial", "TutorialOn", "0")
            End If

            If TutorialManager.bFirstFactoryBuilt = True Then
                oINI.WriteString("Tutorial", "FirstFactoryBuilt", "1")
            Else : oINI.WriteString("Tutorial", "FirstFactoryBuilt", "0")
            End If

            oINI.WriteString("Tutorial", "Steps", sSteps)
            oINI.WriteString("Tutorial", "Triggers", sTriggers)
            oINI = Nothing
        End If
        MyBase.Finalize()
    End Sub

    Public Sub ForcefullyMoveNext()
        Dim lStepID As Int32 = Int32.MaxValue

        For X As Int32 = 0 To mlStepUB
            If muSteps(X).lPreqStepID = mlCurrentStepID Then
                lStepID = muSteps(X).lStepID
                Exit For
            End If
        Next X

        If lStepID <> Int32.MaxValue Then
            ExecuteTutorialStep(lStepID)
        End If
    End Sub

    Public Sub MovePrevious()
        If mlCurrentStepID <> -1 Then

            Dim lPreqStepID As Int32 = -1
            For X As Int32 = 0 To mlStepUB
                If muSteps(X).lStepID = mlCurrentStepID Then
                    lPreqStepID = muSteps(X).lPreqStepID
                    If lPreqStepID <> -1 Then
                        muSteps(X).bStepExecuted = False
                        mlCurrentStepID = lPreqStepID
                    End If
                    Exit For
                End If
            Next X
            If lPreqStepID <> -1 Then ExecuteTutorialStep(lPreqStepID)

        End If
    End Sub

    Public Sub ForcefulSetTrigger(ByVal lTrigger As TutorialTriggerType)
        For X As Int32 = 0 To mlTriggerUB
            If mlTriggers(X) = lTrigger Then Return
        Next X
        mlTriggerUB += 1
        ReDim Preserve mlTriggers(mlTriggerUB)
        mlTriggers(mlTriggerUB) = lTrigger
    End Sub

    Public Function ContinueTutorial() As Boolean
        If TutorialOn = True Then

            Dim lCurStep As Int32 = -1
            Dim bDone As Boolean = False

            For X As Int32 = 0 To mlStepUB
                If muSteps(X).lPreqActionID = TutorialTriggerType.FirstTimePlayer Then
                    lCurStep = muSteps(X).lStepID
                    If muSteps(X).bStepExecuted = False Then Return False
                    Exit For
                End If
            Next X

            Dim bFound As Boolean = False
            While bDone = False
                'ok, now, find the item that requires the current step...
                bDone = True
                For X As Int32 = 0 To mlStepUB
                    If muSteps(X).lPreqStepID = lCurStep Then
                        'ok, now, is this step executed?
                        If muSteps(X).bStepExecuted = True Then
                            'yes, ok, set our curstep
                            lCurStep = muSteps(X).lStepID
                            bDone = False
                            Exit For
                        Else
                            'no, ok, execute our current step
                            bFound = True
                            Exit For
                        End If
                    End If
                Next X
            End While
            If bFound = True Then
                ExecuteTutorialStep(lCurStep)
                Return True
            End If

            'For X As Int32 = 0 To mlStepUB
            '	If muSteps(X).bStepExecuted = False Then
            '		If muSteps(X).lStepID = 1 Then Return False
            '		If muSteps(X).lPreqActionID = -1 OrElse EventTriggered(muSteps(X).lPreqActionID) Then
            '			ExecuteTutorialStep(muSteps(X).lStepID)
            '			Return True
            '		End If
            '		Return False
            '	End If
            'Next X
        End If
        Return False
    End Function

    Private moAlertControl As UIControl = Nothing
    Private mlAlertCount As Int32
    Private mlLastAlertUpdate As Int32
    Private mbRepeatAlertControl As Boolean = False
    Private msAlertControlName As String
    Private mbAlertSounded As Boolean = False

    Private Sub AlertControl(ByVal sControlName As String)
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

    Public Sub HandleAlertControlUpdate()
        'If moAlertControl Is Nothing = False Then
        '    If mbAlertSounded = False Then If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
        '    mbAlertSounded = True
        '    mbRepeatAlertControl = False
        '    If glCurrentCycle - mlLastAlertUpdate > 3 Then
        '        moAlertControl.Visible = Not moAlertControl.Visible
        '        moAlertControl.IsDirty = False
        '        moAlertControl.IsDirty = True
        '        mlAlertCount -= 1
        '        mlLastAlertUpdate = glCurrentCycle
        '        If mlAlertCount < 1 Then
        '            moAlertControl.Visible = False
        '            moAlertControl.Visible = True
        '            moAlertControl.IsDirty = False
        '            moAlertControl.IsDirty = True
        '            moAlertControl = Nothing
        '        End If
        '    End If

        'ElseIf mbRepeatAlertControl = True Then
        '    AlertControl(msAlertControlName)
        'End If
    End Sub

    Public Sub BeginNewStepType(ByVal yStepType As TutorialStepType)
        myCurrentStepType = yStepType
        Dim lGroupID As Int32 = -1
        Dim lLowestPreqStepID As Int32 = Int32.MaxValue
        For X As Int32 = 0 To mlStepUB
            If muSteps(X).yStepType = yStepType Then
                muSteps(X).bStepExecuted = False
                If muSteps(X).lPreqActionID <> -1 Then
                    ResetTrigger(muSteps(X).lPreqActionID)
                End If
                lGroupID = muSteps(X).GroupID
                If muSteps(X).lPreqStepID > 0 AndAlso muSteps(X).lPreqStepID < lLowestPreqStepID Then
                    lLowestPreqStepID = muSteps(X).lPreqStepID
                End If
            End If
        Next X
        If TutorialOn = False Then lTemporaryOnGroupID = lGroupID
        TutorialOn = True

        If lLowestPreqStepID <> Int32.MaxValue Then ExecuteTutorialStep(lLowestPreqStepID) Else HandleTutorialTrigger(TutorialTriggerType.TutorialStepEnded)
    End Sub

    Public Sub CancelCustomStepType()
        myCurrentStepType = TutorialStepType.eNormalTutorial
        If lTemporaryOnGroupID <> -1 Then
            Me.TutorialOn = False
            lTemporaryOnGroupID = -1
        End If

        'If Me.TutorialOn = True Then
        '    HandleTutorialTrigger(TutorialTriggerType.TutorialStepEnded)
        'End If
    End Sub

    Public ReadOnly Property CurrentStepType() As TutorialStepType
        Get
            Return CType(myCurrentStepType, TutorialStepType)
        End Get
    End Property
End Class
