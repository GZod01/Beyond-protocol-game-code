Option Strict On

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
Public Class Agent
    Inherits Base_GUID

    Public sAgentName As String
    Public Infiltration As Byte
    Public Dagger As Byte
    Public Resourcefulness As Byte
    Public Luck As Byte
    Public Loyalty As Byte
    Public Suspicion As Byte

	Public lUpfrontCost As Int32
	Public lMaintCost As Int32

    Public InfiltrationLevel As Byte
    Public InfiltrationType As eInfiltrationType = 0

    Public lTargetID As Int32 = -1
    Public iTargetTypeID As Int16 = -1
    Public lOwnerID As Int32 = -1
    Public lMissionID As Int32 = -1

    Public bMale As Boolean
    Public yReportFreq As Byte

    Public Skills() As Skill
    Public SkillProf() As Byte
    Public SkillUB As Int32 = -1

    Public uHistory() As AgentMissionHistory
    Public lHistoryUB As Int32 = -1
    Private mbHistoryNeedsSorted As Boolean = True

	Public dtRecruited As Date

	Public mlArrivalCycles As Int32 = -1
	Public mlReceivedCycle As Int32 = -1

    'Mission Specifics....
	Public lAgentStatus As AgentStatus = AgentStatus.NormalStatus

    Private mbReceivedSkillList As Boolean = False
    Private mbRequestedSkillList As Boolean = False
    Public ReadOnly Property bRequestedSkillList() As Boolean
        Get
            Return mbRequestedSkillList
        End Get
    End Property

    Public Sub CheckRequestSkillList()
        If mbRequestedSkillList = True Then Return
        If goUILib Is Nothing Then Return
        mbRequestedSkillList = True

        Dim yOut(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yOut, 0)
        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yOut, 2)
        goUILib.SendMsgToPrimary(yOut)
    End Sub

    Public Sub FillFromMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 'for msgcode
        With Me
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.sAgentName = GetStringFromBytes(yData, lPos, 30) : lPos += 30
            .Infiltration = yData(lPos) : lPos += 1
            .Dagger = yData(lPos) : lPos += 1
            .Resourcefulness = yData(lPos) : lPos += 1
            .Suspicion = yData(lPos) : lPos += 1
            .InfiltrationLevel = yData(lPos) : lPos += 1
            .InfiltrationType = CType(yData(lPos), eInfiltrationType) : lPos += 1

            Dim lSecondsFromRecruit As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            dtRecruited = Now.Subtract(New TimeSpan(0, 0, lSecondsFromRecruit))

            lTargetID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iTargetTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lOwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lMissionID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			bMale = (yData(lPos) <> 0) : lPos += 1

			.lUpfrontCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lMaintCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.yReportFreq = yData(lPos) : lPos += 1
			.lAgentStatus = CType(System.BitConverter.ToInt32(yData, lPos), AgentStatus) : lPos += 4
			.mlArrivalCycles = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		End With

		mlReceivedCycle = glCurrentCycle

	End Sub

	Public Sub HandleSkillList(ByRef yData() As Byte)
		Dim lPos As Int32 = 6	'2 for msgcode, 4 for agentid

		SkillUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
		ReDim Skills(SkillUB)
		ReDim SkillProf(SkillUB)

		For X As Int32 = 0 To SkillUB
			Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			SkillProf(X) = yData(lPos) : lPos += 1

			For Y As Int32 = 0 To glSkillUB
				If goSkills(Y).ObjectID = lID Then
					Skills(X) = goSkills(Y)
					Exit For
				End If
			Next Y
        Next X

        mbReceivedSkillList = True
	End Sub

    Public Sub HandleAgentHistoryMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 6   'for msgcode and agentid

        lHistoryUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
        ReDim uHistory(lHistoryUB)
        For X As Int32 = 0 To lHistoryUB
            With uHistory(X)
                .lTargetPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lMissionUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4

                ReDim .lMissionID(.lMissionUB)
                ReDim .yResult(.lMissionUB)
                For Y As Int32 = 0 To .lMissionUB
                    .lMissionID(Y) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .yResult(Y) = yData(lPos) : lPos += 1
                Next Y
            End With
        Next X
    End Sub

    Private Sub SortHistory()

        mbHistoryNeedsSorted = False
        'ensure all of our names are in place
        For X As Int32 = 0 To lHistoryUB
            Dim sName As String = GetCacheObjectValue(uHistory(X).lTargetPlayerID, ObjectType.ePlayer)
            If sName.ToUpper = "UNKNOWN" Then mbHistoryNeedsSorted = True
        Next X
        If mbHistoryNeedsSorted = True Then Return

        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        For X As Int32 = 0 To lHistoryUB
            Dim lIdx As Int32 = -1

            Dim sName As String = GetCacheObjectValue(uHistory(X).lTargetPlayerID, ObjectType.ePlayer)

            For Y As Int32 = 0 To lSortedUB
                Dim sOtherName As String = GetCacheObjectValue(uHistory(lSorted(Y)).lTargetPlayerID, ObjectType.ePlayer)
                If sOtherName > sName Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedUB += 1
            ReDim Preserve lSorted(lSortedUB)
            If lIdx = -1 Then
                lSorted(lSortedUB) = X
            Else
                For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                    lSorted(Y) = lSorted(Y - 1)
                Next Y
                lSorted(lIdx) = X
            End If
        Next X

        'Ok, lsorted should have our values
        Dim uTemp() As AgentMissionHistory = uHistory
        For X As Int32 = 0 To lHistoryUB
            uHistory(X) = uTemp(lSorted(X))
        Next X
    End Sub

    Public Sub SmartFillHistoryList(ByRef oLB As UIListBox)
        If mbHistoryNeedsSorted = True Then SortHistory()
        If mbHistoryNeedsSorted = True Then Return

        If oLB.ListCount - 1 <> lHistoryUB Then
            FillHistoryList(oLB)
            Return
        End If


        For X As Int32 = 0 To oLB.ListCount - 1
            Dim lID As Int32 = oLB.ItemData(X)
            Dim lID2 As Int32 = oLB.ItemData2(X)
            Dim sName As String = ""
            If lID2 = -1 Then
                'player
                sName = GetCacheObjectValue(lID, ObjectType.ePlayer)
            Else
                'Mission
                sName = GetMissionName(lID).PadRight(21, " "c)
                If lID2 = 0 Then sName &= "FAILED" Else sName &= "SUCCESS"
            End If
            If oLB.List(X) <> sName Then oLB.List(X) = sName
        Next X
    End Sub

    Private Sub FillHistoryList(ByRef oLB As UIListBox)
        oLB.Clear()
        For X As Int32 = 0 To lHistoryUB
            With uHistory(X)
                oLB.AddItem(GetCacheObjectValue(.lTargetPlayerID, ObjectType.ePlayer), True)
                oLB.ItemData(oLB.NewIndex) = .lTargetPlayerID
                oLB.ItemData2(oLB.NewIndex) = -1
                For Y As Int32 = 0 To .lMissionUB
                    Dim sText As String = GetMissionName(.lMissionID(Y)).PadRight(21, " "c)
                    If .yResult(Y) = 0 Then sText &= "FAILED" Else sText &= "SUCCESS"
                    oLB.AddItem(sText, False)
                    oLB.ItemData(oLB.NewIndex) = .lMissionID(Y)
                    oLB.ItemData2(oLB.NewIndex) = .yResult(Y)
                    oLB.ApplyIconOffset(oLB.NewIndex) = True
                Next Y
            End With
        Next X
    End Sub

    Private Function GetMissionName(ByVal lMissionID As Int32) As String
        For X As Int32 = 0 To glMissionUB
            If glMissionIdx(X) = lMissionID Then Return goMissions(X).sMissionName
        Next X
        Return GetCacheObjectValue(lMissionID, ObjectType.eMission)
    End Function

    Public Shared Function GetStatusText(ByVal lStatus As Int32, ByVal lTargetID As Int32, ByVal iTargetTypeID As Int16, ByVal yInfiltrationType As Byte) As String

        If (lStatus And AgentStatus.IsDead) <> 0 Then
            Return "Killed"
        ElseIf (lStatus And AgentStatus.Dismissed) <> 0 Then
            Return "Dismissed"
        ElseIf (lStatus And AgentStatus.NewRecruit) <> 0 Then
            Return "New Recruit"
        ElseIf (lStatus And AgentStatus.HasBeenCaptured) <> 0 Then
            Return "Captured by " & GetCacheObjectValue(lTargetID, iTargetTypeID)
        ElseIf (lStatus And AgentStatus.ReturningHome) <> 0 Then
            Return "Returning Home"
        ElseIf (lStatus And AgentStatus.CounterAgent) <> 0 Then
            Dim sIType As String
            Select Case yInfiltrationType
                Case eInfiltrationType.eGeneralInfiltration
                    sIType = "General"
                Case eInfiltrationType.eMilitaryInfiltration
                    sIType = "Military"
                Case eInfiltrationType.eFederalInfiltration
                    sIType = "Government"
                Case eInfiltrationType.eColonialInfiltration
                    sIType = "Colonial"
                Case eInfiltrationType.eProductionInfiltration
                    sIType = "Production"
                Case eInfiltrationType.eResearchInfiltration
                    sIType = "Research"
                Case eInfiltrationType.eTradeInfiltration
                    sIType = "Trade"
                Case eInfiltrationType.ePowerCenterInfiltration
                    sIType = "Power Center"
                Case eInfiltrationType.eMiningInfiltration
                    sIType = "Mining"
                Case eInfiltrationType.eSolarSystemInfiltration
                    sIType = "System"
                Case eInfiltrationType.eCapitalShipInfiltration
                    sIType = "Fleet"
                Case eInfiltrationType.ePlanetInfiltration
                    sIType = "Planet"
                Case eInfiltrationType.eCombatUnitInfiltration
                    sIType = "Unit"
                Case eInfiltrationType.eAgencyInfiltration
                    sIType = "Agency"
                Case eInfiltrationType.eCorporationInfiltration
                    sIType = "Corporate"
                Case Else
                    sIType = "Unknown"



            End Select


            Return "Counter Agent(" & sIType & ")"
        ElseIf (lStatus And AgentStatus.OnAMission) <> 0 Then
            Return "On a Mission"
        ElseIf (lStatus And AgentStatus.IsInfiltrated) <> 0 Then
            Dim sName As String = ""
            If lTargetID > -1 AndAlso iTargetTypeID > -1 Then
                sName = " " & GetCacheObjectValue(lTargetID, iTargetTypeID)
            End If
            Dim sIType As String
            Select Case yInfiltrationType
                Case eInfiltrationType.eGeneralInfiltration
                    sIType = "General"
                Case eInfiltrationType.eMilitaryInfiltration
                    sIType = "Military"
                Case eInfiltrationType.eFederalInfiltration
                    sIType = "Government"
                Case eInfiltrationType.eColonialInfiltration
                    sIType = "Colonial"
                Case eInfiltrationType.eProductionInfiltration
                    sIType = "Production"
                Case eInfiltrationType.eResearchInfiltration
                    sIType = "Research"
                Case eInfiltrationType.eTradeInfiltration
                    sIType = "Trade"
                Case eInfiltrationType.ePowerCenterInfiltration
                    sIType = "Power Center"
                Case eInfiltrationType.eMiningInfiltration
                    sIType = "Mining"
                Case eInfiltrationType.eSolarSystemInfiltration
                    sIType = "System"
                Case eInfiltrationType.eCapitalShipInfiltration
                    sIType = "Fleet"
                Case eInfiltrationType.ePlanetInfiltration
                    sIType = "Planet"
                Case eInfiltrationType.eCombatUnitInfiltration
                    sIType = "Unit"
                Case eInfiltrationType.eAgencyInfiltration
                    sIType = "Agency"
                Case eInfiltrationType.eCorporationInfiltration
                    sIType = "Corporate"
                Case Else
                    sIType = "Unknown"



            End Select

            Return "Infiltrated " & sName & " (" & sIType & ")"
        ElseIf (lStatus And AgentStatus.Infiltrating) <> 0 Then
            Return "Infiltrating"
        Else : Return "Waiting"
        End If

    End Function

	Public Function GetStatusIconRect() As Rectangle
		If (lAgentStatus And AgentStatus.IsDead) <> 0 Then
			Return New Rectangle(16, 48, 16, 16)
		ElseIf (lAgentStatus And AgentStatus.Dismissed) <> 0 Then
			Return New Rectangle(16, 0, 16, 16)
		ElseIf (lAgentStatus And AgentStatus.NewRecruit) <> 0 Then
			Return New Rectangle(48, 16, 16, 16)
		ElseIf (lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
			Return New Rectangle(0, 48, 16, 16)
		ElseIf (lAgentStatus And AgentStatus.ReturningHome) <> 0 Then
			Return New Rectangle(48, 32, 16, 16)
		ElseIf (lAgentStatus And AgentStatus.CounterAgent) <> 0 Then
			Return New Rectangle(0, 32, 16, 16)
		ElseIf (lAgentStatus And AgentStatus.OnAMission) <> 0 Then
			Return New Rectangle(0, 0, 16, 16)
		ElseIf (lAgentStatus And AgentStatus.IsInfiltrated) <> 0 Then
			Return New Rectangle(16, 32, 16, 16)
		ElseIf (lAgentStatus And AgentStatus.Infiltrating) <> 0 Then
			Return New Rectangle(32, 32, 16, 16)
		Else : Return New Rectangle(64, 0, 16, 16)
		End If
	End Function

	Public Function GetSecondsTillArrival() As Int32
		If mlArrivalCycles = -1 Then Return 0
		Dim lCycles As Int32 = mlArrivalCycles - (glCurrentCycle - mlReceivedCycle)
		If lCycles < 0 Then
			mlArrivalCycles = -1
			lCycles = 0
		End If
		Return lCycles \ 30
    End Function

    Public Function GetSkillProficiency(ByVal lSkillID As Int32) As Byte
        For X As Int32 = 0 To SkillUB
            If Skills(X).ObjectID = lSkillID Then Return SkillProf(X)
        Next X
        Return 0
    End Function


    Private Shared mbInExport As Boolean = False
    Public Shared Sub Export_AgentInfo()
        If goCurrentPlayer Is Nothing Then Return
        If mbInExport = True Then Return
        mbInExport = True

        Dim oThread As New Threading.Thread(AddressOf DoTheExport)
        oThread.Start()
    End Sub
    Private Shared Sub DoTheExport()
        Try
            Dim bDone As Boolean = False
            While bDone = False
                bDone = True
                For X As Int32 = 0 To goCurrentPlayer.AgentUB
                    If goCurrentPlayer.AgentIdx(X) > -1 Then
                        Dim oAgent As Agent = goCurrentPlayer.Agents(X)
                        If oAgent.bRequestedSkillList = False Then
                            oAgent.CheckRequestSkillList()
                            bDone = False
                            Exit For
                        ElseIf oAgent.mbReceivedSkillList = False Then
                            bDone = False
                            Exit For
                        End If
                    End If
                Next X

                If bDone = False Then
                    Threading.Thread.Sleep(10)
                End If
            End While

            If muSettings.ExportedDataFormat = 1 Then
                Export_AgentInfo_Csv()
            ElseIf muSettings.ExportedDataFormat = 2 Then
                Export_AgentInfo_Xml()
            End If
        Catch
            'Do nothing... we could alert the user that Export Agent failed
        End Try

        'set our inExport flag before thread dies
        mbInExport = False
    End Sub
    Private Shared Sub Export_AgentInfo_Xml()
        Dim sExportData As String = ""
        Dim bShortFmt As Boolean = True
        sExportData &= "<Agents>" & vbCrLf
        Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName)
        If lSorted Is Nothing = False Then
            For i As Int32 = 0 To lSorted.GetUpperBound(0)
                If goCurrentPlayer.AgentIdx(lSorted(i)) <> -1 Then
                    With goCurrentPlayer.Agents(lSorted(i))
                        sExportData &= vbTab & vbTab & "<Agent>" & vbCrLf
                        sExportData &= vbTab & vbTab & vbTab & "<AgentName>" & .sAgentName & "</AgentName>" & vbCrLf
                        sExportData &= vbTab & vbTab & vbTab & "<Status>" & Agent.GetStatusText(.lAgentStatus, .lTargetID, .iTargetTypeID, .InfiltrationType) & "</Status>" & vbCrLf
                        If bShortFmt = True Then
                            sExportData &= vbTab & vbTab & vbTab & "<Proficiency><Name>Dagger</Name><Value>" & .Dagger.ToString & "</Value></Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "<Proficiency><Name>Infiltration</Name><Value>" & .Infiltration.ToString & "</Value></Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "<Proficiency><Name>Resourcefullness</Name><Value>" & .Resourcefulness.ToString & "</Value></Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "<Proficiency><Name>Suspicion</Name><Value>" & .Suspicion.ToString & "</Value></Proficiency>" & vbCrLf
                        Else
                            sExportData &= vbTab & vbTab & vbTab & "<Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & vbTab & "<Name>Dagger</Name>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & vbTab & "<Value>" & .Dagger.ToString & "</Value>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "</Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "<Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & vbTab & "<Name>Infiltration</Name>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & vbTab & "<Value>" & .Infiltration.ToString & "</Value>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "</Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "<Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & vbTab & "<Name>Resourcefullness</Name>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & vbTab & "<Value>" & .Resourcefulness.ToString & "</Value>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "</Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "<Proficiency>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & vbTab & "<Name>Suspicion</Name>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & vbTab & "<Value>" & .Suspicion.ToString & "</Value>" & vbCrLf
                            sExportData &= vbTab & vbTab & vbTab & "</Proficiency>" & vbCrLf
                        End If

                        'TODO: Not sure you want to begin an async call here
                        'If goCurrentPlayer.Agents(lSorted(i)) Is Nothing = False AndAlso goCurrentPlayer.Agents(lSorted(i)).bRequestedSkillList = False Then
                        '    .bRequestedSkillList = True
                        '    Dim yOut(5) As Byte
                        '    System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yOut, 0)
                        '    System.BitConverter.GetBytes(.ObjectID).CopyTo(yOut, 2)
                        '    MyBase.moUILib.SendMsgToPrimary(yOut)
                        'End If

                        'Sort by proficiency
                        Dim llSorted() As Int32 = Nothing
                        Dim llSortedUB As Int32 = -1
                        For X As Int32 = 0 To .SkillUB
                            Dim lIdx As Int32 = -1

                            For Y As Int32 = 0 To llSortedUB
                                If .SkillProf(llSorted(Y)) < .SkillProf(X) Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Next Y
                            llSortedUB += 1
                            ReDim Preserve llSorted(llSortedUB)
                            If lIdx = -1 Then
                                llSorted(llSortedUB) = X
                            Else
                                For Y As Int32 = llSortedUB To lIdx + 1 Step -1
                                    llSorted(Y) = llSorted(Y - 1)
                                Next Y
                                llSorted(lIdx) = X
                            End If
                        Next X

                        For X As Int32 = 0 To llSortedUB
                            If bShortFmt = True Then
                                sExportData &= vbTab & vbTab & vbTab & "<Skill><Name>" & .Skills(llSorted(X)).SkillName & "</Name><Value>" & .SkillProf(llSorted(X)).ToString & "</Value></Skill>" & vbCrLf
                            Else
                                sExportData &= vbTab & vbTab & vbTab & "<Skill>" & vbCrLf
                                sExportData &= vbTab & vbTab & vbTab & vbTab & "<Name>" & .Skills(llSorted(X)).SkillName & "</Name>" & vbCrLf
                                sExportData &= vbTab & vbTab & vbTab & vbTab & "<Value>" & .SkillProf(llSorted(X)).ToString & "</Value>" & vbCrLf
                                sExportData &= vbTab & vbTab & vbTab & "</Skill>" & vbCrLf
                            End If
                        Next X
                    End With
                End If
                sExportData &= vbTab & vbTab & "</Agent>" & vbCrLf
            Next i
        End If
        sExportData &= "</Agents>" & vbCrLf

        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile = sFile & "\"
        sFile &= "ExportedData"
        If Exists(sFile) = False Then MkDir(sFile)
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "Agents_" & goCurrentPlayer.PlayerName & "_" & Now.ToString("MM_dd_yyyy_hhmmss") & ".xml"

        Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Create)

        Dim info As Byte() = New System.Text.UTF8Encoding(True).GetBytes(sExportData)
        fsFile.Write(info, 0, info.Length)
        fsFile.Close()
        fsFile.Dispose()

        If goUILib Is Nothing = False Then
            goUILib.AddNotification("Agents Exported.", Color.White, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub
    Private Shared Sub Export_AgentInfo_Csv()
        Dim sExportData As String = ""
        sExportData &= Chr(34) & "AgentName" & Chr(34) & "," & Chr(34) & "Status" & Chr(34) & ",Dagger,Infiltration,Resourcefullness,Suspicion"

        For X As Int32 = 0 To glSkillUB 'oAgent.SkillUB
            sExportData &= "," & goSkills(X).SkillName
        Next
        sExportData &= vbCrLf
        Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName)
        If lSorted Is Nothing = False Then
            For i As Int32 = 0 To lSorted.GetUpperBound(0)
                If goCurrentPlayer.AgentIdx(lSorted(i)) <> -1 Then
                    With goCurrentPlayer.Agents(lSorted(i))
                        sExportData &= Chr(34) & .sAgentName & Chr(34)
                        sExportData &= "," & Chr(34) & Agent.GetStatusText(.lAgentStatus, .lTargetID, .iTargetTypeID, .InfiltrationType) & Chr(34)
                        sExportData &= "," & .Dagger.ToString
                        sExportData &= "," & .Infiltration.ToString
                        sExportData &= "," & .Resourcefulness.ToString
                        sExportData &= "," & .Suspicion.ToString

                        'TODO: Not sure you want to begin an async call here
                        'If goCurrentPlayer.Agents(lSorted(i)) Is Nothing = False AndAlso goCurrentPlayer.Agents(lSorted(i)).bRequestedSkillList = False Then
                        '    .bRequestedSkillList = True
                        '    Dim yOut(5) As Byte
                        '    System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yOut, 0)
                        '    System.BitConverter.GetBytes(.ObjectID).CopyTo(yOut, 2)
                        '    MyBase.moUILib.SendMsgToPrimary(yOut)
                        'End If

                        Dim bFound As Boolean
                        For X As Int32 = 0 To glSkillUB
                            bFound = False
                            sExportData &= ","
                            For j As Int32 = 0 To .SkillUB
                                If .Skills(j).ObjectID = goSkills(X).ObjectID Then
                                    sExportData &= .SkillProf(j).ToString
                                    bFound = True
                                    Exit For
                                End If
                            Next
                            If bFound = False Then
                                sExportData &= "0"
                            End If
                        Next X

                        sExportData &= vbCrLf
                    End With
                End If
            Next i
        End If

        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile = sFile & "\"
        sFile &= "ExportedData"
        If Exists(sFile) = False Then MkDir(sFile)
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "Agents_" & goCurrentPlayer.PlayerName & "_" & Now.ToString("MM_dd_yyyy_hhmmss") & ".csv"

        Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Create)

        Dim info As Byte() = New System.Text.UTF8Encoding(True).GetBytes(sExportData)
        fsFile.Write(info, 0, info.Length)
        fsFile.Close()
        fsFile.Dispose()

        If goUILib Is Nothing = False Then
            goUILib.AddNotification("Agents Exported.", Color.White, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

End Class

Public Structure AgentMissionHistory
    Public lTargetPlayerID As Int32
    Public lMissionID() As Int32
    Public yResult() As Byte
    Public lMissionUB As Int32
End Structure

Public Class CapturedAgent
	Inherits Base_GUID

	Public sAgentName As String = "Unknown"
	Public yInfLevel As Byte = 0
	Public yInfType As Byte = 255
	Public lOwnerID As Int32 = -1
	Public lMissionID As Int32 = -1
	Public lMissionTargetID As Int32 = -1
	Public iMissionTargetTypeID As Int16 = -1
	Public yHealth As Byte = 100

	Public lInterrogatorID As Int32 = -1
	Public yInterrogationProgress As Byte = 0

End Class