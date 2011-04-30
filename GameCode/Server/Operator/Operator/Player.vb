Option Strict On

Public Enum AliasingRights As Int32
    eNoRights = 0
    eMoveUnits = 1
    eDockUndockUnits = 2
    eChangeBehavior = 4
    eAddProduction = 8
    eCancelProduction = 16
    eViewTechDesigns = 32
    eCreateBattleGroups = 64
    eViewBattleGroups = 128
    eModifyBattleGroups = 256
    eViewDiplomacy = 512
    eAlterDiplomacy = 1024
    eViewAgents = 2048
    eAlterAgents = 4096
    eViewColonyStats = 8192
    eAlterColonyStats = 16384
    eViewBudget = 32768
    eViewMining = 65536
    eViewEmail = 131072
    eAlterEmail = 262144
    eViewTreasury = 524288
    eViewTrades = 1048576
    eAlterTrades = 2097152
    eAddResearch = 4194304
    eCancelResearch = 8388608
    eChangeEnvironment = 16777216
    eViewResearch = 33554432
    eViewUnitsAndFacilities = 67108864
    eAlterAutoLaunchPower = 134217728
	eTransferCargo = 268435456
	eCreateDesigns = 536870912
End Enum
Public Structure AliasLogin
    Public yUserName() As Byte
    Public yPassword() As Byte
    Public lRights As Int32
End Structure
Public Enum AccountStatusType As Integer
    eInactiveAccount = 0
    eActiveAccount
    eBannedAccount
    eSuspendedAccount
    eOpenBetaAccount
    eTrialAccount = 99
    eMondelisInactive = 100
    eMondelisActive = 101

End Enum
Public Enum ePlayerSpecialAttributeSetting As Short
    eCelebrationPeriod = 0
    eInPirateSpawn = 1
    eNegativeCashflow = 2
    eDoctrineOfLeadership = 3
    eBadWarDecMoralePenalty = 4
    eBadWarDecCPPenalty = 5
    'TODO: add more here
    ePlayerTitle = 6
End Enum

Public Structure PrimaryTitleVal
    Public lPrimaryIdx As Int32
    Public yTitleValue As Byte
End Structure

Public Class Player
    Inherits Epica_GUID

    Public Enum PlayerRank As Byte
        Magistrate = 0
        Governor = 1
        Overseer = 2
        Duke = 3
        Baron = 4
        King = 5
        Emperor = 6
        ExRankShift = 128
    End Enum

    Public AccountStatus As Int32 = AccountStatusType.eInactiveAccount
    Public lPreviousStatus As Int32 = -1

    Public LastLogin As Date
    Public TotalPlayTime As Int32 = 0

    Public PlayerName(19) As Byte         'displayed to other players
    Public EmpireName(19) As Byte         'displayed to other players
    Public RaceName(19) As Byte           'displayed to other players

    Public oBudget As PlayerBudgetData = Nothing

#Region "  Alias System Lookups  "
    'Aliases that I am allowing to other players
    Public oAliases() As Player
    Public lAliasIdx() As Int32
    Public uAliasLogin() As AliasLogin
    Public lAliasUB As Int32 = -1

    'Aliases that I am allowed to use
    Public oAllowances() As Player
    Public lAllowanceIdx() As Int32
    Public uAllowanceLogin() As AliasLogin
    Public lAllowanceUB As Int32 = -1

    'Indicates the ID that the player is aliasing
    Public lAliasingPlayerID As Int32 = -1

    Public ReadOnly Property HasOnlineAliases(ByVal lRequiredRight As AliasingRights) As Boolean
		Get
			Try
				For X As Int32 = 0 To lAliasUB
					If lAliasIdx(X) <> -1 AndAlso (uAliasLogin(X).lRights And lRequiredRight) <> 0 AndAlso oAliases(X) Is Nothing = False AndAlso oAliases(X).lAliasingPlayerID = Me.ObjectID AndAlso oAliases(X).oSocket Is Nothing = False Then Return True
				Next X
				Return False
			Catch
				Return False
			End Try
		End Get
    End Property

	Public Sub AddPlayerAlias(ByRef oPlayer As Player, ByVal yUserName() As Byte, ByVal yPassword() As Byte, ByVal lRights As Int32)
		Try
			Dim lIdx As Int32 = -1
			Dim lFirstIdx As Int32 = -1

			'Ok, now, find a place in that for us
			For X As Int32 = 0 To lAliasUB
				If lAliasIdx(X) = oPlayer.ObjectID Then
					lIdx = X
					Exit For
				ElseIf lFirstIdx = -1 AndAlso lAliasIdx(X) = 1 Then
					lFirstIdx = X
				End If
			Next X
			If lIdx = -1 Then
				If lFirstIdx = -1 Then
					lIdx = lAliasUB + 1
					ReDim Preserve oAliases(lIdx)
					ReDim Preserve lAliasIdx(lIdx)
					ReDim Preserve uAliasLogin(lIdx)
					lAliasUB += 1
				Else : lIdx = lFirstIdx
				End If
			End If
			oAliases(lIdx) = oPlayer
			lAliasIdx(lIdx) = oPlayer.ObjectID
			With uAliasLogin(lIdx)
				ReDim .yUserName(19)
				ReDim .yPassword(19)
				Array.Copy(yUserName, 0, .yUserName, 0, 20)
				Array.Copy(yPassword, 0, .yPassword, 0, 20)
				.lRights = lRights
			End With
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "AddPlayerAlias: " & ex.Message)
		End Try
	End Sub
	Public Sub AddPlayerAliasAllowance(ByRef oPlayer As Player, ByVal yUserName() As Byte, ByVal yPassword() As Byte, ByVal lRights As Int32)
		Try
			Dim lIdx As Int32 = -1
			Dim lFirstIdx As Int32 = -1

			'Ok, now, find a place in that for us
			For X As Int32 = 0 To lAllowanceUB
				If lAllowanceIdx(X) = oPlayer.ObjectID Then
					lIdx = X
					Exit For
				ElseIf lFirstIdx = -1 AndAlso lAllowanceIdx(X) = 1 Then
					lFirstIdx = X
				End If
			Next X
			If lIdx = -1 Then
				If lFirstIdx = -1 Then
					lIdx = lAllowanceUB + 1
					ReDim Preserve oAllowances(lIdx)
					ReDim Preserve lAllowanceIdx(lIdx)
					ReDim Preserve uAllowanceLogin(lIdx)
					lAllowanceUB += 1
				Else : lIdx = lFirstIdx
				End If
			End If
			oAllowances(lIdx) = oPlayer
			lAllowanceIdx(lIdx) = oPlayer.ObjectID
			With uAllowanceLogin(lIdx)
				ReDim .yUserName(19)
				ReDim .yPassword(19)
				Array.Copy(yUserName, 0, .yUserName, 0, 20)
				Array.Copy(yPassword, 0, .yPassword, 0, 20)
				.lRights = lRights
			End With
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "AddPlayerAliasAllowance: " & ex.Message)
		End Try
	End Sub
#End Region

    Public ExternalEmailAddress(254) As Byte        'sent to email srvr
    
    Public sPlayerName As String
    Public sPlayerNameProper As String
    Public lPlayerIcon As Int32
    Public yGender As Byte
    'NOTE: THE NEXT TWO ITEMS (USERNAME AND PASSWORD) ARE TO NEVER BE TRANSMITTED!!!!
    Public PlayerUserName(19) As Byte
    Public PlayerPassword(19) As Byte

    Public oSocket As NetSock           'the socket this user is connected on

    Public lLastViewedEnvir As Int32
    Public iLastViewedEnvirType As Int16
    Public lStartedEnvirID As Int32
    Public iStartedEnvirTypeID As Int16
    Public lStartLocX As Int32 = Int32.MinValue
    Public lStartLocZ As Int32 = Int32.MinValue

    Public blCredits As Int64            'credits the player currently has 
  
    'Always available
    'Public oBudget As Budget = New Budget()

    Public swInDyingProcess As Stopwatch = Nothing
    Public lPrimaryReturned() As Int32      'index of the primary returning the death msg

    Public PirateStartLocX As Int32 = Int32.MinValue
    Public PirateStartLocZ As Int32 = Int32.MinValue

    Public yPlayerTitle As Byte = 0
    Public lControlledPlanetIdx() As Int32
    Public lControlledPlanetUB As Int32 = -1
    Public lOwnerPrimaryIdx As Int32 = -1
    Public lConnectedPrimaryIdx As Int32 = -1

    Public lPlayerScoreRequestCnt As Int32 = 0
    Public lTechnologyScore As Int32 = 0
    Public lDiplomacyScore As Int32 = 0
    Public lPopulationScore As Int32 = 0
    Public lMilitaryScore As Int32 = 0
    Public lProductionScore As Int32 = 0
    Public lWealthScore As Int32 = 0
    Public lTotalScore As Int32 = 0

#Region "  Attached Primary Server Management  "
    Public lAttachedPrimaryIdx() As Int32
    Public lAttachedPrimaryUB As Int32 = -1
    Public uTitleVals(-1) As PrimaryTitleVal
    Public Sub PrimaryReportingTitleValueChange(ByVal lPrimaryIdx As Int32, ByVal yNewValue As Byte)
        SyncLock uTitleVals
            Dim bFound As Boolean = False
            For X As Int32 = 0 To uTitleVals.GetUpperBound(0)
                If uTitleVals(X).lPrimaryIdx = lPrimaryIdx Then
                    uTitleVals(X).yTitleValue = yNewValue
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False Then
                ReDim Preserve uTitleVals(uTitleVals.GetUpperBound(0) + 1)
                With uTitleVals(uTitleVals.GetUpperBound(0))
                    .yTitleValue = yNewValue
                    .lPrimaryIdx = lPrimaryIdx
                End With
            End If
        End SyncLock

        'Now, we retest our title...
        Dim yStart As Byte = Me.yPlayerTitle
        If (yStart And Player.PlayerRank.ExRankShift) <> 0 Then
            yStart = yStart Xor Player.PlayerRank.ExRankShift
        End If
        Dim yMax As Byte = 0
        Dim lUB As Int32 = uTitleVals.GetUpperBound(0)
        For X As Int32 = 0 To lUB
            Dim yTemp As Byte = uTitleVals(X).yTitleValue
            If (yTemp And Player.PlayerRank.ExRankShift) <> 0 Then
                yTemp = yTemp Xor Player.PlayerRank.ExRankShift
            End If
            If yTemp > yMax Then yMax = yTemp
        Next X

        Me.yPlayerTitle = yMax

        'Ok, is our max value higher than oru start?
        If yMax > yStart OrElse ((Me.yPlayerTitle And Player.PlayerRank.ExRankShift) <> 0 AndAlso yStart = yMax) Then
            'yes, ok, promotion time/end of ex title time
            Dim yMsg(11) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.ePlayerTitle).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(CInt(yNewValue)).CopyTo(yMsg, lPos) : lPos += 4
            SendToAttachedPrimaries(yMsg)
        End If
    End Sub

    Public Sub SendToAttachedPrimaries(ByVal yMsg() As Byte)
        If Me.lOwnerPrimaryIdx > -1 AndAlso Me.lOwnerPrimaryIdx <= goMsgSys.mlServerUB Then
            goMsgSys.SendToServerIndex(Me.lOwnerPrimaryIdx, yMsg)
        End If
        For X As Int32 = 0 To Me.lAttachedPrimaryUB
            If Me.lAttachedPrimaryIdx(X) > -1 AndAlso Me.lAttachedPrimaryIdx(X) <= goMsgSys.mlServerUB Then
                goMsgSys.SendToServerIndex(Me.lAttachedPrimaryIdx(X), yMsg)
            End If
        Next X
    End Sub

#End Region

    Private mbPlayerIsDead As Boolean = False
    Public Property PlayerIsDead() As Boolean
        Get
            Return mbPlayerIsDead
        End Get
        Set(ByVal value As Boolean)
            If value <> mbPlayerIsDead Then
                mbPlayerIsDead = value
                If oSocket Is Nothing = False Then
                    Dim yMsg(2) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yMsg, 0)
                    If mbPlayerIsDead = True Then yMsg(2) = 1 Else yMsg(2) = 0
                    oSocket.SendData(yMsg)
                End If
            End If
        End Set
    End Property

    Private mySendString() As Byte
   
    '#Region "  Player Communication (Email and Socket)  "
    '    Public EmailFolders() As PlayerCommFolder
    '    Public EmailFolderIdx() As Int32
    '    Public EmailFolderUB As Int32 = -1

    '    Private muEnvirAlerts() As EnvirAlert
    '    Private mlEnvirAlertUB As Int32 = -1

    '    Public Function GetOrAddEmailFolder(ByVal lPCF_ID As Int32) As Int32
    '        Dim lIdx As Int32 = -1

    '        'Ok, find our folder first...
    '        For X As Int32 = 0 To EmailFolderUB
    '            If EmailFolderIdx(X) = lPCF_ID Then
    '                lIdx = X
    '                Exit For
    '            End If
    '        Next X

    '        If lIdx = -1 Then
    '            Dim sName As String = "Unnamed"
    '            Select Case lPCF_ID
    '                Case PlayerCommFolder.ePCF_ID_HARDCODES.eDeleted_PCF
    '                    sName = "Deleted"
    '                Case PlayerCommFolder.ePCF_ID_HARDCODES.eDrafts_PCF
    '                    sName = "Drafts"
    '                Case PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF
    '                    sName = "Inbox"
    '                Case PlayerCommFolder.ePCF_ID_HARDCODES.eOutbox_PCF
    '                    sName = "Outbox"
    '            End Select
    '            'Now, add the folder
    '            lIdx = AddFolder(lPCF_ID, sName)
    '            If lIdx <> -1 Then EmailFolderIdx(lIdx) = lIdx
    '        End If
    '        Return lIdx
    '    End Function

    '    Public Function AddEmailMsg(ByVal lPC_ID As Int32, ByVal lPCF_ID As Int32, ByVal iEncryption As Short, ByVal sBody As String, ByVal sTitle As String, ByVal lSentByID As Int32, ByVal lSentOn As Int32, ByVal bRead As Boolean, ByVal sRecipients As String) As PlayerComm
    '        If Me.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then Return Nothing

    '        Dim lIdx As Int32 = GetOrAddEmailFolder(lPCF_ID)
    '        If lIdx <> -1 Then
    '            Return EmailFolders(lIdx).AddPlayerComm(lPC_ID, iEncryption, sBody, sTitle, lSentByID, lSentOn, bRead, sRecipients)
    '        End If
    '        Return Nothing
    '    End Function

    '    Public Function AddFolder(ByVal lBaseID As Int32, ByVal sName As String) As Int32
    '        Dim lIdx As Int32 = -1
    '        Dim lFirstIdx As Int32 = -1

    '        'Ok, first, see if the base ID is 0 or -1
    '        If lBaseID = 0 OrElse lBaseID = -1 Then
    '            'Ok, no base ID... so, we'll go through and look for an empty slot
    '            For X As Int32 = 0 To EmailFolderUB
    '                If EmailFolderIdx(X) = -1 Then
    '                    lFirstIdx = X
    '                    Exit For
    '                End If
    '            Next X
    '        Else
    '            'check if we already have it...
    '            For X As Int32 = 0 To EmailFolderUB
    '                If EmailFolderIdx(X) = lBaseID Then
    '                    lIdx = X
    '                    Exit For
    '                ElseIf lFirstIdx = -1 AndAlso EmailFolderIdx(X) = -1 Then
    '                    lFirstIdx = X
    '                End If
    '            Next X
    '        End If

    '        'Now, check our result
    '        If lIdx = -1 Then
    '            'Did not find it... did we find an empty slot?
    '            If lFirstIdx = -1 Then
    '                'No, so expand our array
    '                EmailFolderUB += 1
    '                ReDim Preserve EmailFolders(EmailFolderUB)
    '                ReDim Preserve EmailFolderIdx(EmailFolderUB)
    '                lIdx = EmailFolderUB
    '            Else : lIdx = lFirstIdx
    '            End If
    '            EmailFolders(lIdx) = New PlayerCommFolder()
    '        End If

    '        'Now, make sure the folder has our stuff
    '        If sName.Length > 20 Then sName = sName.Substring(0, 20)

    '        If EmailFolders(lIdx) Is Nothing Then EmailFolders(lIdx) = New PlayerCommFolder
    '        With EmailFolders(lIdx)
    '            StringToBytes(sName).CopyTo(.FolderName, 0)
    '            If lBaseID <> 0 AndAlso lBaseID <> -1 Then .PCF_ID = lBaseID Else .PCF_ID = -1
    '            .PlayerID = Me.ObjectID
    '            If .PCF_ID = -1 Then
    '                .SaveFolder()
    '            End If
    '            EmailFolderIdx(lIdx) = .PCF_ID
    '        End With

    '        Return lIdx
    '    End Function

    '    Public Function DeleteFolder(ByVal lPCF_ID As Int32) As Boolean
    '        If lPCF_ID > 0 Then
    '            For X As Int32 = 0 To EmailFolderUB
    '                If EmailFolderIdx(X) = lPCF_ID Then
    '                    EmailFolderIdx(X) = -1
    '                    EmailFolders(X) = Nothing
    '                    Return True
    '                End If
    '            Next X
    '        End If
    '        Return False
    '    End Function

    '    Public Sub RemoveEmailMsg(ByVal lMsgID As Int32, ByVal lPCF_ID As Int32)
    '        Dim lFolderIdx As Int32 = -1
    '        Dim lTrashFolderIdx As Int32 = GetOrAddEmailFolder(PlayerCommFolder.ePCF_ID_HARDCODES.eDeleted_PCF)

    '        For X As Int32 = 0 To EmailFolderUB
    '            If EmailFolderIdx(X) = lPCF_ID Then
    '                lFolderIdx = X
    '                Exit For
    '            End If
    '        Next X

    '        If lFolderIdx <> -1 Then
    '            With EmailFolders(lFolderIdx)

    '                'Ok, are we removing this message from the trash can?
    '                If lPCF_ID = PlayerCommFolder.ePCF_ID_HARDCODES.eDeleted_PCF Then
    '                    'yes, so, delete it
    '                    For X As Int32 = 0 To .PlayerMsgsUB
    '                        If .PlayerMsgsIdx(X) = lMsgID Then
    '                            'Ok, found the message...
    '                            .PlayerMsgs(X).DeleteComm()
    '                            .PlayerMsgsIdx(X) = -1
    '                            .PlayerMsgs(X) = Nothing
    '                            Exit For
    '                        End If
    '                    Next X
    '                Else
    '                    'No, so put the message in the trash can
    '                    .MoveMessageToFolder(lMsgID, EmailFolders(lTrashFolderIdx))
    '                End If
    '            End With
    '        End If
    '    End Sub

    '    Public Sub SendPlayerMessage(ByRef yMsg() As Byte, ByVal bAddEmail As Boolean, ByVal lRights As AliasingRights)
    '        If Me.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then Return

    '        For X As Int32 = 0 To lAliasUB
    '            Try
    '                If lAliasIdx(X) <> -1 AndAlso (lRights = 0 OrElse (uAliasLogin(X).lRights And lRights) <> 0) AndAlso oAliases(X) Is Nothing = False AndAlso oAliases(X).lAliasingPlayerID = Me.ObjectID AndAlso oAliases(X).oSocket Is Nothing = False Then
    '                    oAliases(X).oSocket.SendData(yMsg)
    '                End If
    '            Catch
    '            End Try
    '        Next X

    '        If oSocket Is Nothing = False Then
    '            oSocket.SendData(yMsg)
    '        ElseIf bAddEmail = True Then
    '            Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yMsg, 0)

    '            Dim sBody As String
    '            Dim sTitle As String
    '            Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder
    '            Dim lPos As Int32 = 2

    '            Select Case iMsgCode
    '                Case EpicaMessageCode.eAddObjectCommand
    '                    Dim lObjID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
    '                    If (iEmailSettings And eEmailSettings.eResearchComplete) <> 0 Then
    '                        Select Case iObjTypeID
    '                            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
    '                                Dim oTech As Epica_Tech = Me.GetTech(lObjID, iObjTypeID)
    '                                If oTech Is Nothing = False Then
    '                                    oSB.Length = 0
    '                                    oSB.AppendLine("Your scientists want to inform you")
    '                                    oSB.AppendLine("that research has been completed on")

    '                                    Dim sTypeStr As String = ""
    '                                    Select Case iObjTypeID
    '                                        Case ObjectType.eAlloyTech : sTypeStr = " (Alloy)"
    '                                        Case ObjectType.eArmorTech : sTypeStr = " (Armor)"
    '                                        Case ObjectType.eEngineTech : sTypeStr = " (Engine)"
    '                                        Case ObjectType.eHullTech : sTypeStr = " (Hull)"
    '                                        Case ObjectType.ePrototype : sTypeStr = " (Prototype)"
    '                                        Case ObjectType.eRadarTech : sTypeStr = " (Radar)"
    '                                        Case ObjectType.eShieldTech : sTypeStr = " (Shield)"
    '                                        Case ObjectType.eSpecialTech : sTypeStr = " (Special Project)"
    '                                        Case ObjectType.eWeaponTech : sTypeStr = " (Weapon)"
    '                                    End Select
    '                                    oSB.AppendLine(BytesToString(oTech.GetTechName) & sTypeStr)
    '                                    sBody = oSB.ToString
    '                                    sTitle = "Research Complete"

    '                                    Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName))
    '                                    If oPC Is Nothing = False Then goMsgSys.SendOutboundEmail(oPC, Me, iMsgCode, 0, 0, 0, 0, 0, 0, "")
    '                                End If
    '                            Case Else
    '                                'TODO: what else? for now return
    '                                Return
    '                        End Select
    '                    End If
    '                Case EpicaMessageCode.eSetPlayerRel
    '                    oSB.Length = 0
    '                    oSB.AppendLine("Player Relationship Altered" & vbCrLf)
    '                    lPos = 2
    '                    Dim lPlayer1ID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim lPlayer2ID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim yScore As Byte = yMsg(lPos) : lPos += 1
    '                    Dim lExt1 As Int32
    '                    Dim bSendExt As Boolean = (iEmailSettings And eEmailSettings.ePlayerRels) <> 0

    '                    If Me.ObjectID = lPlayer1ID Then
    '                        Dim oPlayer As Player = GetEpicaPlayer(lPlayer2ID)
    '                        lExt1 = lPlayer2ID
    '                        If oPlayer Is Nothing = False Then
    '                            oSB.AppendLine("New Relationship Value with " & oPlayer.sPlayerName & ": " & yScore)
    '                        End If
    '                    Else
    '                        Dim oPlayer As Player = GetEpicaPlayer(lPlayer1ID)
    '                        lExt1 = lPlayer1ID
    '                        If oPlayer Is Nothing = False Then
    '                            oSB.AppendLine("New Relationship Value with " & oPlayer.sPlayerName & ": " & yScore)
    '                        End If
    '                    End If

    '                    Dim oSBExtra As New System.Text.StringBuilder()
    '                    oSBExtra.Length = 0
    '                    If bSendExt = True Then
    '                        oSBExtra.AppendLine()
    '                        oSBExtra.AppendLine("How would you like to respond?" & vbCrLf)
    '                        oSBExtra.AppendLine("Set Relationship to <enter value>")
    '                    End If

    '                    sBody = oSB.ToString
    '                    sTitle = "Foreign Policy Change"
    '                    Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper)
    '                    If oPC Is Nothing = False AndAlso bSendExt = True Then goMsgSys.SendOutboundEmail(oPC, Me, iMsgCode, lExt1, yScore, 0, 0, 0, 0, oSBExtra.ToString)
    '                Case EpicaMessageCode.eRequestDockFail
    '                    oSB.Length = 0
    '                    oSB.AppendLine("Docking Request Rejected")
    '                    oSB.AppendLine()


    '                    Dim lObjID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
    '                    Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
    '                    lPos += 8       'loc
    '                    Dim yReason As Byte = yMsg(lPos) : lPos += 1

    '                    'Now, construct the rest
    '                    If iEnvirTypeID = ObjectType.ePlanet Then
    '                        Dim oP As Planet = GetEpicaPlanet(lEnvirID)
    '                        If oP Is Nothing = False Then
    '                            oSB.AppendLine("Location: " & BytesToString(oP.PlanetName) & " (Planet side)")
    '                        End If
    '                    ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
    '                        Dim oS As SolarSystem = GetEpicaSystem(lEnvirID)
    '                        If oS Is Nothing = False Then
    '                            oSB.AppendLine("Location: " & BytesToString(oS.SystemName) & " (System)")
    '                        End If
    '                    End If
    '                    Select Case yReason
    '                        Case DockRejectType.eDockeeMoving
    '                            oSB.AppendLine("Reason: Dock Target was Moving")
    '                        Case DockRejectType.eDockerCannotEnterEnvir
    '                            oSB.AppendLine("Reason: Undock Environment was inhospitable")
    '                        Case DockRejectType.eDoorSizeExceeded
    '                            oSB.AppendLine("Reason: Target's Hangar Bay Doors were not big enough")
    '                        Case DockRejectType.eHangarInoperable
    '                            oSB.AppendLine("Reason: Target's Hangar was inoperable")
    '                        Case DockRejectType.eInsufficientHangarCap
    '                            oSB.AppendLine("Reason: Target's Hangar Capacity was Insufficient")
    '                    End Select

    '                    sBody = oSB.ToString
    '                    sTitle = "Docking Request Rejection"

    '                    Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName))
    '                Case EpicaMessageCode.eColonyLowResources
    '                    Dim yTemp(19) As Byte
    '                    Array.Copy(yMsg, lPos, yTemp, 0, 20)
    '                    lPos += 20
    '                    Dim sColonyName As String = BytesToString(yTemp)
    '                    Dim lColonyID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    lPos += 2       'colony guid
    '                    Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
    '                    Dim yReason As Byte = yMsg(lPos) : lPos += 1

    '                    Dim lItemID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim iTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2

    '                    Dim lProdID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2

    '                    Dim bSendExt As Boolean = (iEmailSettings And eEmailSettings.eLowResources) <> 0

    '                    Dim lExt1 As Int32 = iTypeID

    '                    oSB.Length = 0
    '                    Select Case yReason
    '                        Case ProductionType.eColonists
    '                            oSB.AppendLine("Insufficient Colonists")
    '                            lExt1 = Int32.MinValue
    '                        Case ProductionType.eEnlisted
    '                            oSB.AppendLine("Insufficient Enlisted")
    '                            lExt1 = ObjectType.eEnlisted
    '                        Case ProductionType.eFood
    '                            oSB.AppendLine("Insufficient Food Supplies")
    '                            lExt1 = Int32.MinValue
    '                        Case ProductionType.eOfficers
    '                            oSB.AppendLine("Insufficient Officers")
    '                            lExt1 = ObjectType.eOfficers
    '                        Case ProductionType.ePowerCenter
    '                            oSB.AppendLine("Insufficient Power")
    '                            lExt1 = Int32.MinValue
    '                        Case ProductionType.eCredits
    '                            oSB.AppendLine("Insufficient Credits")
    '                            lExt1 = Int32.MinValue
    '                        Case ProductionType.eWareHouse
    '                            oSB.AppendLine("Insufficient Cargo Capacity (Colony-wide)")
    '                            lExt1 = Int32.MinValue
    '                        Case ProductionType.eProduction
    '                            oSB.AppendLine("Insufficient Hangar Capacity")
    '                            lExt1 = Int32.MinValue
    '                        Case Else
    '                            oSB.AppendLine("Insufficient Resources")

    '                            Dim oObj As Object = GetEpicaObject(lItemID, iTypeID)
    '                            If oObj Is Nothing = False Then
    '                                oSB.AppendLine("Not enough " & BytesToString(GetEpicaObjectName(iTypeID, oObj)))
    '                                If CType(oObj, Epica_GUID).ObjTypeID = ObjectType.eMineral Then
    '                                    If CType(oObj, Mineral).lAlloyTechID <> -1 Then
    '                                        lExt1 = ObjectType.eAlloyTech
    '                                    End If
    '                                End If
    '                            End If
    '                            oObj = Nothing
    '                    End Select
    '                    oSB.AppendLine()
    '                    If sColonyName <> "" Then
    '                        oSB.AppendLine("Colony: " & sColonyName)
    '                    End If
    '                    If iEnvirTypeID = ObjectType.ePlanet Then
    '                        Dim oP As Planet = GetEpicaPlanet(lEnvirID)
    '                        If oP Is Nothing = False Then
    '                            oSB.AppendLine("Location: " & BytesToString(oP.PlanetName) & " (Planet side)")
    '                        End If
    '                    ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
    '                        Dim oS As SolarSystem = GetEpicaSystem(lEnvirID)
    '                        If oS Is Nothing = False Then
    '                            oSB.AppendLine("Location: " & BytesToString(oS.SystemName) & " (System)")
    '                        End If
    '                    End If

    '                    If iProdTypeID = ObjectType.eEnlisted Then
    '                        oSB.AppendLine("Attempting to Build Enlisted")
    '                    ElseIf iProdTypeID = ObjectType.eOfficers Then
    '                        oSB.AppendLine("Attempting to Build Officers")
    '                    ElseIf iProdTypeID = ObjectType.eUnitDef OrElse iProdTypeID = ObjectType.eFacilityDef Then
    '                        Dim oUnit As Epica_Entity_Def = GetEpicaUnitDef(lProdID)
    '                        If oUnit Is Nothing = False Then
    '                            oSB.AppendLine("Attempting to Build " & BytesToString(oUnit.DefName))
    '                        ElseIf iProdTypeID = ObjectType.eUnitDef Then
    '                            oSB.AppendLine("Attempting to Build a Unit")
    '                        Else : oSB.AppendLine("Attempting to Build a Facility")
    '                        End If
    '                    Else
    '                        Select Case iProdTypeID
    '                            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHangarTech, ObjectType.eHullTech, ObjectType.eMineralTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eSpecialTech
    '                                Dim oTech As Epica_Tech = GetTech(lProdID, iProdTypeID)
    '                                If oTech Is Nothing = False Then
    '                                    oSB.AppendLine("Attempting to Build a " & BytesToString(oTech.GetTechName))
    '                                Else : oSB.AppendLine("Attempting to Build a Component")
    '                                End If
    '                        End Select
    '                    End If

    '                    Dim oSBExtra As New System.Text.StringBuilder
    '                    oSBExtra.Length = 0
    '                    If bSendExt = True AndAlso lExt1 <> Int32.MinValue Then
    '                        oSBExtra.AppendLine()
    '                        oSBExtra.AppendLine("How would you like to respond?" & vbCrLf)
    '                        oSBExtra.AppendLine("Build/Gather <enter number>")
    '                    End If

    '                    sBody = oSB.ToString
    '                    sTitle = "Low Resources"

    '                    Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName))
    '                    If oPC Is Nothing = False AndAlso bSendExt = True Then goMsgSys.SendOutboundEmail(oPC, Me, iMsgCode, iTypeID, lColonyID, lItemID, 0, 0, 0, oSBExtra.ToString)
    '                Case EpicaMessageCode.ePlayerAlert
    '                    'MsgCode, Type(1), EntityGUID(6), EnvirGUID(6), PlayerID(4), EnemyID(4), Optional Name(20)
    '                    Dim yType As Byte = yMsg(lPos) : lPos += 1
    '                    Dim lObjID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
    '                    Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                    Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
    '                    lPos += 4       'playerid
    '                    Dim lEnemyID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4

    '                    'Now, we store this specifically in an muEnvirAlert object...
    '                    Dim lIdx As Int32 = -1
    '                    For X As Int32 = 0 To mlEnvirAlertUB
    '                        With muEnvirAlerts(X)
    '                            If .EnvirID = lEnvirID AndAlso .EnvirTypeID = iEnvirTypeID AndAlso .AlertType = yType Then
    '                                lIdx = X
    '                                Exit For
    '                            End If
    '                        End With
    '                    Next X

    '                    If lIdx = -1 Then
    '                        mlEnvirAlertUB += 1
    '                        ReDim Preserve muEnvirAlerts(mlEnvirAlertUB)
    '                        lIdx = mlEnvirAlertUB
    '                        With muEnvirAlerts(lIdx)
    '                            .EnvirID = lEnvirID
    '                            .EnvirTypeID = iEnvirTypeID
    '                            .AlertType = yType
    '                            .lNameUB = -1
    '                        End With
    '                    End If

    '                    'Now, check our type
    '                    Select Case yType
    '                        Case PlayerAlertType.eColonyLost, PlayerAlertType.eFacilityLost, PlayerAlertType.eUnitLost
    '                            'Ok, for these, a name will be appended to the message
    '                            Dim yTemp(19) As Byte
    '                            Array.Copy(yMsg, lPos, yTemp, 0, 20)
    '                            lPos += 20
    '                            Dim sName As String = BytesToString(yTemp)

    '                            Dim lLocX As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                            Dim lLocZ As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4

    '                            muEnvirAlerts(lIdx).AddName(sName, lLocX, lLocZ)

    '                            If muEnvirAlerts(lIdx).OriginalAlert = 0 Then
    '                                muEnvirAlerts(lIdx).OriginalAlert = GetDateAsNumber(Now)
    '                            End If
    '                        Case PlayerAlertType.eEngagedEnemy, PlayerAlertType.eUnderAttack, PlayerAlertType.eBuyOrderAccepted

    '                            If muEnvirAlerts(lIdx).EmailSent = False Then
    '                                'Ok, let's construct our email
    '                                oSB.Length = 0

    '                                Dim oPlayer As Player = GetEpicaPlayer(lEnemyID)

    '                                Dim bSendExtEmail As Boolean = False
    '                                Dim oSBExtra As New System.Text.StringBuilder
    '                                oSBExtra.Length = 0

    '                                If yType = PlayerAlertType.eEngagedEnemy Then
    '                                    If oPlayer Is Nothing = False Then
    '                                        oSB.AppendLine("Our forces are reporting an engagement with the enemy! (" & BytesToString(oPlayer.PlayerName) & ")")
    '                                    Else : oSB.AppendLine("Our forces are reporting an engagement with the enemy!")
    '                                    End If
    '                                    sTitle = "Enemy Engaged"

    '                                    bSendExtEmail = (iEmailSettings And eEmailSettings.eEngagedEnemy) <> 0
    '                                ElseIf yType = PlayerAlertType.eBuyOrderAccepted Then
    '                                    If oPlayer Is Nothing = False Then
    '                                        oSB.AppendLine(BytesToString(oPlayer.PlayerName) & " has accepted one of our buy order.")
    '                                    Else : oSB.AppendLine("One of our buy orders have been accepted.")
    '                                    End If
    '                                    sTitle = "Buy Order Accepted"
    '                                    bSendExtEmail = (iEmailSettings And eEmailSettings.eBuyOrderAccepted) <> 0
    '                                Else
    '                                    If oPlayer Is Nothing = False Then
    '                                        oSB.AppendLine(BytesToString(oPlayer.PlayerName) & " is attacking our forces!")
    '                                    Else : oSB.AppendLine("Our forces are under attack!")
    '                                    End If
    '                                    sTitle = "We Are Under Attack"

    '                                    bSendExtEmail = (iEmailSettings And eEmailSettings.eUnderAttack) <> 0

    '                                    If bSendExtEmail = True Then
    '                                        oSBExtra.AppendLine()
    '                                        oSBExtra.AppendLine("How would you like to respond?" & vbCrLf)

    '                                        For X As Int32 = 0 To glUnitGroupUB
    '                                            If glUnitGroupIdx(X) <> -1 Then
    '                                                'AndAlso goUnitGroup(X).oOwner.ObjectID = Me.ObjectID Then
    '                                                Try
    '                                                    Dim oGroup As UnitGroup = goUnitGroup(X)
    '                                                    If oGroup.oOwner.ObjectID = Me.ObjectID Then
    '                                                        oSBExtra.AppendLine("Attack Using Battlegroup: " & BytesToString(oGroup.UnitGroupName))
    '                                                    End If
    '                                                Catch
    '                                                End Try
    '                                            End If
    '                                        Next X

    '                                        oSBExtra.AppendLine("Launch To Attack")
    '                                        oSBExtra.AppendLine("Attack Using All")
    '                                        oSBExtra.AppendLine("Attack Using Half")
    '                                    End If
    '                                End If

    '                                If yType <> PlayerAlertType.eBuyOrderAccepted Then
    '                                    If iEnvirTypeID = ObjectType.ePlanet Then
    '                                        Dim oP As Planet = GetEpicaPlanet(lEnvirID)
    '                                        If oP Is Nothing = False Then
    '                                            oSB.AppendLine("Location: " & BytesToString(oP.PlanetName) & " (Planet side)")
    '                                        End If
    '                                    ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
    '                                        Dim oS As SolarSystem = GetEpicaSystem(lEnvirID)
    '                                        If oS Is Nothing = False Then
    '                                            oSB.AppendLine("Location: " & BytesToString(oS.SystemName) & " (System)")
    '                                        End If
    '                                    End If
    '                                Else
    '                                    oSB.AppendLine(goGTC.GetBuyOrderAcceptedEmailBody(lObjID))
    '                                End If
    '                                sBody = oSB.ToString

    '                                muEnvirAlerts(lIdx).EmailSent = True

    '                                Dim lLocX As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
    '                                Dim lLocZ As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4

    '                                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName))
    '                                If oPC Is Nothing = False Then
    '                                    oPC.AddEmailAttachment(0, lEnvirID, iEnvirTypeID, lLocX, lLocZ, "Event Location")
    '                                    If bSendExtEmail = True Then

    '                                        'If iObjTypeID = ObjectType.eUnit Then
    '                                        '    Dim oUnit As Unit = GetEpicaUnit(lObjID)
    '                                        '    If oUnit Is Nothing = False Then
    '                                        '        lLocX = oUnit.LocX
    '                                        '        lLocZ = oUnit.LocZ
    '                                        '    End If
    '                                        'ElseIf iObjTypeID = ObjectType.eFacility Then
    '                                        '    Dim oFacility As Facility = GetEpicaFacility(lObjID)
    '                                        '    If oFacility Is Nothing = False Then
    '                                        '        lLocX = oFacility.LocX
    '                                        '        lLocZ = oFacility.LocZ
    '                                        '    End If
    '                                        'End If

    '                                        goMsgSys.SendOutboundEmail(oPC, Me, CShort(yType), yType, lEnemyID, lEnvirID, iEnvirTypeID, lLocX, lLocZ, oSBExtra.ToString)
    '                                    End If
    '                                End If

    '                            End If
    '                    End Select

    '                    'TODO: However, if the player has msg alerts set up for external devices, we need to send them now

    '                Case Else
    '                    'TODO: Queue the message up to send to the player when they come online... for now, return
    '            End Select
    '        End If
    '    End Sub

    '    Public Sub CreateAndSendPlayerAlert(ByVal yType As PlayerAlertType, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lEnemyID As Int32, ByVal sName As String, ByVal lLocX As Int32, ByVal lLocZ As Int32)
    '        'MsgCode, Type(1), EntityGUID(6), EnvirGUID(6), PlayerID(4), EnemyID (4) , Name (20), LocX (4), LocZ (4)
    '        Dim yMsg(50) As Byte
    '        Dim lPos As Int32 = 0
    '        System.BitConverter.GetBytes(EpicaMessageCode.ePlayerAlert).CopyTo(yMsg, lPos) : lPos += 2
    '        yMsg(lPos) = yType : lPos += 1
    '        System.BitConverter.GetBytes(lEntityID).CopyTo(yMsg, lPos) : lPos += 4
    '        System.BitConverter.GetBytes(iEntityTypeID).CopyTo(yMsg, lPos) : lPos += 2
    '        System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
    '        System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
    '        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
    '        System.BitConverter.GetBytes(lEnemyID).CopyTo(yMsg, lPos) : lPos += 4

    '        If sName Is Nothing = False AndAlso sName <> "" Then
    '            Dim sTemp As String
    '            If sName.Length > 20 Then sTemp = sName.Substring(0, 20) Else sTemp = sName
    '            StringToBytes(sTemp).CopyTo(yMsg, lPos)
    '        End If
    '        lPos += 20

    '        System.BitConverter.GetBytes(lLocX).CopyTo(yMsg, lPos) : lPos += 4
    '        System.BitConverter.GetBytes(lLocZ).CopyTo(yMsg, lPos) : lPos += 4

    '        Me.SendPlayerMessage(yMsg, True, 0)
    '    End Sub

    '    Public Sub FinalizeEnvirAlerts()
    '        'Ok, go through all environment alerts
    '        For X As Int32 = 0 To mlEnvirAlertUB
    '            If muEnvirAlerts(X).EmailSent = False Then
    '                With muEnvirAlerts(X)
    '                    Dim sBody As String
    '                    Dim sTitle As String

    '                    Dim oSB As New System.Text.StringBuilder

    '                    If .EnvirTypeID = ObjectType.ePlanet Then
    '                        Dim oP As Planet = GetEpicaPlanet(.EnvirID)
    '                        If oP Is Nothing = False Then
    '                            oSB.AppendLine("Location: " & BytesToString(oP.PlanetName) & " (Planet side)")
    '                        End If
    '                    ElseIf .EnvirTypeID = ObjectType.eSolarSystem Then
    '                        Dim oS As SolarSystem = GetEpicaSystem(.EnvirID)
    '                        If oS Is Nothing = False Then
    '                            oSB.AppendLine("Location: " & BytesToString(oS.SystemName) & " (System)")
    '                        End If
    '                    End If

    '                    Select Case .AlertType
    '                        Case PlayerAlertType.eColonyLost
    '                            oSB.AppendLine("The following Colonies were lost:")
    '                            sTitle = "Colony Lost"
    '                        Case PlayerAlertType.eFacilityLost
    '                            oSB.AppendLine("The following Facilities were lost:")
    '                            sTitle = "Facility Lost"
    '                        Case PlayerAlertType.eUnitLost
    '                            oSB.AppendLine("The following Units were lost:")
    '                            sTitle = "Unit Lost"
    '                        Case Else
    '                            Continue For
    '                    End Select

    '                    oSB.AppendLine()
    '                    oSB.Append(.GetNameList)

    '                    sBody = oSB.ToString

    '                    Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, .OriginalAlert, False, BytesToString(Me.PlayerName))
    '                    If oPC Is Nothing = False Then
    '                        For Y As Int32 = 0 To Math.Min(.lNameUB, 254)
    '                            Dim sName As String = .GetNameItem(Y)
    '                            If sName <> "" Then
    '                                Dim ptLoc As Point = .GetNameLoc(Y)
    '                                If ptLoc <> Point.Empty Then
    '                                    oPC.AddEmailAttachment(CByte(Y + 1), .EnvirID, .EnvirTypeID, ptLoc.X, ptLoc.Y, sName & " Loc")
    '                                End If
    '                            End If
    '                        Next Y
    '                        oPC.SaveObject()
    '                    End If
    '                End With
    '            End If
    '        Next X

    '    End Sub

    '    Private Structure EnvirAlert
    '        Public EnvirID As Int32
    '        Public EnvirTypeID As Int16
    '        Public AlertType As Byte

    '        Public EmailSent As Boolean

    '        Public OriginalAlert As Int32

    '        Private msNames() As String
    '        Private mlLocX() As Int32
    '        Private mlLocZ() As Int32
    '        Public lNameUB As Int32

    '        Public Sub AddName(ByVal sName As String, ByVal lLocX As Int32, ByVal lLocZ As Int32)
    '            If msNames Is Nothing Then
    '                ReDim msNames(-1)
    '                lNameUB = -1
    '            End If

    '            lNameUB += 1
    '            ReDim Preserve msNames(lNameUB)
    '            ReDim Preserve mlLocX(lNameUB)
    '            ReDim Preserve mlLocZ(lNameUB)
    '            msNames(lNameUB) = Mid$(sName, 1, 20)
    '            mlLocX(lNameUB) = lLocX
    '            mlLocZ(lNameUB) = lLocZ
    '        End Sub

    '        Public Function GetNameList() As String
    '            Dim oSB As New System.Text.StringBuilder
    '            For X As Int32 = 0 To lNameUB
    '                oSB.AppendLine(msNames(X))
    '            Next X
    '            Return oSB.ToString
    '        End Function

    '        Public Function GetNameLoc(ByVal lIndex As Int32) As Point
    '            Try
    '                Return New Point(mlLocX(lIndex), mlLocZ(lIndex))
    '            Catch ex As Exception
    '                Return Point.Empty
    '            End Try
    '        End Function
    '        Public Function GetNameItem(ByVal lIndex As Int32) As String
    '            Try
    '                Return msNames(lIndex)
    '            Catch ex As Exception
    '                Return ""
    '            End Try
    '        End Function
    '    End Structure

    '    Public Sub AddEmailAttachment(ByVal PC_ID As Int32, ByVal PCF_ID As Int32, ByVal AttachNumber As Byte, ByVal EnvirID As Int32, ByVal EnvirTypeID As Int16, ByVal LocX As Int32, ByVal LocZ As Int32, ByVal sName As String)
    '        For X As Int32 = 0 To EmailFolderUB
    '            If EmailFolderIdx(X) = PCF_ID Then
    '                With EmailFolders(X)
    '                    For Y As Int32 = 0 To .PlayerMsgsUB
    '                        If .PlayerMsgsIdx(Y) = PC_ID Then
    '                            .PlayerMsgs(Y).AddEmailAttachment(AttachNumber, EnvirID, EnvirTypeID, LocX, LocZ, sName)
    '                            Exit For
    '                        End If
    '                    Next
    '                End With
    '                Exit For
    '            End If
    '        Next X
    '    End Sub
    '#End Region

    '#Region "  Chat Room Management  "
    '    'Player Chat Rooms
    '    Public ChatRooms() As Int32
    '    Public ChatRoomAlias() As String
    '    Public ChatRoomUB As Int32 = -1

    '    Public Sub AssociateChatRoom(ByVal lChatRoomID As Int32)
    '        Dim X As Int32
    '        Dim lIdx As Int32 = -1

    '        For X = 0 To ChatRoomUB
    '            If lIdx = -1 AndAlso ChatRooms(X) = -1 Then
    '                lIdx = X
    '            ElseIf ChatRooms(X) = lChatRoomID Then
    '                Exit Sub
    '            End If
    '        Next X

    '        If lIdx = -1 Then
    '            ChatRoomUB += 1
    '            ReDim Preserve ChatRooms(ChatRoomUB)
    '            ReDim Preserve ChatRoomAlias(ChatRoomUB)
    '            lIdx = ChatRoomUB
    '        End If

    '        ChatRooms(lIdx) = lChatRoomID
    '    End Sub

    '    Public Sub RemoveChatRoom(ByVal lChatRoomID As Int32)
    '        Dim X As Int32

    '        For X = 0 To ChatRoomUB
    '            If ChatRooms(X) = lChatRoomID Then
    '                ChatRooms(X) = -1
    '            End If
    '        Next X
    '    End Sub

    '    Public Sub AddChatRoomAlias(ByVal lChatRoomID As Int32, ByVal sAlias As String)
    '        Dim X As Int32

    '        For X = 0 To ChatRoomUB
    '            If ChatRooms(X) = lChatRoomID Then
    '                ChatRoomAlias(X) = UCase$(sAlias)
    '                Exit For
    '            End If
    '        Next X
    '    End Sub
    '#End Region

    '#Region "  Player Scores  "

    '    Public ReadOnly Property TechnologyScore() As Int32
    '        Get
    '            Dim lResult As Int32 = 0
    '            For X As Int32 = 0 To mlAlloyUB
    '                If mlAlloyIdx(X) <> -1 AndAlso moAlloy(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moAlloy(X).TechnologyScore
    '            Next X
    '            For X As Int32 = 0 To mlArmorUB
    '                If mlArmorIdx(X) <> -1 AndAlso moArmor(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moArmor(X).TechnologyScore
    '            Next X
    '            For X As Int32 = 0 To mlEngineUB
    '                If mlEngineIdx(X) <> -1 AndAlso moEngine(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moEngine(X).TechnologyScore
    '            Next X
    '            For X As Int32 = 0 To mlHullUB
    '                If mlHullIdx(X) <> -1 AndAlso moHull(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moHull(X).TechnologyScore
    '            Next X
    '            For X As Int32 = 0 To mlPrototypeUB
    '                If mlPrototypeIdx(X) <> -1 AndAlso moPrototype(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moPrototype(X).TechnologyScore
    '            Next X
    '            For X As Int32 = 0 To mlRadarUB
    '                If mlRadarIdx(X) <> -1 AndAlso moRadar(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moRadar(X).TechnologyScore
    '            Next X
    '            For X As Int32 = 0 To mlShieldUB
    '                If mlShieldIdx(X) <> -1 AndAlso moShield(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moShield(X).TechnologyScore
    '            Next X
    '            For X As Int32 = 0 To mlWeaponUB
    '                If mlWeaponIdx(X) <> -1 AndAlso moWeapon(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moWeapon(X).TechnologyScore
    '            Next X
    '            For X As Int32 = 0 To mlSpecialTechUB
    '                If mlSpecialTechIdx(X) <> -1 AndAlso moSpecialTech(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moSpecialTech(X).TechnologyScore
    '            Next X
    '            Return lResult \ 10 '100
    '        End Get
    '    End Property

    '    Public ReadOnly Property DiplomacyScore() As Int32
    '        Get
    '            Dim lResult As Int32 = 0

    '            For X As Int32 = 0 To Me.PlayerRelUB
    '                If Me.mlPlayerRelIdx(X) <> -1 Then
    '                    lResult += (CInt(moPlayerRels(X).WithThisScore) - 40I)
    '                End If
    '            Next X

    '            If lGuildID <> -1 AndAlso Me.oGuild Is Nothing = False Then
    '                If oGuild.yForeignPolicyShare = Guild.ForeignPolicyShare.eFullSharing Then
    '                    lResult += (oGuild.MemberCount * 1000)
    '                Else : lResult += 1000
    '                End If
    '            End If

    '            Return lResult * 20     'added the *20
    '        End Get
    '    End Property

    '    'Military Score needs to be managed on its own (in AddUnit and Entity.DeleteEntity)
    '    Public lMilitaryScore As Int32 = 0

    '    Public ReadOnly Property PopulationScore() As Int32
    '        Get
    '            Dim lResult As Int32 = 0
    '            For X As Int32 = 0 To mlColonyUB
    '                If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) <> -1 Then
    '                    lResult += (goColony(mlColonyIdx(X)).Population \ 10000)
    '                End If
    '            Next X
    '            Return lResult * 10     'added the x10
    '        End Get
    '    End Property

    '    Public ReadOnly Property ProductionScore() As Int32
    '        Get
    '            Dim lResult As Int32 = 0
    '            For X As Int32 = 0 To mlColonyUB
    '                If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) <> -1 Then
    '                    With goColony(mlColonyIdx(X))
    '                        For Y As Int32 = 0 To .ChildrenUB
    '                            If .lChildrenIdx(Y) <> -1 AndAlso (.oChildren(Y).yProductionType And ProductionType.eProduction) <> 0 Then
    '                                'lResult += .oChildren(Y).mlProdPoints
    '                                lResult += .oChildren(Y).EntityDef.ProdFactor
    '                            End If
    '                        Next Y
    '                    End With
    '                End If
    '            Next X
    '            Return lResult \ 5          'added the \5
    '        End Get
    '    End Property

    '    Public ReadOnly Property WealthScore() As Int32
    '        Get
    '            'Dim blCashFlow As Int64 = oBudget.GetCashFlow
    '            'Dim blVal As Int64 = (Me.blCredits \ (blCashFlow * 17280)) * 10000
    '            'If Me.blCredits < 0 AndAlso blCashFlow < 0 Then blVal = -(Math.Abs(blVal))

    '            If blCredits < 0 Then Return 0

    '            Dim dVal As Double = Math.Pow(blCredits, 0.3F) * 3      'added the x3
    '            If dVal > Int32.MaxValue Then Return Int32.MaxValue
    '            If dVal < Int32.MinValue Then Return Int32.MinValue
    '            Return CInt(dVal)
    '        End Get
    '    End Property

    '    Public ReadOnly Property TotalScore() As Int32
    '        Get
    '            Return (TechnologyScore + DiplomacyScore + PopulationScore + ProductionScore + WealthScore + (lMilitaryScore \ 100)) \ 6
    '        End Get
    '    End Property

    '    Public Sub ReTestTitle()
    '        'Ok, the layer by default has Magistrate
    '        Dim yLevel As Byte = 0

    '        Dim blPop As Int64 = 0
    '        Dim bSpaceColony As Boolean = False

    '        Dim lOwnedPlanets As Int32 = 0
    '        Dim lPlanetColonies As Int32 = 0

    '        Dim lSystemID() As Int32 = Nothing      'ID of the system (for grouping)
    '        Dim lSystemPCnt() As Int32 = Nothing    'Total planets in the system
    '        Dim lSystemOwned() As Int32 = Nothing   'Total planets in the system that I own
    '        Dim lSystemUB As Int32 = -1

    '        For X As Int32 = 0 To mlColonyUB
    '            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) <> -1 Then
    '                blPop += goColony(mlColonyIdx(X)).Population
    '                If CType(goColony(mlColonyIdx(X)).ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
    '                    bSpaceColony = True
    '                Else

    '                    lPlanetColonies += 1

    '                    Dim oPlanet As Planet = CType(goColony(mlColonyIdx(X)).ParentObject, Planet)
    '                    Dim bGood As Boolean = True
    '                    Dim lMyPop As Int32 = goColony(mlColonyIdx(X)).Population

    '                    If oPlanet Is Nothing = False Then

    '                        Dim lIdx As Int32 = -1
    '                        For Y As Int32 = 0 To lSystemUB
    '                            If oPlanet.ParentSystem.ObjectID = lSystemID(Y) Then
    '                                lIdx = Y
    '                                Exit For
    '                            End If
    '                        Next Y
    '                        If lIdx = -1 Then
    '                            lSystemUB += 1
    '                            ReDim Preserve lSystemID(lSystemUB)
    '                            ReDim Preserve lSystemPCnt(lSystemUB)
    '                            ReDim Preserve lSystemOwned(lSystemUB)
    '                            lSystemID(lSystemUB) = oPlanet.ParentSystem.ObjectID
    '                            lSystemPCnt(lSystemUB) = oPlanet.ParentSystem.mlPlanetUB + 1
    '                            lSystemOwned(lSystemUB) = 0
    '                            lIdx = lSystemUB
    '                        End If

    '                        Dim blTotalPop As Int64 = lMyPop
    '                        For Y As Int32 = 0 To oPlanet.lColonysHereUB
    '                            If oPlanet.lColonysHereIdx(Y) <> -1 AndAlso oPlanet.lColonysHereIdx(Y) <> mlColonyIdx(X) Then
    '                                If glColonyIdx(oPlanet.lColonysHereIdx(Y)) <> -1 Then
    '                                    blTotalPop += goColony(oPlanet.lColonysHereIdx(Y)).Population
    '                                End If
    '                            End If
    '                        Next Y

    '                        If blTotalPop = 0 Then
    '                            bGood = False
    '                        Else
    '                            bGood = CSng(lMyPop / blTotalPop) > 0.75F
    '                            goColony(mlColonyIdx(X)).GovScore = CByte(Math.Floor((lMyPop / blTotalPop) * 100.0F))
    '                        End If

    '                        If bGood = True Then
    '                            lOwnedPlanets += 1
    '                            lSystemOwned(lIdx) += 1
    '                        End If

    '                    End If
    '                End If
    '            End If
    '        Next X

    '        If blPop > 300000 Then yLevel = 1
    '        If blPop > 1000000 AndAlso bSpaceColony = True Then yLevel = 2
    '        If lOwnedPlanets > 0 AndAlso bSpaceColony = True AndAlso lPlanetColonies > 1 Then yLevel = 3
    '        If lOwnedPlanets > 1 AndAlso bSpaceColony = True AndAlso lPlanetColonies > 2 Then yLevel = 4

    '        If lOwnedPlanets > 0 Then
    '            For X As Int32 = 0 To lSystemUB
    '                If lSystemOwned(X) = lSystemPCnt(X) Then
    '                    yLevel = 6
    '                    Exit For
    '                Else
    '                    Dim fTemp As Single = CSng(lSystemOwned(X) / lSystemPCnt(X))
    '                    If fTemp > 0.5F Then yLevel = 5
    '                End If
    '            Next X
    '        End If

    '        If yPlayerTitle <> yLevel Then
    '            Dim yPrevTitle As Byte = yPlayerTitle

    '            yPlayerTitle = yLevel
    '            Dim yMsg(6) As Byte
    '            System.BitConverter.GetBytes(EpicaMessageCode.ePlayerTitleChange).CopyTo(yMsg, 0)
    '            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
    '            yMsg(6) = yPlayerTitle

    '            Me.SendPlayerMessage(yMsg, False, 0)

    '            If yPlayerTitle > yPrevTitle Then

    '                lCelebrationEnds = glCurrentCycle + 2592000

    '                Dim oSB As New System.Text.StringBuilder
    '                oSB.AppendLine(Me.sPlayerNameProper & ",")
    '                oSB.AppendLine()
    '                oSB.AppendLine("  In honor of your recent title recognition, festivities have begun at all major colonies! The festivities will last for 24 hours. This is only good news for our civilization and we are honored to be alive during your reign.")
    '                oSB.AppendLine()
    '                oSB.AppendLine()
    '                oSB.AppendLine("Your Loyal Citizens and Servants")
    '                Dim sSubject As String = "Title Promotion to " & GetPlayerTitle(yPlayerTitle, Me.yGender <> 2) & " from " & GetPlayerTitle(yPrevTitle, Me.yGender <> 2) & "."

    '                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, sSubject, Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper)
    '                If Me.oSocket Is Nothing = False Then
    '                    If oPC Is Nothing = False Then Me.oSocket.SendData(goMsgSys.GetAddObjectMessage(oPC, EpicaMessageCode.eAddObjectCommand))

    '                    Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
    '                    CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eCelebrationPeriod, lTemp)
    '                End If
    '            End If
    '        End If

    '    End Sub

    '    Public Enum PlayerRank As Byte
    '        Magistrate = 0
    '        Governor = 1
    '        Overseer = 2
    '        Duke = 3
    '        Baron = 4
    '        King = 5
    '        Emperor = 6
    '    End Enum
    '    Public Shared Function GetPlayerTitle(ByVal pyTitle As Byte, ByVal pbIsMale As Boolean) As String
    '        Dim sTemp As String = ""
    '        Select Case pyTitle
    '            Case PlayerRank.Baron
    '                If pbIsMale = True Then sTemp = "Baron " Else sTemp = "Baroness "
    '            Case PlayerRank.Duke
    '                If pbIsMale = True Then sTemp = "Duke " Else sTemp = "Lady "
    '            Case PlayerRank.Emperor
    '                If pbIsMale = True Then sTemp = "Emperor " Else sTemp = "Empress "
    '            Case PlayerRank.Governor
    '                If pbIsMale = True Then sTemp = "Governor " Else sTemp = "Governess "
    '            Case PlayerRank.King
    '                If pbIsMale = True Then sTemp = "King " Else sTemp = "Queen "
    '            Case PlayerRank.Magistrate
    '                sTemp = "Magistrate "
    '            Case PlayerRank.Overseer
    '                sTemp = "Overseer "
    '        End Select

    '        Return sTemp '& sName
    '    End Function
    '#End Region
 
    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        Dim lPos As Int32

        'No longer cache this... 
        ReDim mySendString(140)

        GetGUIDAsString.CopyTo(mySendString, 0)
        PlayerName.CopyTo(mySendString, 6)
        EmpireName.CopyTo(mySendString, 26)
        RaceName.CopyTo(mySendString, 46)
        lPos = 66

        PlayerUserName.CopyTo(mySendString, lPos)
        lPos += 20
        PlayerPassword.CopyTo(mySendString, lPos)
        lPos += 20

        System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, lPos)
        lPos += 4
        System.BitConverter.GetBytes(0S).CopyTo(mySendString, lPos)
        lPos += 2
        mySendString(lPos) = 0
        lPos += 1

        System.BitConverter.GetBytes(blCredits).CopyTo(mySendString, lPos)
        lPos += 8

        System.BitConverter.GetBytes(lStartedEnvirID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lStartLocX).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lStartLocZ).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(PirateStartLocX).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(PirateStartLocZ).CopyTo(mySendString, lPos) : lPos += 4

        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblPlayer (PlayerName, EmpireName, RaceName, PlayerUserName, PlayerPassword, SenateID, CommEncryptLevel, " & _
                  "EmpireTaxRate, Credits, LastViewedID, LastViewedTypeID, BaseMorale, StartedEnvirID, StartedEnvirTypeID, TotalPlayTime, LastLogin, " & _
                  "WarSentiment, CelebrationEnds, StartLocX, StartLocZ, PirateStartLocX, PirateStartLocZ, GuildID, GuildRankID, " & _
                  "JoinedGuildOn, PlayerTitle, PlayerIcon, PlayerGender, DeathBudgetBalance, AccountStatus, EmailAddress, EmailSettings) VALUES ('" & _
                  MakeDBStr(BytesToString(PlayerName)) & "', '" & MakeDBStr(BytesToString(EmpireName)) & "', '" & MakeDBStr(BytesToString(RaceName)) & "', '" & _
                  MakeDBStr(BytesToString(PlayerUserName)) & "', '" & MakeDBStr(BytesToString(PlayerPassword)) & "', "
                sSQL = sSQL & "-1, "
                sSQL &= "0, 0, " & blCredits & ", " & lLastViewedEnvir & ", " & _
                  iLastViewedEnvirType & ", 100, " & lStartedEnvirID & ", " & iStartedEnvirTypeID & ", " & _
                  TotalPlayTime & ", "
                If LastLogin <> Date.MinValue Then
                    Dim lVal As Int32 = CInt(Val(LastLogin.ToString("yyMMddHHmm")))
                    sSQL &= lVal.ToString
                Else : sSQL &= "0"
                End If
                sSQL &= ", 0, -1, " & lStartLocX & ", " & lStartLocZ & ", " & _
                  PirateStartLocX & ", " & PirateStartLocZ & ", -1, -1, -1, 0, 0, " & Me.yGender & ", 0, " & Me.AccountStatus & ", '" & MakeDBStr(BytesToString(ExternalEmailAddress)) & "', 0)"
            Else
                'UPDATE
                sSQL = "UPDATE tblPlayer SET PlayerName = '" & MakeDBStr(BytesToString(PlayerName)) & "', EmpireName = '" & MakeDBStr(BytesToString(EmpireName)) & _
                  "', RaceName = '" & MakeDBStr(BytesToString(RaceName)) & "', PlayerUserName = '" & MakeDBStr(BytesToString(PlayerUserName)) & "', PlayerPassword = '" & _
                  MakeDBStr(BytesToString(PlayerPassword)) & "', SenateID = "
                sSQL = sSQL & "-1"
                sSQL = sSQL & ", Credits = " & blCredits & ", LastViewedID = " & lLastViewedEnvir & ", LastViewedTypeID = " & iLastViewedEnvirType & _
                  ", BaseMorale = 100, StartedEnvirID = " & lStartedEnvirID & ", StartedEnvirTypeID = " & iStartedEnvirTypeID & _
                  ", LastLogin = "
                If LastLogin <> Date.MinValue Then
                    Dim lVal As Int32 = CInt(Val(LastLogin.ToString("yyMMddHHmm")))
                    sSQL &= lVal.ToString
                Else : sSQL &= "0"
                End If
                sSQL &= ", WarSentiment = 0, CelebrationEnds = -1, StartLocX = " & lStartLocX & ", StartLocZ = " & lStartLocZ
                sSQL &= ", PirateStartLocX = " & PirateStartLocX & ", PirateStartLocZ = " & PirateStartLocZ & _
                  ", AccountStatus = " & Me.AccountStatus & ", EmailAddress = '" & _
                  MakeDBStr(BytesToString(ExternalEmailAddress)) & "', PlayerGender = " & yGender & " WHERE PlayerID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PlayerID) FROM tblPlayer WHERE PlayerUserName = '" & MakeDBStr(BytesToString(PlayerUserName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Public Sub SaveStatusOnly()
		Try
			Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand("UPDATE tblPlayer SET AccountStatus = " & CInt(Me.AccountStatus) & " WHERE PlayerID = " & Me.ObjectID, goCN)
			oComm.ExecuteNonQuery()
		Catch
		End Try
    End Sub

    Public Function GetControlledPlanetList() As Planet()
        Dim oResult(-1) As Planet

        For X As Int32 = 0 To lControlledPlanetUB
            If lControlledPlanetIdx(X) > -1 AndAlso glPlanetIdx(lControlledPlanetIdx(X)) > -1 Then
                Dim oPlanet As Planet = goPlanet(lControlledPlanetIdx(X))
                If oPlanet Is Nothing = False Then
                    If oPlanet.ParentSystem Is Nothing OrElse oPlanet.ParentSystem.SystemType = 128 OrElse oPlanet.ParentSystem.SystemType = 255 Then Continue For
                    If oPlanet.OwnerID = Me.ObjectID Then
                        ReDim Preserve oResult(oResult.GetUpperBound(0) + 1)
                        oResult(oResult.GetUpperBound(0)) = oPlanet
                    End If
                End If
            End If
        Next X
        Return oResult
    End Function

    Public Sub AddPlanetControl(ByVal lIdx As Int32)
        Dim lFirstIdx As Int32 = -1
        For X As Int32 = 0 To lControlledPlanetUB
            If lControlledPlanetIdx(X) = lIdx Then Return
            If lFirstIdx = -1 AndAlso lControlledPlanetIdx(X) = -1 Then lFirstIdx = X
        Next X
        If lFirstIdx = -1 Then
            lFirstIdx = lControlledPlanetUB + 1
            ReDim Preserve lControlledPlanetIdx(lFirstIdx)
            lControlledPlanetUB += 1
        End If
        lControlledPlanetIdx(lFirstIdx) = lIdx
    End Sub
    Public Sub RemovePlanetControl(ByVal lIdx As Int32)
        For X As Int32 = 0 To lControlledPlanetUB
            If lControlledPlanetIdx(X) = lIdx Then lControlledPlanetIdx(X) = -1
        Next X
    End Sub
End Class

Public Class PlayerBudgetData
    Public oPlayer As Player
    Public blAgentMaintCost As Int64
    Public lIronCurtainPlanetID As Int32
    Public lMaxDeathBudget As Int32

    'a primary may request this for a player requesting it
    'a admin console may request this for a csr
    Private mlPrimReqs(-1) As Int32
    Private mlAdminReqs(-1) As Int32
    Private miPrimRequestorCode() As Int16
    Private miAdminRequestorCode() As Int16
    Private mlPrimReqUB As Int32 = -1
    Private mlAdminReqUB As Int32 = -1

    Private mlRequested() As Int32      'servers requested from

    Private mbDetailsRequested As Boolean = False

    Public oItems(-1) As BudgetDataItem
    Public lItemUB As Int32
    Public Class BudgetDataItem
        Public lEnvirID As Int32
        Public iEnvirTypeID As Int16
        Public yFlags As Byte
        Public lColonyID As Int32
        Public lTaxIncome As Int32
        Public lPopUpkeep As Int32
        Public lResearchCost As Int32
        Public lFactoryCost As Int32
        Public lSpaceportCost As Int32
        Public lOtherFacCost As Int32
        Public lUnemploymentCost As Int32
        Public lNonAirCost As Int32
        Public lAirUnitCost As Int32
        Public yTaxRate As Byte
        Public lJumpToEnvirID As Int32
        Public lTurretCost As Int32
        Public lDockedAirUnitCost As Int32
        Public lDockedNonAirUnitCost As Int32
        Public lExcessStorageCost As Int32
        Public blMiningBidIncome As Int64
        Public blMiningBidExpense As Int64
        Public GovScore As Byte
    End Class
    Public lTradePlayerID() As Int32
    Public lTradeValue() As Int32

    Private Function ComposeResponseMsg(ByVal bForAdminConsole As Boolean) As Byte()
        Dim lTradeItemCnt As Int32 = 0
        'For X As Int32 = 0 To lItemUB
        '    With oItems(X)
        '        If .lTradePlayerID Is Nothing = False Then
        '            lTradeItemCnt += .lTradePlayerID.Length
        '        End If
        '    End With
        'Next X
        If lTradePlayerID Is Nothing = False Then
            lTradeItemCnt += lTradePlayerID.Length
        End If

        Dim lItemLen As Int32 = 85 '72
        Dim lTradeItemLen As Int32 = 8
        If bForAdminConsole = True Then
            lItemLen += 20
            lTradeItemLen += 20
        End If

        Dim yMsg(31 + ((lItemUB + 1) * lItemLen) + (lTradeItemCnt * lTradeItemLen)) As Byte
        Dim lPos As Int32 = 0

        lPos += 2 'MsgCode placeholder (2)
        If bForAdminConsole = True Then
            System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Else
            System.BitConverter.GetBytes(-oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        End If
        System.BitConverter.GetBytes(ObjectType.eBudget).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(blAgentMaintCost).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(lIronCurtainPlanetID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMaxDeathBudget).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lItemUB + 1).CopyTo(yMsg, lPos) : lPos += 4


        For X As Int32 = 0 To lItemUB
            With oItems(X)
                System.BitConverter.GetBytes(.lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2

                If bForAdminConsole = True Then
                    If .iEnvirTypeID = ObjectType.ePlanet Then
                        Dim oEnvir As Planet = GetEpicaPlanet(.lEnvirID)
                        If oEnvir Is Nothing = False Then oEnvir.PlanetName.CopyTo(yMsg, lPos)
                    ElseIf .iEnvirTypeID = ObjectType.eSolarSystem Then
                        Dim oEnvir As SolarSystem = GetEpicaSystem(.lEnvirID)
                        If oEnvir Is Nothing = False Then oEnvir.SystemName.CopyTo(yMsg, lPos)
                    End If
                    lPos += 20
                End If

                yMsg(lPos) = .yFlags : lPos += 1
                System.BitConverter.GetBytes(.lColonyID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lTaxIncome).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lPopUpkeep).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lResearchCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lFactoryCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lSpaceportCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lOtherFacCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lUnemploymentCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lNonAirCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lAirUnitCost).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = .yTaxRate : lPos += 1
                System.BitConverter.GetBytes(.lJumpToEnvirID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lTurretCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lDockedAirUnitCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lDockedNonAirUnitCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lExcessStorageCost).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.blMiningBidExpense).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(.blMiningBidIncome).CopyTo(yMsg, lPos) : lPos += 8
                yMsg(lPos) = .GovScore : lPos += 1
            End With
        Next X

        If lTradePlayerID Is Nothing = False Then
            System.BitConverter.GetBytes(lTradePlayerID.Length).CopyTo(yMsg, lPos) : lPos += 4
            For Y As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
                System.BitConverter.GetBytes(lTradePlayerID(Y)).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lTradeValue(Y)).CopyTo(yMsg, lPos) : lPos += 4
                If bForAdminConsole = True Then
                    Dim oTrade As Player = GetEpicaPlayer(lTradePlayerID(Y))
                    If oTrade Is Nothing = False Then
                        oTrade.PlayerName.CopyTo(yMsg, lPos)
                    End If
                    lPos += 20
                End If
            Next Y
        Else
            System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4
        End If

        Return yMsg

    End Function

    Public Sub HandlePrimaryResponse(ByVal yData() As Byte, ByVal lIdx As Int32)
        Dim lPos As Int32 = 2     'for msgcode
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'iok, if the id is negative, then the primary has nothing to report
        If lID > 0 Then
            'ok, let's parse our msg
            lPos += 2
            blAgentMaintCost += System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            If lIronCurtainPlanetID <> -1 Then lIronCurtainPlanetID = System.BitConverter.ToInt32(yData, lPos)
            lPos += 4

            lMaxDeathBudget = Math.Max(lMaxDeathBudget, System.BitConverter.ToInt32(yData, lPos)) : lPos += 4

            Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            For X As Int32 = 0 To lCnt - 1
                Dim oTmp As New BudgetDataItem
                With oTmp
                    .lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .iEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .yFlags = yData(lPos) : lPos += 1
                    .lColonyID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lTaxIncome = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lPopUpkeep = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lResearchCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lFactoryCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lSpaceportCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lOtherFacCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lUnemploymentCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lNonAirCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lAirUnitCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .yTaxRate = yData(lPos) : lPos += 1
                    .lJumpToEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lTurretCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lDockedAirUnitCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lDockedNonAirUnitCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lExcessStorageCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    .blMiningBidExpense = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                    .blMiningBidIncome = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                    .GovScore = yData(lPos) : lPos += 1
                End With

                SyncLock oItems
                    lItemUB += 1
                    ReDim Preserve oItems(lItemUB)
                    oItems(lItemUB) = oTmp
                End SyncLock

            Next X
        End If

        Dim lUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
        If lTradePlayerID Is Nothing Then
            ReDim lTradePlayerID(-1)
            ReDim lTradeValue(-1)
        End If

        For X As Int32 = 0 To lUB
            Dim lPlayer As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim lTmpIdx As Int32 = -1
            For Y As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
                If lTradePlayerID(Y) = lPlayer Then
                    lTmpIdx = Y
                    Exit For
                End If
            Next Y
            If lTmpIdx = -1 Then
                lTmpIdx = lTradePlayerID.GetUpperBound(0) + 1
                ReDim Preserve lTradePlayerID(lTmpIdx)
                ReDim Preserve lTradeValue(lTmpIdx)
                lTradeValue(lTmpIdx) = 0
            End If
            lTradePlayerID(lTmpIdx) = lPlayer
            lTradeValue(lTmpIdx) += lValue
        Next X

        'Now, find our index
        If mlRequested Is Nothing = False Then
            For X As Int32 = 0 To mlRequested.GetUpperBound(0)
                If mlRequested(X) = lIdx Then
                    mlRequested(X) = -1
                    Exit For
                End If
            Next X
        End If

        'Now, check if we are ready to send our response(s)
        Dim bSend As Boolean = True
        If mlRequested Is Nothing = False Then
            For X As Int32 = 0 To mlRequested.GetUpperBound(0)
                If mlRequested(X) > -1 Then
                    bSend = False
                    Exit For
                End If
            Next X
        End If

        If bSend = True Then
            'Ok, need to send our final response
            Dim yAdminMsg() As Byte = Nothing
            Dim yPrimMsg() As Byte = Nothing

            For X As Int32 = 0 To mlPrimReqUB
                Dim lReq As Int32 = mlPrimReqs(X)
                If lReq > -1 Then
                    'primary idx
                    If yPrimMsg Is Nothing Then yPrimMsg = ComposeResponseMsg(False)
                    System.BitConverter.GetBytes(miPrimRequestorCode(X)).CopyTo(yPrimMsg, 0)
                    goMsgSys.SendToServerIndex(lReq, yPrimMsg)
                End If
            Next X

            For X As Int32 = 0 To mlAdminReqUB
                Dim lReq As Int32 = mlAdminReqs(X)
                If lReq > -1 Then
                    'admin console
                    If yAdminMsg Is Nothing Then yAdminMsg = ComposeResponseMsg(True)
                    System.BitConverter.GetBytes(miAdminRequestorCode(X)).CopyTo(yAdminMsg, 0)
                    goMsgSys.SendToAdminIndex(lReq, yAdminMsg)
                End If
            Next X

            mlPrimReqUB = -1
            mlAdminReqUB = -1

            mbDetailsRequested = False
        End If
    End Sub

    Public Sub RequestDetails(ByVal lPrimaryIdx As Int32, ByVal lAdminConsoleIdx As Int32, ByVal iMsgCode As Int16)
        If oPlayer Is Nothing Then Return

        If lPrimaryIdx <> -1 Then
            SyncLock mlPrimReqs
                Dim lNewVal As Int32 = -1

                Dim bFound As Boolean = False
                For X As Int32 = 0 To mlPrimReqUB
                    If mlPrimReqs(X) = lNewVal Then
                        bFound = True
                        Exit For
                    End If
                Next X
                If bFound = False Then
                    mlPrimReqUB += 1
                    ReDim Preserve mlPrimReqs(mlPrimReqUB)
                    ReDim Preserve miPrimRequestorCode(mlPrimReqUB)
                    mlPrimReqs(mlPrimReqUB) = lPrimaryIdx
                    miPrimRequestorCode(mlPrimReqUB) = iMsgCode
                End If
            End SyncLock
        Else
            SyncLock mlAdminReqs
                Dim lNewVal As Int32 = -1

                Dim bFound As Boolean = False
                For X As Int32 = 0 To mlAdminReqUB
                    If mlAdminReqs(X) = lNewVal Then
                        bFound = True
                        Exit For
                    End If
                Next X
                If bFound = False Then
                    mlAdminReqUB += 1
                    ReDim Preserve mlAdminReqs(mlAdminReqUB)
                    ReDim Preserve miAdminRequestorCode(mlAdminReqUB)
                    mlAdminReqs(mlAdminReqUB) = lAdminConsoleIdx
                    miAdminRequestorCode(mlAdminReqUB) = iMsgCode
                End If
            End SyncLock
        End If

        If mbDetailsRequested = True Then Return
        mbDetailsRequested = False

        blAgentMaintCost = 0
        lIronCurtainPlanetID = -1
        lMaxDeathBudget = 0
        lItemUB = -1

        Dim yRequest(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerBudget).CopyTo(yRequest, 0)
        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yRequest, 2)

        ReDim mlRequested(oPlayer.lAttachedPrimaryUB + 1)
        For X As Int32 = 0 To mlRequested.GetUpperBound(0)
            mlRequested(X) = -1
        Next X

        Dim lIdx As Int32 = oPlayer.lOwnerPrimaryIdx
        goMsgSys.SendToServerIndex(lIdx, yRequest)
        Dim lTmpIdx As Int32 = 0
        mlRequested(lTmpIdx) = lIdx
        lTmpIdx += 1

        For X As Int32 = 0 To oPlayer.lAttachedPrimaryUB
            lIdx = oPlayer.lAttachedPrimaryIdx(X)
            If lIdx <> oPlayer.lOwnerPrimaryIdx Then
                goMsgSys.SendToServerIndex(lIdx, yRequest)
                mlRequested(lTmpIdx) = lIdx
                lTmpIdx += 1
            End If
        Next X

    End Sub

End Class