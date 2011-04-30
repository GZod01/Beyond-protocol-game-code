Option Strict On


Public Enum eyPublicity As Byte
    ePublicToAll = 0
    eGuildOnly = 1
    eInviteOnly = 2
End Enum

'AKA - ArenaType
Public Enum eyGameMode As Byte
    eDeathMatch = 0
    eCaptureTheFlag = 1
    eControlPoints = 2
    eScenario = 3
    eContestedScenario = 4
    eConquest = 5

    eLastGameMode       'alway the last game mode listed
End Enum

Public Enum eyGeneralArenaFlags As Byte
    eResultsTracked = 1
    eAllowOrbitalBombardment = 2
End Enum

Public Enum eyArenaState As Byte
    eInitial = 0            'client is submitting the config
    eForming = 1            'server approves of config and it is being posted
    eStarting = 2           'Last Player places themselves as READY (and more than minimum player number is met)
    eValidation = 3         'Arena Server is validating the config and players
    eInProgress = 4         'Valid config and arena is in progress
    eRoundFinished = 5      'round time, goal limit, arena time limit, etc... met
    eProcessResults = 6     'end of the Arena, primary is processing/storing results
    eClosed = 255           'arena is over for whatever reason
End Enum

Public MustInherit Class Arena
    Public lArenaID As Int32     'indicates the specific instance ID for this arena. Used when connecting to the ArenaServer
    Public lCreatorID As Int32      'playerID who created the instance

    Public yPublicity As Byte

    Public yBaseArenaFlags As Byte  'see eyGeneralArenaFlags

    Public lDuration As Int32       'in minutes before instance expires

    Public lInstanceCreateTime As Int32 'when instance was created - in GMT
    Public lInstanceStartTime As Int32  'when instance was started - in GMT

    Public lMapID As Int32          'which map is being played
    Private moMap As ArenaMap = Nothing
    Public ReadOnly Property Map() As ArenaMap
        Get
            If moMap Is Nothing OrElse moMap.lMapID <> lMapID Then
                'get the map
                moMap = ArenaMap.AllMapByID(lMapID)
            End If
            Return moMap
        End Get
    End Property

    Public yGameMode As Byte        'see eyGameMode - determines the child class of this class

    Public lSideCnt As Int32 = 0        'number of sides for this arena
    Public oSides() As ArenaSide

    Public lPlayerCnt As Int32      'for list purposes only, number of players in the arena already

    Public lRespawnLimit As Int32   'number of times units respawn, 0 = none, -1 = infinite
    Public lRespawnDelay As Int32   'number of seconds before a unit respawns

    Public yArenaState As eyArenaState = eyArenaState.eInitial

    Private mbDetailsRequested As Boolean = False
    Public Sub RequestDetails()
        If mbDetailsRequested = True Then Return
        mbDetailsRequested = True

        Dim yMsg(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(Me.lArenaID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(ObjectType.eArena).CopyTo(yMsg, 6)
        goUILib.SendMsgToPrimary(yMsg)
    End Sub

#Region "  Unit Limits (Per Player)  "
    Public oUnitLimits() As UnitLimit
    Public lUnitLimitUB As Int32 = -1
#End Region

    Public Function GetPlayerSidePlayer(ByVal lPlayerID As Int32) As ArenaSidePlayer
        If oSides Is Nothing = True Then Return Nothing

        For X As Int32 = 0 To (lSideCnt - 1)
            If oSides(X) Is Nothing = False Then
                If oSides(X).oSidePlayers Is Nothing = False Then
                    For Y As Int32 = 0 To oSides(X).oSidePlayers.GetUpperBound(0)
                        If oSides(X).oSidePlayers(Y).lPlayerID = lPlayerID Then Return oSides(X).oSidePlayers(Y)
                    Next Y
                End If
            End If
        Next X
        Return Nothing
    End Function

    Public Function AddContestant(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lPlayerID As Int32) As ArenaContestant
        Dim oSidePlayer As ArenaSidePlayer = GetPlayerSidePlayer(lPlayerID)
        If oSidePlayer Is Nothing = False Then
            Dim oContestant As ArenaContestant = oSidePlayer.AddContestant(lEntityID, iEntityTypeID)
            oContestant.lRespawns = 0
            oContestant.lNextRespawnCycle = Int32.MinValue
            Return oContestant
        End If
        Return Nothing
    End Function

    Public Function CanAddUnit(ByVal lHullSize As Int32, ByVal yHullType As Byte) As Boolean
        Dim oPlayerSide As ArenaSidePlayer = GetPlayerSidePlayer(glPlayerID)
        If oPlayerSide Is Nothing Then Return False

        Dim lCnts(lUnitLimitUB) As Int32
        Dim lGrndCnt As Int32 = 0
        Dim lFlyCnt As Int32 = 0
        For X As Int32 = 0 To lUnitLimitUB
            lCnts(X) = 0
        Next X

        Dim bBad As Boolean = False
        For X As Int32 = 0 To oPlayerSide.lContestantUB
            If oPlayerSide.oContestants(X) Is Nothing = False Then
                For Y As Int32 = 0 To lUnitLimitUB
                    If oUnitLimits(Y) Is Nothing = False Then
                        Dim yRes As Byte = oUnitLimits(Y).DoesDefFitLimit(oPlayerSide.oContestants(X).lHullSize, oPlayerSide.oContestants(X).yHullType)
                        If yRes = 1 Then
                            lCnts(Y) += 1

                            If lCnts(Y) > oUnitLimits(Y).lMaxCnt Then Return False
                            bBad = False
                        ElseIf yRes = 2 Then
                            bBad = True
                        End If
                    End If
                Next Y
                If bBad = True Then Return False
            End If
        Next X

        For Y As Int32 = 0 To lUnitLimitUB
            Dim yRes As Byte = oUnitLimits(Y).DoesDefFitLimit(lHullSize, yHullType)
            If yRes = 1 Then
                lCnts(Y) += 1
                If lCnts(Y) > oUnitLimits(Y).lMaxCnt Then Return False
                bBad = False
            ElseIf yRes = 2 Then
                bBad = True
            End If
        Next Y
        Return Not bBad
        
    End Function

    Public Sub SetPlayerReadyState(ByVal lPlayerID As Int32, ByVal bValue As Boolean)

        Dim oASP As ArenaSidePlayer = GetPlayerSidePlayer(lPlayerID)
        If oASP Is Nothing = False Then
            If bValue = True Then
                oASP.yFlags = oASP.yFlags Or eyArenaSidePlayerFlag.PlayerReady
            Else

                If (oASP.yFlags And eyArenaSidePlayerFlag.PlayerReady) <> 0 Then
                    oASP.yFlags = oASP.yFlags Xor eyArenaSidePlayerFlag.PlayerReady
                End If
            End If
        End If
 
    End Sub
 
    Public Sub AddPlayer(ByVal lPlayerID As Int32, ByVal lSideID As Int32)
        For X As Int32 = 0 To lSideCnt - 1
            If oSides(X) Is Nothing = False Then
                If oSides(X).lSideID = lSideID Then
                    oSides(X).AddArenaSidePlayer(lPlayerID)
                    Exit For
                End If
            End If
        Next X
    End Sub

    Public Shared Function GetArenaMainListBoxHeader() As String
        Dim sType As String = "Type".PadRight(6, " "c)
        Dim sMapName As String = "Map Name".PadRight(31, " "c)
        Dim sPlayers As String = "Plyrs".PadRight(6, " "c)
        Dim sStatus As String = "Status".PadRight(8, " "c)
        Dim sCreator As String = "Creator"

        Return sType & sMapName & sPlayers & sStatus & sCreator
    End Function
    Public Function GetArenaMainListText() As String
        Dim sGameMode As String = "?"
        Select Case yGameMode
            Case eyGameMode.eCaptureTheFlag
                sGameMode = "CTF"
            Case eyGameMode.eConquest
                sGameMode = "C"
            Case eyGameMode.eContestedScenario
                sGameMode = "CS"
            Case eyGameMode.eControlPoints
                sGameMode = "CP"
            Case eyGameMode.eDeathMatch
                sGameMode = "DM"
            Case eyGameMode.eScenario
                sGameMode = "S"
        End Select
        sGameMode = sGameMode.PadRight(4, " "c)
        If (yBaseArenaFlags And eyGeneralArenaFlags.eResultsTracked) <> 0 Then
            sGameMode &= "T"
        End If
        sGameMode = sGameMode.PadRight(6, " "c)

        Dim sMapName As String = "Unknown Map"
        Dim lMaxPlayer As Int32 = 0
        If Me.Map Is Nothing = False Then
            sMapName = Me.Map.sMapName
            lMaxPlayer = Me.Map.lMaxPlayerCnt
        End If
        If sMapName.Length > 30 Then sMapName = sMapName.Substring(0, 27) & "..."
        sMapName = sMapName.PadRight(31, " "c)

        'CreatorID...
        Dim sCreator As String = GetCacheObjectValue(lCreatorID, ObjectType.ePlayer)

        Dim sPlayers As String = lPlayerCnt.ToString & "/" & lMaxPlayer.ToString
        sPlayers = sPlayers.PadRight(6, " "c)

        Dim sState As String = "Locked"
        Select Case yArenaState
            Case eyArenaState.eForming, eyArenaState.eInitial
                sState = "Forming"
        End Select
        sState = sState.PadRight(8, " "c)

        'NOT INCLUDED BUT SENT DOWN...
        '.lSideCnt
        '.lDuration = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        '.yPublicity = yData(lPos) : lPos += 1

        Return sGameMode & sMapName & sPlayers & sState & sCreator

    End Function

    Public Function GetArenaSummaryText() As String

        Dim sResult As String = ""
        Select Case yGameMode
            Case eyGameMode.eCaptureTheFlag
                sResult = "Capture the Flag"
            Case eyGameMode.eConquest
                sResult = "Conquest"
            Case eyGameMode.eContestedScenario
                sResult = "Contested Scenario"
            Case eyGameMode.eControlPoints
                sResult = "Control Points"
            Case eyGameMode.eDeathMatch
                sResult = "Death Match"
            Case eyGameMode.eScenario
                sResult = "Scenario"
        End Select

        If Me.Map Is Nothing = False Then
            sResult &= vbCrLf & Map.sMapName
        End If

        sResult &= vbCrLf & vbCrLf & "LIMITS:"
        For X As Int32 = 0 To Me.lUnitLimitUB
            sResult &= vbCrLf & Me.oUnitLimits(X).GetDisplayText()
        Next X

        Return sResult
    End Function

End Class
