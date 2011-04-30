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
#Region "  ARENA ID GENERATOR  "
    Private Shared ml_Arena_ID As Int32 = 0
    Private Shared mo_Lock_Arena As DM_Arena
    Public Shared Function GetNextArenaID() As Int32
        If mo_Lock_Arena Is Nothing Then mo_Lock_Arena = New DM_Arena
        SyncLock mo_Lock_Arena
            ml_Arena_ID += 1
            Return ml_Arena_ID
        End SyncLock
    End Function
#End Region

    Public ServerIndex As Int32     'indicates the index if the server's goArena array

    Public lArenaID As Int32     'indicates the specific instance ID for this arena. Used when connecting to the ArenaServer
    Public lCreatorID As Int32      'playerID who created the instance
    Private moCreator As Player = Nothing
    Public ReadOnly Property oCreator() As Player
        Get
            If moCreator Is Nothing OrElse moCreator.ObjectID <> lCreatorID Then
                moCreator = GetEpicaPlayer(lCreatorID)
            End If
            Return moCreator
        End Get
    End Property

    Public yPublicity As Byte

    Public yBaseArenaFlags As Byte  'see eyGeneralArenaFlags

    Public lDuration As Int32       'in minutes before instance expires

    Public lInstanceCreateTime As Int32 'when instance was created
    Public lInstanceStartTime As Int32  'when instance was started

    Public lMapID As Int32          'which map is being played
    Private moMap As ArenaMap = Nothing
    Public ReadOnly Property Map() As ArenaMap
        Get
            If moMap Is Nothing OrElse moMap.lMapID <> lMapID Then
                'get the map
            End If
            Return moMap
        End Get
    End Property

    Public yGameMode As Byte        'see eyGameMode - determines the child class of this class

    Public lSideCnt As Int32        'number of sides for this arena
    Public oSides() As ArenaSide

    Public lRespawnLimit As Int32   'number of times units respawn, 0 = none, -1 = infinite
    Public lRespawnDelay As Int32   'number of seconds before a unit respawns

    Public yArenaState As eyArenaState = eyArenaState.eInitial

    Public MustOverride Sub AppendGameModeSpecificData(ByRef yData() As Byte, ByRef lPos As Int32)
    Public MustOverride Function GetGameModeSpecificDataLen() As Int32

#Region "  Unit Limits (Per Player)  "
    Public oUnitLimits() As UnitLimit
    Public lUnitLimitUB As Int32 = -1
#End Region

    Public Function GetPlayerSidePlayer(ByVal lPlayerID As Int32) As ArenaSidePlayer
        For X As Int32 = 0 To lSideCnt - 1
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

    Public Function AddContestant(ByRef oEntity As Epica_Entity) As ArenaContestant
        'OK, here, add the contestant, then determine if the players are in alignment with the unit limits

        If CanAddUnit(oEntity) = False Then Return Nothing

        Dim oSidePlayer As ArenaSidePlayer = GetPlayerSidePlayer(oEntity.Owner.ObjectID)
        If oSidePlayer Is Nothing = False Then
            Dim oContestant As ArenaContestant = oSidePlayer.AddContestant(oEntity)
            oContestant.lRespawns = 0
            oContestant.lNextRespawnCycle = Int32.MinValue
            Return oContestant
        End If
        Return Nothing
    End Function

    Private Function CanAddUnit(ByVal oEntity As Epica_Entity) As Boolean
        Dim lPlayerID As Int32 = oEntity.Owner.ObjectID
        Dim oPlayerSide As ArenaSidePlayer = GetPlayerSidePlayer(lPlayerID)
        If oPlayerSide Is Nothing Then Return False

        'TODO: Check the Max Unit count, max ground unit, max flying unit counts here... based on Map (it has those values)
        'TODO: Include Max Unit (254), Max Ground Unit (253), Max Flying Unit (252) here...


        Dim lCnts(lUnitLimitUB) As Int32
        Dim lGrndCnt As Int32 = 0
        Dim lFlyCnt As Int32 = 0
        For X As Int32 = 0 To lUnitLimitUB
            lCnts(X) = 0
        Next X

        For X As Int32 = 0 To oPlayerSide.lContestantUB
            If oPlayerSide.oContestants(X) Is Nothing = False Then
                If oPlayerSide.oContestants(X).oEntity Is Nothing = False Then
                    Dim oDef As Epica_Entity_Def = Nothing
                    If oPlayerSide.oContestants(X).oEntity.ObjTypeID = ObjectType.eUnit Then
                        oDef = CType(oPlayerSide.oContestants(X).oEntity, Unit).EntityDef
                    ElseIf oPlayerSide.oContestants(X).oEntity.ObjTypeID = ObjectType.eFacility Then
                        oDef = CType(oPlayerSide.oContestants(X).oEntity, Facility).EntityDef
                    End If

                    If oDef Is Nothing = False Then
                        If (oDef.yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then lGrndCnt += 1 Else lFlyCnt += 1

                        Dim bBad As Boolean = False
                        For Y As Int32 = 0 To lUnitLimitUB
                            If oUnitLimits(Y) Is Nothing = False Then
                                Dim yRes As Byte = oUnitLimits(Y).DoesDefFitLimit(oDef)
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
                End If
            End If
        Next X

        'now, determine if the unit we are adding is ok
        With oEntity
            Dim oDef As Epica_Entity_Def = Nothing
            If .ObjTypeID = ObjectType.eUnit Then
                oDef = CType(oEntity, Unit).EntityDef
            ElseIf .ObjTypeID = ObjectType.eFacility Then
                oDef = CType(oEntity, Facility).EntityDef
            End If
            If oDef Is Nothing Then Return False

            Dim bBad As Boolean = False
            For Y As Int32 = 0 To lUnitLimitUB
                Dim yRes As Byte = oUnitLimits(Y).DoesDefFitLimit(oDef)
                If yRes = 1 Then
                    lCnts(Y) += 1
                    If lCnts(Y) > oUnitLimits(Y).lMaxCnt Then Return False
                    bBad = False
                ElseIf yRes = 2 Then
                    bBad = True
                End If
            Next Y
            If bBad = True Then Return False
        End With

        Return True

    End Function

    Public Sub PlayerReady(ByVal lPlayerID As Int32)
        'Mark the player as ready then check if all players are ready
        For X As Int32 = 0 To lSideCnt - 1
            If oSides(X) Is Nothing = False Then
                If oSides(X).oSidePlayers Is Nothing = False Then
                    Dim bFound As Boolean = False
                    For Y As Int32 = 0 To oSides(X).oSidePlayers.GetUpperBound(0)
                        If oSides(X).oSidePlayers(Y) Is Nothing = False Then
                            If oSides(X).oSidePlayers(Y).lPlayerID = lPlayerID Then
                                oSides(X).oSidePlayers(Y).yFlags = oSides(X).oSidePlayers(Y).yFlags Or eyArenaSidePlayerFlag.PlayerReady
                                bFound = True
                                Exit For
                            End If
                        End If
                    Next Y
                    If bFound = True Then Exit For
                End If
            End If
        Next X

        'Now, verify our ready states...
        'Dim bGoodToGo As Boolean = True
        'For X As Int32 = 0 To lSideCnt - 1
        '    If oSides(X) Is Nothing = False Then
        '        If oSides(X).oSidePlayers Is Nothing = False Then

        '            For Y As Int32 = 0 To oSides(X).oSidePlayers.GetUpperBound(0)
        '                If oSides(X).oSidePlayers(Y) Is Nothing = False Then
        '                    If oSides(X).oSidePlayers(Y).lPlayerID > 0 Then
        '                        If (oSides(X).oSidePlayers(Y).yFlags And eyArenaSidePlayerFlag.PlayerReady) = 0 Then
        '                            bGoodToGo = False
        '                        End If
        '                    End If
        '                End If
        '            Next Y

        '        End If
        '    End If
        'Next X

    End Sub

    Public Function GetObjAsString() As Byte()

        Dim lLen As Int32 = 33

        'Ok, determine our length
        For X As Int32 = 0 To lSideCnt - 1
            If oSides(X) Is Nothing = False Then
                lLen += 12
                If oSides(X).oSidePlayers Is Nothing = False Then
                    lLen += 9
                End If
            End If
        Next X
        lLen += GetGameModeSpecificDataLen()

        Dim yMsg(lLen) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(Me.lArenaID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(ObjectType.eArena).CopyTo(yMsg, lPos) : lPos += 2
        yMsg(lPos) = yGameMode : lPos += 1
        System.BitConverter.GetBytes(lCreatorID).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = yPublicity : lPos += 1
        yMsg(lPos) = yBaseArenaFlags : lPos += 1
        yMsg(lPos) = yArenaState : lPos += 1
        System.BitConverter.GetBytes(lDuration).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lInstanceCreateTime).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lInstanceStartTime).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMapID).CopyTo(yMsg, lPos) : lPos += 4

        AppendGameModeSpecificData(yMsg, lPos)

        System.BitConverter.GetBytes(lSideCnt).CopyTo(yMsg, lPos) : lPos += 4

        For X As Int32 = 0 To lSideCnt - 1
            If oSides(X) Is Nothing = False Then
                With oSides(X)
                    System.BitConverter.GetBytes(.lSideID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lSideScore).CopyTo(yMsg, lPos) : lPos += 4

                    If .oSidePlayers Is Nothing = False Then
                        System.BitConverter.GetBytes(.oSidePlayers.GetUpperBound(0) + 1).CopyTo(yMsg, lPos) : lPos += 4
                    Else
                        System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4
                    End If
                End With

                For Y As Int32 = 0 To oSides(X).oSidePlayers.GetUpperBound(0)
                    If oSides(X).oSidePlayers(Y) Is Nothing = False Then
                        With oSides(X).oSidePlayers(Y)
                            System.BitConverter.GetBytes(.lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.lCapturePoints).CopyTo(yMsg, lPos) : lPos += 4
                            yMsg(lPos) = .yFlags : lPos += 1
                        End With
                    End If
                Next Y
            End If
        Next X

        Return yMsg

    End Function

    Public Sub AddPlayer(ByRef oPlayer As Player)
        '
    End Sub

    Public Function GetArenasListItemMsg() As Byte()
        Dim yMsg(25) As Byte
        Dim lPos As Int32 = 0

        'msgcode here...
        lPos += 2

        yMsg(lPos) = yGameMode : lPos += 1
        System.BitConverter.GetBytes(Me.lArenaID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCreatorID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMapID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSideCnt).CopyTo(yMsg, lPos) : lPos += 4

        Dim lPCnt As Int32 = 0
        System.BitConverter.GetBytes(lPCnt).CopyTo(yMsg, lPos) : lPos += 4

        yMsg(lPos) = yBaseArenaFlags : lPos += 1
        yMsg(lPos) = yPublicity : lPos += 1
        yMsg(lPos) = yArenaState : lPos += 1

        Return yMsg
    End Function

    Public Function GetArenaDetailMsg(ByVal lPlayerID As Int32) As Byte()
        Dim yBaseMsg() As Byte = GetObjAsString()

        Dim lAddLen As Int32 = 0
        Dim oASP As ArenaSidePlayer = Me.GetPlayerSidePlayer(lPlayerID)
        If oASP Is Nothing = False Then
            lAddLen = oASP.lContestantUB + 1
            lAddLen *= 6
        End If

        Dim yFinal(yBaseMsg.GetUpperBound(0) + lAddLen + 2) As Byte
        Dim lPos As Int32 = 0
        'msgcode here...
        lPos += 2
        yBaseMsg.CopyTo(yFinal, lPos) : lPos += yBaseMsg.Length

        'now, our units
        If oASP Is Nothing = False Then
            For X As Int32 = 0 To oASP.lContestantUB
                If oASP.oContestants(X) Is Nothing = False Then
                    If oASP.oContestants(X).oEntity Is Nothing = False Then
                        oASP.oContestants(X).oEntity.GetGUIDAsString().CopyTo(yFinal, lPos) : lPos += 6
                    End If
                End If
            Next X
        End If

        Return yFinal
    End Function

#Region "  Invitations  "
    Public lInvited(-1) As Int32
    Public Function PlayerInvited(ByVal lPlayerID As Int32) As Boolean
        For X As Int32 = 0 To lInvited.GetUpperBound(0)
            If lInvited(X) = lPlayerID Then Return True
        Next X
        Return False
    End Function

#End Region


#Region "  ALL ARENAS  "
    Private Shared moArenas(-1) As Arena
    Private Shared mlArenaUB As Int32 = -1
    Public Shared Sub AddArena(ByRef oArena As Arena)
        SyncLock moArenas
            Dim lUB As Int32 = -1
            If moArenas Is Nothing = False Then lUB = Math.Min(mlArenaUB, moArenas.GetUpperBound(0))

            For X As Int32 = 0 To lUB
                If moArenas(X) Is Nothing Then
                    oArena.ServerIndex = X
                    moArenas(X) = oArena
                    Return
                End If
            Next X

            mlArenaUB += 1
            ReDim Preserve moArenas(mlArenaUB)
            moArenas(mlArenaUB) = oArena
            oArena.ServerIndex = mlArenaUB
        End SyncLock
    End Sub
    Public Shared Function GetArena(ByVal lArenaID As Int32) As Arena
        Try
            Dim lUB As Int32 = -1
            If moArenas Is Nothing = False Then lUB = Math.Min(mlArenaUB, moArenas.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If moArenas(X) Is Nothing = False AndAlso moArenas(X).lArenaID = lArenaID Then Return moArenas(X)
            Next X
        Catch
        End Try
        Return Nothing
    End Function
    Public Shared Sub RemoveArena(ByVal lArenaID As Int32)
        Try
            Dim lUB As Int32 = -1
            If moArenas Is Nothing = False Then lUB = Math.Min(mlArenaUB, moArenas.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If moArenas(X) Is Nothing = False AndAlso moArenas(X).lArenaID = lArenaID Then
                    moArenas(X) = Nothing
                End If
            Next X
        Catch
        End Try
    End Sub

    Public Shared Sub HandleRequestArenaList(ByRef oSocket As NetSock, ByRef oPlayer As Player)
        Try
            Dim lUB As Int32 = -1
            If moArenas Is Nothing = False Then lUB = Math.Min(mlArenaUB, moArenas.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If moArenas(X) Is Nothing = False Then
                    Dim bGood As Boolean = False

                    'is the arena closed?
                    If moArenas(X).yArenaState = eyArenaState.eClosed Then Continue For

                    'is the player creator of the arena?
                    'is the arena public?
                    'is the arena guild only and creator's guild = player's guild?
                    'is player in invited list?
                    If moArenas(X).lCreatorID = oPlayer.ObjectID Then
                        bGood = True
                    ElseIf moArenas(X).yPublicity = eyPublicity.ePublicToAll Then
                        bGood = True
                    ElseIf moArenas(X).yPublicity = eyPublicity.eGuildOnly AndAlso moArenas(X).oCreator Is Nothing = False AndAlso moArenas(X).oCreator.oGuild Is Nothing = False AndAlso oPlayer.oGuild Is Nothing = False AndAlso moArenas(X).oCreator.oGuild.ObjectID = oPlayer.oGuild.ObjectID Then
                        bGood = True
                    ElseIf moArenas(X).PlayerInvited(oPlayer.ObjectID) = True Then
                        bGood = True
                    End If

                    If bGood = False Then
                        Dim oASP As ArenaSidePlayer = moArenas(X).GetPlayerSidePlayer(oPlayer.ObjectID)
                        If oASP Is Nothing = False Then bGood = True
                        oASP = Nothing
                    End If

                    If bGood = True Then
                        'Ok, send the msg
                        oSocket.SendData(moArenas(X).GetArenasListItemMsg())
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub
#End Region
End Class
