Public Class Planet
    Inherits Epica_GUID

    Public Const PLANET_OBJ_MSG_LEN As Int32 = 71
    Public Const PLANET_ROTATE_ANGLE_MSG_LOC As Int32 = 69

    Public PlanetName(19) As Byte
    Public PlanetTypeID As Byte         'based on the PlanetType enum in EpicaShared
    Public PlanetSizeID As Byte         'used for determining Map Size: 0-Tiny, 1-Small, 2-Medium, 3-Large, 4-Huge
    Public PlanetRadius As Int16        'used for displaying the planet sphere
    Public ParentSystem As SolarSystem
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32
    Public Vegetation As Byte
    Public Atmosphere As Byte
    Public Hydrosphere As Byte
    Public Gravity As Byte
    Public SurfaceTemperature As Byte
	Public RotationDelay As Int16		'cycles between incrementing the rotation angle (not used except on client)
	Public OwnerID As Int32

    Public InnerRingRadius As Int32
    Public OuterRingRadius As Int32
    Public RingDiffuse As Int32
    Public RingMineralID As Int32 = 0
    Public RingMineralConcentration As Int32 = 0
    Private moRingMineral As Mineral = Nothing
    Public ReadOnly Property RingMineral() As Mineral
        Get
            If moRingMineral Is Nothing AndAlso RingMineralID > 0 Then
                moRingMineral = GetEpicaMineral(RingMineralID)
            End If
            Return moRingMineral
        End Get
    End Property

    Public RingMinerID As Int32 = -1
    Private moRingMiner As Facility = Nothing
    Public ReadOnly Property RingMiner() As Facility
        Get
            If moRingMiner Is Nothing OrElse moRingMiner.ObjectID <> RingMinerID Then
                If RingMinerID < 1 Then moRingMiner = Nothing Else moRingMiner = GetEpicaFacility(RingMinerID)
            End If
            Return moRingMiner
        End Get
    End Property

    Public PlayerSpawns As Int32 = 0    'number of player spawns used
    Public SpawnLocked As Boolean = False

	Public AxisAngle As Int32

	Public blOriginalMineralQuantity As Int64
	Public ySentGNSLowRes As Byte
	Public lPrimaryComposition As Int32		'mineralid

    Private Shared msw_Rotate As Stopwatch      'MSC 12/18/2006

    Public oDomain As DomainServer

    Private mySendString() As Byte

    Public lColonysHereIdx() As Int32       'indices to the global array 
    Public lColonysHereUB As Int32 = -1

    Private mbAtColonyLimit As Boolean = False
    Public Property bAtColonyLimit() As Boolean
        Get
            Return mbAtColonyLimit
        End Get
        Set(ByVal value As Boolean)
            If mbAtColonyLimit <> value Then
                'need to inform the parent domain of such
                If oDomain Is Nothing = False AndAlso oDomain.DomainSocket Is Nothing = False Then
                    Dim yMsg(8) As Byte
                    Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eAdvancedEventConfig).CopyTo(yMsg, lPos) : lPos += 2
                    GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    If value = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
                    lPos += 1
                    oDomain.DomainSocket.SendData(yMsg)
                End If
            End If
            mbAtColonyLimit = value
        End Set
    End Property

    Private Structure uGuildBid
        Public lGuildID As Int32
        Public BidAmount As Int32
        Public lDuration As Int32
        Public lSlotID As Int32
        Public oGuild As Guild
        Public Function SaveObject(ByVal lPlanetID As Int32) As Boolean
            Dim oComm As OleDb.OleDbCommand = Nothing
            Try
                Dim sSQL As String = "UPDATE tblGuildBillboard SET BidAmount = " & BidAmount & ", Duration = " & lDuration & " WHERE GuildID = " & lGuildID & " AND PlanetID = " & lPlanetID & " AND SlotID = " & lSlotID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    oComm.Dispose()
                    oComm = Nothing
                    sSQL = "INSERT INTO tblGuildBillboard (GuildID, PlanetID, SlotID, BidAmount, Duration) VALUES (" & lGuildID & ", " & lPlanetID & ", " & lSlotID & ", " & BidAmount & ", " & lDuration & ")"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oComm.ExecuteNonQuery()
                End If
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Save Planet.GuildBid: " & ex.Message)
            Finally
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing
            End Try
        End Function
    End Structure
    Private muBids() As uGuildBid
    Private mlBidUB As Int32 = -1

    Public lGuildBillboardSlotIds(5) As Int32

    Public Function AddBid(ByVal lAmt As Int32, ByVal lDuration As Int32, ByRef oGuild As Guild, ByVal lSlotID As Int32, ByVal bSave As Boolean) As Int32
        Dim lResult As Int32 = -1
        If muBids Is Nothing Then ReDim muBids(-1)

        SyncLock muBids
            'Ok, first, remove my bid...
            Dim lPlacement(mlBidUB) As Int32
            For X As Int32 = 0 To mlBidUB
                lPlacement(X) = muBids(X).lGuildID
            Next X

            For X As Int32 = 0 To mlBidUB
                If muBids(X).lGuildID = oGuild.ObjectID AndAlso muBids(X).lSlotID = lSlotID Then
                    For Y As Int32 = X To mlBidUB - 1
                        muBids(Y) = muBids(Y + 1)
                    Next Y
                    mlBidUB -= 1
                    Exit For
                End If
            Next X

            If lAmt < 1 OrElse lDuration < 1 Then Return -1

            'Now, determine where to add this one...
            mlBidUB += 1
            ReDim Preserve muBids(mlBidUB)
            muBids(mlBidUB).BidAmount = Int32.MinValue
            For X As Int32 = 0 To mlBidUB
                If muBids(X).BidAmount < lAmt Then
                    'Ok, here it is...
                    For Y As Int32 = mlBidUB To X + 1 Step -1
                        muBids(Y) = muBids(Y - 1)
                    Next Y

                    With muBids(X)
                        .BidAmount = lAmt
                        .lDuration = lDuration
                        .lGuildID = oGuild.ObjectID
                        .lSlotID = lSlotID
                        .oGuild = oGuild
                    End With

                    lResult = X
                    Exit For
                End If
            Next X

            If lResult > -1 Then
                'TODO: Alert any guild that may be out of luck now???
            End If

            If bSave = True AndAlso lResult > -1 Then
                muBids(lResult).SaveObject(Me.ObjectID)
            End If
        End SyncLock

        Return lResult
    End Function
    Public Sub RemoveBid(ByRef oGuild As Guild, ByVal lSlotID As Int32)
        AddBid(-1, -1, oGuild, lSlotID, False)
        Try
            Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand("DELETE FROM tblGuildBillboard WHERE GuildID = " & oGuild.ObjectID & " AND SlotID = " & lSlotID & " AND PlanetID = " & Me.ObjectID, goCN)
            oComm.ExecuteNonQuery()
        Catch
        End Try
    End Sub
    Public Sub ProcessBids()
        For X As Int32 = 0 To 5
            lGuildBillboardSlotIds(X) = -1
        Next X

        'Ok, go through the bids... determine which ones are still active, deduct their funds and reduce their durations
        Dim lBidSlotsFilled As Int32 = 0

        For X As Int32 = 0 To mlBidUB
            If lGuildBillboardSlotIds(muBids(X).lSlotID - 1) = -1 Then
                'Ok, this bid is available to take a spot... does the bid have duration?
                If muBids(X).lDuration > 0 Then
                    If muBids(X).oGuild Is Nothing = False AndAlso muBids(X).oGuild.blTreasury > muBids(X).BidAmount Then
                        muBids(X).oGuild.blTreasury -= muBids(X).BidAmount
                        lGuildBillboardSlotIds(muBids(X).lSlotID - 1) = X 'muBids(X).lGuildID
                        lBidSlotsFilled += 1
                        muBids(X).lDuration -= 1
                        If lBidSlotsFilled > 5 Then Exit For
                    End If
                End If
            End If
        Next X

    End Sub
    Public Function GetBillBoardsResponse() As Byte()
        Dim yMsg(11 + (lGuildBillboardSlotIds.Length * 214)) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildBillboards).CopyTo(yMsg, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        If Me.ParentSystem Is Nothing = False Then
            System.BitConverter.GetBytes(Me.ParentSystem.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Else
            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        End If
        For X As Int32 = 0 To lGuildBillboardSlotIds.GetUpperBound(0)
            Dim lIdx As Int32 = lGuildBillboardSlotIds(X)
            If lIdx > -1 Then
                Dim oGuild As Guild = muBids(lIdx).oGuild
                If oGuild Is Nothing = False Then
                    System.BitConverter.GetBytes(oGuild.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oGuild.iRecruitFlags).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(oGuild.lIcon).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(muBids(lIdx).BidAmount).CopyTo(yMsg, lPos) : lPos += 4
                    If oGuild.yBillboard Is Nothing = False Then
                        Array.Copy(oGuild.yBillboard, 0, yMsg, lPos, Math.Min(200, oGuild.yBillboard.Length))
                        'oGuild.yBillboard.CopyTo(yMsg, lPos)
                    End If
                    lPos += 200
                Else : lPos += 214
                End If
            Else : lPos += 214
            End If 
        Next X
        Return yMsg
    End Function


    Public Overloads Sub DataChanged()    'call this sub when changing properties of this object
        mbStringReady = False
		'ParentSystem.mbDetailsReady = False
    End Sub

    Public Function GetObjAsString() As Byte()
        'here we will return the object populated as a string
        If mbStringReady = False Then
            'guid-6
            'name-20
            'typeid-1
            'sizeid-1
            'Radius-2
            'parentid-4
            'locx-4
            'locy-4
            'locz-4
            'veg-1
            'atmos-1
            'hyrdo-1
            'grav-1
            'surftemp-1
            'RotDelay-2
            'InnerRingRadius-4
            'OuterRingRadius-4
            'RingDiffuse-4
            'AxisAngle-4
            'CurrentAngle-2
            ReDim mySendString(PLANET_OBJ_MSG_LEN - 1)
            GetGUIDAsString.CopyTo(mySendString, 0)
            PlanetName.CopyTo(mySendString, 6)
            mySendString(26) = PlanetTypeID
            mySendString(27) = PlanetSizeID
            System.BitConverter.GetBytes(PlanetRadius).CopyTo(mySendString, 28)
            System.BitConverter.GetBytes(ParentSystem.ObjectID).CopyTo(mySendString, 30)
            System.BitConverter.GetBytes(LocX).CopyTo(mySendString, 34)
            System.BitConverter.GetBytes(LocY).CopyTo(mySendString, 38)
            System.BitConverter.GetBytes(LocZ).CopyTo(mySendString, 42)
            mySendString(46) = Vegetation
            mySendString(47) = Atmosphere
            mySendString(48) = Hydrosphere
            mySendString(49) = Gravity
            mySendString(50) = SurfaceTemperature
            System.BitConverter.GetBytes(RotationDelay).CopyTo(mySendString, 51)
            System.BitConverter.GetBytes(InnerRingRadius).CopyTo(mySendString, 53)
            System.BitConverter.GetBytes(OuterRingRadius).CopyTo(mySendString, 57)
            System.BitConverter.GetBytes(RingDiffuse).CopyTo(mySendString, 61)
            System.BitConverter.GetBytes(AxisAngle).CopyTo(mySendString, 65)
            System.BitConverter.GetBytes(GetRotateAngle).CopyTo(mySendString, 69)
            mbStringReady = True
        End If

        Return mySendString
    End Function

    Public Function GetRotateAngle() As Int16
        If msw_Rotate Is Nothing Then
            msw_Rotate = Stopwatch.StartNew
        End If
        'Lastly, the CurrentAngle changes constantly, send it back as of right now
        Return CShort(Math.Floor(msw_Rotate.ElapsedMilliseconds / RotationDelay) Mod 3600)
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
		If Me.InMyDomain = False Then Return True

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblPlanet (PlanetTypeID, PlanetName, PlanetSizeID, ParentID, LocX, LocY, LocZ, OwnerID, SentGNSLowRes " & _
                  "Vegetation, Atmosphere, Hydrosphere, Gravity, SurfaceTemperature, RotationDelay, PlanetRadius, PlayerSpawns, SpawnLocked, OrigMinQty, RingMineralID, RingMineralConcentration) VALUES (" & _
                  PlanetTypeID & ", '" & MakeDBStr(BytesToString(PlanetName)) & "', " & PlanetSizeID & ", " & ParentSystem.ObjectID & ", " & _
                  LocX & ", " & LocY & ", " & LocZ & ", " & OwnerID & ", " & ySentGNSLowRes & ", " & Vegetation & ", " & Atmosphere & ", " & Hydrosphere & ", " & _
                  Gravity & ", " & SurfaceTemperature & ", " & RotationDelay & ", " & PlanetRadius & ", " & PlayerSpawns & ", "
                If SpawnLocked = True Then sSQL &= "1" Else sSQL &= "0"
				sSQL &= ", " & blOriginalMineralQuantity.ToString & ")"

            Else
                'UPDATE
				sSQL = "UPDATE tblPlanet SET PlanetTypeID = " & PlanetTypeID & ", PlanetName = '" & MakeDBStr(BytesToString(PlanetName)) & _
				  "', PlanetSizeID = " & PlanetSizeID & ", ParentID = " & ParentSystem.ObjectID & ", LocX = " & _
				  LocX & ", LocY = " & LocY & ", LocZ = " & LocZ & ", Vegetation = " & Vegetation & ", Atmosphere = " & _
				  Atmosphere & ", Hydrosphere = " & Hydrosphere & ", Gravity = " & Gravity & ", SurfaceTemperature = " & _
				  SurfaceTemperature & ", RotationDelay = " & RotationDelay & ", PlanetRadius = " & PlanetRadius & _
				  ", PlayerSpawns = " & PlayerSpawns & ", OwnerID = " & OwnerID & ", SentGNSLowRes = " & ySentGNSLowRes & ", SpawnLocked = "
                If SpawnLocked = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", OrigMinQty = " & blOriginalMineralQuantity.ToString & ", RingMineralID = " & RingMineralID & _
                ", RingMineralConcentration = " & RingMineralConcentration & " WHERE PlanetID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PlanetID) FROM tblPlanet WHERE PlanetName = '" & MakeDBStr(BytesToString(PlanetName)) & "'"
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

    Private Shared mlWidthTotals() As Int32
    Public ReadOnly Property PlanetWidthTotal() As Int32
        Get
            If mlWidthTotals Is Nothing Then
                ReDim mlWidthTotals(4)
                mlWidthTotals(0) = gl_TINY_PLANET_CELL_SPACING * 240
                mlWidthTotals(1) = gl_SMALL_PLANET_CELL_SPACING * 240
                mlWidthTotals(2) = gl_MEDIUM_PLANET_CELL_SPACING * 240
                mlWidthTotals(3) = gl_LARGE_PLANET_CELL_SPACING * 240
                mlWidthTotals(4) = gl_HUGE_PLANET_CELL_SPACING * 240
            End If
            Return mlWidthTotals(PlanetSizeID)
        End Get
    End Property

    Public Sub AddColonyReference(ByVal lColonyIdx As Int32)
        'Dim lIdx As Int32 = -1
        'For X As Int32 = 0 To lColonysHereUB
        '    If lColonysHereIdx(X) = lColonyIdx Then
        '        Return
        '    ElseIf lColonysHereIdx(X) = -1 AndAlso lIdx = -1 Then
        '        lIdx = X
        '    End If
        'Next X

        'If lIdx = -1 Then
        '    lColonysHereUB += 1
        '    ReDim Preserve lColonysHereIdx(lColonysHereUB)
        '    lIdx = lColonysHereUB
        'End If
        'lColonysHereIdx(lIdx) = lColonyIdx
        CheckForColonyLimits(True, gb_Main_Loop_Running = False)
    End Sub

    Public Sub RemoveColonyReference(ByVal lColonyIdx As Int32)
        'For X As Int32 = 0 To lColonysHereUB
        '    If lColonysHereIdx(X) = lColonyIdx Then
        '        lColonysHereIdx(X) = -1
        '    End If
        'Next X
        CheckForColonyLimits(True, False)
    End Sub

    Public Shared Function GetPlanetMaxPlayerSpawnCnt(ByVal lSize As Int32, ByVal ySystemType As Byte) As Int32
        If ySystemType = SolarSystem.elSystemType.SpawnSystem Then
            Select Case lSize
                'Case 0 : Return 2       'Tiny
                'Case 1 : Return 2       'Small      -was 3
                'Case 2 : Return 3       'Medium     -was 5
                'Case 3 : Return 5       'Large      -was 7
                'Case 4 : Return 6       'Huge       -was 9
                'Case Else
                '    Return 7            'Huge       -was 10
                Case 0 : Return 2       'Tiny
                Case 1 : Return 2       'Small      -was 3
                Case 2 : Return 2       'Medium     -was 5
                Case 3 : Return 2       'Large      -was 7
                Case 4 : Return 2       'Huge       -was 9
                Case Else
                    Return 2            'Huge       -was 10
            End Select
        ElseIf ySystemType = SolarSystem.elSystemType.RespawnSystem Then
            Select Case lSize
                'Case 0 : Return 2       'Tiny
                'Case 1 : Return 3       'Small      
                'Case 2 : Return 5       'Medium     
                'Case 3 : Return 7       'Large      
                'Case 4 : Return 9       'Huge       
                'Case Else
                '    Return 10           'Huge     
                Case 0 : Return 2       'Tiny
                Case 1 : Return 2       'Small      
                Case 2 : Return 2       'Medium     
                Case 3 : Return 2       'Large      
                Case 4 : Return 2       'Huge       
                Case Else
                    Return 2           'Huge    
            End Select
        End If

    End Function

    Public Function GetGuildColonyCount(ByVal lGuildID As Int32) As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lColonysHereUB
            If lColonysHereIdx(X) > -1 AndAlso glColonyIdx(lColonysHereIdx(X)) > -1 Then
                Dim oColony As Colony = goColony(lColonysHereIdx(X))
                If oColony Is Nothing = False Then
                    If oColony.ParentObject Is Nothing = False Then
                        If CType(oColony.ParentObject, Epica_GUID).ObjectID = Me.ObjectID Then
                            If oColony.Owner Is Nothing = False AndAlso oColony.Owner.oGuild Is Nothing = False AndAlso oColony.Owner.oGuild.ObjectID = lGuildID Then
                                lCnt += 1
                            End If
                        End If
                    End If
                End If
            End If
        Next X
        Return lCnt
    End Function

    Public Sub CheckForColonyLimits(ByVal bRebuildList As Boolean, ByVal bSilent As Boolean)

        Dim lUB As Int32 = -1
        If glColonyIdx Is Nothing = False Then lUB = Math.Max(glColonyUB, glColonyIdx.GetUpperBound(0))

        Dim lCount As Int32 = 0

        Dim lNewColonyHereUB As Int32 = -1
        Dim lNewColonyHere(-1) As Int32

        For X As Int32 = 0 To lUB
            If glColonyIdx(X) > 0 Then
                Dim oCol As Colony = goColony(X)
                If oCol Is Nothing = False Then
                    If oCol.ParentObject Is Nothing = False AndAlso CType(oCol.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                        If CType(oCol.ParentObject, Epica_GUID).ObjectID = Me.ObjectID Then
                            'Ok, we're in here
                            lCount += 1

                            If bRebuildList = True Then
                                lNewColonyHereUB += 1
                                ReDim Preserve lNewColonyHere(lNewColonyHereUB)
                                lNewColonyHere(lNewColonyHereUB) = X
                            End If
                        End If
                    End If
                End If
            End If
        Next X

        If bRebuildList = True Then
            'NOTE: can be dangerous if the code elsewhere is accessing the array while adjusting it
            lColonysHereUB = -1
            lColonysHereIdx = lNewColonyHere
            lColonysHereUB = lNewColonyHereUB
        End If

        'Now, check our count
        Dim bNewVal As Boolean = False
        Dim lMax As Int32 = CInt(Me.PlanetSizeID) + 2
        bNewVal = lCount > lMax

        If bNewVal <> bAtColonyLimit Then
            bAtColonyLimit = bNewVal

            If bSilent = False Then
                Dim sPlanet As String = BytesToString(Me.PlanetName)
                Dim sMsg As String
                Dim sTitle As String
                Dim lSentOn As Int32 = GetDateAsNumber(Now)
                'Ok, we are shifting... 
                If bNewVal = True Then
                    'Ok, colony limit reached
                    sMsg = "There are presently " & lCount.ToString & " colonies on " & sPlanet & " which can only have " & lMax.ToString & " colonies. The governments of the planet are breaking down and becoming ineffective causing mass corruption for the entire planet."
                    sMsg &= vbCrLf & vbCrLf & "Due to this corruption:" & vbCrLf & "  * Your colony will no longer generate revenue but will maintain an expense." & vbCrLf
                    sMsg &= "  * Mining has been interrupted and all mines on " & sPlanet & " no longer generate minerals." & vbCrLf
                    sMsg &= "  * Colonies will no longer grow, however they can be reduced in population." & vbCrLf
                    sMsg &= "  * Production has halted." & vbCrLf
                    sMsg &= "  * Invulnerability will not work for this planet."
                    sMsg &= vbCrLf & vbCrLf & "Reduce the number of colonies on the planet in order to return order."

                    sTitle = "Colony Limit Reached on " & sPlanet
                Else
                    'ok, colony limit removed
                    sMsg = "The chaos on " & sPlanet & " has subsided as the governments regain control. All corruption penalties have been removed that were imposed because of the colony limit being exceeded on the planet."
                    sTitle = "Colony Limit Corruption Over on " & sPlanet
                End If

                lUB = -1
                If lColonysHereIdx Is Nothing = False Then lUB = Math.Max(lColonysHereUB, lColonysHereIdx.GetUpperBound(0))
                For X As Int32 = 0 To lUB
                    Dim oCol As Colony = goColony(lColonysHereIdx(X))
                    If oCol Is Nothing = False Then
                        If oCol.ParentObject Is Nothing = False AndAlso CType(oCol.ParentObject, Epica_GUID).ObjectID = Me.ObjectID AndAlso CType(oCol.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                            If oCol.Owner Is Nothing = False Then

                                If oCol.Owner.lStartedEnvirID = Me.ObjectID AndAlso oCol.Owner.iStartedEnvirTypeID = Me.ObjTypeID Then Continue For

                                Dim oPC As PlayerComm = oCol.Owner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sMsg, sTitle, oCol.Owner.ObjectID, lSentOn, False, oCol.Owner.sPlayerNameProper, Nothing)
                                If oPC Is Nothing = False Then
                                    oCol.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                End If
                            End If
                        End If
                    End If
                Next X
            End If

        End If
    End Sub

    Public Function PlanetInCorruption(ByVal lPlayerStartEnvirID As Int32, ByVal iPlayerStartEnvirTypeID As Int16) As Boolean
        Return mbAtColonyLimit = True AndAlso (Me.ObjectID <> lPlayerStartEnvirID OrElse Me.ObjTypeID <> iPlayerStartEnvirTypeID)
    End Function

    Public Sub MicroExplosion()

        LogEvent(LogEventType.Informational, "MicroExplosion event on " & BytesToString(Me.PlanetName) & " (" & Me.ObjectID & ", " & Me.ObjTypeID & ").")

        Dim oSB As New System.Text.StringBuilder()
        oSB.AppendLine("A massive planet-wide explosion has been detected on " & BytesToString(Me.PlanetName) & " eradicating everything on the planet's surface.")
        oSB.AppendLine()
        If Me.Atmosphere > 0 Then
            oSB.AppendLine("Reports are still coming in but there is a massive amount of residual quantum flux in the soil and in the atmosphere.")
        Else
            oSB.AppendLine("Reports are still coming in but there is a massive amount of residual quantum flux in the soil.")
        End If
        oSB.AppendLine()
        oSB.AppendLine("Survey teams from the galactic senate have been dispatched to try and determine the cause of this catastrophe.")

        'notify colony owners
        Dim lUB As Int32 = -1
        If goColony Is Nothing = False Then lUB = Math.Min(glColonyUB, goColony.GetUpperBound(0))
        For X As Int32 = 0 To lUB
            If glColonyIdx(X) > -1 Then
                Dim oColony As Colony = goColony(X)
                If oColony Is Nothing = False Then
                    If oColony.ParentObject Is Nothing Then Continue For
                    With CType(oColony.ParentObject, Epica_GUID)
                        If .ObjectID <> Me.ObjectID OrElse .ObjTypeID <> Me.ObjTypeID Then Continue For
                    End With
                    Dim oPC As PlayerComm = oColony.Owner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString & vbCrLf & vbCrLf & BytesToString(oColony.ColonyName) & " has been wiped out.", "Catastrophic Event", oColony.Owner.ObjectID, GetDateAsNumber(Now), False, oColony.Owner.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then
                        oColony.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                    oColony.Population = 0
                End If
            End If
        Next

        'destroy all units and facilities on this planet
        lUB = -1
        If goUnit Is Nothing = False Then lUB = Math.Min(glUnitUB, goUnit.GetUpperBound(0))
        For X As Int32 = 0 To lUB
            If glUnitIdx(X) > -1 Then
                Dim oUnit As Unit = goUnit(X)
                If oUnit Is Nothing = False AndAlso oUnit.ParentObject Is Nothing = False Then
                    Dim oPG As Epica_GUID = CType(oUnit.ParentObject, Epica_GUID)
                    If oPG.ObjTypeID = ObjectType.ePlanet AndAlso oPG.ObjectID = Me.ObjectID Then
                        DestroyEntity(CType(oUnit, Epica_Entity), True, -1, False, "MicroExplosion")
                    End If
                End If
            End If
        Next X
        lUB = -1
        If goFacility Is Nothing = False Then lUB = Math.Min(glFacilityUB, goFacility.GetUpperBound(0))
        For X As Int32 = 0 To lUB
            If glFacilityIdx(X) > -1 Then
                Dim oFac As Facility = goFacility(X)
                If oFac Is Nothing = False AndAlso oFac.ParentObject Is Nothing = False Then
                    Dim oPG As Epica_GUID = CType(oFac.ParentObject, Epica_GUID)
                    If oPG.ObjTypeID = ObjectType.ePlanet AndAlso oPG.ObjectID = Me.ObjectID Then
                        DestroyEntity(CType(oFac, Epica_Entity), True, -1, False, "MicroExplosion")
                    End If
                End If
            End If
        Next X
    End Sub
End Class
