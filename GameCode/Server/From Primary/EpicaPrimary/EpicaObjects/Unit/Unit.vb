Public Class Unit
    Inherits Epica_Entity

    Public EntityDef As Epica_Entity_Def

	Public lPirate_For_PlayerID As Int32 = 0

    Public oRebuilderFor As Colony = Nothing        'pointer to a colony to rebuild for...

    Private moProdQueue() As EntityProduction
    Private mlProdQueueUB As Int32 = -1
    Public bProdQueueMoveSent As Boolean = False

	Private mbSmallStringReady As Boolean = False
    Private mySendString() As Byte
    Private mySmallSendString() As Byte
    Public Overloads Sub DataChanged()    'call this sub when changing properties of this object
        mbStringReady = False
        mbSmallStringReady = False
        bNeedsSaved = True
    End Sub

    Public bUnitInSellOrder As Boolean = False

#Region "  Cargo Route System  "
	Public uRoute() As RouteItem
	Public lRouteUB As Int32 = -1
	Public lCurrentRouteIdx As Int32 = -1
	Public bRoutePaused As Boolean = False
	Public bRunRouteOnce As Boolean = False
    Private mbInProcessNextRouteItem As Boolean = False
    Public lRetryCount As Int32 = 0

	Public Sub ProcessNextRouteItem()
        'If Me.ObjectID = 2645830 Then LogEvent(LogEventType.Informational, "ProcessNextRouteItem " & mbInProcessNextRouteItem.ToString & ", " & lRouteUB.ToString & ", " & bRoutePaused.ToString)
        lRetryCount = 0

        If mbInProcessNextRouteItem = False Then
            mbInProcessNextRouteItem = True
            AddToQueue(glCurrentCycle + 150, QueueItemType.eProcessNextRouteItem, Me.ObjectID, -1, -1, -1, -1, -1, -1, -1)
            Return
        Else
            mbInProcessNextRouteItem = False
        End If

        If lRouteUB = -1 Then Return
		If bRoutePaused = True Then Return
		lCurrentRouteIdx += 1
		If lCurrentRouteIdx > lRouteUB Then
			lCurrentRouteIdx = 0
			If bRunRouteOnce = True Then
				bRunRouteOnce = False
				Return
			End If
		End If
		CurrentRouteItemAction()
	End Sub
    Public Sub CurrentRouteItemAction()
        'If Me.ObjectID = 2645830 Then LogEvent(LogEventType.Informational, "CurrentRouteItemAction: " & lCurrentRouteIdx & ", " & (Me.ParentObject Is Nothing).ToString)
        If lCurrentRouteIdx > -1 AndAlso lCurrentRouteIdx <= lRouteUB Then
            With uRoute(lCurrentRouteIdx)
                'ok, the step... it is always a move order... determine our parent object
                If Me.ParentObject Is Nothing Then Return
                If .oDest Is Nothing Then Return

                Dim oDomainSock As NetSock = Nothing
                Dim iTemp As Int16 = CType(Me.ParentObject, Epica_GUID).ObjTypeID
                If iTemp = ObjectType.ePlanet Then
                    oDomainSock = CType(Me.ParentObject, Planet).oDomain.DomainSocket
                ElseIf iTemp = ObjectType.eSolarSystem Then
                    oDomainSock = CType(Me.ParentObject, SolarSystem).oDomain.DomainSocket
                End If
                If oDomainSock Is Nothing Then
                    With CType(Me.ParentObject, Epica_GUID)
                        LogEvent(LogEventType.Warning, "CurrentRouteItemAction: oDomainSock is nothing! UnitID: " & Me.ObjectID & ", Current Parent: " & .ObjectID & ", " & .ObjTypeID)
                    End With
                    Return
                End If

                'Ok, check if this is a ground or naval unit...
                If (EntityDef.yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then
                    'Ok, we have a parent... get the parent geo object whether that be a planet or system
                    Dim oPGuid As Epica_GUID = CType(Me.ParentObject, Epica_GUID)
                    If oPGuid Is Nothing = False Then
                        Select Case oPGuid.ObjTypeID
                            'Case ObjectType.ePlanet - nothing to do
                            Case ObjectType.eSolarSystem
                                Return
                            Case ObjectType.eUnit
                                oPGuid = CType(CType(oPGuid, Unit).ParentObject, Epica_GUID)
                            Case ObjectType.eFacility
                                oPGuid = CType(CType(oPGuid, Facility).ParentObject, Epica_GUID)
                        End Select
                    End If
                    If oPGuid Is Nothing OrElse oPGuid.ObjTypeID <> ObjectType.ePlanet Then Return

                    'Then, we have a dest, get the dest geo object whether that be a planet or system
                    Dim oDestP As Epica_GUID = .oDest
                    If .oDest Is Nothing = False Then
                        Select Case .oDest.ObjTypeID
                            Case ObjectType.eSolarSystem
                                Return
                            Case ObjectType.eUnit
                                oDestP = CType(CType(.oDest, Unit).ParentObject, Epica_GUID)
                            Case ObjectType.eFacility
                                oDestP = CType(CType(.oDest, Facility).ParentObject, Epica_GUID)
                        End Select
                    End If
                    If oDestP Is Nothing OrElse oDestP.ObjTypeID <> ObjectType.ePlanet Then Return

                    'make sure the dest is reachable by a land/naval unit. If not, return
                    If oPGuid.ObjectID <> oDestP.ObjectID Then Return
                End If

                If .oDest.ObjTypeID = ObjectType.eFacility Then
                    'Ok, ensure the facility still exists
                    Dim oFac As Facility = CType(.oDest, Facility)
                    Dim bGood As Boolean = False
                    If oFac Is Nothing = False Then
                        If oFac.ServerIndex > -1 AndAlso glFacilityIdx(oFac.ServerIndex) = oFac.ObjectID Then
                            bGood = True
                        End If
                    End If
                    If bGood = False Then
                        lRouteUB = -1
                        lCurrentRouteIdx = -1
                        Return
                    End If
                End If

                If .oDest.ObjTypeID = ObjectType.eWormhole Then
                    Dim yMsg(19) As Byte
                    Dim lPos As Int32 = 0

                    System.BitConverter.GetBytes(GlobalMessageCode.eRouteMoveCommand).CopyTo(yMsg, lPos) : lPos += 2
                    Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    .oDest.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    CType(Me.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    oDomainSock.SendData(yMsg)
                Else
                    'Ok, we're here... we send a special message for this...
                    Dim yMsg(27) As Byte
                    Dim lPos As Int32 = 0

                    System.BitConverter.GetBytes(GlobalMessageCode.eRouteMoveCommand).CopyTo(yMsg, lPos) : lPos += 2
                    Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    .oDest.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    System.BitConverter.GetBytes(.lLocX).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lLocZ).CopyTo(yMsg, lPos) : lPos += 4

                    Dim bParentAdded As Boolean = False
                    If .oDest.ObjTypeID = ObjectType.eUnit OrElse .oDest.ObjTypeID = ObjectType.eFacility Then
                        Dim oTmpP As Epica_GUID = CType(CType(.oDest, Epica_Entity).ParentObject, Epica_GUID)
                        If oTmpP Is Nothing = False Then
                            oTmpP.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                            bParentAdded = True
                        End If
                    End If
                    If bParentAdded = False Then
                        .oDest.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    End If

                    oDomainSock.SendData(yMsg)
                End If

            End With
        End If
    End Sub
	Public Function RouteContainsFacility(ByVal lFacID As Int32) As Boolean
		For X As Int32 = 0 To lRouteUB
			Dim oTmp As Epica_GUID = uRoute(X).oDest
			If oTmp Is Nothing = False Then
				If oTmp.ObjTypeID = ObjectType.eFacility AndAlso oTmp.ObjectID = lFacID Then Return True
			End If
		Next X
		Return False
    End Function
    Public Function AcceptingColonists() As Boolean
        If lCurrentRouteIdx > -1 AndAlso lCurrentRouteIdx <= lRouteUB Then
            Return uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadColonists
        End If
        Return False
    End Function
    Public Function AcceptingEnlisted() As Boolean
        If lCurrentRouteIdx > -1 AndAlso lCurrentRouteIdx <= lRouteUB Then
            Return uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadEnlisted
        End If
        Return False
    End Function
    Public Function AcceptingOfficers() As Boolean
        If lCurrentRouteIdx > -1 AndAlso lCurrentRouteIdx <= lRouteUB Then
            Return uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadOfficers
        End If
        Return False
    End Function
    Public Function AcceptingCargo(ByVal lItemID As Int32, ByVal iParentTypeID As Int16, ByVal iCacheTypeID As Int16) As Boolean
        If lCurrentRouteIdx > -1 AndAlso lCurrentRouteIdx <= lRouteUB Then
            If uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadAbsolutelyNothing Then Return False
            If uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadAnyAllItems Then Return True 'If uRoute(lCurrentRouteIdx).bLoadAllItems = True Then Return True

            Select Case uRoute(lCurrentRouteIdx).yLoadAllItems
                Case eyRouteLoadItemType.eLoadAllArmor
                    If iParentTypeID = ObjectType.eArmorTech Then Return True
                Case eyRouteLoadItemType.eLoadAllComponents
                    If iCacheTypeID = ObjectType.eComponentCache Then Return True
                Case eyRouteLoadItemType.eLoadAllEngines
                    If iParentTypeID = ObjectType.eEngineTech Then Return True
                Case eyRouteLoadItemType.eLoadAllMinerals
                    If iCacheTypeID = ObjectType.eMineralCache Then Return True
                Case eyRouteLoadItemType.eLoadAllRadar
                    If iParentTypeID = ObjectType.eRadarTech Then Return True
                Case eyRouteLoadItemType.eLoadAllShields
                    If iParentTypeID = ObjectType.eShieldTech Then Return True
                Case eyRouteLoadItemType.eLoadAllWeapons
                    If iParentTypeID = ObjectType.eWeaponTech Then Return True
                Case eyRouteLoadItemType.eLoadColonists
                    If iParentTypeID = ObjectType.eColonists Then Return True
                Case eyRouteLoadItemType.eLoadEnlisted
                    If iParentTypeID = ObjectType.eEnlisted Then Return True
                Case eyRouteLoadItemType.eLoadOfficers
                    If iParentTypeID = ObjectType.eOfficers Then Return True
            End Select

            Dim oTmp As Epica_GUID = uRoute(lCurrentRouteIdx).oLoadItem
            If oTmp Is Nothing = False AndAlso (oTmp.ObjectID = lItemID AndAlso oTmp.ObjTypeID = iParentTypeID) Then
                Return True
            End If
        End If
        Return False
    End Function
    Public Function NeedToTryMinerals() As Boolean
        If lCurrentRouteIdx > -1 AndAlso lCurrentRouteIdx <= lRouteUB Then
            If uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadAbsolutelyNothing Then Return False
            If uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadColonists OrElse _
               uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadEnlisted OrElse _
               uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadOfficers Then Return False
            If uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadAnyAllItems Then Return True 'If uRoute(lCurrentRouteIdx).bLoadAllItems = True Then Return True

            Select Case uRoute(lCurrentRouteIdx).yLoadAllItems
                Case eyRouteLoadItemType.eLoadAllMinerals
                    Return True
            End Select

            Dim oTmp As Epica_GUID = uRoute(lCurrentRouteIdx).oLoadItem
            If oTmp Is Nothing = False AndAlso oTmp.ObjTypeID = ObjectType.eMineral Then
                Return True
            End If
        End If
        Return False
    End Function
    Public Function NeedToTryComponents() As Boolean
        If lCurrentRouteIdx > -1 AndAlso lCurrentRouteIdx <= lRouteUB Then
            If uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadAbsolutelyNothing Then Return False
            If uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadColonists OrElse _
              uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadEnlisted OrElse _
              uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadOfficers Then Return False
            If uRoute(lCurrentRouteIdx).yLoadAllItems = eyRouteLoadItemType.eLoadAnyAllItems Then Return True 'If uRoute(lCurrentRouteIdx).bLoadAllItems = True Then Return True

            Select Case uRoute(lCurrentRouteIdx).yLoadAllItems
                Case eyRouteLoadItemType.eLoadAllArmor
                    Return True
                Case eyRouteLoadItemType.eLoadAllComponents
                    Return True
                Case eyRouteLoadItemType.eLoadAllEngines
                    Return True
                Case eyRouteLoadItemType.eLoadAllRadar
                    Return True
                Case eyRouteLoadItemType.eLoadAllShields
                    Return True
                Case eyRouteLoadItemType.eLoadAllWeapons
                    Return True
            End Select

            Dim oTmp As Epica_GUID = uRoute(lCurrentRouteIdx).oLoadItem
            If oTmp Is Nothing = False AndAlso oTmp.ObjTypeID <> ObjectType.eComponentCache Then
                Return True
            End If
        End If
        Return False
    End Function
    Public Sub CheckRouteArrival()
        If (Me.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then Return
        If lCurrentRouteIdx > -1 AndAlso lCurrentRouteIdx <= lRouteUB Then
            Dim oTmpParent As Epica_GUID = CType(Me.ParentObject, Epica_GUID)
            Dim oRouteDest As Epica_GUID = uRoute(lCurrentRouteIdx).oDest
            If oTmpParent Is Nothing = False AndAlso oRouteDest Is Nothing = False Then
                If oTmpParent.ObjectID = oRouteDest.ObjectID AndAlso oTmpParent.ObjTypeID = oRouteDest.ObjTypeID Then
                    'Ok, we are in fact here... Unload our cargo here if our parent's prod type shares colony cargo
                    If oTmpParent.ObjTypeID = ObjectType.eUnit OrElse oTmpParent.ObjTypeID = ObjectType.eFacility Then
                        Dim bDoUnload As Boolean = False
                        Try
                            With CType(oTmpParent, Epica_Entity)
                                If .Owner Is Nothing = False AndAlso .Owner.ObjectID = Me.Owner.ObjectID Then
                                    If (uRoute(lCurrentRouteIdx).yExtraFlags And 1) <> 0 Then
                                        If oRouteDest.ObjTypeID = ObjectType.eFacility Then
                                            If Colony.ProductionTypeSharesColonyCargo(CType(oRouteDest, Epica_Entity).yProductionType) = True Then
                                                bDoUnload = True
                                            End If
                                        ElseIf oRouteDest.ObjTypeID = ObjectType.eUnit Then
                                            bDoUnload = True
                                        End If
                                        If bDoUnload = True Then
                                            'ok, unload...
                                            Me.TransferCargo(CType(oTmpParent, Epica_Entity), False)
                                        End If
                                    End If
                                End If

                                'And then check if we are loading any cargo here...
                                If (.Owner Is Nothing = False AndAlso .Owner.ObjectID = Me.Owner.ObjectID) OrElse .yProductionType = ProductionType.eMining Then
                                    If uRoute(lCurrentRouteIdx).yLoadAllItems <> eyRouteLoadItemType.eNoLoadAllItems Then
                                        .TransferCargo(Me, True)
                                    Else
                                        Dim oPickup As Epica_GUID = uRoute(lCurrentRouteIdx).oLoadItem
                                        If oPickup Is Nothing = False Then
                                            .TransferCargo(Me, True)
                                        End If
                                    End If
                                End If
                            End With
                        Catch ex As Exception
                            LogEvent(LogEventType.Warning, "CheckRouteArrival: " & ex.Message)
                        End Try
                    End If

                    'Ok, we are here, so now, if my parent is a facility or unit...
                    If oTmpParent.ObjTypeID = ObjectType.eFacility OrElse oTmpParent.ObjTypeID = ObjectType.eUnit Then
                        AddToQueue(glCurrentCycle + 30, QueueItemType.eUndockAndReturnToRefinery_QIT, ObjectID, ObjTypeID, oTmpParent.ObjectID, oTmpParent.ObjTypeID, 0, 0, 0, 0)
                    Else
                        ProcessNextRouteItem()
                    End If
                Else
                    LogEvent(LogEventType.Warning, "CheckRouteArrival: Parent/Dest Mismatch")
                End If
            Else
                LogEvent(LogEventType.Warning, "CheckRouteArrival: Parent or Dest Nothing")
            End If
        End If
    End Sub
	Public Sub AddRouteItem(ByVal uNewRouteItem As RouteItem)
		Dim uTmp() As RouteItem = uRoute
		Dim lTmpUB As Int32 = lRouteUB
		Dim lOrderNum As Int32 = uNewRouteItem.lOrderNum
		For X As Int32 = 0 To lTmpUB
			If uTmp(X).lOrderNum > lOrderNum Then
				'ok, this is an insert into the array at this point
				lTmpUB += 1
				ReDim Preserve uTmp(lTmpUB)
				For Y As Int32 = lTmpUB To X + 1 Step -1
					uTmp(Y) = uTmp(Y - 1)
				Next Y
				uTmp(X) = uNewRouteItem

				uRoute = uTmp
				lRouteUB = lTmpUB
				Return
			End If
		Next X

		'Ok, if we are here, then we need to add it to the end
		lTmpUB += 1
		ReDim Preserve uTmp(lTmpUB)
		uTmp(lTmpUB) = uNewRouteItem
		uRoute = uTmp
        lRouteUB = lTmpUB

        bNeedsSaved = True
	End Sub
	Public Sub RemoveRouteItem(ByVal lIdx As Int32)

		If lIdx = Int32.MinValue Then
			lRouteUB = -1
			Return
		End If

		If lIdx < 0 OrElse lIdx > lRouteUB Then Return

		Dim uTmp() As RouteItem = uRoute
		Dim lTmpUB As Int32 = lRouteUB

		For X As Int32 = lIdx To lTmpUB - 1
			uTmp(X) = uTmp(X + 1)
		Next X
		lTmpUB -= 1

		uRoute = uTmp
        lRouteUB = lTmpUB

        bNeedsSaved = True
	End Sub
#End Region

#Region " Properties For Asynchronous Persistence "
	'Private mlRefineryIndex As Int32 = -1      'SERVER INDEX NOT ID
	'Public Property lRefineryIndex() As Int32
	'    Get
	'        Return mlRefineryIndex
	'    End Get
	'    Set(ByVal value As Int32)
	'        mlRefineryIndex = value
	'        bNeedsSaved = True
	'    End Set
	'End Property
	'Private mlMiningFacIndex As Int32 = -1     'SERVER INDEX NOT ID
	'Public Property lMiningFacIndex() As Int32
	'    Get
	'        Return mlMiningFacIndex
	'    End Get
	'    Set(ByVal value As Int32)
	'        mlMiningFacIndex = value
	'        bNeedsSaved = True
	'    End Set
	'End Property
    Private mlDestX As Int32
    Public Property DestX() As Int32
        Get
            Return mlDestX
        End Get
        Set(ByVal value As Int32)
            mlDestX = value
            bNeedsSaved = True
        End Set
    End Property
    Private mlDestZ As Int32
    Public Property DestZ() As Int32
        Get
            Return mlDestZ
        End Get
        Set(ByVal value As Int32)
            mlDestZ = value
            bNeedsSaved = True
        End Set
    End Property
    Private mlDestEnvirID As Int32 = -1
    Public Property DestEnvirID() As Int32
        Get
            Return mlDestEnvirID
        End Get
        Set(ByVal value As Int32)
            mlDestEnvirID = value
            bNeedsSaved = True
        End Set
    End Property
    Private miDestEnvirTypeID As Int16 = -1
    Public Property DestEnvirTypeID() As Int16
        Get
            Return miDestEnvirTypeID
        End Get
        Set(ByVal value As Int16)
            miDestEnvirTypeID = value
            bNeedsSaved = True
        End Set
    End Property
    Private mlFleetID As Int32 = -1
    Public Property lFleetID() As Int32
        Get
            Return mlFleetID
        End Get
        Set(ByVal value As Int32)

            If True = False Then
                If mlFleetID > 0 AndAlso value < 1 AndAlso Me.ParentObject Is Nothing = False AndAlso CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eGalaxy Then
                    Stop
                End If
            End If
            If value > 0 Then lTestFleetID = value

            mlFleetID = value
            bNeedsSaved = True
        End Set
    End Property

    Private lTestFleetID As Int32 = -1

    Private mlTetherPointX As Int32 = Int32.MinValue
    Public Property TetherPointX() As Int32
        Get
            Return mlTetherPointX
        End Get
        Set(ByVal value As Int32)
            mlTetherPointX = value
            bNeedsSaved = True
        End Set
    End Property
    Private mlTetherPointZ As Int32 = Int32.MinValue
    Public Property TetherPointZ() As Int32
        Get
            Return mlTetherPointZ
        End Get
        Set(ByVal value As Int32)
            mlTetherPointZ = value
            bNeedsSaved = True
        End Set
    End Property

    Public Sub DoAsyncPersistence()
        'Dim oComm As OleDb.OleDbCommand = Nothing
        'Dim sSQL As String = ""
        'Try
        If DestEnvirID < 1 OrElse DestEnvirTypeID < 0 Then
            DestX = LocX : DestZ = LocZ : DestEnvirID = CType(ParentObject, Epica_GUID).ObjectID
            DestEnvirTypeID = CType(ParentObject, Epica_GUID).ObjTypeID
        End If
        SaveObject()

        'Dim lSourceID As Int32 = -1
        'Dim iSourceTypeID As Int16 = -1
        'Dim lDropoffID As Int32 = -1
        'Dim iDropoffTypeID As Int16 = -1

        ''are we pulling from a cache (aka mining truck) or from a mining facility (aka cargo truck)
        'If lCacheIndex <> -1 AndAlso glMineralCacheIdx(lCacheIndex) <> -1 Then
        '    lSourceID = glMineralCacheIdx(lCacheIndex)
        '    iSourceTypeID = ObjectType.eMineral
        'ElseIf lMiningFacIndex <> -1 AndAlso glFacilityIdx(lMiningFacIndex) <> -1 Then
        '    lSourceID = glFacilityIdx(lMiningFacIndex)
        '    iSourceTypeID = ObjectType.eFacility
        'End If

        ''Ok, get our refinery
        'If lRefineryIndex <> -1 AndAlso glFacilityIdx(lRefineryIndex) <> -1 Then
        '    lDropoffID = glFacilityIdx(lRefineryIndex)
        '    iDropoffTypeID = ObjectType.eFacility
        'End If

        'sSQL = "UPDATE tblUnit SET DestX = " & DestX & ", DestZ = " & DestZ & ", SourceID = " & lSourceID & _
        '  ", SourceTypeID = " & iSourceTypeID & ", DropOffID = " & lDropoffID & ", DropOffTypeID = " & iDropoffTypeID & _
        '  ", DestEnvirID = " & DestEnvirID & ", DestEnvirTypeID = " & DestEnvirTypeID & ", Shield_HP = " & Shield_HP & _
        '  ", Structure_HP = " & Structure_HP & ", ExpLevel = " & ExpLevel & ", Q1_HP = " & Q1_HP & ", Q2_HP = " & _
        '  Q2_HP & ", Q3_HP = " & Q3_HP & ", Q4_HP = " & Q4_HP & ", CurrentStatus = " & CurrentStatus & ", LocX = " & _
        '  LocX & ", LocY = " & LocZ & ", ParentID = " & CType(Me.ParentObject, Epica_GUID).ObjectID & ", ParentTypeID = " & _
        '  " WHERE UnitID = " & Me.ObjectID
        'oComm = New OleDb.OleDbCommand(sSQL, goCN)
        'oComm.ExecuteNonQuery()
        'Catch ex As Exception
        '    LogEvent(LogEventType.CriticalError, "Unit.DoAsyncPersistence (" & sSQL & "): " & ex.Message)
        'Finally
        '    If oComm Is Nothing = False Then oComm.Dispose()
        '    oComm = Nothing
        'End Try

        'Clear our need save flag
        bNeedsSaved = False
    End Sub
#End Region

    Public Function GetObjAsString() As Byte()

        If mbStringReady = False Then
            'NOTE: because the object we are returning is associated to a UnitDefID... it is assumed that the request
            '  sender has the unit def object. Or in the case of a Client... if the unitDef is not there, it should
            '  send a subsequent request for the Unit Def object based on the UnitDefID.
            ReDim mySendString(95 + (8 * lCurrentAmmo.Length) + 8)  '89
            Dim lPos As Int32
            Dim X As Int32

            GetGUIDAsString.CopyTo(mySendString, 0)

            'TODO: this could be very dangerous... the IDX array stores -1 for objects that no longer exist...
            If ParentObject Is Nothing Then
                System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, 6)
                System.BitConverter.GetBytes(CShort(-1)).CopyTo(mySendString, 10)
            Else
                CType(ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, 6)
            End If
            System.BitConverter.GetBytes(EntityDef.ObjectID).CopyTo(mySendString, 12)
            EntityName.CopyTo(mySendString, 16)
            If Owner Is Nothing Then
                System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, 36)
            Else
                System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySendString, 36)
            End If
            System.BitConverter.GetBytes(LocX).CopyTo(mySendString, 40)
            System.BitConverter.GetBytes(LocZ).CopyTo(mySendString, 44)
            System.BitConverter.GetBytes(LocAngle).CopyTo(mySendString, 48)
            System.BitConverter.GetBytes(Structure_HP).CopyTo(mySendString, 50)
            System.BitConverter.GetBytes(Fuel_Cap).CopyTo(mySendString, 54)
            System.BitConverter.GetBytes(Shield_HP).CopyTo(mySendString, 58)
            mySendString(62) = ExpLevel
            'If oUnitGroup Is Nothing Then
            System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, 63)
            'Else
            '    System.BitConverter.GetBytes(oUnitGroup.ObjectID).CopyTo(mySendString, 59)
            'End If
            System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySendString, 67)

            System.BitConverter.GetBytes(Q1_HP).CopyTo(mySendString, 71)
            System.BitConverter.GetBytes(Q2_HP).CopyTo(mySendString, 75)
            System.BitConverter.GetBytes(Q3_HP).CopyTo(mySendString, 79)
            System.BitConverter.GetBytes(Q4_HP).CopyTo(mySendString, 83)

            mySendString(87) = yProductionType

            System.BitConverter.GetBytes(iCombatTactics).CopyTo(mySendString, 88)
            System.BitConverter.GetBytes(iTargetingTactics).CopyTo(mySendString, 92)

            System.BitConverter.GetBytes(CShort(lCurrentAmmo.Length)).CopyTo(mySendString, 94)
            lPos = 96

            For X = 0 To lCurrentAmmo.Length - 1
                System.BitConverter.GetBytes(EntityDef.WeaponDefs(X).ObjectID).CopyTo(mySendString, lPos) : lPos += 4
                System.BitConverter.GetBytes(lCurrentAmmo(X)).CopyTo(mySendString, lPos) : lPos += 4
            Next X

            System.BitConverter.GetBytes(TetherPointX).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(TetherPointZ).CopyTo(mySendString, lPos) : lPos += 4

            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Sub New()
        ReDim lCurrentAmmo(-1)
    End Sub

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If DestEnvirID < 1 OrElse DestEnvirTypeID < 0 Then
                DestX = LocX : DestZ = LocZ : DestEnvirID = CType(ParentObject, Epica_GUID).ObjectID
                DestEnvirTypeID = CType(ParentObject, Epica_GUID).ObjTypeID
            End If

            If lRouteUB > -1 AndAlso bRoutePaused = False Then
                CurrentSpeed = 1
            Else : CurrentSpeed = 0
            End If

			'Dim lSourceID As Int32 = -1
			'Dim iSourceTypeID As Int16 = -1
			'Dim lDropoffID As Int32 = -1
			'Dim iDropoffTypeID As Int16 = -1

			''are we pulling from a cache (aka mining truck) or from a mining facility (aka cargo truck)
			'If lCacheIndex <> -1 AndAlso glMineralCacheIdx(lCacheIndex) <> -1 Then
			'	lSourceID = glMineralCacheIdx(lCacheIndex)
			'	iSourceTypeID = ObjectType.eMineral
			'ElseIf lMiningFacIndex <> -1 AndAlso glFacilityIdx(lMiningFacIndex) <> -1 Then
			'	lSourceID = glFacilityIdx(lMiningFacIndex)
			'	iSourceTypeID = ObjectType.eFacility
			'End If

			''Ok, get our refinery
			'If lRefineryIndex <> -1 AndAlso glFacilityIdx(lRefineryIndex) <> -1 Then
			'	lDropoffID = glFacilityIdx(lRefineryIndex)
			'	iDropoffTypeID = ObjectType.eFacility
            'End If

            Dim lCol As Int32 = 0
            Dim lEnl As Int32 = 0
            Dim lOff As Int32 = 0

            Try
                For X As Int32 = 0 To Me.lCargoUB
                    If Me.lCargoIdx(X) > -1 AndAlso Me.oCargoContents(X) Is Nothing = False Then
                        Select Case Me.oCargoContents(X).ObjTypeID
                            Case ObjectType.eColonists
                                lCol += oCargoContents(X).ObjectID
                            Case ObjectType.eEnlisted
                                lEnl += oCargoContents(X).ObjectID
                            Case ObjectType.eOfficers
                                lOff += oCargoContents(X).ObjectID
                        End Select
                    End If
                Next X
            Catch
            End Try

			If Me.bProducing = True AndAlso Me.CurrentProduction Is Nothing = False Then
				DestX = LocX
				DestZ = LocZ
				DestX = LocX : DestZ = LocZ : DestEnvirID = CType(ParentObject, Epica_GUID).ObjectID
				DestEnvirTypeID = CType(ParentObject, Epica_GUID).ObjTypeID
				'lSourceID = -1 : iSourceTypeID = -1
				'lDropoffID = -1
				'iDropoffTypeID = -1
			End If

            Dim bIncludeParentObj As Boolean = (ParentObject Is Nothing = False)

            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblUnit (UnitGroupID, UnitName, OwnerID, "
                If bIncludeParentObj = True Then sSQL &= "ParentID, ParentTypeID, "
                sSQL &= "LocX, LocY, UnitDefID, Q1_HP, Q2_HP, Q3_HP, Q4_HP, CurrentSpeed, Structure_HP, Hangar_Cap, " & _
                  "Cargo_Cap, Fuel_Cap, Shield_HP, ExpLevel, CurrentStatus, LocAngle, ProductionTypeID, " & _
                  "CombatTactics, TargetingTactics, DestX, DestZ, DestEnvirID, DestEnvirTypeID, CurrentRouteNum, Enlisted, Officers, Colonists, TestFleetID, TetherPointX, TetherPointZ, RallyPointX, RallyPointZ, RallyPointA, RallyPointEnvirID, RallyPointEnvirTypeID) VALUES ("
				', SourceID, SourceTypeID, " & DropoffID, DropoffTypeID) VALUES ("
                'If oUnitGroup Is Nothing = False Then
                '    sSQL = sSQL & oUnitGroup.ObjectID
                'Else
                sSQL = sSQL & lFleetID
                'End If
                sSQL = sSQL & ", '" & MakeDBStr(BytesToString(EntityName)) & "', "
                If Owner Is Nothing = False Then
                    sSQL = sSQL & Owner.ObjectID & ","
                Else
                    sSQL = sSQL & "-1,"
                End If
                If bIncludeParentObj = True Then sSQL = sSQL & CType(ParentObject, Epica_GUID).ObjectID & ", " & CType(ParentObject, Epica_GUID).ObjTypeID & ", "
                sSQL &= LocX & ", " & LocZ & ", " & EntityDef.ObjectID & ", " & Q1_HP & ", " & Q2_HP & ", " & _
                  Q3_HP & ", " & Q4_HP & ", " & CurrentSpeed & ", " & Structure_HP & ", " & Hangar_Cap & ", " & _
                   Cargo_Cap & ", " & Fuel_Cap & ", " & Shield_HP & ", " & ExpLevel & ", " & _
                  CurrentStatus & ", " & LocAngle & ", " & yProductionType & ", " & iCombatTactics & ", " & _
                  iTargetingTactics & ", " & DestX & ", " & DestZ & ", " & DestEnvirID & ", " & DestEnvirTypeID & ", "
				If Me.lCurrentRouteIdx > -1 AndAlso Me.lCurrentRouteIdx <= Me.lRouteUB Then
					sSQL &= (Me.lCurrentRouteIdx)
				Else : sSQL &= "-1"
				End If
                sSQL &= ", " & lEnl.ToString & ", " & lOff.ToString & ", " & lCol.ToString & ", " & lTestFleetID
                sSQL &= ", " & TetherPointX.ToString & ", " & TetherPointZ.ToString
                sSQL &= ", " & RallyPointX.ToString & ", " & RallyPointZ.ToString & ", " & RallyPointA.ToString & ", " & RallyPointEnvirID.ToString & ", " & RallyPointEnvirTypeID.ToString
                sSQL &= ")"
				'& ", " & lSourceID & ", " & iSourceTypeID & ", " & lDropoffID & ", " & iDropoffTypeID & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblUnit SET UnitGroupID = "
                'If oUnitGroup Is Nothing = False Then
                '    sSQL = sSQL & oUnitGroup.ObjectID
                'Else 
                sSQL = sSQL & lFleetID
                'End If
                sSQL = sSQL & ", UnitName='" & MakeDBStr(BytesToString(EntityName)) & "', OwnerID = "
                If Owner Is Nothing = False Then
                    sSQL = sSQL & Owner.ObjectID
                Else : sSQL = sSQL & "-1"
                End If
                If bIncludeParentObj = True Then sSQL = sSQL & ", ParentID = " & CType(ParentObject, Epica_GUID).ObjectID & ", ParentTypeID = " & CType(ParentObject, Epica_GUID).ObjTypeID
                sSQL &= ", LocX = " & LocX & ", LocY = " & LocZ & ", UnitDefID = " & _
                  EntityDef.ObjectID & ", Q1_HP = " & Q1_HP & ", Q2_HP = " & Q2_HP & ", Q3_HP = " & Q3_HP & _
                  ", Q4_HP = " & Q4_HP & ", CurrentSpeed = " & CurrentSpeed & ", Structure_HP = " & Structure_HP & _
                  ", Hangar_Cap = " & Hangar_Cap & ", Cargo_Cap = " & Cargo_Cap & _
                  ", Fuel_Cap = " & Fuel_Cap & ", Shield_HP = " & Shield_HP & ", ExpLevel = " & ExpLevel & _
                  ", CurrentStatus = " & CurrentStatus & ", LocAngle = " & LocAngle & ", ProductionTypeId = " & _
                  yProductionType & ", CombatTactics = " & iCombatTactics & ", TargetingTactics = " & _
                  iTargetingTactics & ", DestX =  " & DestX & ", DestZ = " & DestZ & ", DestEnvirID = " & DestEnvirID & _
                  ", DestEnvirTypeID = " & DestEnvirTypeID & ", CurrentRouteNum = "
                If Me.lCurrentRouteIdx > -1 AndAlso Me.lCurrentRouteIdx <= Me.lRouteUB Then
                    sSQL &= (Me.lCurrentRouteIdx)
                Else : sSQL &= "-1"
                End If
                sSQL &= ", Enlisted = " & lEnl.ToString & ", Officers = " & lOff.ToString & ", Colonists = " & lCol.ToString & ", TestFleetID = " & lTestFleetID
                sSQL &= ", TetherPointX=" & TetherPointX.ToString & ", TetherPointZ=" & TetherPointZ.ToString
                sSQL &= ", RallyPointX=" & RallyPointX.ToString & ", RallyPointZ=" & RallyPointZ.ToString & ", RallyPointA=" & RallyPointA.ToString & ", RallyPointEnvirID=" & RallyPointEnvirID.ToString & ", RallyPointEnvirTypeID=" & RallyPointEnvirTypeID.ToString
                sSQL &= "  WHERE UnitID = " & ObjectID
                '", SourceID = " & lSourceID & ", SourceTypeID = " & iSourceTypeID & ", DropoffID = " & lDropoffID & ", DropoffTypeID = " & iDropoffTypeID & _

            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(UnitID) FROM tblUnit WHERE UnitName = '" & MakeDBStr(BytesToString(EntityName)) & "' AND OwnerID = "
                If Owner Is Nothing = False Then
                    sSQL = sSQL & Owner.ObjectID
                Else : sSQL = sSQL & "-1"
                End If
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            sSQL = "DELETE FROM tblUnitProduction WHERE UnitID = " & ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm = Nothing
            If Me.bProducing = True AndAlso Me.CurrentProduction Is Nothing = False Then
                RecalcProduction()
                'Now, save it
                sSQL = "INSERT INTO tblUnitProduction (UnitID, PointsProduced, ObjectID, ObjTypeID, ProdAngle) VALUES (" & Me.ObjectID & _
                  ", " & Me.CurrentProduction.PointsProduced & ", " & Me.CurrentProduction.ProductionID & ", " & _
                  Me.CurrentProduction.ProductionTypeID & ", " & Me.CurrentProduction.iProdA & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving Unit's Production object!")
                End If
            End If

            sSQL = "DELETE FROM tblRouteItem WHERE UnitID = " & Me.ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm = Nothing
            For X As Int32 = 0 To lRouteUB
                With uRoute(X)
                    sSQL = "INSERT INTO tblRouteItem (UnitID, LocX, LocZ, OrderNum, DestID, DestTypeID, LoadItemID, LoadItemTypeID, ExtraFlags) VALUES (" & _
                         Me.ObjectID & ", " & .lLocX & ", " & .lLocZ & ", " & (X + 1) & ", "
                    If .oDest Is Nothing = False Then
                        sSQL &= .oDest.ObjectID & ", " & .oDest.ObjTypeID & ", "
                    Else : sSQL &= "-1, -1, "
                    End If
                    If .oLoadItem Is Nothing = False Then
                        sSQL &= .oLoadItem.ObjectID & ", " & .oLoadItem.ObjTypeID
                    Else : sSQL &= "-1, -1"
                    End If
                    sSQL &= ", " & .yExtraFlags & ")"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        LogEvent(LogEventType.CriticalError, "Unit.SaveRoute: No records affected! " & sSQL)
                    End If
                    oComm = Nothing
                End With
            Next X

            MyBase.SaveAgentEffects()

            bResult = True
            bNeedsSaved = False
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    'NOTE: This assumes the following SQL statements were already ran:
    '  DELETE FROM tblAgentEffects WHERE EffectedItemTypeID = " & objecttype.eUnit
    '  DELETE FROM tblUnitProduction
    '  DELETE FROM tblRouteItem
	Public Function GetSaveObjectText() As String
		Dim sSQL As String

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
		If ObjectID = -1 Then
			SaveObject()
			Return ""
		End If

		Try
			If DestEnvirID < 1 OrElse DestEnvirTypeID < 0 Then
				DestX = LocX : DestZ = LocZ : DestEnvirID = CType(ParentObject, Epica_GUID).ObjectID
				DestEnvirTypeID = CType(ParentObject, Epica_GUID).ObjTypeID
			End If

			If Me.bProducing = True AndAlso Me.CurrentProduction Is Nothing = False Then
				DestX = LocX
				DestZ = LocZ
				DestX = LocX : DestZ = LocZ : DestEnvirID = CType(ParentObject, Epica_GUID).ObjectID
				DestEnvirTypeID = CType(ParentObject, Epica_GUID).ObjTypeID
			End If

            Dim lCol As Int32 = 0
            Dim lEnl As Int32 = 0
            Dim lOff As Int32 = 0

            Try
                For X As Int32 = 0 To Me.lCargoUB
                    If Me.lCargoIdx(X) > -1 AndAlso Me.oCargoContents(X) Is Nothing = False Then
                        Select Case Me.oCargoContents(X).ObjTypeID
                            Case ObjectType.eColonists
                                lCol += oCargoContents(X).ObjectID
                            Case ObjectType.eEnlisted
                                lEnl += oCargoContents(X).ObjectID
                            Case ObjectType.eOfficers
                                lOff += oCargoContents(X).ObjectID
                        End Select
                    End If
                Next X
            Catch
            End Try

			Dim oSB As New System.Text.StringBuilder

            Dim bIncludeParentObj As Boolean = (ParentObject Is Nothing = False)

			'UPDATE
			sSQL = "UPDATE tblUnit SET UnitGroupID = "
			'If oUnitGroup Is Nothing = False Then
			'    sSQL = sSQL & oUnitGroup.ObjectID
			'Else 
			sSQL = sSQL & lFleetID
			'End If
			sSQL = sSQL & ", UnitName='" & MakeDBStr(BytesToString(EntityName)) & "', OwnerID = "
			If Owner Is Nothing = False Then
				sSQL = sSQL & Owner.ObjectID
			Else : sSQL = sSQL & "-1"
            End If
            If bIncludeParentObj = True Then sSQL &= ", ParentID = " & CType(ParentObject, Epica_GUID).ObjectID & ", ParentTypeID = " & CType(ParentObject, Epica_GUID).ObjTypeID
            sSQL = sSQL & ", LocX = " & LocX & ", LocY = " & LocZ & ", UnitDefID = " & _
                 EntityDef.ObjectID & ", Q1_HP = " & Q1_HP & ", Q2_HP = " & Q2_HP & ", Q3_HP = " & Q3_HP & _
                 ", Q4_HP = " & Q4_HP & ", CurrentSpeed = " & CurrentSpeed & ", Structure_HP = " & Structure_HP & _
                 ", Hangar_Cap = " & Hangar_Cap & ", Cargo_Cap = " & Cargo_Cap & _
                 ", Fuel_Cap = " & Fuel_Cap & ", Shield_HP = " & Shield_HP & ", ExpLevel = " & ExpLevel & _
                 ", CurrentStatus = " & CurrentStatus & ", LocAngle = " & LocAngle & ", ProductionTypeId = " & _
                 yProductionType & ", CombatTactics = " & iCombatTactics & ", TargetingTactics = " & _
                 iTargetingTactics & ", DestX =  " & DestX & ", DestZ = " & DestZ & ", DestEnvirID = " & DestEnvirID & _
                 ", DestEnvirTypeID = " & DestEnvirTypeID & ", CurrentRouteNum = "
			If Me.lCurrentRouteIdx > -1 AndAlso Me.lCurrentRouteIdx <= Me.lRouteUB Then
				sSQL &= (Me.lCurrentRouteIdx)
			Else : sSQL &= "-1"
			End If
            sSQL &= ", Enlisted = " & lEnl.ToString & ", Officers = " & lOff.ToString & ", Colonists = " & lCol.ToString & ", TestFleetID = " & lTestFleetID & "  WHERE UnitID = " & ObjectID
			'", SourceID = " & lSourceID & ", SourceTypeID = " & iSourceTypeID & ", DropoffID = " & lDropoffID & ", DropoffTypeID = " & iDropoffTypeID & _

			oSB.AppendLine(sSQL)

            'sSQL = "DELETE FROM tblUnitProduction WHERE UnitID = " & ObjectID
            'oSB.AppendLine(sSQL)
			If Me.bProducing = True AndAlso Me.CurrentProduction Is Nothing = False Then
				RecalcProduction()
				'Now, save it
				sSQL = "INSERT INTO tblUnitProduction (UnitID, PointsProduced, ObjectID, ObjTypeID, ProdAngle) VALUES (" & Me.ObjectID & _
				  ", " & Me.CurrentProduction.PointsProduced & ", " & Me.CurrentProduction.ProductionID & ", " & _
				  Me.CurrentProduction.ProductionTypeID & ", " & Me.CurrentProduction.iProdA & ")"
				oSB.AppendLine(sSQL)
			End If

            'sSQL = "DELETE FROM tblRouteItem WHERE UnitID = " & Me.ObjectID
            'oSB.AppendLine(sSQL)
			For X As Int32 = 0 To lRouteUB
				With uRoute(X)
                    sSQL = "INSERT INTO tblRouteItem (UnitID, LocX, LocZ, OrderNum, DestID, DestTypeID, LoadItemID, LoadItemTypeID, ExtraFlags) VALUES (" & _
                                        Me.ObjectID & ", " & .lLocX & ", " & .lLocZ & ", " & (X + 1) & ", "
                    If .oDest Is Nothing = False Then
                        sSQL &= .oDest.ObjectID & ", " & .oDest.ObjTypeID & ", "
                    Else : sSQL &= "-1, -1, "
                    End If
                    If .oLoadItem Is Nothing = False Then
                        sSQL &= .oLoadItem.ObjectID & ", " & .oLoadItem.ObjTypeID
                    Else : sSQL &= "-1, -1"
                    End If
                    sSQL &= ", " & .yExtraFlags & ")"
					oSB.AppendLine(sSQL)
				End With
			Next X

			oSB.AppendLine(MyBase.GetSaveAgentEffectsText())
			Return oSB.ToString

		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		End Try
		Return ""
	End Function

	Public Function VerifyCompletion() As Boolean
		Dim blProduced As Int64

		If bProducing = True Then
			blProduced = CurrentProduction.PointsProduced
			blProduced += ((glCurrentCycle - CurrentProduction.lLastUpdateCycle) * mlProdPoints)
			Return blProduced >= CurrentProduction.PointsRequired
		Else : Return False
		End If
	End Function

	Public Sub RecalcProduction()
		If Me.Owner Is Nothing = False AndAlso Me.Owner.bInFullLockDown = True Then mlProdPoints = 0

		If bProducing = True Then
			'Ok, first, store off are points accumulated thus far
			CurrentProduction.PointsProduced += ((glCurrentCycle - CurrentProduction.lLastUpdateCycle) * mlProdPoints)
		End If
		mlProdPoints = 100		'TODO: This is hacked badly... where does a unit get its production points?

		If Me.Owner Is Nothing = False AndAlso Me.Owner.bInFullLockDown = True Then
			mlProdPoints = 0
			CurrentProduction.lFinishCycle = glCurrentCycle + CInt(CurrentProduction.PointsRequired - CurrentProduction.PointsProduced)
			Return
		End If


		'Now, how much production time are we going to need?
		If bProducing = True Then
			If mlProdPoints <> 0 Then
				CurrentProduction.lFinishCycle = glCurrentCycle + CInt((CurrentProduction.PointsRequired - CurrentProduction.PointsProduced) / mlProdPoints)
				CurrentProduction.lLastUpdateCycle = glCurrentCycle
			Else
				bProducing = False
			End If
		Else
			bProducing = False
		End If
	End Sub

    'Fairly easier than the HasProductionREquirements of Colony
    Public Function PopulateRequirements() As Boolean
        If CurrentProduction Is Nothing Then Return False

        If gb_IS_TEST_SERVER = True Then Return True

        Dim oProdCost As ProductionCost = Nothing

        'Now... check our production type
        If CurrentProduction.ProductionTypeID = ObjectType.eFacilityDef Then
            Dim oTmpDef As FacilityDef = GetEpicaFacilityDef(CurrentProduction.ProductionID)
            CurrentProduction.ProdCost = oTmpDef.ProductionRequirements
            CurrentProduction.PointsRequired = oTmpDef.ProductionRequirements.PointsRequired
            oProdCost = CurrentProduction.ProdCost
            oTmpDef = Nothing
        Else
            'TODO: Not sure about this... what else can a unit produce?
            Return False
        End If

        If oProdCost Is Nothing Then Return False

        If Owner.DeathBudgetEndTime > glCurrentCycle Then
            With oProdCost
                Dim blTotalCost As Int64 = .CreditCost + .ColonistCost + .EnlistedCost + .OfficerCost
                For X As Int32 = 0 To .ItemCostUB
                    blTotalCost += Colony.GetProdCostItemDBCost(.ItemCosts(X), Me.Owner) '.ItemCosts(X).QuantityNeeded
                Next X
                If Owner.DeathBudgetFundsRemaining >= blTotalCost AndAlso Owner.blCredits >= blTotalCost Then
                    Owner.DeathBudgetFundsRemaining -= CInt(blTotalCost)
                    Owner.blCredits -= CInt(blTotalCost)
                    CurrentProduction.PointsRequired = 1
                    Return True
                End If
            End With
        End If

        Dim bCargo As Boolean = False
        bCargo = (Me.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 AndAlso Me.lCargoUB > -1

        'Ok, find the parent colony...
        Dim lColonyIdx As Int32 = -1
        'If the unit is in space, there is no colony...
        With CType(Me.ParentObject, Epica_GUID)
            'Units can only get resources from colonies on a planet surface
			If .ObjTypeID = ObjectType.ePlanet Then
				lColonyIdx = Me.Owner.GetColonyFromParent(.ObjectID, .ObjTypeID)
			Else
				'Get closest space station colony to me
				lColonyIdx = Me.Owner.GetClosestSpaceColonyToLoc(.ObjectID, Me.LocX, Me.LocZ)
			End If
		End With
		Dim oColony As Colony = Nothing
		If lColonyIdx <> -1 Then
			oColony = goColony(lColonyIdx)
			If oColony Is Nothing Then lColonyIdx = -1
		End If

        'Ok, get the UnitGroupIdx
        Dim lUnitGroupIdx As Int32 = -1
        If Me.lFleetID > 0 Then
            For X As Int32 = 0 To glUnitGroupUB
                If glUnitGroupIdx(X) = lFleetID Then
                    lUnitGroupIdx = X
                    Exit For
                End If
            Next X
        End If

        'Now... check if we have the basic required resources
        If Me.Owner.blCredits < oProdCost.CreditCost Then
			If lColonyIdx <> -1 Then
				Owner.SendPlayerMessage(oColony.GetLowResourcesMsg(ProductionType.eCredits, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
			Else : Owner.SendPlayerMessage(GetEnvirLowResourcesMsg(ProductionType.eCredits, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
			End If
			Return False
		End If

		Dim lEnlAvail As Int32 = 0
		Dim lOffAvail As Int32 = 0
		Dim lColAvail As Int32 = 0

		If lColonyIdx <> -1 Then
			lEnlAvail += oColony.ColonyEnlisted
			lOffAvail += oColony.ColonyOfficers
			lColAvail += oColony.Population
		End If

		If bCargo = True Then
			For X As Int32 = 0 To Me.lCargoUB
				If Me.lCargoIdx(X) <> -1 Then
					Select Case Me.oCargoContents(X).ObjTypeID
						Case ObjectType.eColonists
							lColAvail += Me.oCargoContents(X).ObjectID + 1
						Case ObjectType.eEnlisted
							lEnlAvail += Me.oCargoContents(X).ObjectID + 1
						Case ObjectType.eOfficers
							lOffAvail += Me.oCargoContents(X).ObjectID + 1
					End Select
				End If
			Next X
		End If

		If lUnitGroupIdx <> -1 AndAlso glUnitGroupIdx(lUnitGroupIdx) <> -1 Then
			goUnitGroup(lUnitGroupIdx).PopulateAvailablePersonnel(lColAvail, lEnlAvail, lOffAvail, CType(ParentObject, Epica_GUID).ObjectID, CType(ParentObject, Epica_GUID).ObjTypeID, Me.ObjectID)
		End If

		If lEnlAvail < oProdCost.EnlistedCost Then
			If lColonyIdx <> -1 Then
				Owner.SendPlayerMessage(oColony.GetLowResourcesMsg(ProductionType.eEnlisted, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
			Else : Owner.SendPlayerMessage(GetEnvirLowResourcesMsg(ProductionType.eEnlisted, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
			End If
			Return False
		End If
		If lOffAvail < oProdCost.OfficerCost Then
			If lColonyIdx <> -1 Then
				Owner.SendPlayerMessage(oColony.GetLowResourcesMsg(ProductionType.eOfficers, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
			Else : Owner.SendPlayerMessage(GetEnvirLowResourcesMsg(ProductionType.eOfficers, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
			End If
			Return False
		End If
		If lColAvail < oProdCost.ColonistCost Then
			If lColonyIdx <> -1 Then
				Owner.SendPlayerMessage(oColony.GetLowResourcesMsg(ProductionType.eColonists, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
			Else : Owner.SendPlayerMessage(GetEnvirLowResourcesMsg(ProductionType.eColonists, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
			End If
			Return False
		End If

		'Ok, basics are out of the way, now check the component costs...
		Dim lItemID(oProdCost.ItemCostUB) As Int32
		Dim iTypeID(oProdCost.ItemCostUB) As Int16
		Dim lQty(oProdCost.ItemCostUB) As Int32

		'go ahead and give enough room for all of the prod cost items...
		Dim uResSrc(oProdCost.ItemCostUB) As ResourceSource
		Dim lResSrcUB As Int32 = -1
		Dim iCacheTypeID As Int16

		For X As Int32 = 0 To oProdCost.ItemCostUB
			lItemID(X) = oProdCost.ItemCosts(X).ItemID
			iTypeID(X) = oProdCost.ItemCosts(X).ItemTypeID
			lQty(X) = oProdCost.ItemCosts(X).QuantityNeeded
		Next X

		'Let's check my cargo first...
		If bCargo = True Then
			For lCIdx As Int32 = 0 To lCargoUB
				If lCargoIdx(lCIdx) <> -1 AndAlso oCargoContents(lCIdx).ObjTypeID <> ObjectType.eAmmunition Then
					Dim lCargoID As Int32
					Dim iCargoTypeID As Int16
					Dim lCargoQty As Int32

					iCacheTypeID = oCargoContents(lCIdx).ObjTypeID

					If iCacheTypeID = ObjectType.eMineralCache Then
						With CType(oCargoContents(lCIdx), MineralCache)
							lCargoID = .oMineral.ObjectID
							iCargoTypeID = ObjectType.eMineral
							lCargoQty = .Quantity
						End With
					ElseIf iCacheTypeID = ObjectType.eComponentCache Then
						With CType(oCargoContents(lCIdx), ComponentCache)
							lCargoID = .ComponentID
							iCargoTypeID = .ComponentTypeID
							lCargoQty = .Quantity
						End With
					Else : Continue For
					End If

					'Now, loop through the item list.... for a match
					For Y As Int32 = 0 To oProdCost.ItemCostUB
						If lQty(Y) > 0 AndAlso lItemID(Y) = lCargoID AndAlso iTypeID(Y) = iCargoTypeID Then
							'If all of that is true... then this is a resource source... add it to our list...
							lResSrcUB += 1
							If lResSrcUB > uResSrc.GetUpperBound(0) Then
								ReDim Preserve uResSrc(lResSrcUB)
							End If

							With uResSrc(lResSrcUB)
								.iTypeID = iCargoTypeID
								.lCargoIdx = lCIdx
								.lFacilityIdx = -1
								.lItemID = lCargoID
								.iCacheTypeID = iCacheTypeID
							End With

							If lQty(Y) > lCargoQty Then
								uResSrc(lResSrcUB).lQuantity = lCargoQty
								lQty(Y) -= lCargoQty
							Else
								uResSrc(lResSrcUB).lQuantity = lQty(Y)
								lQty(Y) = 0
							End If
							Exit For
						End If
					Next Y
				End If
			Next lCIdx
		End If

		'Now, check if we need to check the colony
		Dim bGood As Boolean = True
		For X As Int32 = 0 To oProdCost.ItemCostUB
			If lQty(X) > 0 Then
				bGood = False
				Exit For
			End If
		Next X

		If bGood = False AndAlso lColonyIdx <> -1 Then
			'Now, loop through the colony's facilities to check their components much like I do for the colony's has required resources
			With oColony

				For X As Int32 = 0 To .mlMineralCacheUB
					If .mlMineralCacheIdx(X) > -1 AndAlso glMineralCacheIdx(.mlMineralCacheIdx(X)) = .mlMineralCacheID(X) Then
						For Y As Int32 = 0 To oProdCost.ItemCostUB
							If lQty(Y) > 0 AndAlso .mlMineralCacheMineralID(X) = lItemID(Y) AndAlso iTypeID(Y) = ObjectType.eMineral Then
								Dim oCache As MineralCache = goMineralCache(.mlMineralCacheIdx(X))
								If oCache Is Nothing Then Exit For

								'If all of that is true... then this is a resource source... add it to our list...
								lResSrcUB += 1
								If lResSrcUB > uResSrc.GetUpperBound(0) Then
									ReDim Preserve uResSrc(lResSrcUB)
								End If

								With uResSrc(lResSrcUB)
									.iTypeID = ObjectType.eMineral
									.lCargoIdx = X
									.lFacilityIdx = Int32.MaxValue
									.lItemID = oCache.oMineral.ObjectID
									.iCacheTypeID = ObjectType.eMineralCache
								End With

								If lQty(Y) > oCache.Quantity Then
									uResSrc(lResSrcUB).lQuantity = oCache.Quantity
									lQty(Y) -= oCache.Quantity
								Else
									uResSrc(lResSrcUB).lQuantity = lQty(Y)
									lQty(Y) = 0
								End If
								Exit For
							End If
						Next Y

						bGood = True
						For Y As Int32 = 0 To oProdCost.ItemCostUB
							If lQty(Y) <> 0 Then
								bGood = False
								Exit For
							End If
						Next Y
						If bGood = True Then Exit For
					End If
                Next X
                If bGood = False Then
                    For X As Int32 = 0 To .mlComponentCacheUB
                        If .mlComponentCacheIdx(X) > -1 AndAlso glComponentCacheIdx(.mlComponentCacheIdx(X)) = .mlComponentCacheID(X) Then
                            For Y As Int32 = 0 To oProdCost.ItemCostUB
                                If lQty(Y) > 0 AndAlso .mlComponentCacheCompID(X) = lItemID(Y) Then
                                    Dim oCache As ComponentCache = goComponentCache(.mlComponentCacheIdx(X))
                                    If oCache Is Nothing Then Exit For
                                    If oCache.ComponentTypeID <> iTypeID(Y) Then Continue For

                                    'If all of that is true... then this is a resource source... add it to our list...
                                    lResSrcUB += 1
                                    If lResSrcUB > uResSrc.GetUpperBound(0) Then
                                        ReDim Preserve uResSrc(lResSrcUB)
                                    End If

                                    With uResSrc(lResSrcUB)
                                        .iTypeID = oCache.ComponentTypeID
                                        .lCargoIdx = oCache.ComponentOwnerID
                                        .lFacilityIdx = Int32.MaxValue
                                        .lItemID = oCache.ComponentID
                                        .iCacheTypeID = ObjectType.eComponentCache
                                    End With

                                    If lQty(Y) > oCache.Quantity Then
                                        uResSrc(lResSrcUB).lQuantity = oCache.Quantity
                                        lQty(Y) -= oCache.Quantity
                                    Else
                                        uResSrc(lResSrcUB).lQuantity = lQty(Y)
                                        lQty(Y) = 0
                                    End If
                                    Exit For
                                End If
                            Next Y

                            bGood = True
                            For Y As Int32 = 0 To oProdCost.ItemCostUB
                                If lQty(Y) <> 0 Then
                                    bGood = False
                                    Exit For
                                End If
                            Next Y
                            If bGood = True Then Exit For
                        End If
                    Next X
                End If

				If bGood = False Then
					For X As Int32 = 0 To .ChildrenUB
                        If .lChildrenIdx(X) <> -1 AndAlso .oChildren(X).Active = True AndAlso (.oChildren(X).yProductionType <> ProductionType.eMining) Then 'AndAlso (.oChildren(X).yProductionType <> ProductionType.eTradePost) Then
                            For lCIdx As Int32 = 0 To .oChildren(X).lCargoUB
                                If .oChildren(X).lCargoIdx(lCIdx) <> -1 AndAlso .oChildren(X).oCargoContents(lCIdx).ObjTypeID <> ObjectType.eAmmunition Then
                                    Dim lCargoID As Int32
                                    Dim iCargoTypeID As Int16
                                    Dim lCargoQty As Int32

                                    iCacheTypeID = .oChildren(X).oCargoContents(lCIdx).ObjTypeID

                                    If iCacheTypeID = ObjectType.eMineralCache Then
                                        With CType(.oChildren(X).oCargoContents(lCIdx), MineralCache)
                                            lCargoID = .oMineral.ObjectID
                                            iCargoTypeID = ObjectType.eMineral
                                            lCargoQty = .Quantity
                                        End With
                                    ElseIf iCacheTypeID = ObjectType.eComponentCache Then
                                        With CType(.oChildren(X).oCargoContents(lCIdx), ComponentCache)
                                            lCargoID = .ComponentID
                                            iCargoTypeID = .ComponentTypeID
                                            lCargoQty = .Quantity
                                        End With
                                    Else : Continue For
                                    End If


                                    For Y As Int32 = 0 To oProdCost.ItemCostUB
                                        If lQty(Y) > 0 AndAlso lItemID(Y) = lCargoID AndAlso iTypeID(Y) = iCargoTypeID Then
                                            'If all of that is true... then this is a resource source... add it to our list...
                                            lResSrcUB += 1
                                            If lResSrcUB > uResSrc.GetUpperBound(0) Then
                                                ReDim Preserve uResSrc(lResSrcUB)
                                            End If

                                            With uResSrc(lResSrcUB)
                                                .iTypeID = iCargoTypeID
                                                .lCargoIdx = lCIdx
                                                .lFacilityIdx = X
                                                .lItemID = lCargoID
                                                .iCacheTypeID = iCacheTypeID
                                            End With

                                            If lQty(Y) > lCargoQty Then
                                                uResSrc(lResSrcUB).lQuantity = lCargoQty
                                                lQty(Y) -= lCargoQty
                                            Else
                                                uResSrc(lResSrcUB).lQuantity = lQty(Y)
                                                lQty(Y) = 0
                                            End If

                                            Exit For
                                        End If
                                    Next Y
                                End If
                            Next lCIdx

                            bGood = True
                            For Y As Int32 = 0 To oProdCost.ItemCostUB
                                If lQty(Y) <> 0 Then
                                    bGood = False
                                    Exit For
                                End If
                            Next Y

                            If bGood = True Then Exit For
                        End If
					Next X
				End If
			End With
		End If

		'Pull from Unit Group Next
		If bGood = False AndAlso lUnitGroupIdx <> -1 AndAlso glUnitGroupIdx(lUnitGroupIdx) <> -1 Then

			Dim lCmpParentID As Int32
			Dim iCmpParentTypeID As Int16
			With CType(ParentObject, Epica_GUID)
				lCmpParentID = .ObjectID
				iCmpParentTypeID = .ObjTypeID
			End With


			For X As Int32 = 0 To goUnitGroup(lUnitGroupIdx).UnitUB
				Dim lIdx As Int32 = goUnitGroup(lUnitGroupIdx).GetUnitIdx(X)
				If lIdx <> -1 Then
					Dim oUnit As Unit = goUnit(lIdx)
					If oUnit Is Nothing = False AndAlso (oUnit.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
						With oUnit

							If .ParentObject Is Nothing Then Continue For
							If CType(.ParentObject, Epica_GUID).ObjectID <> lCmpParentID OrElse CType(.ParentObject, Epica_GUID).ObjTypeID <> iCmpParentTypeID Then Continue For

							For lCIdx As Int32 = 0 To .lCargoUB
								If .lCargoIdx(lCIdx) <> -1 AndAlso .oCargoContents(lCIdx).ObjTypeID <> ObjectType.eAmmunition Then
									Dim lCargoID As Int32
									Dim iCargoTypeID As Int16
									Dim lCargoQty As Int32

									iCacheTypeID = .oCargoContents(lCIdx).ObjTypeID

									If iCacheTypeID = ObjectType.eMineralCache Then
										With CType(.oCargoContents(lCIdx), MineralCache)
											lCargoID = .oMineral.ObjectID
											iCargoTypeID = ObjectType.eMineral
											lCargoQty = .Quantity
										End With
									ElseIf iCacheTypeID = ObjectType.eComponentCache Then
										With CType(.oCargoContents(lCIdx), ComponentCache)
											lCargoID = .ComponentID
											iCargoTypeID = .ComponentTypeID
											lCargoQty = .Quantity
										End With
									Else : Continue For
									End If


									For Y As Int32 = 0 To oProdCost.ItemCostUB
										If lQty(Y) > 0 AndAlso lItemID(Y) = lCargoID AndAlso iTypeID(Y) = iCargoTypeID Then
											'If all of that is true... then this is a resource source... add it to our list...
											lResSrcUB += 1
											If lResSrcUB > uResSrc.GetUpperBound(0) Then
												ReDim Preserve uResSrc(lResSrcUB)
											End If

											With uResSrc(lResSrcUB)
												.iTypeID = iCargoTypeID
												.lCargoIdx = lCIdx
												.lFacilityIdx = Int32.MinValue
												.lUnitIdx = lIdx
												.lItemID = lCargoID
												.iCacheTypeID = iCacheTypeID
											End With

											If lQty(Y) > lCargoQty Then
												uResSrc(lResSrcUB).lQuantity = lCargoQty
												lQty(Y) -= lCargoQty
											Else
												uResSrc(lResSrcUB).lQuantity = lQty(Y)
												lQty(Y) = 0
											End If

											Exit For
										End If
									Next Y
								End If
							Next lCIdx

							bGood = True
							For Y As Int32 = 0 To oProdCost.ItemCostUB
								If lQty(Y) <> 0 Then
									bGood = False
									Exit For
								End If
							Next Y

							If bGood = True Then Exit For
						End With
					End If
				End If
			Next X
		End If

		If bGood = False Then
			If lColonyIdx <> -1 Then
				For Y As Int32 = 0 To oProdCost.ItemCostUB
					If lQty(Y) <> 0 Then
						Owner.SendPlayerMessage(oColony.GetLowResourcesMsg(ProductionType.eMaterials, lItemID(Y), iTypeID(Y), oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
						Exit For
					End If
				Next Y
			Else
				For Y As Int32 = 0 To oProdCost.ItemCostUB
					If lQty(Y) <> 0 Then
						Owner.SendPlayerMessage(GetEnvirLowResourcesMsg(ProductionType.eMaterials, lItemID(Y), iTypeID(Y), oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
						Exit For
					End If
				Next Y
			End If
			Return False
		Else
			'Deduct our basics first...
			Owner.blCredits -= oProdCost.CreditCost

			lEnlAvail = oProdCost.EnlistedCost
			lOffAvail = oProdCost.OfficerCost
			lColAvail = oProdCost.ColonistCost
			If bCargo = True Then
				'Deduct from the unit's cargo first... reducing our EnlAvail, OffAvail, ColAvail as needed
				For X As Int32 = 0 To Me.lCargoUB
					If Me.lCargoIdx(X) <> -1 Then
						Select Case Me.oCargoContents(X).ObjTypeID
							Case ObjectType.eColonists
								Dim lTemp As Int32 = Math.Min(Me.oCargoContents(X).ObjectID, lColAvail)
								Me.oCargoContents(X).ObjectID -= lTemp
								Me.lCargoIdx(X) = Me.oCargoContents(X).ObjectID
								lColAvail -= lTemp
							Case ObjectType.eEnlisted
								Dim lTemp As Int32 = Math.Min(Me.oCargoContents(X).ObjectID, lEnlAvail)
								Me.oCargoContents(X).ObjectID -= lTemp
								Me.lCargoIdx(X) = Me.oCargoContents(X).ObjectID
								lEnlAvail -= lTemp
							Case ObjectType.eOfficers
								Dim lTemp As Int32 = Math.Min(Me.oCargoContents(X).ObjectID, lOffAvail)
								Me.oCargoContents(X).ObjectID -= lTemp
								Me.lCargoIdx(X) = Me.oCargoContents(X).ObjectID
								lOffAvail -= lTemp
						End Select
					End If
				Next X
			End If

			'Now, try the colony for anything else
			If lColonyIdx <> -1 AndAlso (lColAvail <> 0 OrElse lEnlAvail <> 0 OrElse lOffAvail <> 0) Then
				With oColony
					Dim lTempVal As Int32 = Math.Min(lOffAvail, .ColonyOfficers)
					.ColonyOfficers -= lTempVal
					lOffAvail -= lTempVal

					lTempVal = Math.Min(lEnlAvail, .ColonyEnlisted)
					.ColonyEnlisted -= lEnlAvail
					lEnlAvail -= lTempVal

					lTempVal = Math.Min(lColAvail, .Population)
					.Population -= lColAvail
					lColAvail -= lTempVal
				End With
			End If

			'Now, try the unit group
			If lUnitGroupIdx <> -1 AndAlso glUnitGroupIdx(lUnitGroupIdx) <> -1 AndAlso (lColAvail <> 0 OrElse lEnlAvail <> 0 OrElse lOffAvail <> 0) Then
				With CType(Me.ParentObject, Epica_GUID)
					goUnitGroup(lUnitGroupIdx).ReduceUnitGroupsPersonnel(lColAvail, lEnlAvail, lOffAvail, .ObjectID, .ObjTypeID, Me.ObjectID)
				End With
			End If

			'TODO: What sort of race conditions could occur here? It would be potentially unsafe...
			'Now, anything leftover, go back through our uResSrc array and pull out the cargo
			For X As Int32 = 0 To lResSrcUB
				'ok... deduct it...
				If uResSrc(X).iCacheTypeID = ObjectType.eMineralCache Then
					If uResSrc(X).lFacilityIdx = Int32.MinValue Then
						Try
							With CType(goUnit(uResSrc(X).lUnitIdx).oCargoContents(uResSrc(X).lCargoIdx), MineralCache)
								.Quantity -= uResSrc(X).lQuantity
							End With
							'goUnit(uResSrc(X).lUnitIdx).Cargo_Cap += uResSrc(X).lQuantity
						Catch
						End Try
					ElseIf uResSrc(X).lFacilityIdx = -1 Then
						With CType(Me.oCargoContents(uResSrc(X).lCargoIdx), MineralCache)
							.Quantity -= uResSrc(X).lQuantity
						End With
					ElseIf uResSrc(X).lFacilityIdx = Int32.MaxValue Then
						oColony.AdjustColonyMineralCache(uResSrc(X).lItemID, -Math.Abs(uResSrc(X).lQuantity))
					Else
						With CType(oColony.oChildren(uResSrc(X).lFacilityIdx).oCargoContents(uResSrc(X).lCargoIdx), MineralCache)
							.Quantity -= uResSrc(X).lQuantity
						End With
						'goColony(lColonyIdx).oChildren(uResSrc(X).lFacilityIdx).Cargo_Cap += uResSrc(X).lQuantity
					End If
				ElseIf uResSrc(X).iCacheTypeID = ObjectType.eComponentCache Then
					If uResSrc(X).lFacilityIdx = Int32.MinValue Then
						Try
							With CType(goUnit(uResSrc(X).lUnitIdx).oCargoContents(uResSrc(X).lCargoIdx), ComponentCache)
								.Quantity -= uResSrc(X).lQuantity
							End With
							'goUnit(uResSrc(X).lUnitIdx).Cargo_Cap += uResSrc(X).lQuantity
						Catch
						End Try
					ElseIf uResSrc(X).lFacilityIdx = -1 Then
						With CType(Me.oCargoContents(uResSrc(X).lCargoIdx), ComponentCache)
							.Quantity -= uResSrc(X).lQuantity
                        End With
                    ElseIf uResSrc(X).lFacilityIdx = Int32.MaxValue Then
                        oColony.AdjustColonyComponentCache(uResSrc(X).lItemID, uResSrc(X).iTypeID, uResSrc(X).lCargoIdx, -Math.Abs(uResSrc(X).lQuantity))
                    Else
                        With CType(oColony.oChildren(uResSrc(X).lFacilityIdx).oCargoContents(uResSrc(X).lCargoIdx), ComponentCache)
                            .Quantity -= uResSrc(X).lQuantity
                        End With
                        'goColony(lColonyIdx).oChildren(uResSrc(X).lFacilityIdx).Cargo_Cap += uResSrc(X).lQuantity
					End If
				Else
					'TODO: What else could it be? Ammunition?
				End If
			Next X
		End If

		Return bGood

    End Function

	'   Public Sub HandleReturnToRefinery()
	'       Dim X As Int32
	'       Dim yData() As Byte
	'       Dim iTemp As Int16

	'       'Reset our cache if that cache no longer exists
	'       If lCacheIndex <> -1 AndAlso glMineralCacheIdx(lCacheIndex) <> -1 Then
	'           'Ok, tell the mineral cache that we are not mining it anymore so others may mine it
	'		If goMineralCache(lCacheIndex) Is Nothing = False Then goMineralCache(lCacheIndex).BeingMinedBy = Nothing
	'           'lCacheIndex = -1
	'       End If

	'       'set status to not mining or building...
	'	If (CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then CurrentStatus = CurrentStatus Xor elUnitStatus.eMoveLock ' CurrentStatus -= elUnitStatus.eMoveLock
	'       bMining = False

	'	'Send our status update to the region server
	'       ReDim yData(11)
	'	System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yData, 0)
	'	GetGUIDAsString.CopyTo(yData, 2)
	'	System.BitConverter.GetBytes(elUnitStatus.eMoveLock * -1).CopyTo(yData, 8)

	'	iTemp = CType(ParentObject, Epica_GUID).ObjTypeID
	'	If iTemp = ObjectType.ePlanet Then
	'		CType(ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
	'	ElseIf iTemp = ObjectType.eSolarSystem Then
	'		CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
	'	End If

	'	'Now, check the Refinery index
	'	Dim oParentGUID As Epica_GUID = CType(Me.ParentObject, Epica_GUID)

	'	If lRefineryIndex = -1 OrElse glFacilityIdx(lRefineryIndex) = -1 OrElse _
	'	  ((Me.EntityDef.yChassisType And ChassisType.eSpaceBased) = 0 AndAlso (oParentGUID.ObjectID <> CType(goFacility(lRefineryIndex).ParentObject, Epica_GUID).ObjectID OrElse _
	'	  oParentGUID.ObjTypeID <> CType(goFacility(lRefineryIndex).ParentObject, Epica_GUID).ObjTypeID)) Then
	'		Dim fDist As Single
	'		Dim fTemp As Single

	'		lRefineryIndex = -1

	'		Dim lExtent As Int32 = 0
	'		If iTemp = ObjectType.ePlanet Then lExtent = CType(ParentObject, Planet).PlanetWidthTotal

	'		'Ok, find the nearest SAME ENVIRONMENT Refinery that belongs to this player...
	'		Dim lColonyIdx As Int32 = -1
	'		With CType(Me.ParentObject, Epica_GUID)
	'			lColonyIdx = Owner.GetColonyFromParent(.ObjectID, .ObjTypeID)
	'		End With

	'		If lColonyIdx <> -1 AndAlso glColonyIdx(lColonyIdx) <> -1 Then
	'			'Ok, we got a colony here... check its children
	'			Dim oColony As Colony = goColony(lColonyIdx)
	'			If oColony Is Nothing = False Then
	'				With oColony
	'					For X = 0 To .ChildrenUB
	'						If .lChildrenIdx(X) <> -1 AndAlso .oChildren(X).yProductionType = ProductionType.eRefining Then
	'							If .oChildren(X).Active = True AndAlso .oChildren(X).EntityDef.HasHangarDoorSize(Me.EntityDef.HullSize) = True Then
	'								Dim lTmpLocX As Int32 = LocX	'1
	'								If iTemp = ObjectType.ePlanet Then
	'									Dim lTmpX As Int32 = lTmpLocX - .oChildren(X).LocX	'-3
	'									Dim lTmpvar As Int32 = lTmpLocX - (.oChildren(X).LocX - lExtent)	'-1
	'									If Math.Abs(lTmpX) > Math.Abs(lTmpvar) Then
	'										If LocX < .oChildren(X).LocX Then
	'											lTmpLocX += lExtent
	'										Else : lTmpLocX -= lExtent
	'										End If
	'									End If
	'								End If

	'								'Yes, ok... do it to it
	'								fTemp = Distance(.oChildren(X).LocX, .oChildren(X).LocZ, lTmpLocX, LocZ)
	'								If lRefineryIndex = -1 OrElse fTemp < fDist Then
	'									fDist = fTemp
	'									lRefineryIndex = X
	'								End If
	'							End If
	'						End If
	'					Next X


	'					If lRefineryIndex <> -1 Then
	'						Dim lTmpID As Int32 = .oChildren(lRefineryIndex).ObjectID
	'						For X = 0 To glFacilityUB
	'							If glFacilityIdx(X) = lTmpID Then
	'								lRefineryIndex = X
	'								Exit For
	'							End If
	'						Next X
	'					End If
	'				End With
	'			End If
	'		End If
	'	End If

	'	'Ok, now, if there is a refinery to go to...
	'	If lRefineryIndex <> -1 Then
	'		Dim oFac As Facility = goFacility(lRefineryIndex)
	'		If oFac Is Nothing = False Then
	'			'Then, create a Dock Request message and send it to the Region Server
	'			ReDim yData(13)
	'			System.BitConverter.GetBytes(GlobalMessageCode.eRequestDock).CopyTo(yData, 0)
	'			'First, the target GUID
	'			oFac.GetGUIDAsString.CopyTo(yData, 2)
	'			'Now, my GUID
	'			GetGUIDAsString.CopyTo(yData, 8)

	'			iTemp = CType(ParentObject, Epica_GUID).ObjTypeID
	'			If iTemp = ObjectType.ePlanet Then
	'				CType(ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
	'			ElseIf iTemp = ObjectType.eSolarSystem Then
	'				CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
	'			End If
	'		Else : lRefineryIndex = -1
	'		End If
	'	End If

	'End Sub
	'Public Sub HandleReturnToMiningSite()
	'	Dim yData() As Byte
	'	Dim iTemp As Int16

	'	If lCacheIndex <> -1 AndAlso glMineralCacheIdx(lCacheIndex) <> -1 Then
	'		Dim oCache As MineralCache = goMineralCache(lCacheIndex)
	'		If oCache Is Nothing Then
	'			lMiningFacIndex = -1
	'			lCacheIndex = -1
	'			Return
	'		End If

	'		If (Me.EntityDef.yChassisType And ChassisType.eSpaceBased) = 0 Then
	'			Dim lCPID As Int32 = -1
	'			Dim iCPTID As Int16 = -1

	'			With CType(oCache.ParentObject, Epica_GUID)
	'				lCPID = .ObjectID
	'				iCPTID = .ObjTypeID
	'			End With

	'			With CType(Me.ParentObject, Epica_GUID)
	'				If .ObjectID <> lCPID AndAlso .ObjTypeID <> iCPTID Then
	'					lMiningFacIndex = -1
	'					lCacheIndex = -1
	'					Return
	'				End If
	'			End With
	'		End If

	'		ReDim yData(11)
	'		System.BitConverter.GetBytes(GlobalMessageCode.eSetMiningLoc).CopyTo(yData, 0)
	'		System.BitConverter.GetBytes(oCache.ObjectID).CopyTo(yData, 2)

	'		GetGUIDAsString.CopyTo(yData, 6)

	'		'TODO: it is assumed here that the ParentObject is associated to a domain
	'		iTemp = CType(ParentObject, Epica_GUID).ObjTypeID
	'		If iTemp = ObjectType.ePlanet Then
	'			CType(ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
	'		ElseIf iTemp = ObjectType.eSolarSystem Then
	'			CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
	'		End If
	'	ElseIf lMiningFacIndex <> -1 AndAlso glFacilityIdx(lMiningFacIndex) <> -1 Then
	'		Dim oFac As Facility = goFacility(lMiningFacIndex)
	'		If oFac Is Nothing = False Then
	'			If (Me.EntityDef.yChassisType And ChassisType.eSpaceBased) = 0 Then
	'				Dim lCPID As Int32 = -1
	'				Dim iCPTID As Int16 = -1
	'				With CType(oFac.ParentObject, Epica_GUID)
	'					lCPID = .ObjectID
	'					iCPTID = .ObjTypeID
	'				End With
	'				With CType(Me.ParentObject, Epica_GUID)
	'					If .ObjectID <> lCPID AndAlso .ObjTypeID <> iCPTID Then
	'						lMiningFacIndex = -1
	'						lCacheIndex = -1
	'						Return
	'					End If
	'				End With
	'			End If

	'			'Ok... return to the mining site... do that by requesting a dock...
	'			ReDim yData(13)
	'			System.BitConverter.GetBytes(GlobalMessageCode.eRequestDock).CopyTo(yData, 0)
	'			'first the target GUID
	'			oFac.GetGUIDAsString.CopyTo(yData, 2)

	'			'Now, my GUID
	'			GetGUIDAsString.CopyTo(yData, 8)

	'			'TODO: it is assumed here that the parentobject is associated to a domain
	'			iTemp = CType(ParentObject, Epica_GUID).ObjTypeID
	'			If iTemp = ObjectType.ePlanet Then
	'				CType(ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
	'			ElseIf iTemp = ObjectType.eSolarSystem Then
	'				CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
	'			End If
	'		Else
	'			lMiningFacIndex = -1
	'			lCacheIndex = -1
	'		End If
	'	Else
	'		lMiningFacIndex = -1
	'		lCacheIndex = -1
	'	End If
	'End Sub

	Public Function GetEnvirLowResourcesMsg(ByVal lLowVal As ProductionType, ByVal lLowItemID As Int32, ByVal iLowItemTypeID As Int16, ByVal lProdID As Int32, ByVal iProdTypeID As Int16) As Byte()
		Dim yMsg(46) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eColonyLowResources).CopyTo(yMsg, 0)

		GetEpicaObjectName(CType(Me.ParentObject, Epica_GUID).ObjTypeID, Me.ParentObject).CopyTo(yMsg, 2)
		Me.GetGUIDAsString.CopyTo(yMsg, 22)
		CType(Me.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, 28)
		yMsg(34) = lLowVal
		System.BitConverter.GetBytes(lLowItemID).CopyTo(yMsg, 35)
		System.BitConverter.GetBytes(iLowItemTypeID).CopyTo(yMsg, 39)
		System.BitConverter.GetBytes(lProdID).CopyTo(yMsg, 41)
		System.BitConverter.GetBytes(iProdTypeID).CopyTo(yMsg, 45)
		Return yMsg
	End Function

    Public Sub ClearDropoffAndSourceDetails()
        'Reset our cache if that cache no longer exists
		If lCacheIndex <> -1 AndAlso glMineralCacheIdx(lCacheIndex) = lCacheID Then
			'Ok, tell the mineral cache that we are not mining it anymore so others may mine it
			If goMineralCache(lCacheIndex) Is Nothing = False Then
				If goMineralCache(lCacheIndex).BeingMinedBy Is Nothing = False Then
					With CType(goMineralCache(lCacheIndex).BeingMinedBy, Epica_GUID)
						If .ObjectID = Me.ObjectID AndAlso .ObjTypeID = Me.ObjTypeID Then
							goMineralCache(lCacheIndex).BeingMinedBy = Nothing
						End If
					End With
				End If
			End If
		End If
		lCacheIndex = -1
		lCacheID = -1
		bRoutePaused = True
	End Sub

	Public Sub LoadFromDataReader(ByRef oData As OleDb.OleDbDataReader)
		With Me
            .CurrentSpeed = CByte(oData("CurrentSpeed"))
            If .CurrentSpeed = 0 Then .bRoutePaused = True

            .Fuel_Cap = CInt(oData("Fuel_Cap"))

            Dim lCol As Int32 = CInt(oData("Colonists"))
            Dim lEnl As Int32 = CInt(oData("Enlisted"))
            Dim lOff As Int32 = CInt(oData("Officers"))

            If lCol > 0 Then
                Dim oTmp As New Epica_GUID()
                oTmp.ObjectID = lCol : oTmp.ObjTypeID = ObjectType.eColonists
                AddCargoRef(oTmp)
            End If
            If lEnl > 0 Then
                Dim oTmp As New Epica_GUID()
                oTmp.ObjectID = lEnl : oTmp.ObjTypeID = ObjectType.eEnlisted
                AddCargoRef(oTmp)
            End If
            If lOff > 0 Then
                Dim oTmp As New Epica_GUID()
                oTmp.ObjectID = lOff : oTmp.ObjTypeID = ObjectType.eOfficers
                AddCargoRef(oTmp)
            End If


			.LocX = CInt(oData("LocX"))
			.LocZ = CInt(oData("LocY"))
			.LocAngle = CShort(oData("LocAngle"))
			.ObjectID = CInt(oData("UnitID"))
			.ObjTypeID = ObjectType.eUnit
			.EntityDef = GetEpicaUnitDef(CInt(oData("UnitDefID")))
            .Owner = GetEpicaPlayer(CInt(oData("OwnerID")))

            Dim lParentID As Int32 = CInt(oData("ParentID"))
            Dim iParentTypeID As Int16 = CShort(oData("ParentTypeID"))
            .ParentObject = GetEpicaObject(lParentID, iParentTypeID)

			'Ensure we register with a containing object's hangar...
			If .ParentObject Is Nothing = False Then
                If iParentTypeID = ObjectType.eFacility OrElse iParentTypeID = ObjectType.eUnit Then
                    CType(.ParentObject, Epica_Entity).AddHangarRef(CType(Me, Epica_GUID))
                End If
			End If

			.Q1_HP = CInt(oData("Q1_HP"))
			.Q2_HP = CInt(oData("Q2_HP"))
			.Q3_HP = CInt(oData("Q3_HP"))
			.Q4_HP = CInt(oData("Q4_HP"))
			.Shield_HP = CInt(oData("Shield_HP"))
			.Structure_HP = CInt(oData("Structure_HP"))
			.EntityName = StringToBytes(CStr(oData("UnitName")))
			.ExpLevel = CByte(oData("ExpLevel"))
			.CurrentStatus = CInt(oData("CurrentStatus"))
            .iCombatTactics = CInt(oData("CombatTactics"))
			.iTargetingTactics = CShort(oData("TargetingTactics"))

			.DestX = CInt(oData("DestX"))
			.DestZ = CInt(oData("DestZ"))
			.DestEnvirID = CInt(oData("DestEnvirID"))
            .DestEnvirTypeID = CShort(oData("DestEnvirTypeID"))

			.yProductionType = CByte(oData("ProductionTypeID"))
			.lCurrentRouteIdx = CInt(oData("CurrentRouteNum"))

			Dim lTemp As Int32
			If .EntityDef Is Nothing = False Then
				ReDim .lCurrentAmmo(.EntityDef.WeaponDefUB)
				For lTemp = 0 To .lCurrentAmmo.Length - 1
					.lCurrentAmmo(lTemp) = -1
				Next lTemp
			End If

            If .Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                If CLng(.Owner.lMilitaryScore) + CLng(.EntityDef.CombatRating) < Int32.MaxValue Then
                    .Owner.lMilitaryScore += .EntityDef.CombatRating
                End If
            End If

			If (.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
				.CurrentStatus = .CurrentStatus Xor elUnitStatus.eMoveLock
            End If

            'Ok, let's see if this unit's dest is invalid
            If .EntityDef Is Nothing = False Then
                If (.EntityDef.yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then
                    If .DestEnvirID <> lParentID OrElse .DestEnvirTypeID <> iParentTypeID Then
                        .DestEnvirID = lParentID
                        .DestEnvirTypeID = iParentTypeID
                        .DestX = .LocX
                        .DestZ = .LocZ
                    End If
                End If
            End If

            .TetherPointX = CInt(oData("TetherPointX"))
            .TetherPointZ = CInt(oData("TetherPointZ"))

            .RallyPointX = CInt(oData("RallyPointX"))
            .RallyPointZ = CInt(oData("RallyPointZ"))
            .RallyPointA = CShort(oData("RallyPointA"))
            .RallyPointEnvirID = CInt(oData("RallyPointEnvirID"))
            .RallyPointEnvirTypeID = CShort(oData("RallyPointEnvirTypeID"))

            .bNeedsSaved = False

		End With
    End Sub

    Public Sub CheckProdQueue()
        If CurrentProduction Is Nothing = False AndAlso CurrentProduction.lProdCount > 0 Then Return
        CurrentProduction = Nothing

        If bProdQueueMoveSent = True Then Return
        bProdQueueMoveSent = True

        If mlProdQueueUB > -1 Then
            If moProdQueue Is Nothing OrElse Me.ParentObject Is Nothing Then
                mlProdQueueUB = -1
                bProdQueueMoveSent = False
                Return
            End If

            For X As Int32 = 0 To mlProdQueueUB
                If moProdQueue(X) Is Nothing = False Then
                    'ok, let's attempt to build it... send off our order
                    Try
                        Dim yMsg(23) As Byte
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yMsg, lPos) : lPos += 2
                        GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                        With moProdQueue(X)
                            System.BitConverter.GetBytes(.ProductionID).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.ProductionTypeID).CopyTo(yMsg, lPos) : lPos += 2
                            System.BitConverter.GetBytes(.lProdX).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.lProdZ).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.iProdA).CopyTo(yMsg, lPos) : lPos += 2
                        End With

                        Dim iTemp As Int16 = CType(ParentObject, Epica_GUID).ObjTypeID
                        If iTemp = ObjectType.ePlanet Then
                            CType(ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                        ElseIf iTemp = ObjectType.eSolarSystem Then
                            CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                        Else
                            LogEvent(LogEventType.CriticalError, "Parent Object of unit is not planet or system with build queue. Type: " & iTemp)
                        End If
                    Catch
                    End Try

                    'Now, remove it from the queue
                    Try
                        Dim lIdx As Int32 = -1
                        For Y As Int32 = X + 1 To mlProdQueueUB
                            lIdx += 1
                            moProdQueue(lIdx) = moProdQueue(Y)
                        Next Y
                        mlProdQueueUB = lIdx
                        ReDim Preserve moProdQueue(mlProdQueueUB)
                    Catch
                        mlProdQueueUB = -1
                    End Try

                    Exit For
                End If
            Next X
        Else
            bProdQueueMoveSent = False
        End If
    End Sub
    Public Sub ClearProdQueue()
        mlProdQueueUB = -1
        bProdQueueMoveSent = False
    End Sub
    Public Sub AddToProdQueue(ByRef oProd As EntityProduction)
        mlProdQueueUB += 1
        ReDim Preserve moProdQueue(mlProdQueueUB)
        moProdQueue(mlProdQueueUB) = oProd

        CheckProdQueue()
    End Sub
    Public Function GetProdQueueList() As Byte()
        Dim yResp(3 + ((mlProdQueueUB + 1) * 14)) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(mlProdQueueUB + 1).CopyTo(yResp, lPos) : lPos += 4
        For X As Int32 = 0 To mlProdQueueUB
            With moProdQueue(X)
                System.BitConverter.GetBytes(.ProductionID).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lProdX).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lProdZ).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.iProdA).CopyTo(yResp, lPos) : lPos += 2
            End With
        Next X
        Return yResp
    End Function

    Public Sub DismantleChildUnit(ByRef oUnit As Unit)
        'get our resulting assets list
        Dim col As Collection = oUnit.RecycleMe(oUnit.EntityDef)
        If col.Count > 0 Then

            Dim oFleet As UnitGroup = Nothing
            Dim oColony As Colony = Nothing

            For Each oItem As RecycledPartsList In col
                Dim lCargoSpaceReq As Int32 = oItem.lQuantity
                Dim bHandled As Boolean = False

                'Resulting assets can go to one of the following and in the order:
                '  1) If I am on a planet, a colony belonging to my owner on this planet
                If Me.ParentObject Is Nothing = False AndAlso CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                    If oColony Is Nothing Then
                        Dim lColIdx As Int32 = Me.Owner.GetColonyFromParent(CType(Me.ParentObject, Epica_GUID).ObjectID, ObjectType.ePlanet)
                        If lColIdx > -1 AndAlso lColIdx <= glColonyUB Then
                            oColony = goColony(lColIdx)
                            If oColony.Owner.ObjectID <> Me.Owner.ObjectID Then oColony = Nothing
                        End If
                    End If

                    If oColony Is Nothing = False Then
                        If oColony.Cargo_Cap > 0 Then
                            Dim lQty As Int32 = Math.Min(oColony.Cargo_Cap, lCargoSpaceReq)
                            oColony.AdjustColonyComponentCache(oItem.lObjID, oItem.iObjTypeID, oItem.lExtended, lQty)
                            If lQty <> lCargoSpaceReq Then lCargoSpaceReq -= lQty Else bHandled = True
                        End If
                    End If
                End If

                '  2) My Cargo Bay
                If bHandled = False AndAlso (Me.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                    If Me.Cargo_Cap > 0 Then
                        Dim lQty As Int32 = Math.Min(Cargo_Cap, lCargoSpaceReq)
                        AddComponentCacheToCargo(oItem.lObjID, oItem.iObjTypeID, lQty, oItem.lExtended)

                        If lQty <> lCargoSpaceReq Then lCargoSpaceReq -= lQty Else bHandled = True
                    End If
                End If

                '  3) My Battlegroup's cargo bay - only units in same environment
                If bHandled = False AndAlso Me.lFleetID > 0 Then
                    If oFleet Is Nothing Then oFleet = GetEpicaUnitGroup(Me.lFleetID)
                    If oFleet Is Nothing = False Then
                        Dim lQty As Int32 = oFleet.AdjustComponentCacheForUnitGroup(oItem.lObjID, oItem.iObjTypeID, oItem.lExtended, oItem.lQuantity, CType(Me.ParentObject, Epica_GUID).ObjectID, CType(Me.ParentObject, Epica_GUID).ObjTypeID)
                        If lQty = 0 Then bHandled = True
                    End If
                End If

                'If it was not handled, then we have no place for it
                If bHandled = False Then Exit For
            Next
        End If

    End Sub

    Public Sub DismantleFacility(ByRef oFac As Facility)
        'get our resulting assets list
        Dim col As Collection = oFac.RecycleMe(oFac.EntityDef)
        If col.Count > 0 Then

            Dim oFleet As UnitGroup = Nothing
            Dim oColony As Colony = Nothing

            For Each oItem As RecycledPartsList In col
                Dim lCargoSpaceReq As Int32 = oItem.lQuantity
                Dim bHandled As Boolean = False

                'Resulting assets can go to one of the following and in the order:
                '  1) If I am on a planet, a colony belonging to my owner on this planet
                If Me.ParentObject Is Nothing = False AndAlso CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                    If oColony Is Nothing Then
                        Dim lColIdx As Int32 = Me.Owner.GetColonyFromParent(CType(Me.ParentObject, Epica_GUID).ObjectID, ObjectType.ePlanet)
                        If lColIdx > -1 AndAlso lColIdx <= glColonyUB Then
                            oColony = goColony(lColIdx)
                            If oColony.Owner.ObjectID <> Me.Owner.ObjectID Then oColony = Nothing
                        End If
                    End If

                    If oColony Is Nothing = False Then
                        If oColony.Cargo_Cap > 0 Then
                            Dim lQty As Int32 = Math.Min(oColony.Cargo_Cap, lCargoSpaceReq)
                            oColony.AdjustColonyComponentCache(oItem.lObjID, oItem.iObjTypeID, oItem.lExtended, lQty)
                            If lQty <> lCargoSpaceReq Then lCargoSpaceReq -= lQty Else bHandled = True
                        End If
                    End If
                End If

                '  2) My Cargo Bay
                If bHandled = False AndAlso (Me.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                    If Me.Cargo_Cap > 0 Then
                        Dim lQty As Int32 = Math.Min(Cargo_Cap, lCargoSpaceReq)
                        AddComponentCacheToCargo(oItem.lObjID, oItem.iObjTypeID, lQty, oItem.lExtended)

                        If lQty <> lCargoSpaceReq Then lCargoSpaceReq -= lQty Else bHandled = True
                    End If
                End If

                '  3) My Battlegroup's cargo bay - only units in same environment
                If bHandled = False AndAlso Me.lFleetID > 0 Then
                    If oFleet Is Nothing Then oFleet = GetEpicaUnitGroup(Me.lFleetID)
                    If oFleet Is Nothing = False Then
                        Dim lQty As Int32 = oFleet.AdjustComponentCacheForUnitGroup(oItem.lObjID, oItem.iObjTypeID, oItem.lExtended, oItem.lQuantity, CType(Me.ParentObject, Epica_GUID).ObjectID, CType(Me.ParentObject, Epica_GUID).ObjTypeID)
                        If lQty = 0 Then bHandled = True
                    End If
                End If

                'If it was not handled, then we have no place for it
                If bHandled = False Then Exit For
            Next
        End If

    End Sub
End Class
