Option Strict On

Module GlobalVars

    Public gb_IS_TEST_SERVER As Boolean = False ' True

#Region "  INTERVAL DELAY CONSTANTS  "
    Public Const gl_DELAY_FOUR_HOURS As Int32 = 432000

#End Region

    'Global Variables go here
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

    Public glCurrentCycle As Int32

    Public glServerID As Int32      'as reported by the operator

    Public glBoxOperatorID As Int32
    Public gsOperatorIP As String
    Public glOperatorPort As Int32
    Public gsExternalIP As String
    Public gsEmailSrvrIP As String
	Public glEmailSrvrPort As Int32
    Public gsBackupOperatorIP As String = "10.70.5.143"
    Public glBackupOperatorPort As Int32 = 7701

    Public gyStarTypeMsg() As Byte
    Public gbStarTypeMsgReady As Boolean = False
    Public gyGalaxySystemMsg() As Byte
    Public gbGalaxyMsgReady As Boolean = False

	Public gfrmDisplayForm As Form1

	Public gfResTimeMult As Single = 1.0F

	Public goModelDefs As ModelDefs

	Public goMsgSys As MsgSystem

    Public gsMOTD As String

    Public goInitialPlayer As Player    'an object to contain all of the initial player tech stuff
    Public goGTC As GalacticTradeSystem

    Public Sub InitializeGlobals()
        InitializeEventLogging()
        InitializeConnection()

        goMsgSys = New MsgSystem()
        goMsgSys.bReceivedDomains = False

        goGTC = New GalacticTradeSystem()
    End Sub

    Public Sub FinalizeGlobals()
        CloseEventLogger()
        CloseConn()
        goMsgSys = Nothing
    End Sub

    Public Function BytesToString(ByVal yBytes() As Byte) As String
        Dim lLen As Int32
        Dim X As Int32

        lLen = yBytes.Length
        For X = 0 To yBytes.Length - 1
            If yBytes(X) = 0 Then
                lLen = X
                Exit For
            End If
        Next X
        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yBytes), 1, lLen)
    End Function

    Public Function StringToBytes(ByVal sVal As String) As Byte()
        Return System.Text.ASCIIEncoding.ASCII.GetBytes(sVal)
    End Function

    Public Function MakeDBStr(ByVal sVal As String) As String
        Return Replace$(sVal, "'", "''")
    End Function

    Public Function AddMineralCache(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal yCacheType As Byte, ByVal lConcentration As Int32, ByVal lQuantity As Int32, ByVal lX As Int32, ByVal lZ As Int32, ByVal oMineral As Mineral) As Int32
        Dim lIdx As Int32 = -1

        If lQuantity = 0 Then Return -1

		Dim oNewCache As New MineralCache()
		With oNewCache
			.CacheTypeID = yCacheType
			.Concentration = lConcentration
			.LocX = lX
			.LocZ = lZ
			.ObjTypeID = ObjectType.eMineralCache
			.ObjectID = -1
			.oMineral = oMineral
			.Quantity = lQuantity
            .ParentObject = GetEpicaObject(lEnvirID, iEnvirTypeID)
            .OriginalConcentration = .Concentration
			If .SaveObject() = False Then
				LogEvent(LogEventType.CriticalError, "AddMineralCache.SaveObject() returned false!")
				Return -1
			End If
		End With

        lIdx = AddMineralCacheToGlobalArray(oNewCache)

		If yCacheType <> MineralCacheType.eMineable Then
			Dim oCache As MineralCache = goMineralCache(lIdx)
			If oCache Is Nothing = False Then
				AddToQueue(glCurrentCycle + 9000, QueueItemType.eRemoveMineralCache, oCache.ObjectID, oCache.ObjTypeID, -1, -1, 0, 0, 0, 0)
			End If

		End If

		Return lIdx
    End Function

    Public Function AddMineralCacheToGlobalArray(ByRef oNewCache As MineralCache) As Int32
        Dim lIdx As Int32 = -1
        If goMineralCache Is Nothing Then ReDim goMineralCache(-1)
        SyncLock goMineralCache
            For X As Int32 = 0 To glMineralCacheUB
                If glMineralCacheIdx(X) = -1 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                'glMineralCacheUB += 1
                'ReDim Preserve goMineralCache(glMineralCacheUB)
                'ReDim Preserve glMineralCacheIdx(glMineralCacheUB)
                'lIdx = glMineralCacheUB

                Dim lNewUB As Int32 = glMineralCacheUB + 1
                If glMineralCacheIdx Is Nothing Then ReDim glMineralCacheIdx(-1)
                If lNewUB > glMineralCacheIdx.GetUpperBound(0) Then
                    ReDim Preserve glMineralCacheIdx(lNewUB + 1000)
                    ReDim Preserve goMineralCache(lNewUB + 1000)
                End If
                glMineralCacheUB = lNewUB
                lIdx = glMineralCacheUB
                glMineralCacheIdx(lIdx) = -1
            End If
            oNewCache.lServerIndex = lIdx
            goMineralCache(lIdx) = oNewCache
            glMineralCacheIdx(lIdx) = oNewCache.ObjectID
        End SyncLock
        Return lIdx
    End Function

	Public Function AddComponentCache(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lQuantity As Int32, ByVal lX As Int32, ByVal lZ As Int32, ByVal lComponentID As Int32, ByVal iComponentTypeID As Int16, ByVal lComponentOwnerID As Int32, ByVal yCacheType As Byte) As Int32
        'Dim X As Int32
		Dim lIdx As Int32 = -1

        If lQuantity = 0 Then Return -1

        Dim oCache As New ComponentCache()
        With oCache
            .ParentObject = GetEpicaObject(lEnvirID, iEnvirTypeID)
            .Quantity = lQuantity
            .LocX = lX
            .LocZ = lZ
            .ObjTypeID = ObjectType.eComponentCache
            .ObjectID = -1
            .yCacheTypeID = yCacheType

            .ComponentID = lComponentID
            .ComponentOwnerID = lComponentOwnerID
            .ComponentTypeID = iComponentTypeID

            If .SaveObject() = False Then
                LogEvent(LogEventType.CriticalError, "AddComponentCache.SaveObject() returned false!")
                Return -1
            End If
        End With

        lIdx = AddComponentCacheToGlobalArray(oCache)

        'For X = 0 To glComponentCacheUB
        '	If glComponentCacheIdx(X) = -1 Then
        '		lIdx = X
        '		Exit For
        '	End If
        'Next X

        'If lIdx = -1 Then
        '	glComponentCacheUB += 1
        '	ReDim Preserve goComponentCache(glComponentCacheUB)
        '	ReDim Preserve glComponentCacheIdx(glComponentCacheUB)
        '	lIdx = glComponentCacheUB
        'End If
        'goComponentCache(lIdx) = Nothing
        'goComponentCache(lIdx) = New ComponentCache

        'With goComponentCache(lIdx)

        glComponentCacheIdx(lIdx) = oCache.ObjectID
        'End With
		If yCacheType <> MineralCacheType.eMineable Then
            If oCache Is Nothing = False Then
                AddToQueue(glCurrentCycle + 9000, QueueItemType.eRemoveComponentCache, oCache.ObjectID, oCache.ObjTypeID, -1, -1, 0, 0, 0, 0)
            End If
		End If

		Return lIdx
    End Function

    Public Function AddComponentCacheToGlobalArray(ByRef oNewCache As ComponentCache) As Int32
        Dim lIdx As Int32 = -1
        If goComponentCache Is Nothing Then ReDim goComponentCache(-1)
        SyncLock goComponentCache
            For X As Int32 = 0 To glComponentCacheUB
                If glComponentCacheIdx(X) = -1 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                'glComponentCacheUB += 1
                'ReDim Preserve goComponentCache(glComponentCacheUB)
                'ReDim Preserve glComponentCacheIdx(glComponentCacheUB)
                'lIdx = glComponentCacheUB

                Dim lNewUB As Int32 = glComponentCacheUB + 1
                If glComponentCacheIdx Is Nothing Then ReDim glComponentCacheIdx(-1)
                If lNewUB > glComponentCacheIdx.GetUpperBound(0) Then
                    ReDim Preserve glComponentCacheIdx(lNewUB + 1000)
                    ReDim Preserve goComponentCache(lNewUB + 1000)
                End If
                glComponentCacheUB = lNewUB
                lIdx = glComponentCacheUB
                glComponentCacheIdx(lIdx) = -1
            End If

            goComponentCache(lIdx) = oNewCache
            glComponentCacheIdx(lIdx) = oNewCache.ObjectID
        End SyncLock
        Return lIdx
    End Function

    Public Function AddAmmunitionCache(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lQuantity As Int32, ByVal lX As Int32, ByVal lZ As Int32, ByRef oWeapon As BaseWeaponTech) As AmmunitionCache
        Dim lIdx As Int32 = -1

        If lQuantity = 0 Then Return Nothing

        Dim oNewCache As New AmmunitionCache()

        With oNewCache
            .ParentObject = GetEpicaObject(lEnvirID, iEnvirTypeID)
            .Quantity = lQuantity
            .LocX = lX
            .LocZ = lZ
            .ObjTypeID = ObjectType.eAmmunition
            .ObjectID = -1

            .oWeaponTech = oWeapon

            If .SaveObject() = False Then
                LogEvent(LogEventType.CriticalError, "AddAmmunitionCache.SaveObject() returned false!")
                Return Nothing
            End If
        End With

        Return oNewCache
	End Function

	Public Function AddGuild(ByRef oGuild As Guild) As Int32
		Dim lIdx As Int32 = -1

		If oGuild.SaveObject() = True Then
			If goGuild Is Nothing Then ReDim goGuild(-1)
			SyncLock goGuild
				'Now, find a suitable place...
				For X As Int32 = 0 To glGuildUB
					If glGuildIdx(X) = -1 Then
						lIdx = X
						Exit For
					End If
				Next X

				If lIdx = -1 Then
					glGuildUB += 1
					ReDim Preserve glGuildIdx(glGuildUB)
					ReDim Preserve goGuild(glGuildUB)
					lIdx = glGuildUB
				End If
				goGuild(lIdx) = oGuild
				glGuildIdx(lIdx) = oGuild.ObjectID
			End SyncLock
		End If

		Return lIdx

	End Function

    ''' <summary>
    ''' This always puts the unit in the PRODUCER's HANGAR (or Producer's parent facility's hangar). If the Hangar is too small, the unit is forced to launch immediately
    ''' </summary>
    ''' <param name="oProducer"></param>
    ''' <param name="oUnitDef"></param>
    ''' <param name="yExpLevel"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddUnit(ByRef oProducer As Epica_Entity, ByVal oUnitDef As Epica_Entity_Def, ByVal yExpLevel As Byte) As Int32
        Dim X As Int32
        Dim lIdx As Int32 = -1

        Dim oTmp As Unit = New Unit()

        'Ok, populate our values
        With oTmp
            .bProducing = False
            '.Cargo_Cap = oUnitDef.Cargo_Cap
            .CurrentProduction = Nothing
            .CurrentSpeed = 0

            .CurrentStatus = 0
            For X = 0 To oUnitDef.lSideCrits.Length - 1
                .CurrentStatus = .CurrentStatus Or oUnitDef.lSideCrits(X)
            Next X

            .EntityDef = oUnitDef
            oUnitDef.DefName.CopyTo(.EntityName, 0)
            .ExpLevel = yExpLevel
            .Fuel_Cap = oUnitDef.Fuel_Cap
            '.Hangar_Cap = oUnitDef.Hangar_Cap
            .ObjTypeID = ObjectType.eUnit
            .Owner = oProducer.Owner

            If .Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                If CLng(.Owner.lMilitaryScore) + CLng(oUnitDef.CombatRating) < Int32.MaxValue Then
                    .Owner.lMilitaryScore += oUnitDef.CombatRating
                End If
            End If

            'PRODUCER PROPERTY INHERITENCE BEGINS HERE!!!
            'The entity inherits the producer's tactics and targeting settings
            .iCombatTactics = oProducer.iCombatTactics
            .iTargetingTactics = oProducer.iTargetingTactics
            .lExtendedCT_1 = oProducer.lExtendedCT_1
            .lExtendedCT_2 = oProducer.lExtendedCT_2

            Dim oParentObj As Object = Nothing

            If CType(oProducer.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                'Parent's Parent is a facility, so I am in a space station... my parent will be my parent's parent
                oParentObj = oProducer.ParentObject
                .LocX = CType(oParentObj, Facility).LocX
                .LocZ = CType(oParentObj, Facility).LocZ
                .LocAngle = CType(oParentObj, Facility).LocAngle
            Else
                'Producer's Parent is not a facility, my parent will be my producer
                oParentObj = oProducer
                .LocAngle = oProducer.LocAngle
                .LocX = oProducer.LocX
                .LocZ = oProducer.LocZ
            End If
            .ParentObject = oParentObj


            .Q1_HP = oUnitDef.Q1_MaxHP
            .Q2_HP = oUnitDef.Q2_MaxHP
            .Q3_HP = oUnitDef.Q3_MaxHP
            .Q4_HP = oUnitDef.Q4_MaxHP
            .Shield_HP = oUnitDef.Shield_MaxHP
            .Structure_HP = oUnitDef.Structure_MaxHP
            .yProductionType = oUnitDef.ProductionTypeID

            'Set our current ammo...
            ReDim .lCurrentAmmo(oUnitDef.WeaponDefUB)
            For X = 0 To oUnitDef.WeaponDefUB
                .lCurrentAmmo(X) = oUnitDef.WeaponDefs(X).mlAmmoCap
            Next X

            .DataChanged()
        End With

        If oTmp.SaveObject() = True Then
			'Now, find a suitable place...
            'SyncLock goUnit
            '	For X = 0 To glUnitUB
            '		If glUnitIdx(X) = -1 Then
            '			lIdx = X
            '			Exit For
            '		End If
            '	Next X

            '	If lIdx = -1 Then
            '		glUnitUB += 1
            '		ReDim Preserve glUnitIdx(glUnitUB)
            '		ReDim Preserve goUnit(glUnitUB)
            '		lIdx = glUnitUB
            '		glUnitIdx(lIdx) = -1
            '	End If
            '	goUnit(lIdx) = oTmp
            '	glUnitIdx(lIdx) = oTmp.ObjectID
            'End SyncLock
            lIdx = AddUnitToGlobalArray(oTmp)

			Return lIdx
		Else
			Return -1
		End If


    End Function

    Public Function AddUnitToGlobalArray(ByRef oUnit As Unit) As Int32
        'Now, find a suitable place...
        SyncLock goUnit
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To glUnitUB
                If glUnitIdx(X) = -1 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                Dim lNewUB As Int32 = glUnitUB + 1
                If glUnitIdx Is Nothing Then ReDim glUnitIdx(-1)
                If lNewUB > glUnitIdx.GetUpperBound(0) Then
                    ReDim Preserve glUnitIdx(lNewUB + 1000)
                    ReDim Preserve goUnit(lNewUB + 1000)
                End If
                glUnitUB = lNewUB
                lIdx = glUnitUB
                glUnitIdx(lIdx) = -1
            End If
            goUnit(lIdx) = oUnit
            glUnitIdx(lIdx) = oUnit.ObjectID
            Return lIdx
        End SyncLock
    End Function

    Public Function AddUnitDefToGlobalArray(ByVal oUnitDef As Epica_Entity_Def) As Int32
        SyncLock goUnitDef
            'Now, find a suitable place...
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To glUnitDefUB
                If glUnitDefIdx(X) = -1 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                glUnitDefUB += 1
                ReDim Preserve glUnitDefIdx(glUnitDefUB)
                ReDim Preserve goUnitDef(glUnitDefUB)
                lIdx = glUnitDefUB
            End If
            goUnitDef(lIdx) = oUnitDef
            glUnitDefIdx(lIdx) = oUnitDef.ObjectID
            Return lIdx
        End SyncLock
    End Function

    Public Function AddWeaponDefToGlobalArray(ByVal oWpnDef As WeaponDef) As Int32
        SyncLock goWeaponDefs
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To glWeaponDefUB
                If glWeaponDefIdx(X) = -1 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                glWeaponDefUB += 1
                ReDim Preserve glWeaponDefIdx(glWeaponDefUB)
                ReDim Preserve goWeaponDefs(glWeaponDefUB)
                lIdx = glWeaponDefUB
            End If
            goWeaponDefs(lIdx) = oWpnDef
            glWeaponDefIdx(lIdx) = oWpnDef.ObjectID
            Return lIdx
        End SyncLock
    End Function

    Private Function DoLookupFacilityFunction(ByVal yType As Byte, ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lIdx As Int32) As Int32
        SyncLock gcolFacLookup
            Try
                Dim sKey As String = lID & ";" & iTypeID
                Select Case yType
                    Case 0  'lookup
                        If gcolFacLookup.Contains(sKey) = True Then
                            Return CInt(gcolFacLookup(sKey))
                        Else
                            Return -1
                        End If
                    Case 1  'add
                        If gcolFacLookup.Contains(sKey) = True Then Return -1
                        gcolFacLookup.Add(lIdx, sKey)
                    Case 2  'remove
                        If gcolFacLookup.Contains(sKey) = True Then gcolFacLookup.Remove(sKey)
                End Select
            Catch
            End Try
        End SyncLock
        Return -1
    End Function
    Public Function LookupFacility(ByVal lID As Int32, ByVal iTypeID As Int16) As Int32
        Return DoLookupFacilityFunction(0, lID, iTypeID, -1)
    End Function
    Public Sub AddLookupFacility(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lIdx As Int32)
        DoLookupFacilityFunction(1, lID, iTypeID, lIdx)
    End Sub
    Public Sub RemoveLookupFacility(ByVal lID As Int32, ByVal iTypeID As Int16)
        DoLookupFacilityFunction(2, lID, iTypeID, -1)
    End Sub


    Public Function AddFacility(ByVal oProducer As Epica_Entity, ByVal oFacilityDef As Epica_Entity_Def, ByVal LocX As Int32, ByVal LocZ As Int32, ByVal LocAngle As Int16) As Int32
        Dim X As Int32
        Dim lIdx As Int32 = -1
        'Dim bFound As Boolean = False
        Dim lColonyIdx As Int32 = -1
        Dim lTemp As Int32
        Dim bNewColony As Boolean = False

		Dim oTmp As Facility = New Facility()

        'Find our Parent Colony
        If oProducer.yProductionType = ProductionType.eSpaceStationSpecial Then
            'Ok, the parent will be the producer's parent colony
            lTemp = CType(oProducer, Facility).ParentColony.ObjectID
            For X = 0 To glColonyUB
                If glColonyIdx(X) = lTemp Then
                    lColonyIdx = X
                    Exit For
                End If
            Next X

            'TODO: Are we sure about this?
            'oProducer.Hangar_Cap += oFacilityDef.Hangar_Cap
            'oProducer.Cargo_Cap += oFacilityDef.Cargo_Cap
            oProducer.Fuel_Cap += oFacilityDef.Fuel_Cap
        ElseIf oFacilityDef.ProductionTypeID <> ProductionType.eSpaceStationSpecial AndAlso (oFacilityDef.ProductionTypeID <> ProductionType.eTradePost OrElse CType(oProducer.ParentObject, Epica_GUID).ObjTypeID <> ObjectType.eSolarSystem) Then
            'we have to find the colony for this environment
            With CType(oProducer.ParentObject, Epica_GUID)
                lColonyIdx = oProducer.Owner.GetColonyFromParent(.ObjectID, .ObjTypeID)
            End With
        End If

		Dim oColony As Colony = Nothing
        If lColonyIdx = -1 Then
            If (oFacilityDef.yChassisType And ChassisType.eSpaceBased) = 0 OrElse oFacilityDef.ProductionTypeID = ProductionType.eSpaceStationSpecial OrElse oFacilityDef.ProductionTypeID = ProductionType.eTradePost Then
                lColonyIdx = AddColony(oProducer.Owner, oProducer.ParentObject)
                If lColonyIdx = -1 Then Return -1
                bNewColony = True
            Else
                lColonyIdx = Int32.MinValue
            End If
        End If
        If lColonyIdx = -1 Then Return -1
        If lColonyIdx <> Int32.MinValue Then
            oColony = goColony(lColonyIdx)
            If oColony Is Nothing Then Return -1
        End If

		'Ok, populate our values
		With oTmp
			.bProducing = False
			'.Cargo_Cap = oFacilityDef.Cargo_Cap
			.CurrentProduction = Nothing
			.CurrentSpeed = 0

			.CurrentStatus = 0
			For X = 0 To oFacilityDef.lSideCrits.Length - 1
				.CurrentStatus = .CurrentStatus Or oFacilityDef.lSideCrits(X)
			Next X
            If lColonyIdx = Int32.MinValue Then
                .CurrentStatus = .CurrentStatus Or elUnitStatus.eFacilityPowered
            End If

			.EntityDef = CType(oFacilityDef, FacilityDef)
			.ParentColony = oColony
			oFacilityDef.DefName.CopyTo(.EntityName, 0)
			.ExpLevel = 0
			.Fuel_Cap = oFacilityDef.Fuel_Cap
			'.Hangar_Cap = oFacilityDef.Hangar_Cap
			.iCombatTactics = oProducer.iCombatTactics
			.iTargetingTactics = oProducer.iTargetingTactics
			.LocAngle = LocAngle
			.LocX = LocX
			.LocZ = LocZ
			.ObjTypeID = ObjectType.eFacility
			.Owner = oProducer.Owner
			If oProducer.yProductionType = ProductionType.eSpaceStationSpecial Then
				.ParentObject = oProducer
			Else : .ParentObject = oProducer.ParentObject
			End If
			.Q1_HP = oFacilityDef.Q1_MaxHP
			.Q2_HP = oFacilityDef.Q2_MaxHP
			.Q3_HP = oFacilityDef.Q3_MaxHP
			.Q4_HP = oFacilityDef.Q4_MaxHP
			.Shield_HP = oFacilityDef.Shield_MaxHP
			.Structure_HP = oFacilityDef.Structure_MaxHP
			.yProductionType = oFacilityDef.ProductionTypeID

			If .yProductionType = ProductionType.eRefining OrElse .yProductionType = ProductionType.eMining OrElse (.yProductionType And ProductionType.eProduction) <> 0 OrElse .yProductionType = ProductionType.eCommandCenterSpecial Then
				.AutoLaunch = True
			End If

			'.FacilitySize = CByte(.Height * .Width)
			'If .EntityDef.ProductionRequirements Is Nothing Then
			'    ' .CurrentWorkers = 0
			'Else
			'    'lTemp = .EntityDef.ProductionRequirements.ColonistCost + .EntityDef.ProductionRequirements.EnlistedCost + .EntityDef.ProductionRequirements.OfficerCost
			'    'If lTemp = 0 AndAlso bNewColony = True Then
			'    '    'First attempt to set it as active
			'    '    If .SetActive(True) = True Then
			'    '        lTemp = .MaxWorkers
			'    '        .ParentColony.Population += lTemp
			'    '        '.CurrentWorkers = lTemp
			'    '        '.CurrentColonists = lTemp
			'    '    End If
			'    'End If
			If bNewColony = True Then .ParentColony.Population += 5000

			'    '.CurrentWorkers = Math.Min(lTemp, .MaxWorkers)
			'End If
			'.ParentColony.Population += .CurrentWorkers
			.DataChanged()
		End With

		If oTmp.SaveObject() = True Then
			If oTmp.yProductionType = ProductionType.eSpaceStationSpecial OrElse oTmp.yProductionType = ProductionType.eTradePost Then
				If CType(oTmp.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
					If oColony Is Nothing = False Then
						oColony.ParentObject = oTmp
						oColony.DataChanged()
						oColony.SaveObject()
					End If
				End If
			End If

            'SyncLock goFacility
            '	'Now, find a suitable place...
            '	For X = 0 To glFacilityUB
            '		If glFacilityIdx(X) = -1 Then
            '			lIdx = X
            '			Exit For
            '		End If
            '	Next X

            '	If lIdx = -1 Then
            '		glFacilityUB += 1
            '		ReDim Preserve glFacilityIdx(glFacilityUB)
            '		ReDim Preserve goFacility(glFacilityUB)
            '		lIdx = glFacilityUB
            '	End If
            '	goFacility(lIdx) = oTmp
            '	glFacilityIdx(lIdx) = oTmp.ObjectID
            '	oTmp.ServerIndex = lIdx
            'End SyncLock
            lIdx = AddFacilityToGlobalArray(oTmp)

			If oTmp.ParentColony Is Nothing = False Then
				oTmp.ParentColony.AddChildFacility(oTmp)
				If oTmp.yProductionType = ProductionType.eCommandCenterSpecial Then oTmp.ParentColony.bCCInProduction = False
				If oTmp.yProductionType = ProductionType.eTradePost Then oTmp.ParentColony.bTradepostInProduction = False
				oTmp.SetActive(True)
				oTmp.ParentColony.UpdateAllValues(-1)
			End If

            If oTmp.yProductionType = ProductionType.eMining Then

                'If oTmp.ParentObject Is Nothing = False AndAlso CType(oTmp.ParentObject, Epica_GUID).ObjectID > 500000000 AndAlso oTmp.Owner Is Nothing = False AndAlso oTmp.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                '    oTmp.AddMineralCacheToCargo(157, 3000000)
                'Else

                'Ok, its a mining facility, find if there is a mineral cache underneath it
                'TODO: Make this less generic (hard-coded) and put in code to get the model's size
                Dim rcTemp As Rectangle = Rectangle.FromLTRB(oTmp.LocX - 50, oTmp.LocZ - 50, oTmp.LocX + 50, oTmp.LocZ + 50)
                'TODO: to speed this up, we could separate the mineral caches into the two different types (mineable and surface)
                For X = 0 To glMineralCacheUB
                    If glMineralCacheIdx(X) <> -1 Then
                        Dim oCache As MineralCache = goMineralCache(X)
                        If oCache Is Nothing = False AndAlso oCache.CacheTypeID = MineralCacheType.eMineable AndAlso oCache.ParentObject Is Nothing = False Then
                            With oCache
                                If CType(.ParentObject, Epica_GUID).ObjectID = CType(oTmp.ParentObject, Epica_GUID).ObjectID Then
                                    If CType(.ParentObject, Epica_GUID).ObjTypeID = CType(oTmp.ParentObject, Epica_GUID).ObjTypeID Then
                                        'Ok, check location
                                        If rcTemp.Contains(.LocX, .LocZ) = True Then
                                            'Ok, found one
                                            If .BeingMinedBy Is Nothing Then
                                                If glMineralCacheIdx(X) = .ObjectID Then
                                                    oTmp.lCacheIndex = X
                                                    oTmp.lCacheID = .ObjectID
                                                    If oTmp.Active Then oTmp.bMining = True '???
                                                    AddFacilityMining(lIdx, oTmp.ObjectID)

                                                    If oTmp.oMiningBid Is Nothing Then oTmp.oMiningBid = New MineBuyOrderManager
                                                    With oTmp.oMiningBid
                                                        .oParentFacility = oTmp
                                                        .oMineralCache = oCache
                                                        .lMaxDaysSold = 0
                                                        .lDaysNotSold = 0
                                                        .lCurrentConseqDays = 0
                                                        .bSomethingSold = False
                                                    End With
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End With
                        End If
                    End If
                Next X

                'End If
            ElseIf (oTmp.EntityDef.ModelID And 255) = 148 Then
                oTmp.AutoLaunch = True
                Dim oSystem As SolarSystem = CType(oTmp.ParentObject, SolarSystem)
                Dim oMinedPlanet As Planet = Nothing
                Dim lDist As Int32 = Int32.MaxValue
                For X = 0 To oSystem.mlPlanetUB
                    Dim lPIdx As Int32 = oSystem.GetPlanetIdx(X)
                    Dim oPlanet As Planet = goPlanet(lPIdx)
                    If oPlanet Is Nothing = False Then
                        Dim fDX As Single = goFacility(lIdx).LocX - oPlanet.LocX
                        Dim fDZ As Single = goFacility(lIdx).LocZ - oPlanet.LocZ
                        fDX *= fDX
                        fDZ *= fDZ
                        Dim fTemp As Single = CSng(Math.Sqrt(fDX + fDZ))
                        If fTemp < lDist Then
                            lDist = CInt(fTemp)
                            oMinedPlanet = oPlanet
                        End If
                    End If
                Next X

                If oMinedPlanet Is Nothing = False Then
                    oMinedPlanet.RingMinerID = goFacility(lIdx).ObjectID
                End If
            End If
			oTmp.DataChanged()

			'Now, try to save it again
			'goFacility(lIdx).SaveObject()

			Return lIdx
		Else
			If bNewColony = True Then
				'TODO: Delete the colony object
			End If
			Return -1
		End If

	End Function

	Public Function AddFacilityNoProducer(ByVal oFacilityDef As Epica_Entity_Def, ByRef oOwner As Player, ByVal lParentID As Int32, ByVal iParentTypeID As Int16, ByVal LocX As Int32, ByVal LocZ As Int32, ByVal LocAngle As Int16, ByVal iCombatTactics As eiBehaviorPatterns, ByVal iTargeting As eiTacticalAttrs, ByVal bNoColony As Boolean) As Int32
		Dim X As Int32
		Dim lIdx As Int32 = -1
		Dim lColonyIdx As Int32 = -1
		Dim bNewColony As Boolean = False

		Dim oTmp As Facility = New Facility()

		'Find our Parent Colony
		Dim oColony As Colony = Nothing
		If oOwner Is Nothing = False AndAlso bNoColony = False Then
			If oFacilityDef.ProductionTypeID <> ProductionType.eSpaceStationSpecial Then
				'we have to find the colony for this environment
				lColonyIdx = oOwner.GetColonyFromParent(lParentID, iParentTypeID)
			End If

			If lColonyIdx = -1 Then
				lColonyIdx = AddColony(oOwner, GetEpicaObject(lParentID, iParentTypeID))
				If lColonyIdx = -1 Then Return -1
				bNewColony = True
			End If
			If lColonyIdx = -1 Then Return -1
			oColony = goColony(lColonyIdx)
			If oColony Is Nothing Then Return -1
		End If

		'Ok, populate our values
		With oTmp
			.bProducing = False
			'.Cargo_Cap = oFacilityDef.Cargo_Cap
			.CurrentProduction = Nothing
			.CurrentSpeed = 0

			.CurrentStatus = 0
			For X = 0 To oFacilityDef.lSideCrits.Length - 1
				.CurrentStatus = .CurrentStatus Or oFacilityDef.lSideCrits(X)
			Next X

			.EntityDef = CType(oFacilityDef, FacilityDef)
			.ParentColony = oColony
			oFacilityDef.DefName.CopyTo(.EntityName, 0)
			.ExpLevel = 0
			.Fuel_Cap = oFacilityDef.Fuel_Cap
			.iCombatTactics = iCombatTactics
			.iTargetingTactics = iTargeting
			.LocAngle = LocAngle
			.LocX = LocX
			.LocZ = LocZ
			.ObjTypeID = ObjectType.eFacility
			.Owner = oOwner
			.ParentObject = GetEpicaObject(lParentID, iParentTypeID)
			.Q1_HP = oFacilityDef.Q1_MaxHP
			.Q2_HP = oFacilityDef.Q2_MaxHP
			.Q3_HP = oFacilityDef.Q3_MaxHP
			.Q4_HP = oFacilityDef.Q4_MaxHP
			.Shield_HP = oFacilityDef.Shield_MaxHP
			.Structure_HP = oFacilityDef.Structure_MaxHP
			.yProductionType = oFacilityDef.ProductionTypeID

			If .yProductionType = ProductionType.eRefining OrElse .yProductionType = ProductionType.eMining OrElse (.yProductionType And ProductionType.eProduction) <> 0 OrElse .yProductionType = ProductionType.eCommandCenterSpecial Then
				.AutoLaunch = True
			End If

			If bNewColony = True Then .ParentColony.Population += 5000

			.DataChanged()
		End With

		If oTmp.SaveObject() = True Then
			If oTmp.yProductionType = ProductionType.eSpaceStationSpecial OrElse oTmp.yProductionType = ProductionType.eTradePost Then
				If CType(oTmp.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
					If oColony Is Nothing = False Then
						oColony.ParentObject = oTmp
						oColony.DataChanged()
						oColony.SaveObject()
					End If
				End If
			End If

            'SyncLock goFacility
            '	'Now, find a suitable place...
            '	For X = 0 To glFacilityUB
            '		If glFacilityIdx(X) = -1 Then
            '			lIdx = X
            '			Exit For
            '		End If
            '	Next X

            '	If lIdx = -1 Then
            '		glFacilityUB += 1
            '		ReDim Preserve glFacilityIdx(glFacilityUB)
            '		ReDim Preserve goFacility(glFacilityUB)
            '		lIdx = glFacilityUB
            '	End If
            '	goFacility(lIdx) = oTmp
            '	glFacilityIdx(lIdx) = oTmp.ObjectID
            '	oTmp.ServerIndex = lIdx
            'End SyncLock
            lIdx = AddFacilityToGlobalArray(oTmp)

			If oTmp.ParentColony Is Nothing = False Then
				oTmp.ParentColony.AddChildFacility(oTmp)
				If oTmp.yProductionType = ProductionType.eCommandCenterSpecial Then oTmp.ParentColony.bCCInProduction = False
				If oTmp.yProductionType = ProductionType.eTradePost Then oTmp.ParentColony.bTradepostInProduction = False
				oTmp.SetActive(True)
				oTmp.ParentColony.UpdateAllValues(-1)
			End If

			If oTmp.yProductionType = ProductionType.eMining Then
				'Ok, its a mining facility, find if there is a mineral cache underneath it
				'TODO: Make this less generic (hard-coded) and put in code to get the model's size
				Dim rcTemp As Rectangle = Rectangle.FromLTRB(oTmp.LocX - 50, oTmp.LocZ - 50, oTmp.LocX + 50, oTmp.LocZ + 50)
				'TODO: to speed this up, we could separate the mineral caches into the two different types (mineable and surface)
				For X = 0 To glMineralCacheUB
					If glMineralCacheIdx(X) <> -1 Then
						Dim oCache As MineralCache = goMineralCache(X)
						If oCache Is Nothing = False AndAlso oCache.CacheTypeID = MineralCacheType.eMineable Then
							With oCache
								If CType(.ParentObject, Epica_GUID).ObjectID = CType(oTmp.ParentObject, Epica_GUID).ObjectID Then
									If CType(.ParentObject, Epica_GUID).ObjTypeID = CType(oTmp.ParentObject, Epica_GUID).ObjTypeID Then
										'Ok, check location
										If rcTemp.Contains(.LocX, .LocZ) = True Then
											'Ok, found one
											If .BeingMinedBy Is Nothing Then
												If glMineralCacheIdx(X) = .ObjectID Then
													oTmp.lCacheIndex = X
													oTmp.lCacheID = .ObjectID
													If oTmp.Active Then oTmp.bMining = True '???
													AddFacilityMining(lIdx, oTmp.ObjectID)
													Exit For
												End If
											End If
										End If
									End If
								End If
							End With
						End If
					End If
				Next X
			End If
			oTmp.DataChanged()

			'Now, try to save it again
			'goFacility(lIdx).SaveObject()

			Return lIdx
		Else
			If bNewColony = True Then
				'TODO: Delete the colony object
			End If
			Return -1
		End If
    End Function

    Public Function AddFacilityToGlobalArray(ByRef oFac As Facility) As Int32
        SyncLock goFacility
            'Now, find a suitable place...
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To glFacilityUB
                If glFacilityIdx(X) = -1 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                Dim lNewUB As Int32 = glFacilityUB + 1
                If glFacilityIdx Is Nothing Then ReDim glFacilityIdx(-1)
                If lNewUB > glFacilityIdx.GetUpperBound(0) Then
                    ReDim Preserve glFacilityIdx(lNewUB + 1000)
                    ReDim Preserve goFacility(lNewUB + 1000)
                End If
                glFacilityUB = lNewUB
                'ReDim Preserve glFacilityIdx(glFacilityUB)
                'ReDim Preserve goFacility(glFacilityUB)
                lIdx = glFacilityUB
                glFacilityIdx(lIdx) = -1
            End If
            goFacility(lIdx) = oFac
            glFacilityIdx(lIdx) = oFac.ObjectID
            oFac.ServerIndex = lIdx

            'With oFac
            '    If .Owner Is Nothing = False AndAlso .ParentObject Is Nothing = False AndAlso .Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID Then
            '        If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
            '            .Owner.AddSpaceFacIdx(.ServerIndex)
            '        End If
            '    End If
            'End With
            AddLookupFacility(oFac.ObjectID, oFac.ObjTypeID, oFac.ServerIndex)


            Return lIdx
        End SyncLock

    End Function

    Public Function AddFacilityDefToGlobalArray(ByVal oFacDef As FacilityDef) As Int32
        SyncLock goFacilityDef
            'Now, find a suitable place...
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To glFacilityDefUB
                If glFacilityDefIdx(X) = -1 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                glFacilityDefUB += 1
                ReDim Preserve glFacilityDefIdx(glFacilityDefUB)
                ReDim Preserve goFacilityDef(glFacilityDefUB)
                lIdx = glFacilityDefUB
            End If
            goFacilityDef(lIdx) = oFacDef
            glFacilityDefIdx(lIdx) = oFacDef.ObjectID
            Return lIdx
        End SyncLock
    End Function

    Public Function AddFormationToGlobalArray(ByRef oFormation As FormationDef) As Int32
        Dim lIdx As Int32 = -1
        SyncLock glFormationDefIdx
            For X As Int32 = 0 To glFormationDefUB
                If glFormationDefIdx(X) = -1 Then
                    goFormationDefs(X) = oFormation
                    glFormationDefIdx(X) = oFormation.FormationID
                    lIdx = X
                    Exit For
                End If
            Next X
            If lIdx = -1 Then
                ReDim Preserve goFormationDefs(glFormationDefUB + 1)
                ReDim Preserve glFormationDefIdx(glFormationDefUB + 1)
                goFormationDefs(glFormationDefUB + 1) = oFormation
                glFormationDefIdx(glFormationDefUB + 1) = oFormation.FormationID
                glFormationDefUB += 1
            End If
        End SyncLock
        Return lIdx
    End Function

    Public Function AddColony(ByRef oOwner As Player, ByRef oParent As Object) As Int32
        Dim X As Int32
        Dim lIdx As Int32 = -1
		Dim oTmp As Colony = New Colony()

		Dim yNewTax As Byte = 40

		If oParent Is Nothing Then Return -1
		With CType(oParent, Epica_GUID)
			If .ObjTypeID = ObjectType.eFacility Then
				Dim oTmpParent As Epica_GUID = CType(CType(oParent, Facility).ParentObject, Epica_GUID)
				If oTmpParent.ObjTypeID = ObjectType.ePlanet Then
					'let's get a colonyid for this parent
					Dim lTempVal As Int32 = oOwner.GetColonyFromParent(oTmpParent.ObjectID, oTmpParent.ObjTypeID)
					If lTempVal <> -1 Then Return lTempVal

					'Ok, we are here, set oparent to oTmpParent...
					oParent = oTmpParent
				Else : yNewTax = 0
				End If
			End If
		End With

        With oTmp
            .AverageIncome = 10000
            .ColonyEnlisted = 0

            Dim oSystem As SolarSystem = Nothing
            If CType(oParent, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                .ColonyName = CType(oParent, Planet).PlanetName
                oSystem = CType(oParent, Planet).ParentSystem
            ElseIf CType(oParent, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                .ColonyName = CType(oParent, SolarSystem).SystemName
                oSystem = CType(oParent, SolarSystem)
            End If

            If oSystem Is Nothing = False Then
                oSystem.CheckUnlockSystemWormholes()
                If oSystem.SystemType <> SolarSystem.elSystemType.RespawnSystem Then
                    'Ok, check if the player making this has a colony in a respawn system
                    'We could do it the long way... but instead, we are going to do it the cheater's way...
                    oOwner.UnlockRespawnSystems()
                End If
            End If

            .ColonyOfficers = 0
            .CostOfLiving = 8000
            .DesireForWar = 100
            .ElectronicsLevel = 10
            .GovScore = 100
            'TODO: Locs aren't used anymore, are they?
            .LocX = 0
            .LocY = 0
            .MaterialsLevel = 10
            '.Morale = 100
            .ObjTypeID = ObjectType.eColony
            .Owner = oOwner
            .ParentObject = oParent
            .Population = 0
            .PowerConsumption = 0
            .PowerGeneration = 0
            .PowerLevel = 10
			.TaxRate = yNewTax
			.ColonyStart = GetDateAsNumber(Now)
            .DataChanged()
        End With

        If oTmp.SaveObject = True Then

			'Now, find a place for it
			SyncLock goColony
				For X = 0 To glColonyUB
					If glColonyIdx(X) = -1 Then
						lIdx = X
						Exit For
					End If
				Next X
				If lIdx = -1 Then
					glColonyUB += 1
					ReDim Preserve glColonyIdx(glColonyUB)
					ReDim Preserve goColony(glColonyUB)
					lIdx = glColonyUB
				End If
				goColony(lIdx) = oTmp
				glColonyIdx(lIdx) = oTmp.ObjectID
			End SyncLock

			'Associate this colony to the owner (and vice versa)
			goColony(lIdx).Owner.AddColonyIndex(lIdx)
			If CType(oParent, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
				CType(oParent, Planet).AddColonyReference(lIdx)
			End If

			Return lIdx
		Else
			Return -1
		End If
	End Function

    Public Sub DestroyEntity(ByRef oEntity As Epica_Entity, ByVal bSendDomainRO As Boolean, ByVal lKilledByID As Int32, ByVal bCreateEntityDefMinerals As Boolean, ByVal sFromSource As String)
        'Now, create the mineral deposits... oEntity should have an EntityDef with it... that entitydef should
        '  have the mineral cache details for when a unit of that type is destroyed
        Dim oTmpDef As Epica_Entity_Def = Nothing

        Dim iTemp As Int16 = -1
        If oEntity.ParentObject Is Nothing = False Then iTemp = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID

        If bSendDomainRO = True Then
            Dim yMsg(8) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yMsg, 0)
            oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
            yMsg(8) = RemovalType.eObjectDestroyed
            If iTemp = ObjectType.ePlanet Then
                CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
            ElseIf iTemp = ObjectType.eSolarSystem Then
                CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
            End If
        End If

        LogEvent(LogEventType.ExtensiveLogging, "DestroyEntity (" & sFromSource & "): " & oEntity.ObjectID & ", " & oEntity.ObjTypeID & ", " & BytesToString(oEntity.EntityName) & ", " & lKilledByID & ", OwnerID: " & oEntity.Owner.ObjectID)

        If oEntity.ObjTypeID = ObjectType.eUnit Then
            oTmpDef = CType(oEntity, Unit).EntityDef
        Else : oTmpDef = CType(oEntity, Facility).EntityDef
        End If
        If oTmpDef Is Nothing Then Return

        If bCreateEntityDefMinerals = True Then
            For X As Int32 = 0 To oTmpDef.lEntityDefMineralUB
                Dim yCacheType As Byte = 0
                If (oTmpDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
                    yCacheType = yCacheType Or MineralCacheType.eGround
                ElseIf (oTmpDef.yChassisType And ChassisType.eAtmospheric) <> 0 OrElse (oTmpDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                    yCacheType = yCacheType Or MineralCacheType.eFlying
                ElseIf (oTmpDef.yChassisType And ChassisType.eNavalBased) <> 0 Then
                    yCacheType = yCacheType Or MineralCacheType.eNaval
                End If

                Dim lCacheIdx As Int32 = AddMineralCache(CType(oEntity.ParentObject, Epica_GUID).ObjectID, CType(oEntity.ParentObject, Epica_GUID).ObjTypeID, _
                  yCacheType, oTmpDef.EntityDefMinerals(X).lQuantity, _
                  oTmpDef.EntityDefMinerals(X).lQuantity, oEntity.LocX, oEntity.LocZ, oTmpDef.EntityDefMinerals(X).oMineral)

                If lCacheIdx > -1 Then
                    AddToQueue(glCurrentCycle + 9000, QueueItemType.eRemoveMineralCache, goMineralCache(lCacheIdx).ObjectID, goMineralCache(lCacheIdx).ObjTypeID, -1, -1, 0, 0, 0, 0)
                    If iTemp = ObjectType.ePlanet Then
                        CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(goMineralCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand))
                    ElseIf iTemp = ObjectType.eSolarSystem Then
                        CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(goMineralCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand))
                    End If
                End If

            Next X
        End If

        If oEntity.ObjTypeID = ObjectType.eFacility Then

            With CType(oEntity, Facility)

                Dim sParentName As String = ""
                If .ParentObject Is Nothing = False Then
                    Dim iParentTypeID As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID
                    If iParentTypeID = ObjectType.ePlanet Then
                        sParentName = BytesToString(CType(.ParentObject, Planet).PlanetName)
                    ElseIf iParentTypeID = ObjectType.eSolarSystem Then
                        sParentName = BytesToString(CType(.ParentObject, SolarSystem).SystemName)
                    End If
                End If

                If .oMiningBid Is Nothing = False AndAlso CType(oEntity.ParentObject, Epica_GUID).ObjectID < 500000000 Then
                    .oMiningBid.SendBiddersDeathAlert(-1, False)
                End If

                .RemoveMe()

                .Owner.CreateAndSendPlayerAlert(PlayerAlertType.eFacilityLost, .ObjectID, .ObjTypeID, CType(.ParentObject, Epica_GUID).ObjectID, CType(.ParentObject, Epica_GUID).ObjTypeID, -1, BytesToString(.EntityName), .LocX, .LocZ, sParentName)
                If .Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID AndAlso (.yProductionType And ProductionType.eProduction) <> 0 Then
                    'If they kill the pirate factory on Tutorial Phase 1...
                    If CType(oEntity.ParentObject, Epica_GUID).ObjectID > 500000000 Then
                        AddToQueue(glCurrentCycle + 90, QueueItemType.ePirateSelfDestruct, .lPirate_For_PlayerID, ObjectType.eFacility, gl_HARDCODE_PIRATE_PLAYER_ID, CType(.ParentObject, Epica_GUID).ObjectID, 0, 0, 0, 0)
                        AddToQueue(glCurrentCycle + 150, QueueItemType.ePirateSelfDestruct, .lPirate_For_PlayerID, ObjectType.eUnit, gl_HARDCODE_PIRATE_PLAYER_ID, CType(.ParentObject, Epica_GUID).ObjectID, 0, 0, 0, 0)

                        Dim yGNS(73) As Byte
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2
                        yGNS(lPos) = NewsItemType.ePirateElimination : lPos += 1
                        CType(oEntity.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                        If iTemp = ObjectType.ePlanet Then
                            CType(oEntity.ParentObject, Planet).PlanetName.CopyTo(yGNS, lPos) : lPos += 20
                        Else
                            StringToBytes("a planet").CopyTo(yGNS, lPos) : lPos += 20
                        End If
                        System.BitConverter.GetBytes(.lPirate_For_PlayerID).CopyTo(yGNS, lPos) : lPos += 4

                        If .lPirate_For_PlayerID > 0 Then
                            Dim oTmpPlayer As Player = GetEpicaPlayer(.lPirate_For_PlayerID)
                            If oTmpPlayer Is Nothing = False Then
                                oTmpPlayer.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                                yGNS(lPos) = oTmpPlayer.yGender : lPos += 1
                                oTmpPlayer.EmpireName.CopyTo(yGNS, lPos) : lPos += 20
                            Else
                                StringToBytes("Citizens").CopyTo(yGNS, lPos) : lPos += 20
                                yGNS(lPos) = 1 : lPos += 1
                                StringToBytes("civilized people").CopyTo(yGNS, lPos) : lPos += 20
                            End If
                        Else
                            StringToBytes("Citizens").CopyTo(yGNS, lPos) : lPos += 20
                            yGNS(lPos) = 1 : lPos += 1
                            StringToBytes("civilized people").CopyTo(yGNS, lPos) : lPos += 20
                        End If

                        goMsgSys.SendToEmailSrvr(yGNS)
                    End If
                End If
            End With

            Dim lIdx As Int32 = -1
            Dim lID As Int32 = oEntity.ObjectID
            For X As Int32 = 0 To glFacilityUB
                If glFacilityIdx(X) = lID Then
                    lIdx = X
                    Exit For
                End If
            Next X
            If lIdx <> -1 Then
                glFacilityIdx(lIdx) = -1
                Dim oFac As Facility = goFacility(lIdx)
                If oFac Is Nothing = False Then
                    RemoveLookupFacility(oFac.ObjectID, oFac.ObjTypeID)
                    oFac.DeleteEntity(lIdx)
                    goFacility(lIdx) = Nothing
                End If
            End If

        ElseIf oEntity.ObjTypeID = ObjectType.eUnit Then
            With CType(oEntity, Unit)
                Dim sParentName As String = ""
                If .ParentObject Is Nothing = False Then
                    Dim iParentTypeID As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID
                    If iParentTypeID = ObjectType.ePlanet Then
                        sParentName = BytesToString(CType(.ParentObject, Planet).PlanetName)
                    ElseIf iParentTypeID = ObjectType.eSolarSystem Then
                        sParentName = BytesToString(CType(.ParentObject, SolarSystem).SystemName)
                    End If
                End If

                .Owner.CreateAndSendPlayerAlert(PlayerAlertType.eUnitLost, .ObjectID, .ObjTypeID, CType(.ParentObject, Epica_GUID).ObjectID, CType(.ParentObject, Epica_GUID).ObjTypeID, lKilledByID, BytesToString(.EntityName), .LocX, .LocZ, sParentName)
            End With

            If oEntity.ParentObject Is Nothing = False AndAlso (CType(oEntity.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eUnit OrElse CType(oEntity.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility) Then
                With CType(oEntity.ParentObject, Epica_Entity)
                    For X As Int32 = 0 To .lHangarUB
                        If .lHangarIdx(X) > -1 AndAlso .oHangarContents(X) Is Nothing = False Then
                            If .oHangarContents(X).ObjectID = oEntity.ObjectID AndAlso .oHangarContents(X).ObjTypeID = oEntity.ObjTypeID Then
                                .lHangarIdx(X) = -1
                                .oHangarContents(X) = Nothing
                            End If
                        End If
                    Next X
                End With
            End If

            Dim lIdx As Int32 = -1
            Dim lID As Int32 = oEntity.ObjectID
            For X As Int32 = 0 To glUnitUB
                If glUnitIdx(X) = lID Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If goUnit(lIdx) Is Nothing = False Then goUnit(lIdx).DeleteEntity(lIdx)
            glUnitIdx(lIdx) = -1
            goUnit(lIdx) = Nothing
        End If

    End Sub
 
    Public Function Distance(ByVal fX1 As Single, ByVal fY1 As Single, ByVal fX2 As Single, ByVal fY2 As Single) As Single
        Dim fDX As Single = fX2 - fX1
        Dim fDY As Single = fY2 - fY1
        Return CSng(Math.Sqrt((fDX * fDX) + (fDY * fDY)))
    End Function

    Public Function GetEpicaObjectName(ByVal iObjTypeID As Int16, ByVal oObj As Object) As Byte()
        Select Case iObjTypeID
            Case ObjectType.eAgent
                Return CType(oObj, Agent).AgentName
            Case ObjectType.eAlloyTech
                Return CType(oObj, AlloyTech).AlloyName
            Case ObjectType.eArmorTech
                Return CType(oObj, ArmorTech).ArmorName
            Case ObjectType.eColony
                Return CType(oObj, Colony).ColonyName
            Case ObjectType.eCorporation
                Return CType(oObj, Corporation).CorporationName
            Case ObjectType.eEngineTech
                Return CType(oObj, EngineTech).EngineName
            Case ObjectType.eFacility
                Return CType(oObj, Facility).EntityName
            Case ObjectType.eFacilityDef
                Return CType(oObj, FacilityDef).DefName
            Case ObjectType.eGalaxy
                Return CType(oObj, Galaxy).GalaxyName
            Case ObjectType.eGNS
                Return CType(oObj, GNS).GNS_Name
            Case ObjectType.eGoal
                Return CType(oObj, Goal).GoalName
            Case ObjectType.eHullTech
                Return CType(oObj, HullTech).HullName
            Case ObjectType.eMineral
				Return CType(oObj, Mineral).MineralName
			Case ObjectType.eMission
				Return CType(oObj, Mission).MissionName
            Case ObjectType.eNebula
                Return CType(oObj, Nebula).NebulaName
            Case ObjectType.ePlanet
                Return CType(oObj, Planet).PlanetName
            Case ObjectType.ePlayer
                Return CType(oObj, Player).PlayerName
            Case ObjectType.ePrototype
                Return CType(oObj, Prototype).PrototypeName
            Case ObjectType.eRadarTech
                Return CType(oObj, RadarTech).RadarName
				'Case ObjectType.eSenate
				'    Return CType(oObj, Senate).SenateName
            Case ObjectType.eShieldTech
                Return CType(oObj, ShieldTech).ShieldName
            Case ObjectType.eSkill
                Return CType(oObj, Skill).SkillName
            Case ObjectType.eSolarSystem
                Return CType(oObj, SolarSystem).SystemName
            Case ObjectType.eUnit
                Return CType(oObj, Unit).EntityName
            Case ObjectType.eUnitDef
                Return CType(oObj, Epica_Entity_Def).DefName
            Case ObjectType.eWeaponDef
                Return CType(oObj, WeaponDef).WeaponName
            Case ObjectType.eWeaponTech
                Return CType(oObj, BaseWeaponTech).WeaponName
            Case ObjectType.eMineralProperty
                Return CType(oObj, MineralProperty).MineralPropertyName
            Case ObjectType.eGuild
				Return CType(oObj, Guild).yGuildName
            Case ObjectType.eComponentCache
                Try
                    Return CType(oObj, Epica_Tech).GetTechName()
                Catch
                    Return System.Text.ASCIIEncoding.ASCII.GetBytes("No Name")
                End Try
            Case Else
                Return System.Text.ASCIIEncoding.ASCII.GetBytes("No Name")
        End Select
    End Function

    Public Function AddPlayerFromCommand(ByVal sMsg As String) As Int32
        'Msg structure is:
        '/addplayer <DisplayName>, <UserName>, <Password>
		Dim sTemp As String = sMsg.Substring(10).Trim
        Dim sValues() As String = Split(sTemp, ",")

		If sValues Is Nothing OrElse sValues.GetUpperBound(0) <> 2 Then Return -2

		For X As Int32 = 0 To 2
			If sValues(X).Trim.Length > 20 Then Return -2
		Next X

        Dim oPlayer As Player = New Player()

        With oPlayer
            ReDim .PlayerName(19)
            StringToBytes(sValues(0).Trim).CopyTo(.PlayerName, 0)
            ReDim .EmpireName(19)
            StringToBytes(sValues(0).Trim).CopyTo(.EmpireName, 0)
            ReDim .RaceName(19)
            StringToBytes(sValues(0).Trim).CopyTo(.RaceName, 0)
            ReDim .PlayerUserName(19)
            StringToBytes(sValues(1).Trim.ToUpper).CopyTo(.PlayerUserName, 0)
            ReDim .PlayerPassword(19)
			StringToBytes(sValues(2).Trim.ToUpper).CopyTo(.PlayerPassword, 0)

			Dim sUserName As String = sValues(1).Trim.ToUpper
			Dim lCurUB As Int32 = Math.Min(glPlayerUB, glPlayerIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If glPlayerIdx(X) <> -1 Then
					Dim sTestUserName As String = BytesToString(goPlayer(X).PlayerUserName).Trim.ToUpper
					If sUserName = sTestUserName Then
						Return -3
					End If
				End If
			Next X

			'.oSenate = Nothing
            .CommEncryptLevel = 0
            .EmpireTaxRate = 0
            .lLastViewedEnvir = 0
            .iLastViewedEnvirType = 0
            .blCredits = 0
            .AccountStatus = 1
            .BaseMorale = 100
            .lStartedEnvirID = 0
            .iStartedEnvirTypeID = 0
            .TotalPlayTime = 0
            .lHangarManMult = 1
            .ObjTypeID = ObjectType.ePlayer
            .AccountStatus = 1
        End With

        If oPlayer.SaveObject(False) = True Then
            Dim lIdx As Int32 = -1
            'PQL MODIFICATION

            ' JC removed for the PQL
            'For X As Int32 = 0 To glPlayerUB
            '    If glPlayerIdx(X) = -1 Then
            '        lIdx = X
            '        Exit For
            '    End If
            'Next X

            'If lIdx = -1 Then
            '    ReDim Preserve goPlayer(glPlayerUB + 1)
            '    ReDim Preserve glPlayerIdx(glPlayerUB + 1)
            '    glPlayerIdx(glPlayerUB + 1) = -1
            '    lIdx = glPlayerUB + 1
            '    glPlayerUB += 1
            'End If
            'goPlayer(lIdx) = oPlayer
            'glPlayerIdx(lIdx) = oPlayer.ObjectID
            'oPlayer.ServerIndex = lIdx
            'goPlayer(lIdx).DataChanged()

            ' in order to not redim on release day, we do 500 at a time (kinda like an extent in DB)
            CheckPlayerExtent(oPlayer.ObjectID)

            goPlayer(oPlayer.ObjectID) = oPlayer
            glPlayerIdx(oPlayer.ObjectID) = oPlayer.ObjectID
            oPlayer.ServerIndex = oPlayer.ObjectID
            goPlayer(oPlayer.ObjectID).DataChanged()
            ' end PQL Modification

			Dim yMsg(7) As Byte	'= goMsgSys.GetAddObjectMessage(goPlayer(lIdx), GlobalMessageCode.eAddObjectCommand)
			System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
			System.BitConverter.GetBytes(ObjectType.ePlayer).CopyTo(yMsg, 6)
			'goMsgSys.BroadcastToDomains(yMsg)
			goMsgSys.SendMsgToOperator(yMsg)

			Return goPlayer(lIdx).ObjectID
		Else : Return -1
		End If


    End Function

    Public Sub CheckPlayerExtent(ByVal objectID As Int32)
        ' in order to not redim on release day, we do 500 at a time (kinda like an extent in DB)
        SyncLock goPlayer
            If glPlayerUB < objectID Then
                Dim origPlayerUB As Int32 = glPlayerUB
                If glPlayerUB + 500 < objectID Then
                    'things are moving TOO FAST or there is a mess somewhere need to report this problem.
                    'LogEvent(LogEventType.Warning, "Extent of oPlayer is moving faster than we can redim")
                    glPlayerUB = objectID + 500
                Else
                    glPlayerUB = glPlayerUB + 500
                End If
                ReDim Preserve goPlayer(glPlayerUB)
                ReDim Preserve glPlayerIdx(glPlayerUB)
                Dim i As Int32
                For i = origPlayerUB + 1 To glPlayerUB
                    glPlayerIdx(i) = -1
                Next

            End If
        End SyncLock

    End Sub

	Public Sub InitializePlayer(ByRef oPlayer As Player, ByVal lOriginalPlanetID As Int32)
		Dim X As Int32

		Dim lStartX As Int32 = 0
		Dim lStartZ As Int32 = 0
		Dim lTemp As Int32 = 0
		Dim yMsg() As Byte

        'oPlayer.lLastWPUpkeepTime = GetDateAsNumber(Now)
        'oPlayer.dtLastWPUpkeep = Now

		If lOriginalPlanetID < 1 Then
			'Ok, first thing, let's create our mineral propertys 
            oPlayer.AddKnownProperty(eMinPropID.BoilingPoint, 2, False, True)
            oPlayer.AddKnownProperty(eMinPropID.ChemicalReactance, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Combustiveness, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Compressibility, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Density, 2, False, True)
            oPlayer.AddKnownProperty(eMinPropID.ElectricalResist, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Hardness, 2, False, True)
            oPlayer.AddKnownProperty(eMinPropID.MagneticProduction, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.MagneticReaction, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Malleable, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.MeltingPoint, 2, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Psych, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Quantum, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Reflection, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Refraction, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.SuperconductivePoint, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.TemperatureSensitivity, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.ThermalConductance, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.ThermalExpansion, 1, False, True)
            oPlayer.AddKnownProperty(eMinPropID.Toxicity, 1, False, True)


			'Now, create the special link for the Schools of Science
            oPlayer.oSpecials.PerformLinkTest()
		End If
		If lOriginalPlanetID = Int32.MinValue Then Return

		'NOTE: DO NOT PUT ANY PLAYER INITIALIZATION STUFF BELOW THIS LINE EXCEPT FOR START PLANET RELATED
        '================================================================================================
        oPlayer.bSpawnAtSameLoc = False

        If oPlayer.AccountStatus = AccountStatusType.eTrialAccount AndAlso oPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase Then
            'Ok, send to the player the respawn list msg with -1, this will open the window for respawn selection... that window will send a msg
            '  back here to request the respawn list
            ReDim yMsg(5)
            System.BitConverter.GetBytes(GlobalMessageCode.eRespawnSelection).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 2)
            oPlayer.SendPlayerMessage(yMsg, False, 0)
            Return
        End If

        If oPlayer.yRespawnWithGuild <> 0 AndAlso oPlayer.bSpawnAtSameLoc = False Then
            'ok, player wants to respawn with their guild, check the guild
            If oPlayer.oGuild Is Nothing = False Then

                'Ok, new method for this to work... we determine the system where most members have colonies
                Dim bSpawnSystem As Boolean = False
                If (oPlayer.ySpawnSystemSetting And eySpawnSettings.eSpawnProcessOver) = 0 Then
                    If (oPlayer.ySpawnSystemSetting And (eySpawnSettings.eSpawnAccept Or eySpawnSettings.eOfferedSpawn)) = (eySpawnSettings.eSpawnAccept Or eySpawnSettings.eOfferedSpawn) Then
                        bSpawnSystem = oPlayer.AccountStatus = AccountStatusType.eActiveAccount
                        oPlayer.ySpawnSystemSetting = oPlayer.ySpawnSystemSetting Or eySpawnSettings.eSpawnProcessOver
                    End If
                End If
                Dim lTargetSystemID As Int32 = oPlayer.oGuild.GetNextHighestSystemByColonyCount(-1, bSpawnSystem)

                If oPlayer.AccountStatus = AccountStatusType.eTrialAccount Then
                    lTargetSystemID = 36 'NOTE: Hardcoded 36 Aurelium system
                ElseIf lTargetSystemID = 36 AndAlso oPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase Then
                    lTargetSystemID = -1
                End If

                If lTargetSystemID <> -1 Then

                    Dim oSystem As SolarSystem = GetEpicaSystem(lTargetSystemID)

                    'ok, we tried the guild bank work... let's get the system
                    Dim lMaxCnt As Int32 = 0
                    Dim bTrySpawnLoc As Boolean = (oPlayer.yRespawnWithGuild And 128) = 0
                    Dim lIdx As Int32 = 0
                    If bTrySpawnLoc = True Then
                        lIdx = CInt(oPlayer.yRespawnWithGuild) - 1
                    Else
                        lIdx = CInt(oPlayer.yRespawnWithGuild Xor 128) - 1
                    End If

                    If bTrySpawnLoc = True Then

                        If lTargetSystemID = 36 Then

                            For X = lIdx To oSystem.mlPlanetUB
                                lIdx = oSystem.GetPlanetIdx(X)
                                If lIdx <> -1 Then
                                    Dim oPlanet As Planet = goPlanet(lIdx)
                                    If oPlanet Is Nothing Then Continue For

                                    If oPlanet.GetGuildColonyCount(oPlayer.oGuild.ObjectID) > 0 Then
                                        oPlayer.yRespawnWithGuild = CByte(X + 2)    'mark our spot for next try

                                        ReDim yMsg(11)
                                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yMsg, 0)
                                        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                                        oPlanet.GetGUIDAsString.CopyTo(yMsg, 6)
                                        oPlanet.oDomain.DomainSocket.SendData(yMsg)
                                        Return
                                    End If
                                End If
                            Next X
                            lIdx = 0
                        Else
                            For X = lIdx To oSystem.mlPlanetUB
                                lIdx = oSystem.GetPlanetIdx(X)
                                If lIdx <> -1 Then
                                    Dim oPlanet As Planet = goPlanet(lIdx)
                                    If oPlanet Is Nothing Then Continue For
                                    lMaxCnt = Planet.GetPlanetMaxPlayerSpawnCnt(oPlanet.PlanetSizeID, oSystem.SystemType)
                                    If oPlanet.PlayerSpawns >= lMaxCnt Then Continue For

                                    oPlayer.yRespawnWithGuild = CByte(X + 2)    'mark our spot for next try

                                    ReDim yMsg(11)
                                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yMsg, 0)
                                    System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                                    oPlanet.GetGUIDAsString.CopyTo(yMsg, 6)
                                    oPlanet.oDomain.DomainSocket.SendData(yMsg)
                                    Return
                                End If
                            Next X
                            lIdx = 0
                        End If
                    End If

                    'Ok, we are here... try non spawn loc way
                    For X = lIdx To oSystem.mlPlanetUB
                        lIdx = oSystem.GetPlanetIdx(X)
                        If lIdx <> -1 Then
                            Dim oPlanet As Planet = goPlanet(lIdx)
                            If oPlanet Is Nothing Then Continue For

                            oPlayer.yRespawnWithGuild = CByte((X + 2) Or 128)    'mark our spot for next try

                            ReDim yMsg(11)
                            System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yMsg, 0)
                            System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                            oPlanet.GetGUIDAsString.CopyTo(yMsg, 6)
                            oPlanet.oDomain.DomainSocket.SendData(yMsg)
                            Return
                        End If
                    Next X

                    'if we are here, respawn with guild failed
                    oPlayer.yRespawnWithGuild = 0
                    Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Unable to spawn near guild. That region of space is too full.", "Guild Spawn Failed", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                    oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If


            End If
        End If

        'TODO: Instead of a generic Player Count... perhaps we should do an age cnt... how long a player has existed there...
        'TODO: This is just plain UGLY!

        Dim lSystemType As Int32 = 0        'for normal spawn systems
        Dim lOriginalSystem As Int32 = -1
        If oPlayer.yPlayerPhase = eyPlayerPhase.eSecondPhase Then 'OrElse oPlayer.ObjectID = 221 OrElse oPlayer.ObjectID = 131 Then
            lSystemType = 255

            oPlayer.blCredits = 3100000000
            oPlayer.DeathBudgetFundsRemaining = 100000000
            oPlayer.DeathBudgetEndTime = glCurrentCycle + 2592000
            oPlayer.DeathBudgetBalance = 0

            Dim lFirstValidIdx As Int32 = -1
            For X = 0 To glPlanetUB
                If glPlanetIdx(X) <> -1 AndAlso goPlanet(X).InMyDomain = True AndAlso goPlanet(X) Is Nothing = False AndAlso goPlanet(X).ParentSystem Is Nothing = False AndAlso goPlanet(X).ParentSystem.SystemType = lSystemType AndAlso goPlanet(X).SpawnLocked = False Then
                    If lFirstValidIdx = -1 Then lFirstValidIdx = X
                    If X > AureliusAI.lSpawnSystemLastPlanetPlacement Then
                        AureliusAI.lSpawnSystemLastPlanetPlacement = X
                        'Ok, this planet has room! So, here we go...
                        ReDim yMsg(11)
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yMsg, 0)
                        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                        goPlanet(X).GetGUIDAsString.CopyTo(yMsg, 6)
                        goPlanet(X).oDomain.DomainSocket.SendData(yMsg)
                        Return
                    End If
                End If
            Next X

            If lFirstValidIdx <> -1 Then
                AureliusAI.lSpawnSystemLastPlanetPlacement = lFirstValidIdx
                'Ok, this planet has room! So, here we go...
                ReDim yMsg(11)
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                goPlanet(lFirstValidIdx).GetGUIDAsString.CopyTo(yMsg, 6)
                goPlanet(lFirstValidIdx).oDomain.DomainSocket.SendData(yMsg)
                Return
            End If

            'default, spawn them in a normal system
            lSystemType = 0
            LogEvent(LogEventType.Warning, "Had to spawn a tutorial 2 player in the normal game.")
            oPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase
        Else

            If oPlayer.lStartedEnvirID > 0 Then lSystemType = 2
            If oPlayer.lIronCurtainSystem = 36 OrElse oPlayer.lIronCurtainSystem = -1 Then lSystemType = 0
            If lSystemType = 2 Then
                'Ok, verify that the player in fact needs to be placed in a non spawn
                If oPlayer.lIronCurtainPlanet > 0 Then
                    If oPlayer.lIronCurtainPlanet > 500000000 Then
                        lSystemType = 0
                    Else
                        Dim oPlanet As Planet = GetEpicaPlanet(oPlayer.lIronCurtainPlanet)
                        If oPlanet Is Nothing = False Then
                            If oPlanet.ParentSystem Is Nothing = False Then
                                If oPlanet.ParentSystem.SystemType = 255 Then lSystemType = 0
                            End If
                        End If
                    End If
                Else
                    lSystemType = 0
                End If
            End If
            If lOriginalPlanetID > 500000000 Then lOriginalPlanetID = -1
            If lOriginalPlanetID > 0 Then
                Dim oTempPlanet As Planet = GetEpicaPlanet(lOriginalPlanetID)
                If oTempPlanet Is Nothing = False Then lOriginalSystem = oTempPlanet.ParentSystem.ObjectID


                If oPlayer.bSpawnAtSameLoc = True Then
                    If oPlayer.bDeclaredWarOn = True Then
                        For X = 0 To oTempPlanet.ParentSystem.mlPlanetUB
                            Dim lTempPlanetIdx As Int32 = oTempPlanet.ParentSystem.GetPlanetIdx(X)
                            If lTempPlanetIdx > -1 AndAlso lTempPlanetIdx <= glPlanetUB Then
                                Dim oPlanet As Planet = goPlanet(lTempPlanetIdx)
                                Dim lMaxCnt As Int32 = 0
                                lMaxCnt = Planet.GetPlanetMaxPlayerSpawnCnt(oPlanet.PlanetSizeID, oPlanet.ParentSystem.SystemType)

                                If oPlanet.PlayerSpawns >= lMaxCnt Then Continue For

                                ReDim yMsg(11)
                                System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yMsg, 0)
                                System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                                oPlanet.GetGUIDAsString.CopyTo(yMsg, 6)
                                oPlanet.oDomain.DomainSocket.SendData(yMsg)
                                Return
                            End If
                        Next X
                        oPlayer.bSpawnAtSameLoc = False
                    Else
                        ReDim yMsg(19)
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yMsg, 0)
                        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                        oTempPlanet.GetGUIDAsString.CopyTo(yMsg, 6)
                        System.BitConverter.GetBytes(oPlayer.lStartLocX).CopyTo(yMsg, 12)
                        System.BitConverter.GetBytes(oPlayer.lStartLocZ).CopyTo(yMsg, 16)
                        goMsgSys.HandleRequestPlayerStartLoc(yMsg)
                        Return
                    End If

                End If

                oTempPlanet = Nothing
            End If
        End If

        If lSystemType = SolarSystem.elSystemType.RespawnSystem Then
            'Ok, send to the player the respawn list msg with -1, this will open the window for respawn selection... that window will send a msg
            '  back here to request the respawn list
            ReDim yMsg(5)
            System.BitConverter.GetBytes(GlobalMessageCode.eRespawnSelection).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 2)
            oPlayer.SendPlayerMessage(yMsg, False, 0)
            Return
        End If

        'If oPlayer.ObjectID = 1 OrElse oPlayer.ObjectID = 2 OrElse oPlayer.ObjectID = 6 OrElse oPlayer.ObjectID = 3253 OrElse oPlayer.ObjectID = 2067 OrElse oPlayer.ObjectID = 2076 OrElse oPlayer.ObjectID = 7570 OrElse oPlayer.ObjectID = 1780 Then
        '    lSystemType = 128
        '    'ElseIf oPlayer.ObjectID = 131 OrElse oPlayer.ObjectID = 221 Then
        '    '    lSystemType = 255
        'End If
        'If oPlayer.ObjectID = 221 OrElse oPlayer.ObjectID = 131 Then lSystemType = 2

        Dim lTotalSpawnLocsAvail As Int32 = 0
        Dim lTotalRespawnLocsAvail As Int32 = 0
        For X = 0 To glPlanetUB
            If glPlanetIdx(X) <> -1 AndAlso goPlanet(X) Is Nothing = False AndAlso goPlanet(X).ObjectID < 500000000 AndAlso goPlanet(X).InMyDomain = True AndAlso goPlanet(X).ParentSystem Is Nothing = False Then
                Dim lMaxCnt As Int32 = Planet.GetPlanetMaxPlayerSpawnCnt(goPlanet(X).PlanetSizeID, goPlanet(X).ParentSystem.SystemType)
                If goPlanet(X).PlayerSpawns >= lMaxCnt OrElse goPlanet(X).SpawnLocked = True Then Continue For

                'ok, now, if over half the spawns for this system are used
                Dim lTotalSpawnsForSystem As Int32 = 0
                Dim lTotalSpawnsInSystem As Int32 = 0
                For Y As Int32 = 0 To goPlanet(X).ParentSystem.mlPlanetUB
                    Dim lTmp As Int32 = goPlanet(X).ParentSystem.GetPlanetIdx(Y)
                    If lTmp > -1 AndAlso lTmp <= glPlanetUB Then
                        lTotalSpawnsForSystem += Planet.GetPlanetMaxPlayerSpawnCnt(goPlanet(lTmp).PlanetSizeID, goPlanet(X).ParentSystem.SystemType)
                        lTotalSpawnsInSystem += goPlanet(lTmp).PlayerSpawns
                    End If
                Next Y

                If lTotalSpawnsInSystem > lTotalSpawnsForSystem \ 2 Then
                    'now, check if the wormhole is breached
                    Dim bBreached As Boolean = False
                    For Y As Int32 = 0 To goPlanet(X).ParentSystem.mlWormholeUB
                        If goPlanet(X).ParentSystem.moWormholes(Y) Is Nothing = False Then
                            With goPlanet(X).ParentSystem.moWormholes(Y)
                                If .System1 Is Nothing = False AndAlso .System1.ObjectID = goPlanet(X).ParentSystem.ObjectID AndAlso (.WormholeFlags And elWormholeFlag.eSystem2Detectable) <> 0 Then
                                    goPlanet(X).SpawnLocked = True
                                    bBreached = True
                                    Exit For
                                End If
                            End With
                        End If
                    Next Y
                    If bBreached = True Then Continue For
                End If

                If goPlanet(X).ParentSystem.SystemType = SolarSystem.elSystemType.SpawnSystem Then
                    lTotalSpawnLocsAvail += (lMaxCnt - goPlanet(X).PlayerSpawns)
                ElseIf goPlanet(X).ParentSystem.SystemType = SolarSystem.elSystemType.RespawnSystem Then
                    lTotalRespawnLocsAvail += (lMaxCnt - goPlanet(X).PlayerSpawns)
                End If
            End If
        Next X

        lTotalRespawnLocsAvail = 10000      'op is never to generate respawn systems

        'Now, inform our operator of the number of spawns i have available
        Dim yOpMsg(9) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yOpMsg, 0)
        System.BitConverter.GetBytes(lTotalSpawnLocsAvail).CopyTo(yOpMsg, 2)
        System.BitConverter.GetBytes(lTotalRespawnLocsAvail).CopyTo(yOpMsg, 6)
        goMsgSys.SendMsgToOperator(yOpMsg)

        For X = 0 To glPlanetUB
            If glPlanetIdx(X) <> lOriginalPlanetID AndAlso goPlanet(X) Is Nothing = False AndAlso goPlanet(X).ObjectID < 500000000 AndAlso goPlanet(X).InMyDomain = True AndAlso goPlanet(X).ParentSystem.SystemType = lSystemType AndAlso goPlanet(X).SpawnLocked = False AndAlso goPlanet(X).PlanetTypeID <> PlanetType.eWaterWorld AndAlso goPlanet(X).ParentSystem.ObjectID <> lOriginalSystem Then

                If goPlanet(X).ParentSystem.SystemType = 128 OrElse goPlanet(X).ParentSystem.SystemType = 255 Then goPlanet(X).PlayerSpawns = 0
                If lSystemType <> 128 AndAlso goPlanet(X).ParentSystem.SystemType = 128 Then Continue For

                'Now, check our colonist count
                Dim blTotalPop As Int64 = 0
                For Y As Int32 = 0 To goPlanet(X).lColonysHereUB
                    If goPlanet(X).lColonysHereIdx(Y) <> -1 Then
                        If glColonyIdx(goPlanet(X).lColonysHereIdx(Y)) <> -1 Then
                            Dim oTmpColony As Colony = goColony(goPlanet(X).lColonysHereIdx(Y))
                            If oTmpColony Is Nothing = False AndAlso oTmpColony.ParentObject Is Nothing = False Then
                                If CType(oTmpColony.ParentObject, Epica_GUID).ObjectID = goPlanet(X).ObjectID Then blTotalPop += oTmpColony.Population
                            End If
                        End If
                    End If
                Next Y
                'TODO: Remarked this out because its bullshit
                'If blTotalPop > 3800000 Then
                '    If oPlayer.DeathBudgetBalance < 50000000 Then
                '        Continue For
                '    End If
                'End If

                Dim lMaxCnt As Int32 = Planet.GetPlanetMaxPlayerSpawnCnt(goPlanet(X).PlanetSizeID, goPlanet(X).ParentSystem.SystemType)

                If goPlanet(X).PlayerSpawns >= lMaxCnt AndAlso goPlanet(X).ParentSystem.SystemType <> 128 Then Continue For

                'Now, go through the players to find who started here
                'For Y As Int32 = 0 To glPlayerUB
                '    If glPlayerIdx(Y) <> -1 AndAlso goPlayer(Y).iStartedEnvirTypeID = ObjectType.ePlanet AndAlso goPlayer(Y).lStartedEnvirID = glPlanetIdx(X) Then
                '        lMaxCnt -= 1

                '        If lMaxCnt = 0 Then Exit For
                '    End If
                'Next Y

                If lMaxCnt <> 0 OrElse goPlanet(X).ParentSystem.SystemType = 128 Then

                    If goPlanet(X).oDomain Is Nothing Then
                        LogEvent(LogEventType.Warning, "InitializePlayer, Planet's domain is nothing, entering contingency.")
                        If goPlanet(X).ParentSystem.oDomain Is Nothing Then
                            LogEvent(LogEventType.Warning, "InitializePlayer, Planet's Parent's domain is nothing, rechecking primary.")
                            ReDim yMsg(1)
                            System.BitConverter.GetBytes(GlobalMessageCode.eCheckPrimaryReady).CopyTo(yMsg, 0)
                            System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yMsg, 2)
                            goMsgSys.SendMsgToOperator(yMsg)
                            Return
                        Else
                            LogEvent(LogEventType.Warning, "InitializePlayer, Planet's parent domain is set, using it.")
                            goPlanet(X).oDomain = goPlanet(X).ParentSystem.oDomain
                        End If
                    End If

                    'Ok, this planet has room! So, here we go...
                    ReDim yMsg(11)
                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                    goPlanet(X).GetGUIDAsString.CopyTo(yMsg, 6)
                    goPlanet(X).oDomain.DomainSocket.SendData(yMsg)
                    Return
                End If
            End If
        Next X

        'Ok, if we're here, send 0,0
        If lSystemType <> 128 AndAlso lSystemType <> 255 Then
            ReDim yOpMsg(9)
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yOpMsg, 0)
            System.BitConverter.GetBytes(0I).CopyTo(yOpMsg, 2)
            System.BitConverter.GetBytes(0I).CopyTo(yOpMsg, 6)
            goMsgSys.SendMsgToOperator(yOpMsg)
        End If

        'Ok, if we are here, we need to do it the old way.......
        Dim lP_PCnt() As Int32
        Dim lSmallestIdx As Int32
        Dim lSmallestVal As Int32 = Int32.MaxValue

        ReDim lP_PCnt(glPlanetUB)
        LogEvent(LogEventType.Warning, "Needed to use the old way in InitializePlayer for finding a start planet!")
        For X = 0 To glPlanetUB
            If glPlanetIdx(X) = -1 OrElse goPlanet(X).InMyDomain = False OrElse goPlanet(X).ParentSystem Is Nothing = True OrElse goPlanet(X).ParentSystem.SystemType = 128 Then lP_PCnt(X) = 10000
        Next X

        'Let's find a place to start the player... first, go through all planets and count players who started there
        For X = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True AndAlso goPlayer(X).iStartedEnvirTypeID = ObjectType.ePlanet Then
                For Y As Int32 = 0 To glPlanetUB
                    If glPlanetIdx(Y) = goPlayer(X).lStartedEnvirID Then
                        lP_PCnt(Y) += 1
                        Exit For
                    End If
                Next Y
            End If
        Next X

        'Now, find the smallest of those
        lSmallestIdx = -1
        For X = 0 To lP_PCnt.GetUpperBound(0)
            If lP_PCnt(X) < lSmallestVal Then
                lSmallestIdx = X
                lSmallestVal = lP_PCnt(X)
            End If
        Next X
        If lSmallestIdx = -1 Then
            LogEvent(LogEventType.CriticalError, "Could not determine least inhabited planet, using index 0.")
            lSmallestIdx = 0
        Else
            goPlanet(lSmallestIdx).PlayerSpawns += 1        'we increment here to stop massive clumping when things break
        End If

        'Get the starting location... we do this by making a call to the region server controlling that object
        ReDim yMsg(11)
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerStartLoc).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
        goPlanet(lSmallestIdx).GetGUIDAsString.CopyTo(yMsg, 6)
        goPlanet(lSmallestIdx).oDomain.DomainSocket.SendData(yMsg)

    End Sub

	Public Sub CreatePirateFacility(ByVal lPlayerID As Int32, ByVal lInstanceID As Int32)
		Const l_HARDCODE_FACTORY As Int32 = 49
		Dim oDef As FacilityDef = GetEpicaFacilityDef(l_HARDCODE_FACTORY)
		Dim oPirate As Player = GetEpicaPlayer(gl_HARDCODE_PIRATE_PLAYER_ID)
		If oPirate Is Nothing Then Return

		Dim oNew As New Facility
		Dim oPlanet As Planet = GetEpicaPlanet(lInstanceID)
		If oPlanet Is Nothing Then Return

		With oNew
			.EntityDef = oDef
			.ObjTypeID = ObjectType.eFacility
			.AutoLaunch = True
			.bProducing = False

			.CurrentProduction = Nothing
			.CurrentSpeed = 0

			.CurrentStatus = 0
			For Y As Int32 = 0 To oDef.lSideCrits.Length - 1
				.CurrentStatus = .CurrentStatus Or oDef.lSideCrits(Y)
			Next Y

			oDef.DefName.CopyTo(.EntityName, 0)
			.ExpLevel = 0
			.Fuel_Cap = oDef.Fuel_Cap
			.iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage
			.iTargetingTactics = 0
			.LocAngle = 0
			.LocX = -15000
			.LocZ = -16600

			.Owner = oPirate
			.ParentObject = oPlanet

			.Q1_HP = oDef.Q1_MaxHP
			.Q2_HP = oDef.Q2_MaxHP
			.Q3_HP = oDef.Q3_MaxHP
			.Q4_HP = oDef.Q4_MaxHP
			.Shield_HP = oDef.Shield_MaxHP
            .Structure_HP = 6000 ' oDef.Structure_MaxHP
			.yProductionType = oDef.ProductionTypeID
			.DataChanged()
		End With

		Dim bResult As Boolean = oNew.SaveObject()
		If bResult = True Then
			'Now, find a suitable place...
			Dim lIdx As Int32 = -1
            'SyncLock goFacility
            '	For Y As Int32 = 0 To glFacilityUB
            '		If glFacilityIdx(Y) = -1 Then
            '			lIdx = Y
            '			Exit For
            '		End If
            '	Next Y

            '	If lIdx = -1 Then
            '		ReDim Preserve glFacilityIdx(glFacilityUB + 1)
            '		ReDim Preserve goFacility(glFacilityUB + 1)
            '		glFacilityUB += 1
            '		lIdx = glFacilityUB
            '	End If

            '	goFacility(lIdx) = CType(oNew, Facility)
            '	glFacilityIdx(lIdx) = oNew.ObjectID
            'End SyncLock
            lIdx = AddFacilityToGlobalArray(oNew)

			With oNew
				.CurrentStatus = goFacility(lIdx).CurrentStatus Or elUnitStatus.eFacilityPowered
				.ServerIndex = lIdx
				.DataChanged()
				.SaveObject()

				If (.yProductionType And ProductionType.eProduction) <> 0 Then
					.lPirate_Counter = 0
					.lPirate_Counter_Max = 4
					.lPirate_For_PlayerID = lPlayerID
					AddPirateProductionItem(oNew)
					AddEntityProducing(lIdx, ObjectType.eFacility, .ObjectID)
				End If

				oPlanet.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oNew, GlobalMessageCode.eAddObjectCommand_CE))
			End With
		End If

	End Sub

    'Public Function CreatePirateSpawn(ByVal lPlayerID As Int32, ByVal lLocX As Int32, ByVal lLocZ As Int32) As Byte()
    '	Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
    '	Dim oPirate As Player = GetEpicaPlayer(gl_HARDCODE_PIRATE_PLAYER_ID)
    '	If oPlayer Is Nothing = False Then
    '		oPlayer.PirateStartLocX = lLocX
    '		oPlayer.PirateStartLocZ = lLocZ
    '	End If
    '	If oPirate Is Nothing Then Return Nothing

    '	Dim oEnvir As Object = GetEpicaObject(oPlayer.lStartedEnvirID, oPlayer.iStartedEnvirTypeID)
    '	If oEnvir Is Nothing Then Return Nothing

    '	'TODO: Probably a better way to do this...
    '	Const l_HARDCODE_POWER_GEN As Int32 = 53
    '	Const l_HARDCODE_FACTORY As Int32 = 49
    '	Const l_HARDCODE_BARRACKS As Int32 = 47
    '	'Const l_HARDCODE_WAREHOUSE As Int32 = 6
    '	Const l_HARDCODE_TURRET As Int32 = 46

    '	Const l_HARDCODE_RAIDER As Int32 = 19
    '	Const l_HARDCODE_MARAUDER As Int32 = 20
    '	Const l_HARDCODE_PROWLER As Int32 = 21

    '	Dim lDefUB As Int32 = 7
    '	Dim oDefs(lDefUB) As Epica_Entity_Def
    '	Dim lDefCnt(lDefUB) As Int32

    '	'NOTE: If we were to ever change the pirate spawn, here is where we do it...
    '	oDefs(0) = GetEpicaFacilityDef(l_HARDCODE_POWER_GEN) : lDefCnt(0) = 2
    '	oDefs(1) = GetEpicaFacilityDef(l_HARDCODE_FACTORY) : lDefCnt(1) = 1
    '	oDefs(2) = GetEpicaFacilityDef(l_HARDCODE_BARRACKS) : lDefCnt(2) = 1
    '	'oDefs(3) = GetEpicaFacilityDef(l_HARDCODE_WAREHOUSE) : lDefCnt(3) = 1
    '	oDefs(4) = GetEpicaFacilityDef(l_HARDCODE_TURRET) : lDefCnt(4) = 6
    '	oDefs(5) = GetEpicaUnitDef(l_HARDCODE_RAIDER) : lDefCnt(5) = 0
    '	oDefs(6) = GetEpicaUnitDef(l_HARDCODE_MARAUDER) : lDefCnt(6) = 0
    '	oDefs(7) = GetEpicaUnitDef(l_HARDCODE_PROWLER) : lDefCnt(7) = 0

    '	For X As Int32 = 0 To 1
    '		Dim lTmpVal As Int32 = CInt(Rnd() * 100)
    '		If lTmpVal < 33 Then
    '			lDefCnt(5) += 1
    '		ElseIf lTmpVal < 66 Then
    '			lDefCnt(6) += 1
    '		Else : lDefCnt(7) += 1
    '		End If
    '	Next X

    '	'Ok, we need to create a message for the region server to place these spawns...
    '	'MsgCode, EnvirGuid, LocXZ, PlayerID, ItemCnt, Items (objGuid, DefGuid)
    '	Dim lCnt As Int32 = 13
    '	Dim yResp(23 + (lCnt * 12)) As Byte
    '	Dim lPos As Int32 = 0
    '	Dim lCntPos As Int32

    '	System.BitConverter.GetBytes(GlobalMessageCode.ePlacePirateAssets).CopyTo(yResp, lPos) : lPos += 2
    '	CType(oEnvir, Epica_GUID).GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
    '	System.BitConverter.GetBytes(lLocX).CopyTo(yResp, lPos) : lPos += 4
    '	System.BitConverter.GetBytes(lLocZ).CopyTo(yResp, lPos) : lPos += 4
    '	System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4

    '	'Skip Cnt for now...
    '	lCntPos = lPos
    '	lPos += 4

    '	For lDefIdx As Int32 = 0 To lDefUB
    '		For X As Int32 = 0 To lDefCnt(lDefIdx) - 1
    '			Dim oNew As Epica_Entity

    '			If oDefs(lDefIdx).ObjTypeID = ObjectType.eUnitDef Then
    '				oNew = New Unit()
    '				oNew.ObjTypeID = ObjectType.eUnit
    '				CType(oNew, Unit).EntityDef = oDefs(lDefIdx)
    '			Else
    '				oNew = New Facility()
    '				CType(oNew, Facility).EntityDef = CType(oDefs(lDefIdx), FacilityDef)
    '				oNew.ObjTypeID = ObjectType.eFacility
    '				CType(oNew, Facility).AutoLaunch = True
    '			End If

    '			With oNew
    '				.bProducing = False
    '				'.Cargo_Cap = oDefs(lDefIdx).Cargo_Cap
    '				.CurrentProduction = Nothing
    '				.CurrentSpeed = 0

    '				.CurrentStatus = 0
    '				For Y As Int32 = 0 To oDefs(lDefIdx).lSideCrits.Length - 1
    '					.CurrentStatus = .CurrentStatus Or oDefs(lDefIdx).lSideCrits(Y)
    '				Next Y

    '				oDefs(lDefIdx).DefName.CopyTo(.EntityName, 0)
    '				.ExpLevel = 0
    '				.Fuel_Cap = oDefs(lDefIdx).Fuel_Cap
    '				'.Hangar_Cap = oDefs(lDefIdx).Hangar_Cap
    '				.iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage
    '				.iTargetingTactics = 0
    '				.LocAngle = 0
    '				.LocX = lLocX
    '				.LocZ = lLocZ

    '				.Owner = oPirate
    '				.ParentObject = oEnvir

    '				.Q1_HP = oDefs(lDefIdx).Q1_MaxHP
    '				.Q2_HP = oDefs(lDefIdx).Q2_MaxHP
    '				.Q3_HP = oDefs(lDefIdx).Q3_MaxHP
    '				.Q4_HP = oDefs(lDefIdx).Q4_MaxHP
    '				.Shield_HP = oDefs(lDefIdx).Shield_MaxHP
    '				.Structure_HP = oDefs(lDefIdx).Structure_MaxHP
    '				.yProductionType = oDefs(lDefIdx).ProductionTypeID

    '				If .ObjTypeID = ObjectType.eUnit Then
    '					CType(oNew, Unit).DataChanged()
    '				Else : CType(oNew, Facility).DataChanged()
    '				End If
    '			End With

    '			Dim bResult As Boolean = False
    '			If oNew.ObjTypeID = ObjectType.eUnit Then
    '				bResult = CType(oNew, Unit).SaveObject()

    '				If bResult = True Then
    '					'Now, find a suitable place...
    '                       'SyncLock goUnit
    '                       '	Dim lIdx As Int32 = -1
    '                       '	For Y As Int32 = 0 To glUnitUB
    '                       '		If glUnitIdx(Y) = -1 Then
    '                       '			lIdx = Y
    '                       '			Exit For
    '                       '		End If
    '                       '	Next Y

    '                       '	If lIdx = -1 Then
    '                       '		ReDim Preserve goUnit(glUnitUB + 1)
    '                       '		ReDim Preserve glUnitIdx(glUnitUB + 1)
    '                       '		glUnitUB += 1
    '                       '		lIdx = glUnitUB
    '                       '	End If

    '                       '	goUnit(lIdx) = CType(oNew, Unit)
    '                       '	glUnitIdx(lIdx) = oNew.ObjectID
    '                       '	goUnit(lIdx).DataChanged()
    '                       'End SyncLock
    '                       AddUnitToGlobalArray(CType(oNew, Unit))
    '                       oNew.DataChanged()
    '					CType(oNew, Unit).SaveObject()
    '				End If
    '			Else
    '				bResult = CType(oNew, Facility).SaveObject()

    '				If bResult = True Then
    '					'Now, find a suitable place...
    '					Dim lIdx As Int32 = -1
    '                       'SyncLock goFacility
    '                       '	For Y As Int32 = 0 To glFacilityUB
    '                       '		If glFacilityIdx(Y) = -1 Then
    '                       '			lIdx = Y
    '                       '			Exit For
    '                       '		End If
    '                       '	Next Y

    '                       '	If lIdx = -1 Then
    '                       '		ReDim Preserve glFacilityIdx(glFacilityUB + 1)
    '                       '		ReDim Preserve goFacility(glFacilityUB + 1)
    '                       '		glFacilityUB += 1
    '                       '		lIdx = glFacilityUB
    '                       '	End If

    '                       '	goFacility(lIdx) = CType(oNew, Facility)
    '                       '	glFacilityIdx(lIdx) = oNew.ObjectID
    '                       'End SyncLock
    '                       lIdx = AddFacilityToGlobalArray(CType(oNew, Facility))

    '					With CType(oNew, Facility)
    '						.CurrentStatus = goFacility(lIdx).CurrentStatus Or elUnitStatus.eFacilityPowered
    '						.ServerIndex = lIdx
    '						.DataChanged()
    '						.SaveObject()

    '						If (.yProductionType And ProductionType.eProduction) <> 0 Then
    '							.lPirate_Counter = 0
    '							.lPirate_Counter_Max = 4
    '							.lPirate_For_PlayerID = lPlayerID
    '							AddPirateProductionItem(CType(oNew, Facility))
    '							AddEntityProducing(lIdx, ObjectType.eFacility, .ObjectID)
    '						End If
    '					End With
    '				End If
    '			End If

    '			If bResult = True Then
    '				oNew.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
    '				oDefs(lDefIdx).GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
    '			Else : lCnt -= 1
    '			End If
    '		Next X
    '	Next lDefIdx

    '	'Now, put in our cnt...
    '	System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lCntPos)

    '	ReDim Preserve yResp(23 + (lCnt * 12))

    '	Return yResp
    'End Function

	Public Function SendPirateRaidForce(ByRef oCreator As Facility) As Boolean
		'Ok, get the player...
		Dim bResult As Boolean = False
		Dim oTargetPlayer As Player = GetEpicaPlayer(oCreator.lPirate_For_PlayerID)
		Dim lEnvirID As Int32
		Dim iEnvirTypeID As Int16

		Dim oDomainSocket As NetSock = Nothing

		With CType(oCreator.ParentObject, Epica_GUID)
			lEnvirID = .ObjectID
			iEnvirTypeID = .ObjTypeID

			If iEnvirTypeID = ObjectType.ePlanet Then
				oDomainSocket = CType(oCreator.ParentObject, Planet).oDomain.DomainSocket
			ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
				oDomainSocket = CType(oCreator.ParentObject, SolarSystem).oDomain.DomainSocket
			End If
		End With

		If oDomainSocket Is Nothing Then Return False

		If oTargetPlayer Is Nothing = False Then

			Dim lTargetColonyIdx As Int32 = oTargetPlayer.GetColonyFromParent(lEnvirID, iEnvirTypeID)

			If lTargetColonyIdx <> -1 AndAlso glColonyIdx(lTargetColonyIdx) <> -1 Then
				Dim lMoveToX As Int32 = Int32.MinValue
				Dim lMoveToY As Int32 = Int32.MinValue
				Dim oColony As Colony = goColony(lTargetColonyIdx)
				If oColony Is Nothing Then Return False

				For X As Int32 = 0 To oColony.ChildrenUB
					If oColony.lChildrenIdx(X) <> -1 Then
						If oColony.oChildren(X).yProductionType = ProductionType.eCommandCenterSpecial OrElse lMoveToX = Int32.MinValue Then
							lMoveToX = oColony.oChildren(X).LocX + 1000
							lMoveToY = oColony.oChildren(X).LocZ + 1000

							If oColony.oChildren(X).yProductionType = ProductionType.eCommandCenterSpecial Then Exit For
						End If
					End If
				Next X

				For X As Int32 = 0 To oCreator.lPirate_ItemUB
					Dim yMove(23) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMove, 0)
					System.BitConverter.GetBytes(lMoveToX).CopyTo(yMove, 2)
					System.BitConverter.GetBytes(lMoveToY).CopyTo(yMove, 6)
					System.BitConverter.GetBytes(0S).CopyTo(yMove, 10)
					System.BitConverter.GetBytes(lEnvirID).CopyTo(yMove, 12)
					System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMove, 16)
					System.BitConverter.GetBytes(oCreator.lPirate_Items(X)).CopyTo(yMove, 18)
					System.BitConverter.GetBytes(ObjectType.eUnit).CopyTo(yMove, 22)
					oDomainSocket.SendData(yMove)
				Next X

				bResult = True
			End If
		End If

		Return bResult
	End Function

	Public Sub AddPirateProductionItem(ByRef oFac As Facility)
		Const l_HARDCODE_RAIDER As Int32 = 19
		Const l_HARDCODE_MARAUDER As Int32 = 20
		Const l_HARDCODE_PROWLER As Int32 = 21

		Dim lTmpVal As Int32 = CInt(Rnd() * 100)
		If lTmpVal < 33 Then
			oFac.AddProduction(l_HARDCODE_RAIDER, ObjectType.eUnitDef, 254, 1, 0)
		ElseIf lTmpVal < 66 Then
			oFac.AddProduction(l_HARDCODE_MARAUDER, ObjectType.eUnitDef, 254, 1, 0)
		Else : oFac.AddProduction(l_HARDCODE_PROWLER, ObjectType.eUnitDef, 254, 1, 0)
		End If
	End Sub

    '#Region " Chat Engine "
    '	Public guChatRooms() As ChatRoom
    '	Public glChatRoomUB As Int32 = -1
    '    Private Enum eyChatRoomAttr As Byte
    '        PublicChannel = 1
    '        UserIsAdmin = 2
    '        AllAreAdmin = 4
    '        PasswordProtected = 8
    '        PlayerIsPermitted = 16
    '        PlayerInChannel = 32
    '    End Enum
    '	Public Structure ChatRoom
    '		Public sRoomName As String
    '		Public sPassword As String
    '		Public lID As Int32			'ID From database
    '		Public bAllAreAdmin As Boolean
    '		Public bPublic As Boolean

    '		Private mlPlayersInRoom() As Int32
    '		Private mlPlayerIDs() As Int32
    '        Private mbHasAdminRights() As Boolean

    '        Public lPlayerCount As Int32

    '		Public Sub SendMsgToPlayersInRoom(ByVal yMsg() As Byte)
    '			Dim X As Int32

    '            If mlPlayersInRoom Is Nothing Then Exit Sub

    '            lPlayerCount = 0

    '			For X = 0 To mlPlayersInRoom.Length - 1
    '				If mlPlayersInRoom(X) > -1 AndAlso glPlayerIdx(mlPlayersInRoom(X)) > -1 Then
    '					Try
    '						Dim oPlayer As Player = goPlayer(mlPlayersInRoom(X))
    '                        If oPlayer Is Nothing = False AndAlso oPlayer.bInPlayerRequestDetails = False Then
    '                            lPlayerCount += 1
    '                            If oPlayer.oSocket Is Nothing = False Then oPlayer.oSocket.SendData(yMsg)
    '                        End If
    '					Catch
    '					End Try
    '				End If
    '			Next X
    '		End Sub

    '		Public Sub AddPlayerToRoom(ByVal lPlayerID As Int32)
    '			Dim X As Int32
    '			Dim lIdx As Int32 = -1
    '			Dim bAdmin As Boolean = False
    '            Dim lPlayerIndex As Int32 = -1

    '            lPlayerCount += 1

    '			For X = 0 To glPlayerUB
    '				If glPlayerIdx(X) = lPlayerID Then
    '					lPlayerIndex = X
    '					Exit For
    '				End If
    '			Next X

    '			If lPlayerIndex > -1 AndAlso glPlayerIdx(lPlayerIndex) > -1 Then
    '				If mlPlayersInRoom Is Nothing Then
    '					ReDim mlPlayersInRoom(0)
    '					ReDim mlPlayerIDs(0)
    '					ReDim mbHasAdminRights(0)
    '					bAdmin = True
    '					lIdx = 0
    '				Else
    '					For X = 0 To mlPlayersInRoom.Length - 1
    '						If mlPlayersInRoom(X) = lPlayerIndex Then
    '							Exit Sub	'we already have this one, leave
    '						ElseIf lIdx = -1 AndAlso mlPlayersInRoom(X) = -1 Then
    '							lIdx = X
    '						End If
    '					Next X
    '					If lIdx = -1 Then
    '						lIdx = mlPlayersInRoom.Length
    '						ReDim Preserve mlPlayersInRoom(lIdx)
    '						ReDim Preserve mlPlayerIDs(lIdx)
    '						ReDim Preserve mbHasAdminRights(lIdx)
    '					End If
    '				End If

    '				'Now, associate
    '				mlPlayersInRoom(lIdx) = lPlayerIndex
    '				mlPlayerIDs(lIdx) = glPlayerIdx(lPlayerIndex)

    '				If Me.bAllAreAdmin = True Then bAdmin = True
    '				mbHasAdminRights(lIdx) = bAdmin

    '				goPlayer(lPlayerIndex).AssociateChatRoom(lID)

    '				SendMsgToPlayersInRoom(MsgSystem.CreateChatMsg(-1, goPlayer(lPlayerIndex).sPlayerNameProper & " has entered " & Me.sRoomName & ".", ChatMessageType.eSysAdminMessage))
    '			End If

    '		End Sub

    '		Public Sub SetPlayerAdmin(ByVal lPlayerID As Int32, ByVal bAdmin As Boolean)
    '			Dim X As Int32
    '			If mlPlayerIDs Is Nothing Then Exit Sub

    '			For X = 0 To mlPlayerIDs.Length - 1
    '				If mlPlayerIDs(X) = lPlayerID Then
    '					mbHasAdminRights(X) = bAdmin
    '					Exit For
    '				End If
    '			Next X
    '		End Sub

    '		Public Sub RemovePlayerFromRoom(ByVal lPlayerID As Int32)
    '			Dim X As Int32

    '            If mlPlayerIDs Is Nothing Then Exit Sub

    '            For X = 0 To mlPlayerIDs.Length - 1
    '                If mlPlayerIDs(X) = lPlayerID Then
    '                    goPlayer(mlPlayersInRoom(X)).RemoveChatRoom(lID)

    '                    mlPlayerIDs(X) = -1
    '                    mlPlayersInRoom(X) = -1
    '                    mbHasAdminRights(X) = False

    '                    lPlayerCount -= 1

    '                    'TODO: notify player they were removed from room

    '                End If
    '            Next X
    '		End Sub

    '		Public Function PlayerHasAdmin(ByVal lPlayerID As Int32) As Boolean
    '			Dim X As Int32
    '			If mlPlayerIDs Is Nothing Then Exit Function
    '			For X = 0 To mlPlayerIDs.Length - 1
    '				If mlPlayerIDs(X) = lPlayerID Then
    '					Return mbHasAdminRights(X)
    '				End If
    '			Next X
    '			Return False
    '		End Function

    '		Public Function GetChannelWhoList() As String
    '			Dim oSB As New System.Text.StringBuilder
    '			oSB.AppendLine("CHANNEL WHO FOR " & sRoomName)
    '			If mlPlayersInRoom Is Nothing = False Then
    '				For X As Int32 = 0 To mlPlayersInRoom.GetUpperBound(0)
    '					If mlPlayersInRoom(X) <> -1 AndAlso glPlayerIdx(mlPlayersInRoom(X)) <> -1 Then
    '						oSB.AppendLine(goPlayer(mlPlayersInRoom(X)).sPlayerNameProper)
    '					End If
    '				Next X
    '			End If
    '			Return oSB.ToString
    '        End Function

    '        Public Function GetChatRoomDetailsMsg() As Byte()
    '            lPlayerCount = 0
    '            If mlPlayersInRoom Is Nothing = False Then
    '                For X As Int32 = 0 To mlPlayersInRoom.GetUpperBound(0)
    '                    If mlPlayersInRoom(X) > -1 AndAlso glPlayerIdx(mlPlayersInRoom(X)) > -1 Then
    '                        lPlayerCount += 1
    '                    End If
    '                Next X
    '            End If
    '            Dim lPCnt As Int32 = lPlayerCount
    '            Dim yResp(70 + (lPCnt * 4)) As Byte
    '            Dim lPos As Int32 = 0
    '            System.BitConverter.GetBytes(GlobalMessageCode.eRequestChannelDetails).CopyTo(yResp, lPos) : lPos += 2
    '            System.BitConverter.GetBytes(lID).CopyTo(yResp, lPos) : lPos += 4
    '            StringToBytes(sRoomName).CopyTo(yResp, lPos) : lPos += 30
    '            If sPassword Is Nothing = False Then StringToBytes(sPassword).CopyTo(yResp, lPos)
    '            lPos += 30
    '            yResp(lPos) = GetChatRoomAttr() : lPos += 1
    '            System.BitConverter.GetBytes(lPCnt).CopyTo(yResp, lPos) : lPos += 4
    '            If mlPlayersInRoom Is Nothing = False Then
    '                For X As Int32 = 0 To mlPlayersInRoom.GetUpperBound(0)
    '                    If mlPlayersInRoom(X) > -1 AndAlso glPlayerIdx(mlPlayersInRoom(X)) > -1 Then
    '                        lPlayerCount -= 1
    '                        If lPlayerCount < 0 Then Exit For

    '                        Dim lPlayerID As Int32 = glPlayerIdx(mlPlayersInRoom(X))
    '                        If mbHasAdminRights(X) = True Then lPlayerID = -lPlayerID
    '                        System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
    '                    End If
    '                Next X
    '            End If
    '            Return yResp
    '        End Function

    '        Public Function GetChatRoomAttr() As Byte
    '            Dim yAttr As Byte = 0
    '            If bPublic = True Then yAttr = yAttr Or eyChatRoomAttr.PublicChannel
    '            If bAllAreAdmin = True Then yAttr = yAttr Or eyChatRoomAttr.AllAreAdmin
    '            If sPassword Is Nothing = False AndAlso sPassword <> "" Then yAttr = yAttr Or eyChatRoomAttr.PasswordProtected
    '            Return yAttr
    '        End Function
    '	End Structure

    '	Public Function AddChatRoom(ByVal sRoomName As String, ByVal lPlayerID As Int32) As Boolean
    '		Dim X As Int32
    '		Dim lIdx As Int32 = -1

    '		Dim sRoom As String = UCase$(sRoomName)

    '		For X = 0 To glChatRoomUB
    '			If guChatRooms(X).sRoomName = sRoom Then
    '				'already found, so return false
    '				Return False
    '			ElseIf lIdx = -1 AndAlso guChatRooms(X).sRoomName = "" Then
    '				'ok, found a good one
    '				lIdx = X
    '			End If
    '		Next X

    '		If lIdx = -1 Then
    '			glChatRoomUB += 1
    '			ReDim Preserve guChatRooms(glChatRoomUB)
    '			lIdx = glChatRoomUB
    '		End If

    '		With guChatRooms(lIdx)
    '			.sRoomName = sRoom
    '			.lID = lIdx		'???
    '			If lPlayerID <> -1 Then
    '				.AddPlayerToRoom(lPlayerID)
    '				.SetPlayerAdmin(lPlayerID, True)
    '				.bPublic = False
    '				'	.bAllAreAdmin = False
    '				'Else
    '				'	.bAllAreAdmin = True
    '			End If
    '		End With

    '		Return True

    '	End Function

    '	Public Sub CloseChatRoom(ByVal sRoomName As String)
    '		Dim X As Int32
    '		Dim sRoom As String = UCase$(sRoomName)

    '		For X = 0 To glChatRoomUB
    '			If guChatRooms(X).sRoomName = sRoom Then
    '				guChatRooms(X).sRoomName = ""
    '				'TODO: if players are here, remove them
    '				Exit For
    '			End If
    '		Next X
    '    End Sub

    '    Public Function HandleRequestChannelList(ByVal lPlayerID As Int32) As Byte()
    '        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
    '        If oPlayer Is Nothing Then Return Nothing

    '        Dim lCnt As Int32 = 0

    '        Dim lCurUB As Int32 = -1
    '        If guChatRooms Is Nothing = False Then lCurUB = Math.Min(glChatRoomUB, guChatRooms.GetUpperBound(0))
    '        For X As Int32 = 0 To lCurUB
    '            If guChatRooms(X).bPublic = True OrElse oPlayer.InChatRoom(X) = True Then
    '                lCnt += 1
    '            End If
    '        Next X

    '        Dim yResp(5 + (lCnt * 39)) As Byte
    '        Dim lPos As Int32 = 0
    '        System.BitConverter.GetBytes(GlobalMessageCode.eRequestChannelList).CopyTo(yResp, lPos) : lPos += 2
    '        System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lPos) : lPos += 4

    '        For X As Int32 = 0 To lCurUB
    '            If guChatRooms(X).bPublic = True OrElse oPlayer.InChatRoom(X) = True Then
    '                lCnt -= 1
    '                If lCnt < 0 Then Exit For
    '                With guChatRooms(X)
    '                    StringToBytes(.sRoomName).CopyTo(yResp, lPos) : lPos += 30
    '                    System.BitConverter.GetBytes(.lID).CopyTo(yResp, lPos) : lPos += 4
    '                    Dim yAttr As Byte = .GetChatRoomAttr()
    '                    If .PlayerHasAdmin(lPlayerID) = True Then yAttr = yAttr Or eyChatRoomAttr.UserIsAdmin
    '                    If oPlayer.InChatRoom(X) = True Then yAttr = yAttr Or eyChatRoomAttr.PlayerInChannel

    '                    yResp(lPos) = yAttr : lPos += 1

    '                    System.BitConverter.GetBytes(.lPlayerCount).CopyTo(yResp, lPos) : lPos += 4
    '                End With
    '            End If
    '        Next X

    '        Return yResp
    '    End Function

    '#End Region

#Region " OLE DB Connection Stuff "
	Private moCN As OleDb.OleDbConnection

	Public Function InitializeConnection() As Boolean
		'go ahead and connect
		Dim sConnStr As String
		Dim oINI As New InitFile()
		Dim sUDL As String
		Dim bNoErrors As Boolean

		bNoErrors = True
		Try
			sUDL = oINI.GetString("SETTINGS", "Perm_Save_UDL", "")
			If sUDL <> "" Then
				sConnStr = "FILE NAME=" & AppDomain.CurrentDomain.BaseDirectory & sUDL
				moCN = New OleDb.OleDbConnection(sConnStr)
				moCN.Open()
			End If
		Catch
			LogEvent(LogEventType.CriticalError, "InitializeConnection: " & Err.Description)
		Finally
			If bNoErrors = False Then
				moCN = Nothing
			End If
		End Try

		Return bNoErrors
	End Function

	Public ReadOnly Property goCN() As OleDb.OleDbConnection
		Get
			Dim bNoErrors As Boolean = True
			Try
                If moCN Is Nothing = True OrElse moCN.State = ConnectionState.Broken OrElse moCN.State = ConnectionState.Closed Then
                    Dim oINI As New InitFile
                    Dim sUDL As String = oINI.GetString("SETTINGS", "Perm_Save_UDL", "")
                    If sUDL <> "" Then
                        Dim sConnStr As String = "FILE NAME=" & AppDomain.CurrentDomain.BaseDirectory & sUDL
                        moCN = New OleDb.OleDbConnection(sConnStr)
                        moCN.Open()
                    End If
                End If
			Catch ex As Exception
				LogEvent(LogEventType.CriticalError, "GetConnection: " & ex.Message)
				moCN = Nothing
			Finally
				If bNoErrors = False Then moCN = Nothing
			End Try
			Return moCN
		End Get
	End Property

	Private moWebCN As Odbc.OdbcConnection
	Public Function GetWebConnection() As Odbc.OdbcConnection
		Try
			If moWebCN Is Nothing = True OrElse moWebCN.State = ConnectionState.Broken OrElse moWebCN.State = ConnectionState.Closed Then
				Dim sConnStr As String = "DSN=WebDB"
				moWebCN = New Odbc.OdbcConnection(sConnStr)
				moWebCN.Open()
			End If
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "GetWebConnection: " & ex.Message)
			moWebCN = Nothing
		End Try
		Return moWebCN
	End Function

	Public Sub CloseConn()
		If Not moCN Is Nothing Then
			moCN.Close()
		End If
		If moWebCN Is Nothing = False Then moWebCN.Close()
		moCN = Nothing
		moWebCN = Nothing
	End Sub
#End Region

#Region " Event Logging "

	Private moFileStream As System.IO.FileStream
	Private moFileWrite As System.IO.StreamWriter
	Public Enum LogEventType
		CriticalError = 1
		PossibleCheat = 2
		Warning = 3
		Informational = 4
		ExtensiveLogging = 5
	End Enum

	Public Function InitializeEventLogging() As Boolean
		Dim oINI As New InitFile()
		Dim sLog As String

		sLog = oINI.GetString("SETTINGS", "EventLogFilePath", "")
		If sLog = "" Then
			sLog = AppDomain.CurrentDomain.BaseDirectory()
			If Right$(sLog, 1) <> "\" Then sLog = sLog & "\"
            sLog = sLog & "Events_.Log"
            Dim sNewFileName As String = "Events_" & Now.ToString("MMddyyyyHHmmss") & ".log"
            sLog &= sNewFileName
		End If
		moFileStream = System.IO.File.Open(sLog, IO.FileMode.Append)
		moFileWrite = New System.IO.StreamWriter(moFileStream)
		moFileWrite.AutoFlush = True
	End Function

	Public Sub LogEvent(ByVal lEventType As LogEventType, ByVal sValue As String)

		'Public Shared Function StringToBytes(ByVal sData As String) As Byte()
		'    Return System.Text.ASCIIEncoding.ASCII.GetBytes(sData)
		'End Function
		'Dim bAddEventLine As Boolean = True
		Select Case lEventType
			Case LogEventType.CriticalError
				sValue = "CRITICAL: " & sValue
			Case LogEventType.PossibleCheat
				sValue = "POSSIBLE CHEAT: " & sValue
			Case LogEventType.Warning
				sValue = "WARNING: " & sValue
				'Case Else
				'	bAddEventLine = False
		End Select

		'If bAddEventLine = True Then
		If lEventType <> LogEventType.ExtensiveLogging Then gfrmDisplayForm.AddEventLine(sValue)

        If moFileStream Is Nothing = False Then
            If moFileStream.Length > 8000000 Then
                CloseEventLogger()
                InitializeEventLogging()
            End If
            If moFileStream.CanWrite() Then
                moFileWrite.WriteLine(Now.ToString("HH:mm:ss") & "|" & sValue)
            End If
        End If
	End Sub

	Public Sub CloseEventLogger()
		moFileWrite.Close()
		moFileStream.Close()
	End Sub
#End Region

#Region "  Quick Technology Lookup  "
	Private mlCompIDs() As Int32
	Private miCompTypeIDs() As Int16
	Private mlOwnerIDs() As Int32
	Private mlCompUB As Int32 = -1
	Public Sub AddQuickTechnologyLookup(ByVal lCompID As Int32, ByVal iCompTypeID As Int16, ByVal lOwnerID As Int32)
		Dim lIdx As Int32 = -1

		For X As Int32 = 0 To mlCompUB
			If mlCompIDs(X) = lCompID AndAlso miCompTypeIDs(X) = iCompTypeID AndAlso mlOwnerIDs(X) = lOwnerID Then
				Return
			ElseIf lIdx = -1 AndAlso mlCompIDs(X) = -1 Then
				lIdx = X
			End If
		Next X

		If lIdx = -1 Then
			mlCompUB += 1
			ReDim Preserve mlCompIDs(mlCompUB)
			ReDim Preserve mlOwnerIDs(mlCompUB)
			ReDim Preserve miCompTypeIDs(mlCompUB)
			lIdx = mlCompUB
		End If

		mlCompIDs(lIdx) = lCompID
		miCompTypeIDs(lIdx) = iCompTypeID
		mlOwnerIDs(lIdx) = lOwnerID
	End Sub

	Public Function QuickLookupTechnology(ByVal lCompID As Int32, ByVal iCompTypeID As Int16) As Epica_Tech
		Dim lOwnerID As Int32 = -1
		For X As Int32 = 0 To mlCompUB
			If mlCompIDs(X) = lCompID AndAlso miCompTypeIDs(X) = iCompTypeID Then
				lOwnerID = mlOwnerIDs(X)
				Exit For
			End If
		Next X

		If lOwnerID <> -1 Then
			Dim oPlayer As Player = GetEpicaPlayer(lOwnerID)
			If oPlayer Is Nothing Then oPlayer = goInitialPlayer
			If oPlayer Is Nothing = False Then
				Return oPlayer.GetTech(lCompID, iCompTypeID)
			End If
		End If
		Return Nothing
	End Function
#End Region

#Region " Object Arrays and Loading "
	Public gbServerInitializing As Boolean = False

	Public goAgent() As Agent
	Public glAgentIdx() As Int32
	Public glAgentUB As Int32 = -1

	Public goColony(-1) As Colony
	Public glColonyIdx() As Int32
	Public glColonyUB As Int32 = -1

	'TODO: May want to make these locational specific/contained to reduce array search times
	Public goComponentCache() As ComponentCache
	Public glComponentCacheIdx() As Int32
	Public glComponentCacheUB As Int32 = -1

	Public goCorporation() As Corporation
	Public glCorporationIdx() As Int32
	Public glCorporationUB As Int32 = -1

	Public goFormationDefs(-1) As FormationDef
	Public glFormationDefIdx(-1) As Int32
	Public glFormationDefUB As Int32 = -1

	Public goGalaxy() As Galaxy
	Public glGalaxyIdx() As Int32
	Public glGalaxyUB As Int32 = -1

	Public goGNS() As GNS
	Public glGNSIdx() As Int32
	Public glGNSUB As Int32 = -1

	Public goGoal() As Goal
	Public glGoalIdx() As Int32
	Public glGoalUB As Int32 = -1
	Private myGoalListResp() As Byte

	Public goGuild() As Guild
	Public glGuildIdx() As Int32
	Public glGuildUB As Int32 = -1

	Public goMineralProperty() As MineralProperty
	Public glMineralPropertyIdx() As Int32
	Public glMineralPropertyUB As Int32 = -1

	Public goMineral() As Mineral
	Public glMineralIdx() As Int32
	Public glMineralUB As Int32 = -1

	'TODO: May want to make these locational specific/contained... ie: inside of a system, planet, unit, etc...
	Public goMineralCache(-1) As MineralCache
	Public glMineralCacheIdx() As Int32
	Public glMineralCacheUB As Int32 = -1

	Public goMission() As Mission
	Public glMissionIdx() As Int32
	Public glMissionUB As Int32 = -1

	Public guMissionMethods() As uMissionMethod
	Public glMissionMethodUB As Int32 = -1

	Public goNebula() As Nebula
	Public glNebulaIdx() As Int32
	Public glNebulaUB As Int32 = -1

	Public goPlanet() As Planet
	Public glPlanetIdx() As Int32
	Public glPlanetUB As Int32 = -1

    Public goPlayer(-1) As Player
    Public glPlayerIdx(-1) As Int32
	Public glPlayerUB As Int32 = -1

	Public goPlayerMission() As PlayerMission
	Public glPlayerMissionIdx(-1) As Int32
	Public glPlayerMissionUB As Int32 = -1

	'Public goSenate() As Senate
	'Public glSenateIdx() As Int32
	'Public glSenateUB As Int32 = -1

	Public goSkill() As Skill
	Public glSkillIdx() As Int32
	Public glSkillUB As Int32 = -1
	Private mySkillListResp() As Byte

	Public goSpecialTechs() As SpecialTech
	Public glSpecialTechIdx() As Int32
    Public glSpecialTechUB As Int32 = -1
    Public gyGuaranteedList() As Byte

	Public goSpecTechPreq() As SpecialTechPrerequisite
	Public glSpecTechPreqIdx() As Int32
	Public glSpecTechPreqUB As Int32 = -1

	Public goStarType() As StarType
	Public glStarTypeIdx() As Int32
	Public glStarTypeUB As Int32 = -1

	Public goSystem() As SolarSystem
	Public glSystemIdx() As Int32
	Public glSystemUB As Int32 = -1

	Public goFacility(-1) As Facility
	Public glFacilityIdx() As Int32
    Public glFacilityUB As Int32 = -1
    Public gcolFacLookup As New Collection

    Public goFacilityDef(-1) As FacilityDef
	Public glFacilityDefIdx() As Int32
	Public glFacilityDefUB As Int32 = -1
 
	Public goUnit(-1) As Unit
	Public glUnitIdx() As Int32
	Public glUnitUB As Int32 = -1

    Public goUnitDef(-1) As Epica_Entity_Def
	Public glUnitDefIdx() As Int32
	Public glUnitDefUB As Int32 = -1

	Public goUnitGroup() As UnitGroup
	Public glUnitGroupIdx() As Int32
	Public glUnitGroupUB As Int32 = -1

    Public goWeaponDefs(-1) As WeaponDef
	Public glWeaponDefIdx() As Int32
	Public glWeaponDefUB As Int32 = -1

	Public goWormhole() As Wormhole
	Public glWormholeIdx() As Int32
	Public glWormholeUB As Int32 = -1

	'Public goBugs() As EpicaBug
	'Public glBugUB As Int32 = -1

	Public Function LoadAllStaticData() As Boolean
		Dim bResult As Boolean = LoadGeographical()
		If bResult = True Then bResult = LoadGoalsAndSkills()
		If bResult = True Then bResult = LoadStaticMineralData()

        'AddChatRoom("Testers", -1)

		Return bResult
	End Function

	Public Function LoadNeighborhoodData() As Boolean
		Dim sPlayerInStr As String = ""
		'Dim bFound As Boolean = False
		'While bFound = False
		'	'sPlayerInStr = "(SELECT PlayerID FROM tblPlayer WHERE StartedEnvirTypeID = 3 AND StartedEnvirID IN ("
		'	For X As Int32 = 0 To glPlanetUB
		'		If glPlanetIdx(X) <> -1 AndAlso goPlanet(X).InMyDomain = True Then
		'			If sPlayerInStr = "" Then sPlayerInStr = "(SELECT PlayerID FROM tblPlayer WHERE StartedEnvirTypeID = 3 AND StartedEnvirID IN (" Else sPlayerInStr &= ", "
		'			'If sPlayerInStr.EndsWith("(") = False Then sPlayerInStr &= ", "
		'			sPlayerInStr &= goPlanet(X).ObjectID
		'			bFound = True
		'		End If
		'	Next X
		'End While
		'sPlayerInStr &= "))"
		sPlayerInStr = "(SELECT PlayerID FROM tblPlayer)"

		Dim bResult As Boolean = LoadPlayers()
		If bResult = True Then bResult = LoadPlayerDetails(sPlayerInStr)
        If bResult = True Then bResult = LoadTechs(" OwnerID = 0 OR OwnerID IN " & sPlayerInStr, True)
        If bResult = True Then bResult = LoadDefs(sPlayerInStr)
        'If True = True Then DoCheckForMicroTechs()
        If bResult = True Then bResult = LoadRemaining(sPlayerInStr)
		Return bResult
    End Function

    Private Sub DoCheckForMicroTechs()
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) > -1 Then
                Dim oPlayer As Player = goPlayer(X)
                If oPlayer Is Nothing = False Then
                    'For Y As Int32 = 0 To oPlayer.mlEngineUB
                    '    If oPlayer.mlEngineIdx(Y) > -1 AndAlso oPlayer.moEngine(Y) Is Nothing = False Then
                    '        oPlayer.moEngine(Y).CheckForMicroTech()
                    '    End If
                    'Next Y
                    'For Y As Int32 = 0 To oPlayer.mlShieldUB
                    '    If oPlayer.mlShieldIdx(Y) > -1 AndAlso oPlayer.moShield(Y) Is Nothing = False Then
                    '        oPlayer.moShield(Y).CheckForMicroTech()
                    '    End If
                    'Next Y
                    'For Y As Int32 = 0 To oPlayer.mlRadarUB
                    '    If oPlayer.mlRadarIdx(Y) > -1 AndAlso oPlayer.moRadar(Y) Is Nothing = False Then
                    '        oPlayer.moRadar(Y).CheckForMicroTech()
                    '    End If
                    'Next Y
                    For Y As Int32 = 0 To oPlayer.mlWeaponUB
                        If oPlayer.mlWeaponIdx(Y) > -1 AndAlso oPlayer.moWeapon(Y) Is Nothing = False Then
                            Select Case oPlayer.moWeapon(Y).WeaponClassTypeID
                                Case WeaponClassType.eBomb
                                    'CType(oPlayer.moWeapon(Y), BombWeaponTech).CheckForMicroTech()
                                Case WeaponClassType.eEnergyBeam
                                    CType(oPlayer.moWeapon(Y), BeamWeaponTech).CheckForMicroTech()
                                Case WeaponClassType.eEnergyPulse
                                    CType(oPlayer.moWeapon(Y), PulseWeaponTech).CheckForMicroTech()
                                Case WeaponClassType.eMissile
                                    'CType(oPlayer.moWeapon(Y), MissileWeaponTech).CheckForMicroTech()
                                Case WeaponClassType.eProjectile
                                    'CType(oPlayer.moWeapon(Y), ProjectileWeaponTech).CheckForMicroTech()
                            End Select
                        End If
                    Next Y
                    For Y As Int32 = 0 To oPlayer.mlPrototypeUB
                        If oPlayer.mlPrototypeIdx(Y) > -1 AndAlso oPlayer.moPrototype(Y) Is Nothing = False Then
                            oPlayer.moPrototype(Y).CheckForMicroTech()
                        End If
                    Next Y
                End If
            End If
        Next X
    End Sub
 
	Private Function LoadGeographical() As Boolean
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim oComm As OleDb.OleDbCommand
		Dim sSQL As String
		Dim bResult As Boolean = False

		Try
			LogEvent(LogEventType.Informational, "Loading Galaxy Objects...")
			'Galaxy objects
			sSQL = "SELECT * FROM tblGalaxy"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glGalaxyUB = glGalaxyUB + 1
				ReDim Preserve goGalaxy(glGalaxyUB)
				ReDim Preserve glGalaxyIdx(glGalaxyUB)
				goGalaxy(glGalaxyUB) = New Galaxy()
				With goGalaxy(glGalaxyUB)
					.GalaxyName = StringToBytes(CStr(oData("GalaxyName")))
					.ObjectID = CInt(oData("GalaxyID"))
					.ObjTypeID = ObjectType.eGalaxy
					.CycleNumber = CInt(oData("CycleNumber"))
					glGalaxyIdx(glGalaxyUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Startypes
			LogEvent(LogEventType.Informational, "Loading Star Types...")
			sSQL = "SELECT * FROM tblStarType"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glStarTypeUB += 1
				ReDim Preserve goStarType(glStarTypeUB)
				ReDim Preserve glStarTypeIdx(glStarTypeUB)
				goStarType(glStarTypeUB) = New StarType()
				With goStarType(glStarTypeUB)
					.HeatIndex = CByte(oData("HeatIndex"))
					.LightAmbient = CInt(oData("LightAmbient"))
					.LightAtt1 = CSng(oData("LightAtt1"))
					.LightAtt2 = CSng(oData("LightAtt2"))
					.LightAtt3 = CSng(oData("LightAtt3"))
					.LightDiffuse = CInt(oData("LightDiffuse"))
					.LightRange = CInt(oData("LightRange"))
					.LightSpecular = CInt(oData("LightSpecular"))
					.MatDiffuse = CInt(oData("MatDiffuse"))
					.MatEmissive = CInt(oData("MatEmissive"))
					.StarMapRectIdx = CByte(oData("StarMapRectIdx"))
					.StarRadius = CInt(oData("StarRadius"))
					.StarRarity = CByte(oData("StarRarity"))
					.StarTexture = StringToBytes(CStr(oData("StarTexture")))
					.StarTypeAttrs = CInt(oData("StarTypeAttrs"))
					.StarTypeID = CByte(oData("StarTypeID"))
					.StarTypeName = StringToBytes(CStr(oData("StarTypeName")))
					glStarTypeIdx(glStarTypeUB) = .StarTypeID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Nebula Objects
			LogEvent(LogEventType.Informational, "Loading Nebulae...")
			sSQL = "SELECT * FROM tblNebula"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glNebulaUB = glNebulaUB + 1
				ReDim Preserve goNebula(glNebulaUB)
				ReDim Preserve glNebulaIdx(glNebulaUB)
				goNebula(glNebulaUB) = New Nebula()
				With goNebula(glNebulaUB)
					.Effects = CInt(oData("Effects"))
					.Height = CByte(oData("Height"))
					.LocX = CByte(oData("LocX"))
					.LocY = CByte(oData("LocY"))
					.NebulaDesc = CStr(oData("NebulaDesc"))
					.NebulaName = StringToBytes(CStr(oData("NebulaName")))
					.ObjectID = CInt(oData("NebulaID"))
					.ObjTypeID = ObjectType.eNebula
					.Width = CByte(oData("Width"))
					glNebulaIdx(glNebulaUB) = .ObjectID
					.ParentGalaxy = GetEpicaGalaxy(CInt(oData("galaxyid")))

					'TODO: May want to 'subscribe' the nebula to the galaxy object at this point
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'System Objects
			LogEvent(LogEventType.Informational, "Loading Solar Systems...")
			sSQL = "SELECT * FROM tblSolarSystem ORDER BY SystemID"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glSystemUB = glSystemUB + 1
				ReDim Preserve goSystem(glSystemUB)
				ReDim Preserve glSystemIdx(glSystemUB)
				goSystem(glSystemUB) = New SolarSystem()
				With goSystem(glSystemUB)
					.LocX = CInt(oData("LocX"))
					.LocY = CInt(oData("LocY"))
					.LocZ = CInt(oData("LocZ"))
					.ObjectID = CInt(oData("SystemID"))
					.ObjTypeID = ObjectType.eSolarSystem
					.ParentGalaxy = GetEpicaGalaxy(CInt(oData("galaxyid")))
					.StarType1ID = CByte(oData("StarType1ID"))
					.StarType2ID = CByte(oData("StarType2ID"))
					.StarType3ID = CByte(oData("StarType3ID"))
					.SystemName = StringToBytes(CStr(oData("SystemName")))
					.SystemType = CByte(oData("SystemType"))
					.OwnerID = CInt(oData("OwnerID"))
                    .FleetJumpPointX = CInt(oData("FleetJumpPointX"))
                    .FleetJumpPointZ = CInt(oData("FleetJumpPointZ"))

                    If .FleetJumpPointX = -1 AndAlso .FleetJumpPointZ = -1 Then
                        '.FleetJumpPointX = CInt((4000 * Rnd()) - 2000)
                        '.FleetJumpPointZ = CInt((4000 * Rnd()) - 2000)
                        Dim dDX As Single = goStarType(.StarType1ID).StarRadius
                        Dim dDY As Single = 0
                        Dim oRandom As New Random
                        Dim lAngle As Single = oRandom.Next(0, 360)
                        RotatePoint(0, 0, dDX, dDY, lAngle)
                        .FleetJumpPointX = CInt(dDX)
                        .FleetJumpPointZ = CInt(dDY)
                        .QueueMeToSave()
                    End If

					glSystemIdx(glSystemUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Planet Objects
			LogEvent(LogEventType.Informational, "Loading Planets...")
			sSQL = "SELECT * FROM tblPlanet ORDER BY PlanetID"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glPlanetUB = glPlanetUB + 1
				ReDim Preserve goPlanet(glPlanetUB)
				ReDim Preserve glPlanetIdx(glPlanetUB)
				goPlanet(glPlanetUB) = New Planet()
				With goPlanet(glPlanetUB)
					.Atmosphere = CByte(oData("Atmosphere"))
					.AxisAngle = CInt(oData("AxisAngle"))
					.Gravity = CByte(oData("Gravity"))
					.Hydrosphere = CByte(oData("Hydrosphere"))
					.InnerRingRadius = CInt(oData("RingInnerRadius"))
					.LocX = CInt(oData("LocX"))
					.LocY = CInt(oData("LocY"))
					.LocZ = CInt(oData("LocZ"))
					.ObjectID = CInt(oData("PlanetID"))
					.ObjTypeID = ObjectType.ePlanet
					.OuterRingRadius = CInt(oData("RingOuterRadius"))
					.ParentSystem = GetEpicaSystem(CInt(oData("parentID")))
					'Set the parent system's Planet Index
					.ParentSystem.AddPlanetIndex(glPlanetUB)
					.PlanetName = StringToBytes(CStr(oData("PlanetName")))
					.PlanetRadius = CShort(oData("PlanetRadius"))
					.PlanetSizeID = CByte(oData("PlanetSizeID"))
					.PlanetTypeID = CByte(oData("PlanetTypeID"))
					.RingDiffuse = CInt(oData("RingDiffuse"))
					.RotationDelay = CShort(oData("RotationDelay"))
					.SurfaceTemperature = CByte(oData("SurfaceTemperature"))
					.Vegetation = CByte(oData("Vegetation"))
					.blOriginalMineralQuantity = CLng(oData("OrigMinQty"))
					.ySentGNSLowRes = CByte(oData("SentGNSLowRes"))
					.lPrimaryComposition = CInt(oData("PrimaryComposition"))

                    .RingMineralConcentration = CInt(oData("RingMineralConcentration"))
                    .RingMineralID = CInt(oData("RingMineralID"))

					.PlayerSpawns = CInt(oData("PlayerSpawns"))
					.SpawnLocked = CByte(oData("SpawnLocked")) <> 0
					.OwnerID = CInt(oData("OwnerID"))

					glPlanetIdx(glPlanetUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Wormhole objects
			LogEvent(LogEventType.Informational, "Loading Wormholes...")
			sSQL = "SELECT * FROM tblWormhole"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glWormholeUB = glWormholeUB + 1
				ReDim Preserve goWormhole(glWormholeUB)
				ReDim Preserve glWormholeIdx(glWormholeUB)
				goWormhole(glWormholeUB) = New Wormhole()
				With goWormhole(glWormholeUB)
                    .WormholeFlags = CInt(oData("WormholeFlag"))
					.LocX1 = CInt(oData("LocX1"))
					.LocY1 = CInt(oData("LocY1"))
					.LocX2 = CInt(oData("LocX2"))
					.LocY2 = CInt(oData("LocY2"))
					.ObjectID = CInt(oData("WormholeID"))
					.ObjTypeID = ObjectType.eWormhole
					.StartCycle = CInt(oData("StartCycle"))
					.System1 = GetEpicaSystem(CInt(oData("System1ID")))
					.System2 = GetEpicaSystem(CInt(oData("System2ID")))
					glWormholeIdx(glWormholeUB) = .ObjectID

					If .System1 Is Nothing = False Then .System1.AddWormholeReference(goWormhole(glWormholeUB))
					If .System2 Is Nothing = False Then .System2.AddWormholeReference(goWormhole(glWormholeUB))
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			If goAureliusAI Is Nothing Then goAureliusAI = New AureliusAI()
			For X As Int32 = 0 To glSystemUB
				If glSystemIdx(X) <> -1 AndAlso goSystem(X).SystemType = 255 Then goAureliusAI.AddSystemAIEnvir(goSystem(X))
			Next X

			bResult = True
		Catch
			LogEvent(LogEventType.CriticalError, "LoadGeographical Error: " & Err.Description)
		Finally
			If oData Is Nothing = False Then
				oData.Close()
			End If
			oData = Nothing
			oComm = Nothing
		End Try
		Return bResult
    End Function

    Private Function LoadPlayers() As Boolean
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim oComm As OleDb.OleDbCommand
        Dim sSQL As String
        Dim X As Int32
        Dim bResult As Boolean = False

        Try
            'Load our guild objects
            LogEvent(LogEventType.Informational, "Loading Guilds...")
            sSQL = "SELECT * FROM tblGuild"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                glGuildUB += 1
                ReDim Preserve glGuildIdx(glGuildUB)
                ReDim Preserve goGuild(glGuildUB)
                goGuild(glGuildUB) = New Guild()
                With goGuild(glGuildUB)
                    .ObjectID = CInt(oData("GuildID"))
                    .ObjTypeID = ObjectType.eGuild

                    .yGuildName = StringToBytes(CStr(oData("GuildName")))
                    .sSearchableName = BytesToString(.yGuildName).Trim.ToUpper
                    .yMOTD = StringToBytes(CStr(oData("MOTD")))
                    .yBillboard = StringToBytes(CStr(oData("Billboard")))
                    Dim lValue As Int32 = CInt(oData("DateFormed"))
                    If lValue = 0 Then .dtFormed = Date.MinValue Else .dtFormed = GetDateFromNumber(lValue)
                    lValue = CInt(oData("DateLastRuleChange"))
                    If lValue = 0 Then .dtLastGuildRuleChange = Date.MinValue Else .dtLastGuildRuleChange = GetDateFromNumber(lValue)
                    lValue = CInt(oData("DateTaxed"))
                    If lValue = 0 Then .dtLastTaxInterval = Date.MinValue Else .dtLastTaxInterval = GetDateFromNumber(lValue)

                    .lGuildHallID = CInt(oData("GuildHallID"))
                    .lBaseGuildFlags = CType(CInt(oData("BaseGuildFlags")), elGuildFlags)
                    .yState = CType(CByte(oData("GuildState")), eyGuildState)
                    .lIcon = CInt(oData("GuildIcon"))
                    .blTreasury = CLng(oData("Treasury"))

                    .iRecruitFlags = CType(CShort(oData("RecruitFlags")), eiRecruitmentFlags)
                    .yVoteWeightType = CType(CByte(oData("VoteWeightType")), eyVoteWeightType)
                    .yGuildTaxRateInterval = CType(CByte(oData("TaxRateInterval")), eyGuildInterval)
                    .yGuildTaxBaseMonth = CByte(oData("TaxBaseMonth"))
                    .yGuildTaxBaseDay = CByte(oData("TaxBaseDay"))

                    glGuildIdx(glGuildUB) = .ObjectID
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Load our ranks
            LogEvent(LogEventType.Informational, "Loading Ranks...")
            sSQL = "SELECT * FROM tblRank ORDER BY GuildID, PeckingOrder"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            Dim oGuild As Guild = Nothing
            While oData.Read
                Dim lGuildID As Int32 = CInt(oData("GuildID"))
                If oGuild Is Nothing OrElse oGuild.ObjectID <> lGuildID Then oGuild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing = False Then
                    Dim oRank As New GuildRank()
                    With oRank
                        .lRankID = CInt(oData("RankID"))
                        .lRankPermissions = CInt(oData("RankRights"))
                        .lVoteStrength = CInt(oData("VotingWeight"))
                        .yPosition = CByte(oData("PeckingOrder"))
                        .yRankName = StringToBytes(CStr(oData("RankName")))
                        .TaxRateFlat = CInt(oData("TaxFlat"))
                        .TaxRatePercType = CType(CInt(oData("TaxPercType")), eyGuildTaxPercType)
                        .TaxRatePercentage = CByte(oData("TaxPerc"))
                    End With
                    oGuild.AddRank(oRank)
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Load our guild events
            LogEvent(LogEventType.Informational, "Loading Guild Events...")
            sSQL = "SELECT * FROM tblGuildEvent ORDER BY GuildID, StartsAt"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim lGuildID As Int32 = CInt(oData("GuildID"))
                If oGuild Is Nothing OrElse oGuild.ObjectID <> lGuildID Then oGuild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing = False Then
                    Dim oEvent As GuildEvent = New GuildEvent()
                    With oEvent
                        Dim lValue As Int32 = CInt(oData("PostedOn"))
                        If lValue = 0 Then .dtPostedOn = Date.MinValue Else .dtPostedOn = GetDateFromNumber(lValue)
                        lValue = CInt(oData("StartsAt"))
                        If lValue = 0 Then .dtStartsAt = Date.MinValue Else .dtStartsAt = GetDateFromNumber(lValue)
                        .EventID = CInt(oData("EventID"))

                        .lDuration = CInt(oData("Duration"))
                        .lPostedBy = CInt(oData("PostedByID"))
                        .ParentGuild = oGuild
                        .yDetails = StringToBytes(CStr(oData("Details")))
                        .yEventIcon = CByte(oData("EventIcon"))
                        .yEventType = CByte(oData("EventType"))
                        .yMembersCanAccept = CByte(oData("MembersCanAccept"))
                        .yRecurrence = CByte(oData("Recurrence"))
                        .ySendAlerts = CByte(oData("SendAlerts"))
                        .yTitle = StringToBytes(CStr(oData("EventTitle")))
                    End With
                    oGuild.AddEvent(oEvent)
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'And now the player objects
            LogEvent(LogEventType.Informational, "Loading Players...")
            sSQL = "SELECT * FROM tblPlayer ORDER BY PlayerID"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim i As Int32 = CInt(oData("PlayerID"))
                CheckPlayerExtent(i)
                'glPlayerUB = glPlayerUB + 1
                'ReDim Preserve goPlayer(glPlayerUB)
                'ReDim Preserve glPlayerIdx(glPlayerUB)
                'goPlayer(glPlayerUB) = New Player()
                goPlayer(i) = New Player()
                With goPlayer(i) 'With goPlayer(glPlayerUB)
                    .ServerIndex = i 'glPlayerUB
                    .CommEncryptLevel = CShort(oData("CommEncryptLevel"))
                    .EmpireName = StringToBytes(CStr(oData("EmpireName")))
                    .EmpireTaxRate = CByte(oData("EmpireTaxRate"))
                    .ObjectID = i 'CInt(oData("PlayerID"))
                    .ObjTypeID = ObjectType.ePlayer
                    '.oSenate = GetEpicaSenate(CInt(oData("SenateID")))
                    .sPlayerName = UCase$(CStr(oData("PlayerName")))
                    .sPlayerNameProper = CStr(oData("PlayerName"))
                    .PlayerName = StringToBytes(CStr(oData("PlayerName")))
                    StringToBytes(CStr(oData("PlayerPassword"))).CopyTo(.PlayerPassword, 0)
                    StringToBytes(CStr(oData("PlayerUserName"))).CopyTo(.PlayerUserName, 0)
                    .RaceName = StringToBytes(CStr(oData("RaceName")))
                    .lLastViewedEnvir = CInt(oData("LastViewedID"))
                    .iLastViewedEnvirType = CShort(oData("LastViewedTypeID"))
                    .BaseMorale = CByte(oData("BaseMorale"))
                    .lStartedEnvirID = CInt(oData("StartedEnvirID"))
                    .iStartedEnvirTypeID = CShort(oData("StartedEnvirTypeID"))
                    glPlayerIdx(i) = .ObjectID 'glPlayerIdx(glPlayerUB) = .ObjectID
                    .blCredits = CLng(oData("Credits"))
                    .AccountStatus = CInt(oData("AccountStatus"))
                    If oData("StatusFlags") Is DBNull.Value Then
                        .lStatusFlags = 0
                    Else
                        .lStatusFlags = CInt(oData("StatusFlags"))
                    End If

                    .yPlayerPhase = CType(CInt(oData("PlayerPhase")), eyPlayerPhase)
                    .lTutorialStep = CInt(oData("TutorialStep"))

                    .LastLogin = GetDateFromNumber(CInt(oData("LastLogin")))
                    Dim lTempVal As Int32 = CInt(oData("LastGuildMembership"))
                    If lTempVal = 0 Then .dtLastGuildMembership = Date.MinValue Else .dtLastGuildMembership = GetDateFromNumber(lTempVal)
                    .TotalPlayTime = CInt(oData("TotalPlayTime"))

                    lTempVal = CInt(oData("DateAccountWentActive"))
                    Dim dtActive As Date = Date.MinValue
                    If lTempVal = 0 Then dtActive = Date.MinValue Else dtActive = GetDateFromNumber(lTempVal)

                    lTempVal = CInt(oData("DateAccountWentInactive"))
                    If lTempVal = 0 Then .dtAccountWentInactive = Date.MinValue Else .dtAccountWentInactive = GetDateFromNumber(lTempVal)
                    If .dtAccountWentInactive < dtActive Then .dtAccountWentInactive = Date.MinValue

                    .lStartLocX = CInt(oData("StartLocX"))
                    .lStartLocZ = CInt(oData("StartLocZ"))
                    .PirateStartLocX = CInt(oData("PirateStartLocX"))
                    .PirateStartLocZ = CInt(oData("PirateStartLocZ"))

                    .lCelebrationEnds = glCurrentCycle + CInt(oData("CelebrationEnds"))
                    .lWarSentiment = CInt(oData("WarSentiment"))
                    .BadWarDecCPIncrease = CInt(oData("BadWarDecCPIncrease"))
                    .BadWarDecCPIncreaseEndCycle = glCurrentCycle + CInt(oData("BadWarDecCPIncreaseEndCycle"))
                    .BadWarDecMoralePenalty = CInt(oData("BadWarDecMoralePenalty"))
                    .BadWarDecMoralePenaltyEndCycle = glCurrentCycle + CInt(oData("BadWarDecMoralePenaltyEndCycle"))

                    .PlayedTimeAtEndOfWaves = CInt(oData("PlayedTimeAtEndOfWaves"))
                    .PlayedTimeInTutorialOne = CInt(oData("PlayedTimeInTutorialOne"))
                    .PlayedTimeWhenFirstWave = CInt(oData("PlayedTimeWhenFirstWave"))
                    '.PlayedTimeWhenTimerStarted = CInt(oData("PlayedTimeWhenTimerStarted"))
                    .TutorialPhaseWaves = CInt(oData("TutorialPhaseWaves"))

                    .lHangarManMult = CInt(oData("HangarManMult"))
                    .SpecTechCostMult = CSng(oData("SpecTechCostMult"))

                    .DeathBudgetBalance = CInt(oData("DeathBudgetBalance"))
                    .DeathBudgetFundsRemaining = CInt(oData("DeathBudgetFunds"))
                    Dim lTemp As Int32 = CInt(oData("DeathBudgetCycles"))
                    If lTemp > 0 Then
                        .DeathBudgetEndTime = glCurrentCycle + lTemp
                    Else : .DeathBudgetEndTime = Int32.MinValue
                    End If

                    .lGuildID = CInt(oData("GuildID"))
                    .lGuildRankID = CInt(oData("GuildRankID"))
                    .lJoinedGuildOn = CInt(oData("JoinedGuildOn"))
                    .yPlayerTitle = CByte(oData("PlayerTitle"))
                    .yExTitleNew = CByte(oData("ExTitleNew"))
                    .yCustomTitle = CByte(oData("CustomTitle"))
                    .lCustomTitlePermission = CInt(oData("CustomTitlePermission"))
                    lTemp = CInt(oData("ExTitleEnd"))
                    If lTemp = 0 Then .dtExTitleEnd = Date.MinValue Else .dtExTitleEnd = GetDateFromNumber(lTemp)
                    .lPlayerIcon = CInt(oData("PlayerIcon"))
                    .GuaranteedSpecialTechID = CInt(oData("GuaranteedSpecialTechID"))
                    .ySpawnSystemSetting = CByte(oData("SpawnSystemSetting"))
                    '.lLastWPUpkeep = CInt(oData("LastWPUpkeep"))
                    '.lLastGuildShareUpkeep = CInt(oData("LastGuildShareUpkeep"))
                    '.blWarpoints = CLng(oData("Warpoints"))
                    '.blWarpointsAllTime = CLng(oData("WarpointsAllTime"))
                    '.lLastWPUpkeepTime = CInt(oData("LastWPUpkeepTime"))
                    'If .lLastWPUpkeepTime <> 0 Then .dtLastWPUpkeep = GetDateFromNumber(.lLastWPUpkeepTime)

                    .yGender = CByte(oData("PlayerGender"))
                    .blMaxPopulation = CLng(oData("MaxTotalPop"))
                    .blDBPopulation = CLng(oData("DBTotalPop"))

                    If .iStartedEnvirTypeID = ObjectType.ePlanet Then
                        Dim oPlanet As Planet = GetEpicaPlanet(.lStartedEnvirID)
                        .InMyDomain = (oPlanet Is Nothing = False) AndAlso (oPlanet.InMyDomain = True)
                    ElseIf .iStartedEnvirTypeID = ObjectType.eSolarSystem Then
                        Dim oSystem As SolarSystem = GetEpicaSystem(.lStartedEnvirID)
                        .InMyDomain = (oSystem Is Nothing = False) AndAlso (oSystem.InMyDomain = True)
                    End If

                    If .InMyDomain = True Then
                        Dim yMsg(5) As Byte
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerPrimaryOwner).CopyTo(yMsg, lPos) : lPos += 2
                        System.BitConverter.GetBytes(i).CopyTo(yMsg, lPos) : lPos += 4
                        goMsgSys.SendMsgToOperator(yMsg)
                    End If

                    ReDim .ExternalEmailAddress(254)
                    If oData("EmailAddress") Is DBNull.Value = False Then
                        StringToBytes(CStr(oData("EmailAddress"))).CopyTo(.ExternalEmailAddress, 0)
                        .iEmailSettings = CShort(oData("EmailSettings"))
                    Else : .iEmailSettings = 0
                    End If
                    .iInternalEmailSettings = CShort(oData("InternalEmailSettings"))

                    ReDim .lSlotID(4)
                    ReDim .ySlotState(4)
                    ReDim .lFactionID(2)

                    For X = 0 To 4
                        .lSlotID(X) = -1
                        .ySlotState(X) = 0
                        .lSlotID(X) = CInt(oData("Slot" & (X + 1) & "ID"))
                        .ySlotState(X) = CByte(oData("Slot" & (X + 1) & "State"))
                    Next X
                    For X = 0 To 2
                        .lFactionID(X) = -1
                        .lFactionID(X) = CInt(oData("Faction" & (X + 1) & "ID"))
                    Next X

                    .lIronCurtainPlanet = CInt(oData("IronCurtainPlanetID"))

                    'Set whether the player is dead or not... when colonies are added later they will reset the player death indicator
                    'If .blCredits <> 0 AndAlso .blCredits <> 100000 Then
                    '	.HandleCheckForPlayerDeath()
                    'End If
                    If .lPlayerIcon <> 0 Then .PlayerIsDead = True
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            LogEvent(LogEventType.Informational, "Loading Player CP Penalties...")
            sSQL = "SELECT * FROM tblPlayerCPPenalty"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True
                Dim lPlayerID As Int32 = CInt(oData("PlayerID"))
                Dim lDecPlayerID As Int32 = CInt(oData("DecPlayerID"))
                Dim yPenalty As Byte = CByte(oData("Penalty"))

                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing Then Continue While
                Dim oDecPlayer As Player = GetEpicaPlayer(lDecPlayerID)
                If oDecPlayer Is Nothing Then Continue While

                Dim oItem As New PlayerCPPenalty()
                With oItem
                    .oPlayer = oPlayer
                    .oDecPlayer = oDecPlayer
                    .yPenalty = yPenalty
                End With
                oPlayer.AddCPPenalty(oItem)
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            LogEvent(LogEventType.Informational, "Loading Player Claimables...")
            sSQL = "SELECT * FROM tblClaimable ORDER BY PlayerID"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            Dim oLinkToPlayer As Player = Nothing
            While oData.Read
                Dim lPlayerID As Int32 = CInt(oData("PlayerID"))
                If oLinkToPlayer Is Nothing OrElse oLinkToPlayer.ObjectID <> lPlayerID Then
                    oLinkToPlayer = GetEpicaPlayer(lPlayerID)
                    If oLinkToPlayer Is Nothing Then Continue While
                End If
                Dim oClaimable As New Claimable()
                With oClaimable
                    .lID = CInt(oData("ObjectID"))
                    .iTypeID = CShort(oData("ObjTypeID"))
                    .lOfferCode = CInt(oData("OfferCode"))
                    .lPlayerID = lPlayerID
                    .yClaimFlag = CByte(oData("ClaimFlag"))
                    .yItemName = StringToBytes(CStr(oData("ItemName")))
                    .blQuantity = CLng(oData("ClaimQuantity"))
                End With
                ReDim Preserve oLinkToPlayer.Claimables(oLinkToPlayer.Claimables.GetUpperBound(0) + 1)
                oLinkToPlayer.Claimables(oLinkToPlayer.Claimables.GetUpperBound(0)) = oClaimable
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            For X = 0 To glPlayerUB
                If glPlayerIdx(X) > -1 AndAlso goPlayer(X).InMyDomain = True Then goPlayer(X).ReverifySlots()
            Next X

            'load our guild members
            LogEvent(LogEventType.Informational, "Loading Guild Members...")
            sSQL = "SELECT * FROM tblGuildMember ORDER BY GuildID"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim lGuildID As Int32 = CInt(oData("GuildID"))
                If oGuild Is Nothing OrElse oGuild.ObjectID <> lGuildID Then oGuild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing = False Then
                    Dim oMember As New GuildMember()
                    With oMember
                        .lMemberID = CInt(oData("MemberID"))
                        .yMemberState = CType(CByte(oData("MemberState")), GuildMemberState)
                    End With
                    oGuild.AddMember(oMember)
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Load our guild votes
            LogEvent(LogEventType.Informational, "Loading Guild Votes...")
            sSQL = "SELECT * FROM tblGuildVote ORDER BY GuildID, VoteStarts"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim lGuildID As Int32 = CInt(oData("GuildID"))
                If oGuild Is Nothing OrElse oGuild.ObjectID <> lGuildID Then oGuild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing = False Then
                    Dim oVote As GuildVote = New GuildVote()
                    With oVote
                        Dim lValue As Int32 = CInt(oData("VoteStarts"))
                        If lValue = 0 Then .dtVoteStarts = Date.MinValue Else .dtVoteStarts = GetDateFromNumber(lValue)
                        .lNewValue = CInt(oData("NewValueNumber"))
                        .lSelectedItem = CInt(oData("SelectedItem"))
                        .lVoteDuration = CInt(oData("VoteDuration"))
                        .ProposedByID = CInt(oData("ProposedByID"))
                        .VoteID = CInt(oData("VoteID"))
                        .yNewValueText = StringToBytes(CStr(oData("NewValueText")))
                        .ySummary = StringToBytes(CStr(oData("VoteSummary")))
                        .yTypeOfVote = CByte(oData("TypeOfVote"))
                        .yVoteState = CByte(oData("VoteState"))
                    End With
                    oGuild.AddVote(oVote)
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Load our vote values
            LogEvent(LogEventType.Informational, "Loading Guild Vote Values...")
            sSQL = "SELECT tblGuildMemberVote.VoteID, tblGuildMemberVote.MemberID, tblGuildMemberVote.VoteValue, " & _
             "tblGuildVote.GuildID FROM tblGuildMemberVote LEFT OUTER JOIN tblGuildVote ON tblGuildMemberVote.VoteID = " & _
             "tblGuildVote.VoteID ORDER BY tblGuildVote.GuildID, tblGuildMemberVote.VoteID"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim lGuildID As Int32 = CInt(oData("GuildID"))
                If oGuild Is Nothing OrElse oGuild.ObjectID <> lGuildID Then oGuild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing = False Then
                    Dim lVoteID As Int32 = CInt(oData("VoteID"))
                    Dim lMemberID As Int32 = CInt(oData("MemberID"))
                    Dim yVoteValue As eyVoteValue = CType(CByte(oData("VoteValue")), eyVoteValue)
                    oGuild.AddMemberVoteValue(lVoteID, lMemberID, yVoteValue)
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            LogEvent(LogEventType.Informational, "Loading Guild Trans Logs...")
            sSQL = "SELECT * FROM tblGuildTransLog ORDER BY TransDate DESC, TDMilli DESC"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim lGuildID As Int32 = CInt(oData("GuildID"))
                oGuild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing Then Continue While

                oGuild.lBankLogUB += 1
                ReDim Preserve oGuild.oBankLog(oGuild.lBankLogUB)
                oGuild.oBankLog(oGuild.lBankLogUB) = New GuildTransLog()
                With oGuild.oBankLog(oGuild.lBankLogUB)
                    .lGuildID = lGuildID
                    .lPlayerID = CInt(oData("PlayerID"))
                    .lTransDate = CInt(oData("TransDate"))
                    .blBalance = CLng(oData("BalanceAmount"))
                    .blAmount = CLng(oData("TransAmount"))
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            ''load our guild rels
            'LogEvent(LogEventType.Informational, "Loading Guild Rels...")
            'sSQL = "SELECT * FROM tblGuildRel ORDER BY GuildID"
            'oComm = New OleDb.OleDbCommand(sSQL, goCN)
            'oData = oComm.ExecuteReader(CommandBehavior.Default)
            'While oData.Read
            '    'EntityID, EntityTypeID, GuildID, RelTowardsUs, RelTowardsThem, WhoMadeFirstContact, " & _
            '    '"WhoFirstContactWasMadeWith, LocationIDOfFC, LocTypeIDOfFC, LocXOfFC, LocZOfFC, DateOfFC, CustomNotes
            '    Dim lGuildID As Int32 = CInt(oData("GuildID"))
            '    If oGuild Is Nothing OrElse oGuild.ObjectID <> lGuildID Then oGuild = GetEpicaGuild(lGuildID)
            '    If oGuild Is Nothing = False Then
            '        Dim oRel As GuildRel = New GuildRel()
            '        With oRel
            '            Dim lValue As Int32 = CInt(oData("DateOfFC"))
            '            If lValue = 0 Then .dtWhenFirstContactMade = Date.MinValue Else .dtWhenFirstContactMade = GetDateFromNumber(lValue)
            '            .iEntityTypeID = CShort(oData("EntityTypeID"))
            '            .iLocationTypeIDOfFC = CShort(oData("LocTypeIDOfFC"))
            '            .lEntityID = CInt(oData("EntityID"))
            '            .lLocationIDOfFC = CInt(oData("LocationIDOfFC"))
            '            .lLocXOfFC = CInt(oData("LocXOfFC"))
            '            .lLocZOfFC = CInt(oData("LocZOfFC"))
            '            .lWhoFirstContactWasMadeWith = CInt(oData("WhoFirstContactWasMadeWith"))
            '            .lWhoMadeFirstContact = CInt(oData("WhoMadeFirstContact"))
            '            .yNote = StringToBytes(CStr(oData("CustomNotes")))
            '            .yRelTowardsThem = CByte(oData("RelTowardsThem"))
            '            .yRelTowardsUs = CByte(oData("RelTowardsUs"))
            '        End With
            '        oGuild.AddRel(oRel)
            '    End If
            'End While
            'oData.Close()
            'oData = Nothing
            'oComm = Nothing

            'Guild Event Acceptance
            LogEvent(LogEventType.Informational, "Loading Guild Event Acceptance...")
            sSQL = "SELECT * FROM tblGuildEventAcceptance"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                'EventID, PlayerID, Acceptance
                Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
                If oPlayer Is Nothing = False Then
                    If oPlayer.oGuild Is Nothing = False Then
                        Dim oEvent As GuildEvent = oPlayer.oGuild.GetEvent(CInt(oData("EventID")))
                        If oEvent Is Nothing = False Then
                            oEvent.SetPlayerAcceptance(oPlayer.ObjectID, CByte(oData("Acceptance")))
                        End If
                    End If
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            Dim oSketchPads(-1) As SketchPad
            Dim lSketchPadEvents(-1) As Int32
            Dim lSketchPadGuilds(-1) As Int32
            Dim lSketchPadUB As Int32 = -1
            'Sketchpad base class
            LogEvent(LogEventType.Informational, "Loading SketchPads...")
            sSQL = "SELECT COUNT(*) FROM tblSketchPad"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            If oData.Read = True Then
                Dim lTmp As Int32 = CInt(oData(0))
                ReDim oSketchPads(lTmp)
                ReDim lSketchPadEvents(lTmp)
                ReDim lSketchPadGuilds(lTmp)
            End If
            oData.Close()
            oData = Nothing
            oComm = Nothing
            sSQL = "SELECT tblSketchPad.*, tblGuildEvent.GuildID FROM tblSketchPad LEFT OUTER JOIN tblGuildEvent ON tblSketchPad.EventID = tblGuildEvent.EventID ORDER BY GuildID"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lSketchPadUB += 1
                If lSketchPadUB > oSketchPads.GetUpperBound(0) Then
                    ReDim Preserve oSketchPads(lSketchPadUB)
                    ReDim Preserve lSketchPadEvents(lSketchPadUB)
                    ReDim Preserve lSketchPadGuilds(lSketchPadUB)
                End If
                oSketchPads(lSketchPadUB) = New SketchPad
                lSketchPadEvents(lSketchPadUB) = CInt(oData("EventID"))
                lSketchPadGuilds(lSketchPadUB) = CInt(oData("GuildID"))
                With oSketchPads(lSketchPadUB)
                    .CameraAtX = CInt(oData("CameraAtX"))
                    .CameraAtY = CInt(oData("CameraAtY"))
                    .CameraAtZ = CInt(oData("CameraAtZ"))
                    .CameraX = CInt(oData("CameraX"))
                    .CameraY = CInt(oData("CameraY"))
                    .CameraZ = CInt(oData("CameraZ"))
                    .iEnvirTypeID = CShort(oData("EnvirTypeID"))
                    .lEnvirID = CInt(oData("EnvirID"))
                    .lID = CInt(oData("SketchPadID"))
                    .ViewID = CInt(oData("ViewID"))
                    ReDim .yName(19)
                    StringToBytes(CStr(oData("SketchPadName"))).CopyTo(.yName, 0)
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Now, load the sketchpad items
            LogEvent(LogEventType.Informational, "Loading Sketchpad Items...")
            sSQL = "SELECT * FROM tblSketchPadItem"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                'EventID, PlayerID, Acceptance
                Dim lSketchPadID As Int32 = CInt(oData("SketchpadID"))
                For X = 0 To lSketchPadUB
                    If oSketchPads(X) Is Nothing = False AndAlso oSketchPads(X).lID = lSketchPadID Then
                        oSketchPads(X).AddSketchPadItem(CByte(oData("ItemType")), CSng(oData("PtAX")), CSng(oData("PtAY")), _
                         CSng(oData("PtBX")), CSng(oData("PtBY")), CByte(oData("ClrVal")), CStr(oData("ItemText")))
                        Exit For
                    End If
                Next X
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Now, associate our sketchpads to events
            For X = 0 To lSketchPadUB
                Dim lGuildID As Int32 = lSketchPadGuilds(X)
                If oGuild Is Nothing OrElse oGuild.ObjectID <> lGuildID Then oGuild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing = False Then
                    Dim oEvent As GuildEvent = oGuild.GetEvent(lSketchPadEvents(X))
                    If oEvent Is Nothing = False Then oEvent.AddOrSetAttachment(oSketchPads(X))
                End If
            Next X

            'Now, the event advanced config items
            LogEvent(LogEventType.Informational, "Loading Event Advanced Configs...")
            sSQL = "SELECT tblEventAdvancedConfig.*, tblGuildEvent.GuildID FROM tblEventAdvancedConfig LEFT OUTER JOIN tblGuildEvent ON tblEventAdvancedConfig.EventID = tblGuildEvent.EventID ORDER BY GuildID"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim lGuildID As Int32 = CInt(oData("GuildID"))
                If oGuild Is Nothing OrElse oGuild.ObjectID <> lGuildID Then oGuild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing = False Then
                    Dim oEvent As GuildEvent = oGuild.GetEvent(CInt(oData("EventID")))
                    If oEvent Is Nothing = False Then
                        Dim lLaps As Int32 = 0
                        If oData("Laps") Is DBNull.Value = False Then lLaps = CInt(oData("Laps"))
                        If lLaps <> 0 Then
                            'race
                            Dim oRC As New RaceConfig
                            With oRC
                                .blEntryFee = CLng(oData("EntryFee"))
                                .lEventID = oEvent.EventID
                                .lMaxHull = CInt(oData("MaxHull"))
                                .lMinHull = CInt(oData("MinHull"))
                                .yFirstPlace = CByte(oData("FirstPlace"))
                                .yGroundOnly = CByte(oData("GroundOnly"))
                                .yGuildTake = CByte(oData("GuildTake"))
                                .yLaps = CByte(oData("Laps"))
                                .yMinRacers = CByte(oData("MinRacers"))
                                .ySecondPlace = CByte(oData("SecondPlace"))
                                .yThirdPlace = CByte(oData("ThirdPlace"))
                            End With
                            oEvent.oAdvancedConfig = oRC
                        Else
                            'tournament
                            Dim oTrn As New TournamentConfig
                            With oTrn
                                .blEntryFee = CLng(oData("EntryFee"))
                                .lEventID = oEvent.EventID
                                .lMapID = CInt(oData("MapID"))
                                .yGuildTake = CByte(oData("GuildTake"))
                                .yMaxAir = CByte(oData("MaxAir"))
                                .yMaxGround = CByte(oData("MaxGround"))
                                .yMaxPlayers = CByte(oData("MaxPlayers"))
                                .yMaxUnits = CByte(oData("MaxUnits"))
                            End With
                            oEvent.oAdvancedConfig = oTrn
                        End If
                    End If
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            LogEvent(LogEventType.Informational, "Loading Event Advanced Config Items...")
            sSQL = "SELECT tblEventAdvancedConfigItem.*, tblGuildEvent.GuildID FROM tblEventAdvancedConfigItem LEFT OUTER JOIN tblGuildEvent ON tblEventAdvancedConfigItem.EventID = tblGuildEvent.EventID ORDER BY GuildID, ItemPos"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim lGuildID As Int32 = CInt(oData("GuildID"))
                If oGuild Is Nothing OrElse oGuild.ObjectID <> lGuildID Then oGuild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing = False Then
                    Dim oEvent As GuildEvent = oGuild.GetEvent(CInt(oData("EventID")))
                    If oEvent Is Nothing = False Then
                        Try
                            With CType(oEvent.oAdvancedConfig, RaceConfig)
                                .lWPUB += 1
                                ReDim Preserve .uWP(.lWPUB)
                                With .uWP(.lWPUB)
                                    .EnvirID = CInt(oData("EnvirID"))
                                    .EnvirTypeID = CShort(oData("EnvirTypeID"))
                                    .lX = CInt(oData("XPos"))
                                    .lZ = CInt(oData("ZPos"))
                                End With
                            End With
                        Catch
                        End Try
                    End If
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Now, set our Domain Indications
            LogEvent(LogEventType.Informational, "Setting Player Domain Responsibilities...")
            For X = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    If goPlayer(X).lStartedEnvirID > 0 AndAlso goPlayer(X).iStartedEnvirTypeID > 0 Then
                        If goPlayer(X).iStartedEnvirTypeID = ObjectType.ePlanet Then
                            Dim oPlanet As Planet = GetEpicaPlanet(goPlayer(X).lStartedEnvirID)
                            If oPlanet Is Nothing = False Then goPlayer(X).InMyDomain = oPlanet.InMyDomain
                        ElseIf goPlayer(X).iStartedEnvirTypeID = ObjectType.eSolarSystem Then
                            Dim oSystem As SolarSystem = GetEpicaSystem(goPlayer(X).lStartedEnvirID)
                            If oSystem Is Nothing = False Then goPlayer(X).InMyDomain = oSystem.InMyDomain
                        End If
                    End If
                End If
            Next X

            ''Set our Senate Leaders
            'LogEvent(LogEventType.Informational, "Associating Senate Leaders to Senates...")
            'For X = 0 To glSenateUB
            '	goSenate(X).SenateLeader = GetEpicaPlayer(lSenateLeaders(X))
            'Next X

            'Load our Player Wormholes
            LogEvent(LogEventType.Informational, "Loading Player Wormholes...")
            sSQL = "SELECT * FROM tblPlayerWormhole"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                If oData("PlayerID") Is DBNull.Value = True OrElse oData("WormholeID") Is DBNull.Value = True Then Continue While
                Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
                If oPlayer Is Nothing = False Then
                    Dim oWormhole As Wormhole = GetEpicaWormhole(CInt(oData("WormholeID")))
                    If oWormhole Is Nothing = False Then
                        oPlayer.AddWormholeKnowledge(oWormhole, False, Nothing, False)
                    End If
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            bResult = True
        Catch
            LogEvent(LogEventType.CriticalError, "LoadPlayers Error: " & Err.Description)
        Finally
            If oData Is Nothing = False Then
                oData.Close()
            End If
            oData = Nothing
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Private Function LoadPlayerDetails(ByVal sPlayerInStr As String) As Boolean
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim oComm As OleDb.OleDbCommand
		Dim sSQL As String
		Dim bResult As Boolean = False

		Try
			'Load our Player Intels
			LogEvent(LogEventType.Informational, "Loading Player Intel...")
			sSQL = "SELECT * FROM tblPlayerIntel WHERE PlayerID IN " & sPlayerInStr & " OR TargetID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
				If oPlayer Is Nothing = False Then
					Dim oPlayerIntel As New PlayerIntel()
					With oPlayerIntel
						.DiplomacyScore = CInt(oData("DiplomacyScore"))
						.DiplomacyUpdate = CInt(oData("DiplomacyUpdate"))
						.lPlayerID = CInt(oData("PlayerID"))
						.MilitaryScore = CInt(oData("MilitaryScore"))
						.MilitaryUpdate = CInt(oData("MilitaryUpdate"))
						.ObjectID = CInt(oData("TargetID"))
						.ObjTypeID = ObjectType.ePlayerIntel
						.PopulationScore = CInt(oData("PopulationScore"))
						.PopulationUpdate = CInt(oData("PopulationUpdate"))
						.ProductionScore = CInt(oData("ProductionScore"))
						.ProductionUpdate = CInt(oData("ProductionUpdate"))
						.StaticVariables = CInt(oData("StaticVariables"))
						.TechnologyScore = CInt(oData("TechnologyScore"))
						.TechnologyUpdate = CInt(oData("TechnologyUpdate"))
						.WealthScore = CInt(oData("WealthScore"))
						.WealthUpdate = CInt(oData("WealthUpdate"))
					End With
					oPlayer.SetPlayerIntel(oPlayerIntel)
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Player Mineral Properties...")
			sSQL = "SELECT * FROM tblPlayerMineralProperty WHERE PlayerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
				If oPlayer Is Nothing = False Then
                    Dim lIdx As Int32 = oPlayer.AddKnownProperty(CInt(oData("MineralPropertyID")), 0, False, False) 'CByte(oData("lDiscovered")))
					If lIdx > -1 Then oPlayer.moMinProperties(lIdx).SetDiscoveredLevel_NoEvents(CByte(oData("lDiscovered")))
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Player Minerals...")
			sSQL = "SELECT * FROM tblPlayerMineral WHERE PlayerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
				If oPlayer Is Nothing = False Then
					Dim lIdx As Int32 = oPlayer.AddPlayerMineral(CInt(oData("MineralID")), CInt(oData("bDiscovered")) <> 0, CInt(oData("PlayerMineralID")), True)
                    If lIdx <> -1 Then oPlayer.oPlayerMinerals(lIdx).yArchived = CByte(oData("bArchived"))
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Player Mineral Values (KNOWN)...")
			For X As Int32 = 0 To glPlayerUB
				If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True Then
					Dim sInClause As String = Player.GetPlayerMineralIDWhereClause(goPlayer(X))
					If sInClause Is Nothing = False AndAlso sInClause.Length > 0 Then
						sSQL = "SELECT * FROM tblPlayerMineralPropertyValue WHERE PlayerMineralID IN (" & sInClause & ")"
						oComm = New OleDb.OleDbCommand(sSQL, goCN)
						oData = oComm.ExecuteReader(CommandBehavior.Default)
						While oData.Read
							goPlayer(X).SetPlayerMineralProperty(CInt(oData("PlayerMineralPropertyValueID")), _
							  CInt(oData("PlayerMineralID")), CInt(oData("MineralPropertyID")), CInt(oData("PropertyValue")))
						End While
						oData.Close()
						oData = Nothing
						oComm = Nothing
					End If
				End If
			Next X

			bResult = True
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "LoadPlayerDetails: " & ex.Message)
		Finally
			If oData Is Nothing = False Then
				oData.Close()
			End If
			oData = Nothing
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Private Function LoadGoalsAndSkills() As Boolean
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim bResult As Boolean

		Try
			'Missions...
			LogEvent(LogEventType.Informational, "Loading Missions...")
			sSQL = "SELECT * FROM tblMission"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glMissionUB += 1
				ReDim Preserve goMission(glMissionUB)
				ReDim Preserve glMissionIdx(glMissionUB)
				goMission(glMissionUB) = New Mission()
				With goMission(glMissionUB)
					.MissionName = StringToBytes(CStr(oData("MissionName")))
					.MissionDesc = StringToBytes(CStr(oData("MissionDesc")))
					.lInfiltrationType = CType(oData("InfiltrationType"), eInfiltrationType)
					.ProgramControlID = CShort(oData("ProgramControlID"))
					.BaseEffect = CShort(oData("BaseEffect"))
					.Modifier = CShort(oData("Modifier"))
					.GoalUB = -1
					.ObjectID = CInt(oData("MissionID"))
					.ObjTypeID = ObjectType.eMission
					glMissionIdx(glMissionUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Mission Methods
			LogEvent(LogEventType.Informational, "Loading Mission Methods...")
			sSQL = "SELECT * FROM tblMissionMethod ORDER BY MethodName"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glMissionMethodUB += 1
				ReDim Preserve guMissionMethods(glMissionMethodUB)
				With guMissionMethods(glMissionMethodUB)
					.lID = CInt(oData("MM_ID"))
					ReDim .yName(19)
					StringToBytes(CStr(oData("MethodName"))).CopyTo(.yName, 0)
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Goals...
			LogEvent(LogEventType.Informational, "Loading Goals...")
			sSQL = "SELECT * FROM tblGoal ORDER BY MissionPhase, OrderNum"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glGoalUB = glGoalUB + 1
				ReDim Preserve goGoal(glGoalUB)
				ReDim Preserve glGoalIdx(glGoalUB)
				goGoal(glGoalUB) = New Goal()
				With goGoal(glGoalUB)
					.BaseTime = CInt(oData("BaseTime"))
					.GoalDesc = StringToBytes(CStr(oData("GoalDesc")))
					.GoalName = StringToBytes(CStr(oData("GoalName")))
					.ObjectID = CInt(oData("GoalID"))
					.ObjTypeID = ObjectType.eGoal
					.RiskOfDetection = CByte(oData("RiskOfDetection"))
					.MissionPhase = CType(oData("MissionPhase"), eMissionPhase)
					.SuccessProgCtrlID = CInt(oData("SuccessProgCtrlID"))
					.FailureProgCtrlID = CInt(oData("FailureProgCtrlID"))
					.OrderNum = CInt(oData("OrderNum"))
					glGoalIdx(glGoalUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Mission Goals
			LogEvent(LogEventType.Informational, "Loading Mission Goals...")
			sSQL = "SELECT * FROM tblMissionGoal"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oMission As Mission = GetEpicaMission(CInt(oData("MissionID")))
				If oMission Is Nothing = False Then
					Dim oGoal As Goal = GetEpicaGoal(CInt(oData("GoalID")))
					If oGoal Is Nothing = False Then oMission.AddGoal(oGoal, CInt(oData("MethodID")))
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Skills...
			LogEvent(LogEventType.Informational, "Loading Skills...")
			sSQL = "SELECT * FROM tblSkill"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glSkillUB = glSkillUB + 1
				ReDim Preserve goSkill(glSkillUB)
				ReDim Preserve glSkillIdx(glSkillUB)
				goSkill(glSkillUB) = New Skill()
				With goSkill(glSkillUB)
					.MaxVal = CByte(oData("MaxVal"))
					.MinVal = CByte(oData("MinVal"))
					.ObjectID = CInt(oData("SkillID"))
					.ObjTypeID = ObjectType.eSkill
					.SkillName = StringToBytes(CStr(oData("SkillName")))
					.SkillDesc = StringToBytes(CStr(oData("SkillDesc")))
					.SkillType = CByte(oData("SkillType"))
					glSkillIdx(glSkillUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Need to get goal skill requirements...
			LogEvent(LogEventType.Informational, "Loading Skillsets...")
			sSQL = "SELECT tblSkillSet.SkillSetID, tblSkillSet.GoalID, tblSkillSet.ProgramControlID, tblSkillSet_Skill.SkillID, " & _
			  "tblSkillSet_Skill.SuccessProgCtrlID, tblSkillSet_Skill.FailureProgCtrlID, tblSkillSet_Skill.PointRequirement, " & _
			  "tblSkillSet_Skill.AgentGroupingID, tblSkillSet_Skill.ToHitModifier FROM tblSkillSet LEFT JOIN tblSkillSet_Skill ON " & _
			  "tblSkillSet.SkillSetID = tblSkillSet_Skill.SkillSetID"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oGoal As Goal = GetEpicaGoal(CInt(oData("GoalID")))
				If oGoal Is Nothing = False Then
                    Dim oSkillSet As SkillSet = oGoal.GetOrAddSkillSet(CInt(oData("SkillSetID")), False)
					If oSkillSet Is Nothing = False Then
						oSkillSet.ProgramControlID = CInt(oData("ProgramControlID"))
						Dim oSkill As Skill = GetEpicaSkill(CInt(oData("SkillID")))
						If oSkill Is Nothing = False Then
							oSkillSet.AddSkill(oSkill, CShort(oData("ToHitModifier")), CInt(oData("SuccessProgCtrlID")), _
							  CInt(oData("FailureProgCtrlID")), CShort(oData("PointRequirement")), CShort(oData("AgentGroupingID")))
						End If
					End If
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Load the related skills data
			LogEvent(LogEventType.Informational, "Loading Related Skills...")
			sSQL = "SELECT * FROM tblRelatedSkill"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oTmpSkill As Skill = GetEpicaSkill(CInt(oData("FromSkillID")))
				If oTmpSkill Is Nothing = False Then
					oTmpSkill.lRelatedSkillUB += 1
					ReDim Preserve oTmpSkill.RelatedSkills(oTmpSkill.lRelatedSkillUB)
					With oTmpSkill.RelatedSkills(oTmpSkill.lRelatedSkillUB)
						.oToSkill = GetEpicaSkill(CInt(oData("ToSkillID")))
						.lToHitNumber = CInt(oData("ToHitNumber"))
					End With
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			bResult = True
		Catch
			LogEvent(LogEventType.CriticalError, "LoadStaticData Error: " & Err.Description)
		Finally
			If oData Is Nothing = False Then
				oData.Close()
			End If
			oData = Nothing
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Private Function LoadStaticMineralData() As Boolean
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim oComm As OleDb.OleDbCommand
		Dim sSQL As String
		Dim bResult As Boolean = False

		Try
			LogEvent(LogEventType.Informational, "Loading Mineral Properties...")
			sSQL = "SELECT * FROM tblMineralProperty ORDER BY MineralPropertyName"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glMineralPropertyUB += 1
				ReDim Preserve goMineralProperty(glMineralPropertyUB)
				ReDim Preserve glMineralPropertyIdx(glMineralPropertyUB)
				goMineralProperty(glMineralPropertyUB) = New MineralProperty
				With goMineralProperty(glMineralPropertyUB)
					.ObjectID = CInt(oData("MineralPropertyID"))
					.ObjTypeID = ObjectType.eMineralProperty
					.MaximumValue = CInt(oData("MaxValue"))
					.MinimumValue = CInt(oData("MinValue"))
					.MineralPropertyName = StringToBytes(CStr(oData("MineralPropertyName")))
					glMineralPropertyIdx(glMineralPropertyUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Mineral Property Value Ranges...")
			sSQL = "SELECT * FROM tblMineralValueRange ORDER BY ValueRangeMinValue"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lID As Int32 = CInt(oData("MineralPropertyID"))
				Dim oTmpMinProp As MineralProperty = GetEpicaMineralProperty(lID)
				If oTmpMinProp Is Nothing = False Then
					oTmpMinProp.AddValueRangeName(CInt(oData("ValueRangeMinValue")), CStr(oData("ValueRangeName")))
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Minerals and Alloys...")
			sSQL = "SELECT * FROM tblMineral"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glMineralUB = glMineralUB + 1
				ReDim Preserve goMineral(glMineralUB)
				ReDim Preserve glMineralIdx(glMineralUB)
				goMineral(glMineralUB) = New Mineral()
				With goMineral(glMineralUB)
					.ServerIndex = glMineralUB
                    .MineralName = StringToBytes(CStr(oData("MineralName")))
                    .MineralValue = CInt(oData("MineralValue"))
					.ObjectID = CInt(oData("MineralID"))
					.ObjTypeID = ObjectType.eMineral
					.lAlloyTechID = CInt(oData("AlloyTechID"))
					.lRarity = CInt(oData("Rarity"))
					glMineralIdx(glMineralUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Mineral ACTUAL Values...")
			For X As Int32 = 0 To glMineralUB
				sSQL = "SELECT * FROM tblMineralPropertyValue WHERE MineralID = " & goMineral(X).ObjectID
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				While oData.Read
					goMineral(X).SetPropertyValue(CInt(oData("MineralPropertyValueID")), CInt(oData("MineralPropertyID")), CInt(oData("PropertyValue")))
				End While
				oData.Close()
				oData = Nothing
				oComm = Nothing
			Next X

            LogEvent(LogEventType.Informational, "Loading Tech Version Numbers...")
            sSQL = "SELECT * FROM tblVersionRel"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oRel As New VersionRel(CInt(oData("VersionNumber")), CInt(oData("OtherVersion")), CInt(oData("NoisePerc")))
                Epica_Tech.AddVersionRel(oRel)
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

			bResult = True
		Catch
			LogEvent(LogEventType.CriticalError, "LoadMinerals Error: " & Err.Description)
		Finally
			If oData Is Nothing = False Then
				oData.Close()
			End If
			oData = Nothing
			oComm = Nothing
		End Try

		Return bResult
	End Function

	Public Function LoadSinglePlayer(ByVal lPlayerID As Int32) As Boolean
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim oComm As OleDb.OleDbCommand
		Dim sSQL As String
		Dim bResult As Boolean = False

		If GetEpicaPlayer(lPlayerID) Is Nothing = False Then Return True

		Try
			'And now the player objects
			'LogEvent(LogEventType.Informational, "Loading Players...")
			sSQL = "SELECT * FROM tblPlayer WHERE PlayerID = " & lPlayerID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As New Player()
				With oPlayer
					.ServerIndex = glPlayerUB
					.CommEncryptLevel = CShort(oData("CommEncryptLevel"))
					.EmpireName = StringToBytes(CStr(oData("EmpireName")))
					.EmpireTaxRate = CByte(oData("EmpireTaxRate"))
					.ObjectID = CInt(oData("PlayerID"))
					.ObjTypeID = ObjectType.ePlayer
					'.oSenate = GetEpicaSenate(CInt(oData("SenateID")))
					.sPlayerName = UCase$(CStr(oData("PlayerName")))
					.sPlayerNameProper = CStr(oData("PlayerName"))
					.PlayerName = StringToBytes(CStr(oData("PlayerName")))
					StringToBytes(CStr(oData("PlayerPassword"))).CopyTo(.PlayerPassword, 0)
					StringToBytes(CStr(oData("PlayerUserName"))).CopyTo(.PlayerUserName, 0)
					.RaceName = StringToBytes(CStr(oData("RaceName")))
					.lLastViewedEnvir = CInt(oData("LastViewedID"))
					.iLastViewedEnvirType = CShort(oData("LastViewedTypeID"))
					.BaseMorale = CByte(oData("BaseMorale"))
					.lStartedEnvirID = CInt(oData("StartedEnvirID"))
					.iStartedEnvirTypeID = CShort(oData("StartedEnvirTypeID"))
					.blCredits = CLng(oData("Credits"))
					.AccountStatus = CInt(oData("AccountStatus"))
					.yPlayerPhase = CType(CInt(oData("PlayerPhase")), eyPlayerPhase)
					.lTutorialStep = CInt(oData("TutorialStep"))
                    If oData("StatusFlags") Is DBNull.Value Then
                        .lStatusFlags = 0
                    Else
                        .lStatusFlags = CInt(oData("StatusFlags"))
                    End If

                    .LastLogin = GetDateFromNumber(CInt(oData("LastLogin")))
                    Dim lTempVal As Int32 = CInt(oData("LastGuildMembership"))
                    If lTempVal = 0 Then .dtLastGuildMembership = Date.MinValue Else .dtLastGuildMembership = GetDateFromNumber(lTempVal)
					.TotalPlayTime = CInt(oData("TotalPlayTime"))

                    lTempVal = CInt(oData("DateAccountWentActive"))
                    Dim dtActive As Date = Date.MinValue
                    If lTempVal = 0 Then dtActive = Date.MinValue Else dtActive = GetDateFromNumber(lTempVal)

                    lTempVal = CInt(oData("DateAccountWentInactive"))
                    If lTempVal = 0 Then .dtAccountWentInactive = Date.MinValue Else .dtAccountWentInactive = GetDateFromNumber(lTempVal)
                    If .dtAccountWentInactive < dtActive Then .dtAccountWentInactive = Date.MinValue

					.lStartLocX = CInt(oData("StartLocX"))
					.lStartLocZ = CInt(oData("StartLocZ"))
					.PirateStartLocX = CInt(oData("PirateStartLocX"))
					.PirateStartLocZ = CInt(oData("PirateStartLocZ"))

                    .lCelebrationEnds = glCurrentCycle + CInt(oData("CelebrationEnds"))
					.lWarSentiment = CInt(oData("WarSentiment"))

                    .lHangarManMult = CInt(oData("HangarManMult"))
                    .SpecTechCostMult = CSng(oData("SpecTechCostMult"))

					.DeathBudgetBalance = CInt(oData("DeathBudgetBalance"))

					.BadWarDecCPIncrease = CInt(oData("BadWarDecCPIncrease"))
					.BadWarDecCPIncreaseEndCycle = glCurrentCycle + CInt(oData("BadWarDecCPIncreaseEndCycle"))
					.BadWarDecMoralePenalty = CInt(oData("BadWarDecMoralePenalty"))
					.BadWarDecMoralePenaltyEndCycle = glCurrentCycle + CInt(oData("BadWarDecMoralePenaltyEndCycle"))

					.PlayedTimeAtEndOfWaves = CInt(oData("PlayedTimeAtEndOfWaves"))
					.PlayedTimeInTutorialOne = CInt(oData("PlayedTimeInTutorialOne"))
					.PlayedTimeWhenFirstWave = CInt(oData("PlayedTimeWhenFirstWave"))
                    '.PlayedTimeWhenTimerStarted = CInt(oData("PlayedTimeWhenTimerStarted"))
					.TutorialPhaseWaves = CInt(oData("TutorialPhaseWaves"))

					.lGuildID = CInt(oData("GuildID"))
					.lGuildRankID = CInt(oData("GuildRankID"))
					.lJoinedGuildOn = CInt(oData("JoinedGuildOn"))
                    .yPlayerTitle = CByte(oData("PlayerTitle"))
                    .yExTitleNew = CByte(oData("ExTitleNew"))
                    .lCustomTitlePermission = CInt(oData("CustomTitlePermission"))
                    .yCustomTitle = CByte(oData("CustomTitle"))
                    Dim lTemp As Int32 = CInt(oData("ExTitleEnd"))
                    If lTemp = 0 Then .dtExTitleEnd = Date.MinValue Else .dtExTitleEnd = GetDateFromNumber(lTemp)
					.lPlayerIcon = CInt(oData("PlayerIcon"))

                    .yGender = CByte(oData("PlayerGender"))
                    .GuaranteedSpecialTechID = CInt(oData("GuaranteedSpecialTechID"))
                    .ySpawnSystemSetting = CByte(oData("SpawnSystemSetting"))
                    '.lLastWPUpkeep = CInt(oData("LastWPUpkeep"))
                    '.lLastGuildShareUpkeep = CInt(oData("LastGuildShareUpkeep"))
                    '.blWarpoints = CLng(oData("Warpoints"))
                    '.blWarpointsAllTime = CLng(oData("WarpointsAllTime"))
                    .blMaxPopulation = CLng(oData("MaxTotalPop"))
                    .blDBPopulation = CLng(oData("DBTotalPop"))
                    '.lLastWPUpkeepTime = CInt(oData("LastWPUpkeepTime"))
                    'If .lLastWPUpkeepTime <> 0 Then .dtLastWPUpkeep = GetDateFromNumber(.lLastWPUpkeepTime)

					If .iStartedEnvirTypeID = ObjectType.ePlanet Then
						Dim oPlanet As Planet = GetEpicaPlanet(.lStartedEnvirID)
						.InMyDomain = (oPlanet Is Nothing = False) AndAlso (oPlanet.InMyDomain = True)
					ElseIf .iStartedEnvirTypeID = ObjectType.eSolarSystem Then
						Dim oSystem As SolarSystem = GetEpicaSystem(.lStartedEnvirID)
						.InMyDomain = (oSystem Is Nothing = False) AndAlso (oSystem.InMyDomain = True)
					End If
					.InMyDomain = True

					ReDim .ExternalEmailAddress(254)
					If oData("EmailAddress") Is DBNull.Value = False Then
						StringToBytes(CStr(oData("EmailAddress"))).CopyTo(.ExternalEmailAddress, 0)
						.iEmailSettings = CShort(oData("EmailSettings"))
					Else : .iEmailSettings = 0
					End If
					.iInternalEmailSettings = CShort(oData("InternalEmailSettings"))

					ReDim .lSlotID(4)
					ReDim .ySlotState(4)
					ReDim .lFactionID(2)

					For X As Int32 = 0 To 4
						.lSlotID(X) = CInt(oData("Slot" & (X + 1) & "ID"))
						.ySlotState(X) = CByte(oData("Slot" & (X + 1) & "State"))
					Next X
					For X As Int32 = 0 To 2
						.lFactionID(X) = CInt(oData("Faction" & (X + 1) & "ID"))
					Next X

					.lIronCurtainPlanet = CInt(oData("IronCurtainPlanetID"))

					'Set whether the player is dead or not... when colonies are added later they will reset the player death indicator
					If .blCredits <> 0 AndAlso .blCredits <> 100000 Then
						.HandleCheckForPlayerDeath()
					End If
				End With

                Dim i As Int32
                i = oPlayer.ObjectID
                CheckPlayerExtent(i)

                goPlayer(i) = oPlayer
                glPlayerIdx(i) = oPlayer.ObjectID

			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing
			bResult = True
		Catch
			LogEvent(LogEventType.CriticalError, "LoadSinglePlayer Error: " & Err.Description)
		Finally
			If oData Is Nothing = False Then
				oData.Close()
			End If
			oData = Nothing
			oComm = Nothing
		End Try
		Return bResult
	End Function
 
    Private Function LoadTechs(ByVal sPlayerInStr As String, ByVal bLoadSpecialTechDefs As Boolean) As Boolean
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim oComm As OleDb.OleDbCommand
        Dim sSQL As String
        Dim bResult As Boolean = False

        Dim oPlayer As Player
        Dim lPlayerID As Int32

        If goInitialPlayer Is Nothing Then goInitialPlayer = New Player()
        goInitialPlayer.ObjectID = 0

        Try
            'Alloy
            LogEvent(LogEventType.Informational, "Loading Alloys Techs...")
            sSQL = "SELECT * FROM tblAlloy WHERE " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If

                If oPlayer Is Nothing = False Then
                    oPlayer.mlAlloyUB += 1
                    ReDim Preserve oPlayer.mlAlloyIdx(oPlayer.mlAlloyUB)
                    ReDim Preserve oPlayer.moAlloy(oPlayer.mlAlloyUB)

                    oPlayer.moAlloy(oPlayer.mlAlloyUB) = New AlloyTech
                    With oPlayer.moAlloy(oPlayer.mlAlloyUB)
                        .AlloyName = StringToBytes(CStr(oData("AlloyName")))
                        .ObjectID = CInt(oData("AlloyID"))
                        .ObjTypeID = ObjectType.eAlloyTech
                        .Owner = oPlayer

                        If CInt(oData("ResultMineralID")) > 0 Then .AlloyResult = GetEpicaMineral(CInt(oData("ResultMineralID")))
                        .Mineral1ID = CInt(oData("Mineral1ID"))
                        .Mineral2ID = CInt(oData("Mineral2ID"))
                        .Mineral3ID = CInt(oData("Mineral3ID"))
                        .Mineral4ID = CInt(oData("Mineral4ID"))
                        .lPropertyID1 = CInt(oData("MineralProperty1ID"))
                        .lPropertyID2 = CInt(oData("MineralProperty2ID"))
                        .lPropertyID3 = CInt(oData("MineralProperty3ID"))
                        '.bHigher1 = CBool(oData("Property1Higher"))
                        '.bHigher2 = CBool(oData("Property2Higher"))
                        '.bHigher3 = CBool(oData("Property3Higher"))
                        .yNewVal1 = CByte(oData("Property1Value"))
                        .yNewVal2 = CByte(oData("Property2Value"))
                        .yNewVal3 = CByte(oData("Property3Value"))
                        .ResearchLevel = CByte(oData("ResearchLevel"))
                        .MajorDesignFlaw = CByte(oData("MajorDesignFlaw"))
                        .PopIntel = CInt(oData("PopIntel"))
                        '.lVersionNum = CInt(oData("VersionNumber"))
                        If oData("VersionNumber") Is DBNull.Value Then
                            .lVersionNum = Epica_Tech.TechVersionNum
                        Else
                            .lVersionNum = CInt(oData("VersionNumber"))
                        End If

                        .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                        .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                        .ResearchAttempts = CInt(oData("ResearchAttempts"))
                        .RandomSeed = CSng(oData("RandomSeed"))
                        .ComponentDevelopmentPhase = CType(oData("ResPhase"), Epica_Tech.eComponentDevelopmentPhase)

                        .yArchived = CByte(oData("bArchived"))

                        oPlayer.mlAlloyIdx(oPlayer.mlAlloyUB) = .ObjectID

                        AddQuickTechnologyLookup(.ObjectID, .ObjTypeID, oPlayer.ObjectID)
                    End With
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Alloy Result Properties
            LogEvent(LogEventType.Informational, "Loading Alloy Result Properties...")
            sSQL = "SELECT arp.AlloyID, arp.MineralPropertyID, arp.PropertyValue, a.OwnerID FROM tblAlloyResultProperty arp " & _
              "LEFT OUTER JOIN tblAlloy a ON a.AlloyID = arp.AlloyID " & _
              " WHERE " & sPlayerInStr & " ORDER BY arp.AlloyID"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If
                If oPlayer Is Nothing = False Then
                    Dim oAlloy As AlloyTech = CType(oPlayer.GetTech(CInt(oData("AlloyID")), ObjectType.eAlloyTech), AlloyTech)
                    If oAlloy Is Nothing = False Then
                        oAlloy.SetLoadedAlloyResultProperty(CInt(oData("MineralPropertyID")), CInt(oData("PropertyValue")))
                    End If
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Armor Tech
            sSQL = "SELECT * FROM tblArmor WHERE " & sPlayerInStr
            LogEvent(LogEventType.Informational, "Loading Armor...")
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If

                If oPlayer Is Nothing = False Then
                    oPlayer.mlArmorUB += 1
                    ReDim Preserve oPlayer.mlArmorIdx(oPlayer.mlArmorUB)
                    ReDim Preserve oPlayer.moArmor(oPlayer.mlArmorUB)

                    oPlayer.moArmor(oPlayer.mlArmorUB) = New ArmorTech
                    With oPlayer.moArmor(oPlayer.mlArmorUB)
                        .ArmorName = StringToBytes(CStr(oData("ArmorName")))
                        .yBeamResist = CByte(oData("BeamResist"))
                        .yChemicalResist = CByte(oData("ChemicalResist"))
                        .yRadarResist = CByte(oData("DetectionResist"))
                        .yMagneticResist = CByte(oData("ECMResist"))
                        .yBurnResist = CByte(oData("FlameResist"))
                        .lHPPerPlate = CInt(oData("HitPoints"))
                        .lHullUsagePerPlate = CInt(oData("HullUsagePerPlate"))
                        .yImpactResist = CByte(oData("ImpactResist"))
                        .ObjectID = CInt(oData("ArmorID"))
                        .ObjTypeID = ObjectType.eArmorTech
                        .Owner = oPlayer
                        .lOuterLayerMineralID = CInt(oData("OuterLayerMineralID"))
                        .lMiddleLayerMineralID = CInt(oData("MiddleLayerMineralID"))
                        .lInnerLayerMineralID = CInt(oData("InnerLayerMineralID"))
                        .yPiercingResist = CByte(oData("PiercingResist"))
                        '.lVersionNum = CInt(oData("VersionNumber"))
                        If oData("VersionNumber") Is DBNull.Value Then
                            .lVersionNum = Epica_Tech.TechVersionNum
                        Else
                            .lVersionNum = CInt(oData("VersionNumber"))
                        End If

                        .ErrorReasonCode = CByte(oData("ErrorReasonCode"))
                        .lIntegrity = CInt(oData("Integrity"))
                        .MajorDesignFlaw = CByte(oData("MajorDesignFlaw"))
                        .PopIntel = CInt(oData("PopIntel"))
                        .yArchived = CByte(oData("bArchived"))

                        .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                        .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                        .ResearchAttempts = CInt(oData("ResearchAttempts"))
                        .RandomSeed = CSng(oData("RandomSeed"))
                        .ComponentDevelopmentPhase = CType(oData("ResPhase"), Epica_Tech.eComponentDevelopmentPhase)

                        .lSpecifiedColonists = CInt(oData("SpecifiedColonist"))
                        .lSpecifiedEnlisted = CInt(oData("SpecifiedEnlisted"))
                        .lSpecifiedHull = CInt(oData("SpecifiedHull"))
                        .lSpecifiedMin1 = CInt(oData("SpecifiedMin1"))
                        .lSpecifiedMin2 = CInt(oData("SpecifiedMin2"))
                        .lSpecifiedMin3 = CInt(oData("SpecifiedMin3"))
                        .lSpecifiedMin4 = CInt(oData("SpecifiedMin4"))
                        .lSpecifiedMin5 = CInt(oData("SpecifiedMin5"))
                        .lSpecifiedMin6 = CInt(oData("SpecifiedMin6"))
                        .lSpecifiedOfficers = CInt(oData("SpecifiedOfficer"))
                        .lSpecifiedPower = CInt(oData("SpecifiedPower"))
                        .blSpecifiedProdCost = CLng(oData("SpecifiedProdCost"))
                        .blSpecifiedProdTime = CLng(oData("SpecifiedProdTime"))
                        .blSpecifiedResCost = CLng(oData("SpecifiedResCost"))
                        .blSpecifiedResTime = CLng(oData("SpecifiedResTime"))

                        oPlayer.mlArmorIdx(oPlayer.mlArmorUB) = .ObjectID

                        AddQuickTechnologyLookup(.ObjectID, .ObjTypeID, oPlayer.ObjectID)
                    End With
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Engine Techs
            LogEvent(LogEventType.Informational, "Loading Engines...")
            sSQL = "SELECT * FROM tblEngine WHERE " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If

                If oPlayer Is Nothing = False Then
                    oPlayer.mlEngineUB += 1
                    ReDim Preserve oPlayer.mlEngineIdx(oPlayer.mlEngineUB)
                    ReDim Preserve oPlayer.moEngine(oPlayer.mlEngineUB)

                    oPlayer.moEngine(oPlayer.mlEngineUB) = New EngineTech
                    With oPlayer.moEngine(oPlayer.mlEngineUB)
                        .EngineName = StringToBytes(CStr(oData("EngineName")))
                        .Maneuver = CByte(oData("Maneuver"))
                        .ObjectID = CInt(oData("EngineID"))
                        .ObjTypeID = ObjectType.eEngineTech
                        .Owner = oPlayer
                        .PowerProd = CInt(oData("PowerProd"))
                        .Speed = CByte(oData("Speed"))
                        .Thrust = CInt(oData("Thrust"))
                        '.lVersionNum = CInt(oData("VersionNumber"))
                        If oData("VersionNumber") Is DBNull.Value Then
                            .lVersionNum = Epica_Tech.TechVersionNum
                        Else
                            .lVersionNum = CInt(oData("VersionNumber"))
                        End If

                        .lStructuralBodyMineralID = CInt(oData("StructuralBodyMineralID"))
                        .lStructuralFrameMineralID = CInt(oData("StructuralFrameMineralID"))
                        .lStructuralMeldMineralID = CInt(oData("StructuralMeldMineralID"))
                        .lDriveBodyMineralID = CInt(oData("DriveBodyMineralID"))
                        .lDriveFrameMineralID = CInt(oData("DriveFrameMineralID"))
                        .lDriveMeldMineralID = CInt(oData("DriveMeldMineralID"))
                        '.lFuelCompositionMineralID = CInt(oData("FuelCompositionMineralID"))
                        '.lFuelCatalystMineralID = CInt(oData("FuelCatalystMineralID"))

                        .ColorValue = CByte(oData("ColorValue"))
                        .ErrorReasonCode = CByte(oData("ErrorReasonCode"))
                        .MajorDesignFlaw = CByte(oData("MajorDesignFlaw"))
                        .PopIntel = CInt(oData("PopIntel"))
                        .yArchived = CByte(oData("bArchived"))

                        .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                        .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                        .ResearchAttempts = CInt(oData("ResearchAttempts"))
                        .RandomSeed = CSng(oData("RandomSeed"))
                        .ComponentDevelopmentPhase = CType(oData("ResPhase"), Epica_Tech.eComponentDevelopmentPhase)

                        .HullRequired = CInt(oData("HullRequired"))
                        .HullTypeID = CByte(oData("HullTypeID"))

                        .lSpecifiedColonists = CInt(oData("SpecifiedColonist"))
                        .lSpecifiedEnlisted = CInt(oData("SpecifiedEnlisted"))
                        .lSpecifiedHull = CInt(oData("SpecifiedHull"))
                        .lSpecifiedMin1 = CInt(oData("SpecifiedMin1"))
                        .lSpecifiedMin2 = CInt(oData("SpecifiedMin2"))
                        .lSpecifiedMin3 = CInt(oData("SpecifiedMin3"))
                        .lSpecifiedMin4 = CInt(oData("SpecifiedMin4"))
                        .lSpecifiedMin5 = CInt(oData("SpecifiedMin5"))
                        .lSpecifiedMin6 = CInt(oData("SpecifiedMin6"))
                        .lSpecifiedOfficers = CInt(oData("SpecifiedOfficer"))
                        .lSpecifiedPower = CInt(oData("SpecifiedPower"))
                        .blSpecifiedProdCost = CLng(oData("SpecifiedProdCost"))
                        .blSpecifiedProdTime = CLng(oData("SpecifiedProdTime"))
                        .blSpecifiedResCost = CLng(oData("SpecifiedResCost"))
                        .blSpecifiedResTime = CLng(oData("SpecifiedResTime"))

                        oPlayer.mlEngineIdx(oPlayer.mlEngineUB) = .ObjectID

                        AddQuickTechnologyLookup(.ObjectID, .ObjTypeID, oPlayer.ObjectID)
                    End With
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Hull Techs
            LogEvent(LogEventType.Informational, "Loading Hulls...")
            sSQL = "SELECT * FROM tblHull WHERE " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If

                If oPlayer Is Nothing = False Then
                    oPlayer.mlHullUB += 1
                    ReDim Preserve oPlayer.mlHullIdx(oPlayer.mlHullUB)
                    ReDim Preserve oPlayer.moHull(oPlayer.mlHullUB)

                    oPlayer.moHull(oPlayer.mlHullUB) = New HullTech
                    With oPlayer.moHull(oPlayer.mlHullUB)
                        .HullName = StringToBytes(CStr(oData("HullName")))
                        .HullSize = CInt(oData("HullSize"))
                        .ModelID = CShort(oData("ModelID"))
                        .ObjectID = CInt(oData("HullID"))
                        .ObjTypeID = ObjectType.eHullTech
                        .Owner = oPlayer
                        .StructuralHitPoints = CInt(oData("StructuralHitPoints"))
                        .StructuralMineralID = CInt(oData("StructuralMineralID"))
                        .yTypeID = CByte(oData("TypeID"))
                        .ySubTypeID = CByte(oData("SubTypeID"))
                        .yChassisType = CByte(oData("ChassisType"))
                        '.lVersionNum = CInt(oData("VersionNumber"))
                        If oData("VersionNumber") Is DBNull.Value Then
                            .lVersionNum = Epica_Tech.TechVersionNum
                        Else
                            .lVersionNum = CInt(oData("VersionNumber"))
                        End If

                        .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                        .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                        .ResearchAttempts = CInt(oData("ResearchAttempts"))
                        .RandomSeed = CSng(oData("RandomSeed"))
                        .ComponentDevelopmentPhase = CType(oData("ResPhase"), Epica_Tech.eComponentDevelopmentPhase)
                        .PopIntel = CInt(oData("PopIntel"))
                        .yArchived = CByte(oData("bArchived"))

                        oPlayer.mlHullIdx(oPlayer.mlHullUB) = .ObjectID

                        AddQuickTechnologyLookup(.ObjectID, .ObjTypeID, oPlayer.ObjectID)
                    End With
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Hardpoints for Hulls
            LogEvent(LogEventType.Informational, "Loading Hull Hardpoints...")
            sSQL = "Select hp.HullID, hp.SlotX, hp.SlotY, hp.SlotConfig, hp.SlotGroup, h.OwnerID FROM tblHardPoint hp LEFT JOIN tblHull h ON hp.HullID = h.HullID WHERE " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oHull As HullTech
                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If
                If oPlayer Is Nothing = False Then
                    oHull = CType(oPlayer.GetTech(CInt(oData("HullID")), ObjectType.eHullTech), HullTech)
                    If oHull Is Nothing = False Then
                        oHull.SetHullSlot(CByte(oData("SlotX")), CByte(oData("SlotY")), CType(oData("SlotConfig"), HullTech.SlotConfig), CByte(oData("SlotGroup")))
                    Else : LogEvent(LogEventType.Warning, "Unable to find Hull ID " & CInt(oData("HullID")) & " for player " & CInt(oData("OwnerID")) & ".")
                    End If
                End If
                oHull = Nothing 'clear our pointer
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Radar techs
            LogEvent(LogEventType.Informational, "Loading Radars...")
            sSQL = "SELECT * FROM tblRadar WHERE " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If

                If oPlayer Is Nothing = False Then
                    oPlayer.mlRadarUB += 1
                    ReDim Preserve oPlayer.moRadar(oPlayer.mlRadarUB)
                    ReDim Preserve oPlayer.mlRadarIdx(oPlayer.mlRadarUB)

                    oPlayer.moRadar(oPlayer.mlRadarUB) = New RadarTech
                    With oPlayer.moRadar(oPlayer.mlRadarUB)
                        .DisruptionResistance = CByte(oData("DisruptionResist"))
                        .HullRequired = CInt(oData("HullRequired"))
                        .MaximumRange = CByte(oData("MaximumRange"))
                        .ObjectID = CInt(oData("RadarID"))
                        .ObjTypeID = ObjectType.eRadarTech
                        .OptimumRange = CByte(oData("OptimumRange"))
                        .Owner = oPlayer
                        .PowerRequired = CInt(oData("PowerRequired"))
                        .RadarName = StringToBytes(CStr(oData("RadarName")))
                        .ScanResolution = CByte(oData("ScanResolution"))
                        .WeaponAcc = CByte(oData("WeaponAcc"))
                        '.lVersionNum = CInt(oData("VersionNumber"))
                        If oData("VersionNumber") Is DBNull.Value Then
                            .lVersionNum = Epica_Tech.TechVersionNum
                        Else
                            .lVersionNum = CInt(oData("VersionNumber"))
                        End If

                        .ErrorReasonCode = CByte(oData("ErrorReasonCode"))
                        .RadarType = CByte(oData("RadarType"))
                        .JamImmunity = CByte(oData("JamImmunity"))
                        .JamStrength = CByte(oData("JamStrength"))
                        .JamTargets = CByte(oData("JamTargets"))
                        .JamEffect = CByte(oData("JamEffect"))

                        .lCasingMineralID = CInt(oData("CasingMineralID"))
                        .lCollectionMineralID = CInt(oData("CollectionMineralID"))
                        .lEmitterMineralID = CInt(oData("EmitterMineralID"))
                        .lDetectionMineralID = CInt(oData("DetectionMineralID"))

                        .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                        .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                        .ResearchAttempts = CInt(oData("ResearchAttempts"))
                        .RandomSeed = CSng(oData("RandomSeed"))
                        .MajorDesignFlaw = CByte(oData("MajorDesignFlaw"))
                        .ComponentDevelopmentPhase = CType(oData("ResPhase"), Epica_Tech.eComponentDevelopmentPhase)
                        .PopIntel = CInt(oData("PopIntel"))
                        .yArchived = CByte(oData("bArchived"))

                        .lSpecifiedColonists = CInt(oData("SpecifiedColonist"))
                        .lSpecifiedEnlisted = CInt(oData("SpecifiedEnlisted"))
                        .lSpecifiedHull = CInt(oData("SpecifiedHull"))
                        .lSpecifiedMin1 = CInt(oData("SpecifiedMin1"))
                        .lSpecifiedMin2 = CInt(oData("SpecifiedMin2"))
                        .lSpecifiedMin3 = CInt(oData("SpecifiedMin3"))
                        .lSpecifiedMin4 = CInt(oData("SpecifiedMin4"))
                        .lSpecifiedMin5 = CInt(oData("SpecifiedMin5"))
                        .lSpecifiedMin6 = CInt(oData("SpecifiedMin6"))
                        .lSpecifiedOfficers = CInt(oData("SpecifiedOfficer"))
                        .lSpecifiedPower = CInt(oData("SpecifiedPower"))
                        .blSpecifiedProdCost = CLng(oData("SpecifiedProdCost"))
                        .blSpecifiedProdTime = CLng(oData("SpecifiedProdTime"))
                        .blSpecifiedResCost = CLng(oData("SpecifiedResCost"))
                        .blSpecifiedResTime = CLng(oData("SpecifiedResTime"))

                        oPlayer.mlRadarIdx(oPlayer.mlRadarUB) = .ObjectID

                        AddQuickTechnologyLookup(.ObjectID, .ObjTypeID, oPlayer.ObjectID)
                    End With
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Shield Techs
            LogEvent(LogEventType.Informational, "Loading Shields...")
            sSQL = "SELECT * FROM tblShield WHERE " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If

                If oPlayer Is Nothing = False Then
                    oPlayer.mlShieldUB += 1
                    ReDim Preserve oPlayer.moShield(oPlayer.mlShieldUB)
                    ReDim Preserve oPlayer.mlShieldIdx(oPlayer.mlShieldUB)
                    oPlayer.moShield(oPlayer.mlShieldUB) = New ShieldTech
                    With oPlayer.moShield(oPlayer.mlShieldUB)
                        .HullRequired = CInt(oData("HullRequired"))
                        .MaxHitPoints = CInt(oData("MaxHitPoints"))
                        .ObjectID = CInt(oData("ShieldID"))
                        .ObjTypeID = ObjectType.eShieldTech
                        .Owner = oPlayer
                        .PowerRequired = CInt(oData("PowerRequired"))
                        .RechargeFreq = CInt(oData("RechargeFreq"))
                        .RechargeRate = CInt(oData("RechargeRate"))
                        .ShieldName = StringToBytes(CStr(oData("ShieldName")))
                        .lProjectionHullSize = CInt(oData("ProjectionHullSize"))
                        '.lVersionNum = CInt(oData("VersionNumber"))
                        If oData("VersionNumber") Is DBNull.Value Then
                            .lVersionNum = Epica_Tech.TechVersionNum
                        Else
                            .lVersionNum = CInt(oData("VersionNumber"))
                        End If

                        .ColorValue = CByte(oData("ColorValue"))
                        .ErrorReasonCode = CByte(oData("ErrorReasonCode"))
                        .HullTypeID = CByte(oData("HullTypeID"))

                        .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                        .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                        .ResearchAttempts = CInt(oData("ResearchAttempts"))
                        .RandomSeed = CSng(oData("RandomSeed"))
                        .ComponentDevelopmentPhase = CType(oData("ResPhase"), Epica_Tech.eComponentDevelopmentPhase)
                        .MajorDesignFlaw = CByte(oData("MajorDesignFlaw"))
                        .PopIntel = CInt(oData("PopIntel"))
                        .yArchived = CByte(oData("bArchived"))

                        .lCasingMineralID = CInt(oData("CasingMineralID"))
                        .lCoilMineralID = CInt(oData("CoilMineralID"))
                        .lAcceleratorMineralID = CInt(oData("AcceleratorMineralID"))

                        .lSpecifiedColonists = CInt(oData("SpecifiedColonist"))
                        .lSpecifiedEnlisted = CInt(oData("SpecifiedEnlisted"))
                        .lSpecifiedHull = CInt(oData("SpecifiedHull"))
                        .lSpecifiedMin1 = CInt(oData("SpecifiedMin1"))
                        .lSpecifiedMin2 = CInt(oData("SpecifiedMin2"))
                        .lSpecifiedMin3 = CInt(oData("SpecifiedMin3"))
                        .lSpecifiedMin4 = CInt(oData("SpecifiedMin4"))
                        .lSpecifiedMin5 = CInt(oData("SpecifiedMin5"))
                        .lSpecifiedMin6 = CInt(oData("SpecifiedMin6"))
                        .lSpecifiedOfficers = CInt(oData("SpecifiedOfficer"))
                        .lSpecifiedPower = CInt(oData("SpecifiedPower"))
                        .blSpecifiedProdCost = CLng(oData("SpecifiedProdCost"))
                        .blSpecifiedProdTime = CLng(oData("SpecifiedProdTime"))
                        .blSpecifiedResCost = CLng(oData("SpecifiedResCost"))
                        .blSpecifiedResTime = CLng(oData("SpecifiedResTime"))

                        oPlayer.mlShieldIdx(oPlayer.mlShieldUB) = .ObjectID

                        AddQuickTechnologyLookup(.ObjectID, .ObjTypeID, oPlayer.ObjectID)
                    End With
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Weapon Techs
            LogEvent(LogEventType.Informational, "Loading Weapons...")
            sSQL = "SELECT * FROM tblWeapon WHERE " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If

                If oPlayer Is Nothing = False Then
                    oPlayer.mlWeaponUB += 1
                    ReDim Preserve oPlayer.moWeapon(oPlayer.mlWeaponUB)
                    ReDim Preserve oPlayer.mlWeaponIdx(oPlayer.mlWeaponUB)
                    oPlayer.moWeapon(oPlayer.mlWeaponUB) = BaseWeaponTech.CreateFromDataReader(oData)
                    oPlayer.mlWeaponIdx(oPlayer.mlWeaponUB) = oPlayer.moWeapon(oPlayer.mlWeaponUB).ObjectID

                    AddQuickTechnologyLookup(oPlayer.moWeapon(oPlayer.mlWeaponUB).ObjectID, oPlayer.moWeapon(oPlayer.mlWeaponUB).ObjTypeID, oPlayer.ObjectID)
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            If bLoadSpecialTechDefs = True Then
                'Special Techs
                LogEvent(LogEventType.Informational, "Loading Special Tech Definitions...")
                sSQL = "SELECT * FROM tblSpecialTech ORDER BY TechName"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                Dim lGuaranteedCnt As Int32 = 0
                Dim lPos As Int32 = 0
                Dim lUB As Int32 = 9

                ReDim gyGuaranteedList(9)
                System.BitConverter.GetBytes(GlobalMessageCode.eGetSpecTechGuaranteeList).CopyTo(gyGuaranteedList, lPos) : lPos += 2
                lPos += 4
                lPos += 4

                While oData.Read
                    glSpecialTechUB += 1
                    ReDim Preserve goSpecialTechs(glSpecialTechUB)
                    ReDim Preserve glSpecialTechIdx(glSpecialTechUB)
                    goSpecialTechs(glSpecialTechUB) = New SpecialTech
                    With goSpecialTechs(glSpecialTechUB)
                        .ObjectID = CInt(oData("TechID"))
                        .ObjTypeID = ObjectType.eSpecialTech
                        .InitialSuccessChance = CInt(oData("InitialSuccessChance"))
                        .IncrementalSuccess = CInt(oData("IncrementalSuccess"))
                        .FallOffSuccess = CInt(oData("FallOffSuccess"))
                        .MaxLinkChanceAttempts = CInt(oData("MaxLinkChanceAttempts"))
                        .ProgramControl = CInt(oData("ProgramControl"))
                        .fPercCostValue = CSng(oData("PercCostValue"))
                        Dim sTemp As String
                        If oData("BenefitsDesc") Is DBNull.Value Then sTemp = "" Else sTemp = CStr(oData("BenefitsDesc"))
                        .BenefitsDesc = StringToBytes(sTemp)
                        If oData("RolePlayDesc") Is DBNull.Value Then sTemp = "" Else sTemp = CStr(oData("RolePlayDesc"))
                        .RolePlayDesc = StringToBytes(sTemp)
                        .TechName = StringToBytes(CStr(oData("TechName")))
                        .lNewValue = CInt(oData("NewAttrValue"))
                        .bInGuaranteeList = CInt(oData("GuaranteedSelectable")) <> 0

                        .bHalfOwned = CInt(oData("HalfOwnedAdjusted")) <> 0

                        .bCanBeLinked = CInt(oData("CanBeLinked")) <> 0

                        If .bInGuaranteeList = True Then ' AndAlso .bCanBeLinked = True Then
                            lGuaranteedCnt += 1
                            Dim iNLen As Int16 = 0
                            If .TechName Is Nothing = False Then iNLen = CShort(.TechName.Length)
                            Dim iBLen As Int16 = 0
                            If .BenefitsDesc Is Nothing = False Then iBLen = CShort(.BenefitsDesc.Length)

                            lUB += 8 + iNLen + iBLen
                            ReDim Preserve gyGuaranteedList(lUB)
                            System.BitConverter.GetBytes(.ObjectID).CopyTo(gyGuaranteedList, lPos) : lPos += 4
                            System.BitConverter.GetBytes(iNLen).CopyTo(gyGuaranteedList, lPos) : lPos += 2
                            .TechName.CopyTo(gyGuaranteedList, lPos) : lPos += iNLen
                            System.BitConverter.GetBytes(iBLen).CopyTo(gyGuaranteedList, lPos) : lPos += 2
                            .BenefitsDesc.CopyTo(gyGuaranteedList, lPos) : lPos += iBLen
                        End If


                        glSpecialTechIdx(glSpecialTechUB) = .ObjectID
                    End With
                End While
                oData.Close()
                oData = Nothing
                oComm = Nothing

                System.BitConverter.GetBytes(lGuaranteedCnt).CopyTo(gyGuaranteedList, 2)


                'Special Tech Prerequisites
                LogEvent(LogEventType.Informational, "Loading Special Tech Prerequisites...")
                sSQL = "SELECT * FROM tblSpecialTechPrerequisite"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                While oData.Read
                    glSpecTechPreqUB += 1
                    ReDim Preserve goSpecTechPreq(glSpecTechPreqUB)
                    ReDim Preserve glSpecTechPreqIdx(glSpecTechPreqUB)
                    goSpecTechPreq(glSpecTechPreqUB) = New SpecialTechPrerequisite
                    With goSpecTechPreq(glSpecTechPreqUB)
                        .TechID = CInt(oData("TechID"))
                        .ObjectID = CInt(oData("STP_ID"))
                        .ObjTypeID = ObjectType.eSpecialTechPreq
                        .lPreqID = CInt(oData("ObjectID"))
                        .iPreqTypeID = CShort(oData("ObjTypeID"))
                        .RequiredValue = CInt(oData("RequiredValue"))
                        .ChanceToOpenLink = CInt(oData("ChanceToOpenLink"))
                        .RequiredPrerequisite = (CInt(oData("RequiredPrerequisite")) <> 0)
                    End With

                    Dim oTech As SpecialTech = GetEpicaSpecialTech(goSpecTechPreq(glSpecTechPreqUB).TechID)
                    If oTech Is Nothing = False Then
                        oTech.lPreqUB += 1
                        ReDim Preserve oTech.oPreqs(oTech.lPreqUB)
                        oTech.oPreqs(oTech.lPreqUB) = goSpecTechPreq(glSpecTechPreqUB)
                    End If
                End While
                oData.Close()
                oData = Nothing
                oComm = Nothing
            End If

            'Player Special Tech
            Dim sSearch As String = sPlayerInStr
            sSearch = sSearch.Replace("OwnerID", "PlayerID")
            LogEvent(LogEventType.Informational, "Loading Player Special Techs...")
            sSQL = "SELECT * FROM tblPlayerSpecialTech WHERE " & sSearch
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read

                lPlayerID = CInt(oData("PlayerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If

                If oPlayer Is Nothing = False Then
                    oPlayer.oSpecials.mlSpecialTechUB += 1
                    ReDim Preserve oPlayer.oSpecials.moSpecialTech(oPlayer.oSpecials.mlSpecialTechUB)
                    ReDim Preserve oPlayer.oSpecials.mlSpecialTechIdx(oPlayer.oSpecials.mlSpecialTechUB)

                    oPlayer.oSpecials.moSpecialTech(oPlayer.oSpecials.mlSpecialTechUB) = New PlayerSpecialTech
                    With oPlayer.oSpecials.moSpecialTech(oPlayer.oSpecials.mlSpecialTechUB)
                        .ObjectID = CInt(oData("SpecialTechID"))
                        .ObjTypeID = ObjectType.eSpecialTech
                        .Owner = oPlayer
                        .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                        .ResearchAttempts = CInt(oData("ResearchAttempts"))
                        .CreditResearchAttempts = CInt(oData("CreditResearchAttempts"))
                        .LinkAttempts = CInt(oData("LinkAttempts"))
                        .ComponentDevelopmentPhase = CType(oData("ResPhase"), Epica_Tech.eComponentDevelopmentPhase)
                        Dim oSpecTech As SpecialTech = GetEpicaSpecialTech(.ObjectID)
                        If oSpecTech Is Nothing = False Then
                            .SuccessChanceIncrement = oSpecTech.IncrementalSuccess
                        Else : .SuccessChanceIncrement = 1
                        End If
                        '.lVersionNum = CInt(oData("VersionNumber"))
                        If oData("VersionNumber") Is DBNull.Value Then
                            .lVersionNum = Epica_Tech.TechVersionNum
                        Else
                            .lVersionNum = CInt(oData("VersionNumber"))
                        End If
                        .yFlags = CByte(oData("SuccessfulLink"))
                        oPlayer.oSpecials.mlSpecialTechIdx(oPlayer.oSpecials.mlSpecialTechUB) = .ObjectID
                        .yArchived = CByte(oData("bArchived"))

                        If .bInTheTank = True AndAlso (.yFlags And PlayerSpecialTech.eySpecialTechFlags.eTankThrowBackEvent) = 0 Then
                            AddToQueue(glCurrentCycle + CInt((Rnd() * 200) + 200) * 1000, QueueItemType.ePerformLinkTest, oPlayer.ObjectID, .ObjectID, -1, -1, -1, -1, -1, -1)
                        End If

                        'If .bInTheTank = True Then
                        '    oPlayer.oSpecials.mcolClosed.Add(oPlayer.oSpecials.moSpecialTech(oPlayer.oSpecials.mlSpecialTechUB), "KEY" & .oTech.ObjectID)
                        'ElseIf .bLinked = True Then
                        '    If .ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        '        oPlayer.oSpecials.mcolOpen.Add(oPlayer.oSpecials.moSpecialTech(oPlayer.oSpecials.mlSpecialTechUB), "KEY" & .oTech.ObjectID)
                        '    End If
                        'Else
                        '    If .LinkAttempts > .oTech.MaxLinkChanceAttempts Then oPlayer.oSpecials.mcolClosed.Add(oPlayer.oSpecials.moSpecialTech(oPlayer.oSpecials.mlSpecialTechUB), "KEY" & .oTech.ObjectID)
                        'End If

                        AddQuickTechnologyLookup(.ObjectID, .ObjTypeID, oPlayer.ObjectID)
                    End With
                End If

            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Now, that special techs have been loaded, update all the player's Special tech bonuses
            LogEvent(LogEventType.Informational, "Calculating Player Special Tech Bonus Effects...")
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then goPlayer(X).CalculateSpecialTechs() 'removed the InMyDomain check to see if that fixes our CP loading bug
            Next X


            LogEvent(LogEventType.Informational, "Loading Tech Costs...")
            sSQL = "SELECT Count(*) FROM tblProductionCost WHERE ObjTypeID IN (" & ObjectType.eAlloyTech & ", " & ObjectType.eArmorTech & ", " & ObjectType.eEngineTech & ", " & _
                ObjectType.eHullTech & ", " & ObjectType.ePrototype & ", " & ObjectType.eRadarTech & ", " & ObjectType.eShieldTech & ", " & ObjectType.eWeaponTech & ")"
            Dim lProdCostUB As Int32 = -1
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            If oData.Read = True Then
                lProdCostUB = CInt(oData(0)) - 1
            End If
            oData.Close()
            oData = Nothing
            oComm = Nothing

            Dim oProdCosts(lProdCostUB) As ProductionCost
            Dim lProdCostsID(lProdCostUB) As Int32
            For lTmpVal As Int32 = 0 To lProdCostUB
                lProdCostsID(lTmpVal) = -1
            Next
            sSQL = "SELECT pc.PC_ID, pc.ProductionCostType, pc.PointsRequired, pc.Officers, pc.Colonists, pc.Credits, pc.Enlisted, pc.ObjectID, pc.ObjTypeID, t.OwnerID FROM "
            sSQL &= "tblAlloy t LEFT OUTER JOIN tblProductionCost pc ON pc.ObjectID = t.AlloyID "
            sSQL &= "WHERE pc.ObjTypeID = " & ObjectType.eAlloyTech & " UNION "
            sSQL &= "SELECT pc.PC_ID, pc.ProductionCostType, pc.PointsRequired, pc.Officers, pc.Colonists, pc.Credits, pc.Enlisted, pc.ObjectID, pc.ObjTypeID, t.OwnerID FROM "
            sSQL &= "tblArmor t LEFT OUTER JOIN tblProductionCost pc ON pc.ObjectID = t.ArmorID "
            sSQL &= "WHERE pc.ObjTypeID = " & ObjectType.eArmorTech & " UNION "
            sSQL &= "SELECT pc.PC_ID, pc.ProductionCostType, pc.PointsRequired, pc.Officers, pc.Colonists, pc.Credits, pc.Enlisted, pc.ObjectID, pc.ObjTypeID, t.OwnerID FROM "
            sSQL &= "tblEngine t LEFT OUTER JOIN tblProductionCost pc ON pc.ObjectID = t.EngineID "
            sSQL &= "WHERE pc.ObjTypeID = " & ObjectType.eEngineTech & " UNION "
            sSQL &= "SELECT pc.PC_ID, pc.ProductionCostType, pc.PointsRequired, pc.Officers, pc.Colonists, pc.Credits, pc.Enlisted, pc.ObjectID, pc.ObjTypeID, t.OwnerID FROM "
            sSQL &= "tblHull t LEFT OUTER JOIN tblProductionCost pc ON pc.ObjectID = t.HullID "
            sSQL &= "WHERE pc.ObjTypeID = " & ObjectType.eHullTech & " UNION "
            sSQL &= "SELECT pc.PC_ID, pc.ProductionCostType, pc.PointsRequired, pc.Officers, pc.Colonists, pc.Credits, pc.Enlisted, pc.ObjectID, pc.ObjTypeID, t.OwnerID FROM "
            sSQL &= "tblPrototype t LEFT OUTER JOIN tblProductionCost pc ON pc.ObjectID = t.PrototypeID "
            sSQL &= "WHERE pc.ObjTypeID = " & ObjectType.ePrototype & " UNION "
            sSQL &= "SELECT pc.PC_ID, pc.ProductionCostType, pc.PointsRequired, pc.Officers, pc.Colonists, pc.Credits, pc.Enlisted, pc.ObjectID, pc.ObjTypeID, t.OwnerID FROM "
            sSQL &= "tblRadar t LEFT OUTER JOIN tblProductionCost pc ON pc.ObjectID = t.RadarID "
            sSQL &= "WHERE pc.ObjTypeID = " & ObjectType.eRadarTech & " UNION "
            sSQL &= "SELECT pc.PC_ID, pc.ProductionCostType, pc.PointsRequired, pc.Officers, pc.Colonists, pc.Credits, pc.Enlisted, pc.ObjectID, pc.ObjTypeID, t.OwnerID FROM "
            sSQL &= "tblShield t LEFT OUTER JOIN tblProductionCost pc ON pc.ObjectID = t.ShieldID "
            sSQL &= "WHERE pc.ObjTypeID = " & ObjectType.eShieldTech & " UNION "
            sSQL &= "SELECT pc.PC_ID, pc.ProductionCostType, pc.PointsRequired, pc.Officers, pc.Colonists, pc.Credits, pc.Enlisted, pc.ObjectID, pc.ObjTypeID, t.OwnerID FROM "
            sSQL &= "tblWeapon t LEFT OUTER JOIN tblProductionCost pc ON pc.ObjectID = t.WeaponID "
            sSQL &= "WHERE pc.ObjTypeID = " & ObjectType.eWeaponTech
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)

            Dim lIdx As Int32 = -1
            While oData.Read
                lIdx += 1

                oProdCosts(lIdx) = New ProductionCost
                With oProdCosts(lIdx)
                    .ColonistCost = CInt(oData("Colonists"))
                    .CreditCost = CLng(oData("Credits"))
                    .EnlistedCost = CInt(oData("Enlisted"))
                    .ObjectID = CInt(oData("ObjectID"))
                    .ObjTypeID = CShort(oData("ObjTypeID"))
                    .OfficerCost = CInt(oData("Officers"))
                    .PC_ID = CInt(oData("PC_ID"))
                    .PointsRequired = CLng(oData("PointsRequired"))
                    .ProductionCostType = CByte(oData("ProductionCostType"))
                    .ItemCostUB = -1
                    lProdCostsID(lIdx) = .PC_ID
                End With

                lPlayerID = CInt(oData("OwnerID"))
                If lPlayerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lPlayerID)
                End If

                If oPlayer Is Nothing = False AndAlso (oPlayer.InMyDomain = True OrElse oPlayer.ObjectID = 0) Then
                    Dim oTech As Epica_Tech = oPlayer.GetTech(oProdCosts(lIdx).ObjectID, oProdCosts(lIdx).ObjTypeID)
                    If oTech Is Nothing = False Then
                        oTech.SetTechCost(oProdCosts(lIdx))
                    End If
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'goInitialPlayer.LoadTechProdCosts()
            'For X As Int32 = 0 To glPlayerUB
            '    If glPlayerIdx(X) <> -1 AndAlso (goPlayer(X).InMyDomain OrElse goPlayer(X).ObjectID = 0) Then goPlayer(X).LoadTechProdCosts()
            'Next X

            'Now, do one call to the database for all productioncostitems
            LogEvent(LogEventType.Informational, "Loading ProductionCostItems...")
            sSQL = "select * from tblproductioncostitem where pc_id in (select pc_id from tblproductioncost where objtypeid in ("
            sSQL &= CInt(ObjectType.eAlloyTech) & ", " & CInt(ObjectType.eArmorTech) & ", " & CInt(ObjectType.eEngineTech) & ", " & CInt(ObjectType.eHullTech) & ", "
            sSQL &= CInt(ObjectType.ePrototype) & ", " & CInt(ObjectType.eRadarTech) & ", " & CInt(ObjectType.eShieldTech) & ", " & CInt(ObjectType.eWeaponTech) & "))"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True
                Dim lPCID As Int32 = CInt(oData("PC_ID"))

                For X As Int32 = 0 To lProdCostUB
                    If lProdCostsID(X) = lPCID AndAlso oProdCosts(X) Is Nothing = False AndAlso oProdCosts(X).PC_ID = lPCID Then

                        oProdCosts(X).ItemCostUB += 1
                        ReDim Preserve oProdCosts(X).ItemCosts(oProdCosts(X).ItemCostUB)
                        oProdCosts(X).ItemCosts(oProdCosts(X).ItemCostUB) = New ProductionCostItem()
                        With oProdCosts(X).ItemCosts(oProdCosts(X).ItemCostUB)
                            .ItemID = CInt(oData("ItemID"))
                            .ItemTypeID = CShort(oData("ItemTypeID"))
                            .oProdCost = oProdCosts(X)
                            .PC_ID = lPCID
                            .PCM_ID = CInt(oData("PCM_ID"))
                            .QuantityNeeded = CInt(oData("Quantity"))
                        End With

                        Exit For
                    End If
                Next X
 
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing


            'Now, load our tech knowledge
            LogEvent(LogEventType.Informational, "Loading Player Tech Knowledges...")
            sSQL = "SELECT * FROM tblPlayerTechKnowledge WHERE " & sSearch
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read

                lPlayerID = CInt(oData("PlayerID"))
                Dim lTechID As Int32 = CInt(oData("TechID"))
                Dim iTechTypeID As Int16 = CShort(oData("TechTypeID"))
                Dim lTechOwnerID As Int32 = CInt(oData("TechOwnerID"))
                Dim yKnowType As Byte = CByte(oData("KnowledgeType"))

                oPlayer = GetEpicaPlayer(lTechOwnerID)
                If oPlayer Is Nothing Then Continue While
                Dim oTech As Epica_Tech = oPlayer.GetTech(lTechID, iTechTypeID)
                If oTech Is Nothing Then Continue While

                oPlayer = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing Then Continue While
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oPlayer, oTech, CType(yKnowType, PlayerTechKnowledge.KnowledgeType), True)
                If oPTK Is Nothing = False Then oPTK.yArchived = CByte(oData("bArchived"))
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            bResult = True
        Catch
            LogEvent(LogEventType.CriticalError, "LoadTechs Error: " & Err.Description)
        Finally
            If oData Is Nothing = False Then
                oData.Close()
            End If
            oData = Nothing
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Private Function LoadDefs(ByVal sPlayerInStr As String) As Boolean
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim oComm As OleDb.OleDbCommand
		Dim sSQL As String
		Dim bResult As Boolean = False
		Dim X As Int32
		Dim oTmpWeaponDef As WeaponDef
		Dim oTmpFacilityDef As FacilityDef
		Dim oTmpUnitDef As Epica_Entity_Def

		Try
			'Weapon Definitions....
			LogEvent(LogEventType.Informational, "Loading Weapon Defs...")
			sSQL = "SELECT tblWeaponDef.*, tblWeapon.OwnerID FROM tblWeaponDef LEFT OUTER JOIN tblWeapon ON tblWeaponDef.WeaponID = tblWeapon.WeaponID WHERE tblWeaponDef.WeaponID = -1 OR tblWeapon.OwnerID = 0 OR tblWeapon.OwnerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
                glWeaponDefUB = glWeaponDefUB + 1
                If glWeaponDefUB > goWeaponDefs.GetUpperBound(0) Then
                    ReDim Preserve goWeaponDefs(glWeaponDefUB + 1000)
                    ReDim Preserve glWeaponDefIdx(glWeaponDefUB + 1000)
                End If
				goWeaponDefs(glWeaponDefUB) = New WeaponDef()
				With goWeaponDefs(glWeaponDefUB)
                    .FillFromDataReader(oData)
					glWeaponDefIdx(glWeaponDefUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Prototypes
			LogEvent(LogEventType.Informational, "Loading Prototypes...")
			sSQL = "SELECT * FROM tblPrototype WHERE OwnerID = 0 OR OwnerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = Nothing
				Dim lPlayerID As Int32 = CInt(oData("OwnerID"))
				If lPlayerID = 0 Then
					oPlayer = goInitialPlayer
				Else : oPlayer = GetEpicaPlayer(lPlayerID)
				End If
				oPlayer.mlPrototypeUB += 1
				ReDim Preserve oPlayer.moPrototype(oPlayer.mlPrototypeUB)
				ReDim Preserve oPlayer.mlPrototypeIdx(oPlayer.mlPrototypeUB)
				oPlayer.moPrototype(oPlayer.mlPrototypeUB) = New Prototype
				With oPlayer.moPrototype(oPlayer.mlPrototypeUB)
					.ObjectID = CInt(oData("PrototypeID"))
					.ObjTypeID = ObjectType.ePrototype
					.PrototypeName = StringToBytes(CStr(oData("PrototypeName")))
					.Owner = oPlayer
					.lEngineTech = CInt(oData("EngineID"))
					.lArmorTech = CInt(oData("ArmorID"))
					'.lHangarTech = CInt(oData("HangarID"))
					.lHullTech = CInt(oData("HullID"))
					.lRadarTech = CInt(oData("RadarID"))
					.lShieldTech = CInt(oData("ShieldID"))
					.ForeArmorUnits = CInt(oData("ForeArmorUnits"))
					.AftArmorUnits = CInt(oData("AftArmorUnits"))
					.LeftArmorUnits = CInt(oData("LeftArmorUnits"))
					.RightArmorUnits = CInt(oData("RightArmorUnits"))
					.ProductionHull = CInt(oData("ProductionHull"))
                    .MaxCrew = CInt(oData("MaxCrew"))
                    '.MinCrew = CShort(oData("MinCrew"))

					.CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
					.SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
					.ResearchAttempts = CInt(oData("ResearchAttempts"))
					.RandomSeed = CSng(oData("RandomSeed"))
					.ComponentDevelopmentPhase = CType(oData("ResPhase"), Epica_Tech.eComponentDevelopmentPhase)
					.PopIntel = CInt(oData("PopIntel"))
                    .yArchived = CByte(oData("bArchived"))

					oPlayer.mlPrototypeIdx(oPlayer.mlPrototypeUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Now, to get the prototype weapons...
			LogEvent(LogEventType.Informational, "Loading Prototype Weapons...")
			sSQL = "SELECT pw.PW_ID, pw.WeaponID, pw.SlotX, pw.SlotY, pw.ArcID, pw.WeaponGroupTypeID, pw.PrototypeID, p.OwnerID FROM tblPrototypeWeapon pw LEFT OUTER JOIN tblPrototype p ON pw.PrototypeID = p.PrototypeID WHERE p.OwnerID = 0 OR p.OwnerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = Nothing
				Dim lPlayerID As Int32 = CInt(oData("OwnerID"))
				If lPlayerID = 0 Then
					oPlayer = goInitialPlayer
				Else : oPlayer = GetEpicaPlayer(lPlayerID)
				End If

				If oPlayer Is Nothing = False Then
					Dim oProto As Prototype = CType(oPlayer.GetTech(CInt(oData("PrototypeID")), ObjectType.ePrototype), Prototype)
					If oProto Is Nothing = False Then
						oProto.AddWeaponPlacement(CInt(oData("WeaponID")), CByte(oData("SlotX")), CByte(oData("SlotY")), _
						  CType(CByte(oData("WeaponGroupTypeID")), WeaponGroupType), CType(CByte(oData("ArcID")), UnitArcs))
					End If
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

            System.GC.Collect()

			'Unit Def
			LogEvent(LogEventType.Informational, "Loading Unit Defs...")
			sSQL = "SELECT * FROM tblUnitDef WHERE OwnerID = 0 OR OwnerID = " & gl_HARDCODE_PIRATE_PLAYER_ID & " OR OwnerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            ReDim goUnitDef(glUnitDefUB + 10000)
            ReDim Preserve glUnitDefIdx(glUnitDefUB + 10000)

            While oData.Read
                glUnitDefUB += 1
                If glUnitDefUB > goUnitDef.GetUpperBound(0) Then
                    ReDim Preserve goUnitDef(glUnitDefUB + 10000)
                    ReDim Preserve glUnitDefIdx(glUnitDefUB + 10000)
                End If
                goUnitDef(glUnitDefUB) = New Epica_Entity_Def()
                With goUnitDef(glUnitDefUB)
                    .FillFromDataReader(oData, ObjectType.eUnitDef)
                    glUnitDefIdx(glUnitDefUB) = .ObjectID
                End With
            End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

            System.GC.Collect()

			'Facility Definitions
			LogEvent(LogEventType.Informational, "Loading Structure Defs...")
			sSQL = "SELECT * FROM tblStructureDef WHERE OwnerID = 0 OR OwnerID = " & gl_HARDCODE_PIRATE_PLAYER_ID & " OR OwnerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            ReDim goFacilityDef(glFacilityDefUB + 10000)
            ReDim glFacilityDefIdx(glFacilityDefUB + 10000)

			While oData.Read
                glFacilityDefUB += 1
                If glFacilityDefUB > goFacilityDef.GetUpperBound(0) Then
                    ReDim Preserve goFacilityDef(glFacilityDefUB + 10000)
                    ReDim Preserve glFacilityDefIdx(glFacilityDefUB + 10000)
                End If
				goFacilityDef(glFacilityDefUB) = New FacilityDef()
				With goFacilityDef(glFacilityDefUB)
                    .FillFromDataReader(oData, ObjectType.eFacilityDef)
					glFacilityDefIdx(glFacilityDefUB) = .ObjectID
                End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Unit Def Weapons...")
			sSQL = "SELECT * FROM tblUnitDefWeapon"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				oTmpUnitDef = GetEpicaUnitDef(CInt(oData("UnitDefID")))
				If oTmpUnitDef Is Nothing = False Then
					oTmpWeaponDef = GetEpicaWeaponDef(CInt(oData("WeaponDefID")))
					If oTmpWeaponDef Is Nothing = False Then
						oTmpUnitDef.AddWeaponDef(CInt(oData("UDW_ID")), oTmpWeaponDef, CType(oData("Quadrant"), UnitArcs), CInt(oData("EntityStatusGroup")), CInt(oData("iAmmoCap")))
					End If
				End If
			End While
			oTmpUnitDef = Nothing
			oTmpWeaponDef = Nothing
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Structure Def Weapons...")
			sSQL = "SELECT * FROM tblStructureDefWeapon"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				oTmpFacilityDef = GetEpicaFacilityDef(CInt(oData("StructureDefID")))
				If oTmpFacilityDef Is Nothing = False Then
					oTmpWeaponDef = GetEpicaWeaponDef(CInt(oData("WeaponDefID")))
					If oTmpWeaponDef Is Nothing = False Then
						oTmpFacilityDef.AddWeaponDef(CInt(oData("SDW_ID")), oTmpWeaponDef, CType(oData("Quadrant"), UnitArcs), CInt(oData("EntityStatusGroup")), CInt(oData("iAmmoCap")))
					End If
				End If
			End While
			oTmpFacilityDef = Nothing
			oTmpWeaponDef = Nothing
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Entity Def Hangar Doors...")
			sSQL = "SELECT * FROM tblEntityDefHangarDoor"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oTmpDef As Epica_Entity_Def = CType(GetEpicaObject(CInt(oData("EntityDefID")), CShort(oData("EntityDefTypeID"))), Epica_Entity_Def)
				If oTmpDef Is Nothing = False Then
					oTmpDef.AddHangarDoor(CInt(oData("ED_HD_ID")), CInt(oData("DoorSize")), CByte(oData("SideArc")))
				End If
				oTmpDef = Nothing
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Production Costs...")
			sSQL = "SELECT * FROM tblProductionCost WHERE ObjTypeID IN (" & ObjectType.eFacilityDef & ", " & _
				ObjectType.eUnitDef & ", " & ObjectType.eSpecialTech & ")"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				'Produceable items... units, facilities
				'TODO: Add more as it comes up
				Select Case Val(oData("ObjTypeID"))
					Case ObjectType.eFacilityDef
						oTmpFacilityDef = GetEpicaFacilityDef(CInt(oData("ObjectID")))
						If oTmpFacilityDef Is Nothing = False Then
							oTmpFacilityDef.ProductionRequirements = New ProductionCost()
							With oTmpFacilityDef.ProductionRequirements
								.ColonistCost = CInt(oData("Colonists"))
								.CreditCost = CLng(oData("Credits"))
								.EnlistedCost = CInt(oData("Enlisted"))
								.ObjectID = CInt(oData("ObjectID"))
								.ObjTypeID = CShort(oData("ObjTypeID"))
								.OfficerCost = CInt(oData("Officers"))
								.PointsRequired = CLng(oData("PointsRequired"))
								.ProductionCostType = CByte(oData("ProductionCostType"))
								.PC_ID = CInt(oData("PC_ID"))
							End With
						End If
					Case ObjectType.eUnitDef
						oTmpUnitDef = GetEpicaUnitDef(CInt(oData("ObjectID")))
						If oTmpUnitDef Is Nothing = False Then
							oTmpUnitDef.ProductionRequirements = New ProductionCost()
							With oTmpUnitDef.ProductionRequirements
								.ColonistCost = CInt(oData("Colonists"))
								.CreditCost = CLng(oData("Credits"))
								.EnlistedCost = CInt(oData("Enlisted"))
								.ObjectID = CInt(oData("ObjectID"))
								.ObjTypeID = CShort(oData("ObjTypeID"))
								.OfficerCost = CInt(oData("Officers"))
								.PointsRequired = CLng(oData("PointsRequired"))
								.ProductionCostType = CByte(oData("ProductionCostType"))
								.PC_ID = CInt(oData("PC_ID"))
							End With
						End If
					Case ObjectType.eSpecialTech
						Dim oSpecTech As SpecialTech = GetEpicaSpecialTech(CInt(oData("ObjectID")))
						If oSpecTech Is Nothing = False Then
							If CByte(oData("ProductionCostType")) = 0 Then
								'TODO: Production Cost?
							Else
								oSpecTech.oResearchCost = New ProductionCost
								With oSpecTech.oResearchCost
									.ColonistCost = CInt(oData("Colonists"))
									.CreditCost = CLng(oData("Credits"))
									.EnlistedCost = CInt(oData("Enlisted"))
									.ObjectID = CInt(oData("ObjectID"))
									.ObjTypeID = CShort(oData("ObjTypeID"))
									.OfficerCost = CInt(oData("Officers"))
									.PointsRequired = CLng(oData("PointsRequired"))
									.ProductionCostType = CByte(oData("ProductionCostType"))
									.PC_ID = CInt(oData("PC_ID"))
								End With
							End If
						End If
				End Select
			End While
			oTmpUnitDef = Nothing
			oTmpFacilityDef = Nothing
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Now, go back thru and add our minerals
			LogEvent(LogEventType.Informational, "Loading Production Item Costs...")

			sSQL = "SELECT tblProductionCost.ObjectID, tblProductionCost.ObjTypeID, tblProductionCostItem.PCM_ID, " & _
			  "tblProductionCostItem.ItemID, tblProductionCostItem.ItemTypeID, tblProductionCostItem.Quantity, " & _
			  "tblProductionCostItem.PC_ID, tblProductionCost.ProductionCostType FROM tblProductionCost " & _
			  "RIGHT JOIN tblProductionCostItem ON tblProductionCost.PC_ID = tblProductionCostItem.PC_ID " & _
			  "WHERE tblProductionCost.ObjTypeID IN (" & ObjectType.eFacilityDef & ", " & ObjectType.eUnitDef & ", " & _
			  ObjectType.eSpecialTech & ")"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				'Produceable items... units, facilities
				'TODO: Add more as it comes up
				Select Case Val(oData("ObjTypeID"))
					Case ObjectType.eFacilityDef
						oTmpFacilityDef = GetEpicaFacilityDef(CInt(oData("ObjectID")))
						If oTmpFacilityDef Is Nothing = False Then
							If oTmpFacilityDef.ProductionRequirements Is Nothing = False Then
								With oTmpFacilityDef.ProductionRequirements
									.ItemCostUB += 1
									ReDim Preserve .ItemCosts(.ItemCostUB)
									X = .ItemCostUB
								End With
								oTmpFacilityDef.ProductionRequirements.ItemCosts(X) = New ProductionCostItem()
								With oTmpFacilityDef.ProductionRequirements.ItemCosts(X)
									'.MineralID = CInt(oData("MineralID"))
									.ItemID = CInt(oData("ItemID"))
									.ItemTypeID = CShort(oData("ItemTypeID"))

									.oProdCost = oTmpFacilityDef.ProductionRequirements
									.PC_ID = CInt(oData("PC_ID"))
									.PCM_ID = CInt(oData("PCM_ID"))
									.QuantityNeeded = CInt(oData("Quantity"))
								End With
							End If
						End If
					Case ObjectType.eUnitDef
						oTmpUnitDef = GetEpicaUnitDef(CInt(oData("ObjectID")))
						If oTmpUnitDef Is Nothing = False Then
							If oTmpUnitDef.ProductionRequirements Is Nothing = False Then
								With oTmpUnitDef.ProductionRequirements
									.ItemCostUB += 1
									ReDim Preserve .ItemCosts(.ItemCostUB)
									X = .ItemCostUB
								End With
								oTmpUnitDef.ProductionRequirements.ItemCosts(X) = New ProductionCostItem()
								With oTmpUnitDef.ProductionRequirements.ItemCosts(X)
									'.MineralID = CInt(oData("MineralID"))
									.ItemID = CInt(oData("ItemID"))
									.ItemTypeID = CShort(oData("ItemTypeID"))

									.oProdCost = oTmpUnitDef.ProductionRequirements
									.PC_ID = CInt(oData("PC_ID"))
									.PCM_ID = CInt(oData("PCM_ID"))
									.QuantityNeeded = CInt(oData("Quantity"))
								End With
							End If
						End If
					Case ObjectType.eSpecialTech
						Dim oSpecTech As SpecialTech = GetEpicaSpecialTech(CInt(oData("ObjectID")))
						If oSpecTech Is Nothing = False Then
							If CByte(oData("ProductionCostType")) = 0 Then
								'TODO: Production Cost?
							Else
								If oSpecTech.oResearchCost Is Nothing = False Then
									oSpecTech.oResearchCost.ItemCostUB += 1
									ReDim Preserve oSpecTech.oResearchCost.ItemCosts(oSpecTech.oResearchCost.ItemCostUB)
									With oSpecTech.oResearchCost.ItemCosts(oSpecTech.oResearchCost.ItemCostUB)
										.ItemID = CInt(oData("ItemID"))
										.ItemTypeID = CShort(oData("ItemTypeID"))
										.oProdCost = oSpecTech.oResearchCost
										.PC_ID = CInt(oData("PC_ID"))
										.PCM_ID = CInt(oData("PCM_ID"))
										.QuantityNeeded = CInt(oData("Quantity"))
									End With
								End If
							End If
						End If
				End Select

			End While
			oTmpUnitDef = Nothing
			oTmpFacilityDef = Nothing
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Now, load our EntityDef.EntityMineralDefs :)
			LogEvent(LogEventType.Informational, "Loading Entity Def Minerals...")
			sSQL = "SELECT * FROM tblEntityDefMineral"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Select Case Val(oData("EntityDefTypeID"))
					Case ObjectType.eFacilityDef
						oTmpFacilityDef = GetEpicaFacilityDef(CInt(oData("EntityDefID")))
						If oTmpFacilityDef Is Nothing = False Then

							oTmpFacilityDef.lEntityDefMineralUB += 1
							ReDim Preserve oTmpFacilityDef.EntityDefMinerals(oTmpFacilityDef.lEntityDefMineralUB)
							oTmpFacilityDef.EntityDefMinerals(oTmpFacilityDef.lEntityDefMineralUB) = New EntityDefMineral()
							With oTmpFacilityDef.EntityDefMinerals(oTmpFacilityDef.lEntityDefMineralUB)
								.iEntityDefTypeID = ObjectType.eFacilityDef
								.lQuantity = CInt(oData("Quantity"))
								.oEntityDef = oTmpFacilityDef
								.oMineral = GetEpicaMineral(CInt(oData("MineralID")))
							End With
						End If
					Case ObjectType.eUnitDef
						oTmpUnitDef = GetEpicaUnitDef(CInt(oData("EntityDefID")))
						If oTmpUnitDef Is Nothing = False Then

							oTmpUnitDef.lEntityDefMineralUB += 1
							ReDim Preserve oTmpUnitDef.EntityDefMinerals(oTmpUnitDef.lEntityDefMineralUB)
							oTmpUnitDef.EntityDefMinerals(oTmpUnitDef.lEntityDefMineralUB) = New EntityDefMineral()
							With oTmpUnitDef.EntityDefMinerals(oTmpUnitDef.lEntityDefMineralUB)
								.iEntityDefTypeID = ObjectType.eUnitDef
								.lQuantity = CInt(oData("Quantity"))
								.oEntityDef = oTmpUnitDef
								.oMineral = GetEpicaMineral(CInt(oData("MineralID")))
							End With
						End If
					Case Else
						LogEvent(LogEventType.Warning, "Entity Def Mineral.EntityTypeID of " & CInt(oData("EntityDefTypeID")) & " unexpected. Record skipped.")
				End Select
			End While
			oTmpUnitDef = Nothing
			oTmpFacilityDef = Nothing
			oData.Close()
			oData = Nothing
			oComm = Nothing

			LogEvent(LogEventType.Informational, "Loading Entity Def Critical Hit Chances...")
			For X = 0 To glUnitDefUB
				If glUnitDefIdx(X) <> -1 Then
					If goUnitDef(X).oPrototype Is Nothing = False AndAlso goUnitDef(X).oPrototype.oHullTech Is Nothing = False Then
						goUnitDef(X).oPrototype.oHullTech.GetSlotCriticalChances(goUnitDef(X))
					End If
				End If
			Next X
			For X = 0 To glFacilityDefUB
				If glFacilityDefIdx(X) <> -1 Then
					If goFacilityDef(X).oPrototype Is Nothing = False AndAlso goFacilityDef(X).oPrototype.oHullTech Is Nothing = False Then
						goFacilityDef(X).oPrototype.oHullTech.GetSlotCriticalChances(CType(goFacilityDef(X), Epica_Entity_Def))
					End If
				End If
			Next X

			bResult = True
		Catch
			LogEvent(LogEventType.CriticalError, "LoadDesigns Error: " & Err.Description)
		Finally
			If oData Is Nothing = False Then
				oData.Close()
			End If
			oData = Nothing
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Private Function LoadRemaining(ByVal sPlayerInStr As String) As Boolean
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim oComm As OleDb.OleDbCommand
		Dim sSQL As String
		Dim bResult As Boolean = False
		Dim lTemp As Int32
		Dim oTmpRel As PlayerRel

		Try
			LogEvent(LogEventType.Informational, "Loading Minerals Geography Rels...")
			sSQL = "SELECT * FROM tblMineralGeographyRel"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oTmpMin As Mineral = GetEpicaMineral(CInt(oData("MineralID")))
				If oTmpMin Is Nothing = False Then
					Dim lID As Int32 = CInt(oData("ObjectID"))
					Dim iTypeID As Int16 = CShort(oData("ObjTypeID"))
					If iTypeID = ObjectType.ePlanet Then
						Dim oPlanet As Planet = GetEpicaPlanet(lID)
						If oPlanet Is Nothing OrElse oPlanet.InMyDomain = False Then Continue While
					ElseIf iTypeID = ObjectType.eSolarSystem Then
						Dim oSystem As SolarSystem = GetEpicaSystem(lID)
						If oSystem Is Nothing OrElse oSystem.InMyDomain = False Then Continue While
					End If
					oTmpMin.AddMinGeoRel(CInt(oData("ObjectID")), CShort(oData("ObjTypeID")), CInt(oData("RedistMaxQty")), CInt(oData("RedistMaxConcentration")))
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Colony
			LogEvent(LogEventType.Informational, "Loading Colonies...")
			sSQL = "SELECT tblColony.*, tblColony.ParentID As TrueParentID, tblColony.ParentTypeID As TrueParentTypeID FROM tblColony WHERE ParentTypeID = 3 "
			sSQL &= "UNION SELECT tblColony.*, tblStructure.ParentID As TrueParentID, tblStructure.ParentTypeID As TrueParentTypeID FROM "
			sSQL &= "tblColony LEFT OUTER JOIN tblStructure ON tblColony.ParentID = tblStructure.StructureID WHERE tblColony.ParentTypeID = 10"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            Dim lNeedToDeletes() As Int32 = Nothing
            Dim lNeedToDeleteUB As Int32 = -1
			While oData.Read
                If oData("TrueParentID") Is DBNull.Value = True OrElse oData("TrueParentTypeID") Is DBNull.Value = True Then
                    If oData("ColonyID") Is DBNull.Value = False Then
                        lNeedToDeleteUB += 1
                        ReDim Preserve lNeedToDeletes(lNeedToDeleteUB)
                        lNeedToDeletes(lNeedToDeleteUB) = CInt(oData("ColonyID"))
                    End If
                    Continue While
                End If

				Dim lTrueParentID As Int32 = CInt(oData("TrueParentID"))
				Dim iTrueParentTypeID As Int16 = CShort(oData("TrueParentTypeID"))
				If iTrueParentTypeID = ObjectType.ePlanet Then
					Dim oPlanet As Planet = GetEpicaPlanet(lTrueParentID)
					If oPlanet Is Nothing OrElse oPlanet.InMyDomain = False Then Continue While
				ElseIf iTrueParentTypeID = ObjectType.eSolarSystem Then
					Dim oSystem As SolarSystem = GetEpicaSystem(lTrueParentID)
					If oSystem Is Nothing OrElse oSystem.InMyDomain = False Then Continue While
				End If

				glColonyUB = glColonyUB + 1
				ReDim Preserve goColony(glColonyUB)
				ReDim Preserve glColonyIdx(glColonyUB)
				goColony(glColonyUB) = New Colony()
				With goColony(glColonyUB)
					.LoadFromDataReader(oData)

					glColonyIdx(glColonyUB) = .ObjectID

					'Now, add this colony index to the Owner's Fast Colony Lookup
					.Owner.AddColonyIndex(glColonyUB)

					'We should have our PLANETS by now
					If .ParentObject Is Nothing = False Then
						If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
							CType(.ParentObject, Planet).AddColonyReference(glColonyUB)
						End If
					End If
				End With
			End While
			oData.Close()
			oData = Nothing
            oComm = Nothing

            If lNeedToDeleteUB > -1 Then
                sSQL = "DELETE FROM tblColony WHERE ColonyID IN ("
                Dim sINList As String = ""
                For X As Int32 = 0 To lNeedToDeleteUB
                    If sINList <> "" Then sINList &= ", "
                    sINList &= lNeedToDeletes(X).ToString
                Next X
                If sINList <> "" Then
                    sSQL &= sINList & ")"

                    LogEvent(LogEventType.Informational, "Deleting Error Colonies: " & sSQL)
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oComm.ExecuteNonQuery()
                    oComm.Dispose()
                    oComm = Nothing
                End If
            End If

            'Structure
            System.GC.Collect()

            ReDim Preserve goFacility(100000)
            ReDim Preserve glFacilityIdx(100000)
			LogEvent(LogEventType.Informational, "Loading Structures...")
			sSQL = "SELECT *, 1 As SortByItem FROM tblStructure WHERE ProductionTypeID = " & ProductionType.eSpaceStationSpecial & _
			  " UNION SELECT *, 2 As SortByItem FROM tblStructure WHERE ProductionTypeID = " & ProductionType.ePowerCenter & _
			  " UNION SELECT *, 3 As SortByItem FROM tblStructure WHERE ProductionTypeID <> " & ProductionType.eSpaceStationSpecial & _
			  " AND ProductionTypeID <> " & ProductionType.ePowerCenter & " ORDER BY SortByItem"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lParentID As Int32 = CInt(oData("ParentID"))
				Dim iParentTypeID As Int16 = CShort(oData("ParentTypeID"))

				If iParentTypeID = ObjectType.eSolarSystem Then
					Dim oSystem As SolarSystem = GetEpicaSystem(lParentID)
					If oSystem Is Nothing OrElse oSystem.InMyDomain = False Then Continue While
				ElseIf iParentTypeID = ObjectType.ePlanet Then
					Dim oPlanet As Planet = GetEpicaPlanet(lParentID)
					If oPlanet Is Nothing OrElse oPlanet.InMyDomain = False Then Continue While
				ElseIf iParentTypeID = ObjectType.eFacility Then
					Dim oFac As Facility = GetEpicaFacility(lParentID)
					If oFac Is Nothing Then Continue While
					If oFac.ParentObject Is Nothing Then Continue While
					If CType(oFac.ParentObject, Epica_GUID).InMyDomain = False Then Continue While
				End If

                glFacilityUB = glFacilityUB + 1
                If glFacilityUB > goFacility.GetUpperBound(0) Then
                    System.GC.Collect()
                    Dim lNewUB As Int32 = glFacilityUB + 100000
                    ReDim Preserve goFacility(lNewUB)
                    ReDim Preserve glFacilityIdx(lNewUB)
                End If
				goFacility(glFacilityUB) = New Facility()
				With goFacility(glFacilityUB)
					.ServerIndex = glFacilityUB
					.LoadFromDataReader(oData)
					glFacilityIdx(glFacilityUB) = .ObjectID
                    If .ParentObject Is Nothing Then
                        glFacilityIdx(glFacilityUB) = -1
                    Else
                        AddLookupFacility(.ObjectID, .ObjTypeID, .ServerIndex)
                    End If

                    'If .Owner Is Nothing = False Then
                    '    If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                    '        .Owner.AddSpaceFacIdx(.ServerIndex)
                    '    End If
                    'End If
				End With
			End While
			oData.Close()
			oData = Nothing
            oComm = Nothing

            ReDim Preserve goFacility(glFacilityUB)
            ReDim Preserve glFacilityIdx(glFacilityUB)

            System.GC.Collect()

			'Associate colonies to structures (because of the race condition)
			LogEvent(LogEventType.Informational, "Associating colonies to Parents...")
			sSQL = "SELECT ColonyID, ParentID, ParentTypeID FROM tblColony WHERE ParentTypeID = " & ObjectType.eFacility
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oTmpColony As Colony = CType(GetEpicaObject(CInt(oData("ColonyID")), ObjectType.eColony), Colony)
				If oTmpColony Is Nothing = False Then
					oTmpColony.ParentObject = GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
					oTmpColony.DataChanged()
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

            If LoadAgentDetails(sPlayerInStr) = False Then Return False

			'Now, formation defs
			LogEvent(LogEventType.Informational, "Loading Formation Defs...")
			sSQL = "SELECT * FROM tblFormationDef WHERE OwnerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glFormationDefUB += 1
				ReDim Preserve goFormationDefs(glFormationDefUB)
				ReDim Preserve glFormationDefIdx(glFormationDefUB)
				goFormationDefs(glFormationDefUB) = New FormationDef
				With goFormationDefs(glFormationDefUB)
					.FormationID = CInt(oData("FormationDefID"))
					ReDim .yName(19)
					.yName = StringToBytes(CStr(oData("FormationName")))
					.lOwnerID = CInt(oData("OwnerID"))
					.yCriteria = CType(oData("yCriteria"), CriteriaType)
					.yDefault = CByte(oData("yDefault"))
					.lCellSize = CInt(oData("CellSize"))
					ReDim .mlSlots(FormationDef.ml_GRID_SIZE_WH - 1, FormationDef.ml_GRID_SIZE_WH - 1)
					glFormationDefIdx(glFormationDefUB) = .FormationID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'And the formation def data...
			LogEvent(LogEventType.Informational, "Loading Formation Defs...")
			sSQL = "SELECT * FROM tblFormationDefSlot WHERE FormationDefID IN (Select FormationDefID FROM tblFormationDef WHERE OwnerID IN " & sPlayerInStr & ")"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oFormation As FormationDef = GetEpicaFormation(CInt(oData("FormationDefID")))
				If oFormation Is Nothing = False Then
					Dim lSlotX As Int32 = CInt(oData("SlotX"))
					Dim lSlotY As Int32 = CInt(oData("SlotY"))
					Dim lSlotValue As Int32 = CInt(oData("SlotValue"))
					If oFormation.mlSlots Is Nothing Then ReDim oFormation.mlSlots(FormationDef.ml_GRID_SIZE_WH - 1, FormationDef.ml_GRID_SIZE_WH - 1)
					oFormation.mlSlots(lSlotX, lSlotY) = lSlotValue
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Now, UnitGroups
			LogEvent(LogEventType.Informational, "Loading Unit Groups...")
			sSQL = "SELECT * FROM tblUnitGroup WHERE OwnerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glUnitGroupUB = glUnitGroupUB + 1
				ReDim Preserve goUnitGroup(glUnitGroupUB)
				ReDim Preserve glUnitGroupIdx(glUnitGroupUB)
				goUnitGroup(glUnitGroupUB) = New UnitGroup()
				With goUnitGroup(glUnitGroupUB)
					.ObjectID = CInt(oData("UnitGroupID"))
					.ObjTypeID = ObjectType.eUnitGroup
					.oOwner = GetEpicaPlayer(CInt(oData("OwnerID")))
					.oParentObject = GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
					.UnitGroupName = StringToBytes(CStr(oData("UnitGroupName")))
					.InterSystemMoveCyclesRemaining = CInt(oData("InterSystemCycles"))
					.lLastInterSystemCycleUpdate = glCurrentCycle
					.lInterSystemOriginID = CInt(oData("InterSystemOrigin"))
					.lInterSystemTargetID = CInt(oData("InterSystemTargetID"))
					.yInterSystemSpeed = CByte(oData("InterSystemSpeed"))
					.lOriginX = CInt(oData("OriginX"))
					.lOriginY = CInt(oData("OriginY"))
                    .lOriginZ = CInt(oData("OriginZ"))
                    .lDefaultFormationID = CInt(oData("DefaultFormationID"))
					glUnitGroupIdx(glUnitGroupUB) = .ObjectID

                    If .lInterSystemTargetID <> -1 Then .SetDestSystem(.lInterSystemTargetID, True)
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

            'Now, unit's
            System.GC.Collect()
            ReDim Preserve goUnit(100000)
            ReDim Preserve glUnitIdx(100000)
			LogEvent(LogEventType.Informational, "Loading Units...")
            sSQL = "SELECT * FROM tblUnit ORDER BY ParentTypeID"       'by parenttypeid because this allows a unit to be loaded that may contain another unit first
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read

				Dim lParentID As Int32 = CInt(oData("ParentID"))
				Dim iParentTypeID As Int16 = CShort(oData("ParentTypeID"))

				If iParentTypeID = ObjectType.eSolarSystem Then
					Dim oSystem As SolarSystem = GetEpicaSystem(lParentID)
					If oSystem Is Nothing OrElse oSystem.InMyDomain = False Then Continue While
				ElseIf iParentTypeID = ObjectType.ePlanet Then
					Dim oPlanet As Planet = GetEpicaPlanet(lParentID)
					If oPlanet Is Nothing OrElse oPlanet.InMyDomain = False Then Continue While
				ElseIf iParentTypeID = ObjectType.eFacility Then
					Dim oFac As Facility = GetEpicaFacility(lParentID)
					If oFac Is Nothing Then Continue While
					If oFac.ParentObject Is Nothing Then Continue While
					If CType(oFac.ParentObject, Epica_GUID).InMyDomain = False Then Continue While
				ElseIf iParentTypeID = ObjectType.eGalaxy Then
                    'OK, determine the dest of the unit... is the there a unit group?
                    lTemp = CInt(oData("UnitGroupID"))
                    Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(lTemp)
                    Dim bIsGood As Boolean = False
                    If oUnitGroup Is Nothing = False Then
                        bIsGood = True 'oUnitGroup.InMyDomain
                    End If
                    If bIsGood = False Then
                        'one more thing, is the dest system in my domain?
                        'Stop    'remove me after we load
                        Dim lDestID As Int32 = CInt(oData("DestEnvirID"))
                        Dim iDestTypeID As Int16 = CShort(oData("DestEnvirTypeID"))
                        If iDestTypeID = ObjectType.ePlanet Then
                            Dim oDestPlanet As Planet = GetEpicaPlanet(lDestID)
                            If oDestPlanet Is Nothing = False Then
                                bIsGood = oDestPlanet.InMyDomain
                            End If
                        ElseIf iDestTypeID = ObjectType.eSolarSystem Then
                            Dim oDestSystem As SolarSystem = GetEpicaSystem(lDestID)
                            If oDestSystem Is Nothing = False Then
                                bIsGood = oDestSystem.InMyDomain
                            End If
                        Else
                            bIsGood = True
                        End If
                    End If
                    If bIsGood = False Then Continue While
				ElseIf iParentTypeID = ObjectType.eUnit Then
					Dim oUnit As Unit = GetEpicaUnit(lParentID)
					If oUnit Is Nothing Then Continue While
					If oUnit.ParentObject Is Nothing Then Continue While

					With CType(oUnit.ParentObject, Epica_GUID)
						If .ObjTypeID = ObjectType.ePlanet Then
							If .InMyDomain = False Then Continue While
						ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
							If .InMyDomain = False Then Continue While
						ElseIf .ObjTypeID = ObjectType.eFacility OrElse .ObjTypeID = ObjectType.eUnit Then
							'Try to get to the first geographical object I can find
							Dim iTempVal As Int16 = .ObjTypeID
							Dim oObj As Object = oUnit.ParentObject
							Dim lCnt As Int32 = 0
							While iTempVal = ObjectType.eFacility OrElse iTempVal = ObjectType.eUnit
								lCnt += 1
								With CType(CType(oObj, Epica_Entity).ParentObject, Epica_GUID)
									iTempVal = .ObjTypeID
									oObj = CType(oObj, Epica_Entity).ParentObject
								End With
								If lCnt > 100 Then
									oObj = Nothing
									Exit While
								End If
							End While
							If oObj Is Nothing OrElse CType(oObj, Epica_GUID).InMyDomain = False Then Continue While
						Else
                            'LogEvent(LogEventType.CriticalError, "LoadRemaining.Units Determining InMyDomain was whack!")
                            'Continue While	'TODO:what else could there be?
						End If
					End With
				End If

                glUnitUB += 1
                If glUnitUB > goUnit.GetUpperBound(0) Then
                    System.GC.Collect()
                    Dim lNewUB As Int32 = glUnitUB + 100000
                    ReDim Preserve goUnit(lNewUB)
                    ReDim Preserve glUnitIdx(lNewUB)
                End If
                goUnit(glUnitUB) = New Unit()
				With goUnit(glUnitUB)
					.LoadFromDataReader(oData)

					glUnitIdx(glUnitUB) = .ObjectID
                    lTemp = CInt(oData("UnitGroupID"))

                    Dim lParentTypeID As Int32 = CInt(oData("ParentTypeID"))
                    If lParentTypeID = ObjectType.eGalaxy Then
                        Dim lTestFleetID As Int32 = CInt(oData("TestFleetID"))
                        If lTemp < 1 Then lTemp = lTestFleetID
                    End If

					For X As Int32 = 0 To glUnitGroupUB
						If glUnitGroupIdx(X) = lTemp Then
                            goUnitGroup(X).AddUnit(glUnitUB, False)
							Exit For
						End If
					Next X

					If .ParentObject Is Nothing Then glUnitIdx(glUnitUB) = -1
				End With
			End While
			oData.Close()
			oData = Nothing
            oComm = Nothing

            'If True = True Then Stop 'remove this line and the stop above in the if blocks for parenttype = galaxy

            ReDim Preserve goUnit(glUnitUB)
            ReDim Preserve glUnitIdx(glUnitUB)
            System.GC.Collect()

			LogEvent(LogEventType.Informational, "Loading Unit Routes...")
			sSQL = "SELECT * FROM tblRouteItem ORDER BY OrderNum"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lUnitID As Int32 = CInt(oData("UnitID"))
				Dim oUnit As Unit = GetEpicaUnit(lUnitID)
				If oUnit Is Nothing = False Then
					Dim uNewRouteItem As RouteItem
					With uNewRouteItem
						.lLocX = CInt(oData("LocX"))
						.lLocZ = CInt(oData("LocZ"))
						.lOrderNum = CInt(oData("OrderNum"))
						.oDest = CType(GetEpicaObject(CInt(oData("DestID")), CShort(oData("DestTypeID"))), Epica_GUID)
						Dim lTmpVal As Int32 = CInt(oData("LoadItemID"))
                        Dim iTmpVal As Int16 = CShort(oData("LoadItemTypeID"))
                        .oLoadItem = Nothing

                        Dim yFlag As Byte = CByte(oData("ExtraFlags"))
                        .SetLoadItem(lTmpVal, iTmpVal, oUnit.Owner, yFlag)
                    End With
					oUnit.AddRouteItem(uNewRouteItem)
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Structure Production
			LogEvent(LogEventType.Informational, "Loading Structure Productions...")
			sSQL = "SELECT * FROM tblStructureProduction ORDER BY StructureID, OrderNum"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
                Dim lTmpID As Int32 = CInt(oData("StructureID"))

                Dim lFacIdx As Int32 = LookupFacility(lTmpID, ObjectType.eFacility)
                If lFacIdx < 0 OrElse lFacIdx > glFacilityUB Then
                    lFacIdx = -1
                    For X As Int32 = 0 To glFacilityUB
                        If glFacilityIdx(X) = lTmpID Then
                            lFacIdx = X
                            Exit For
                        End If
                    Next X
                    If lFacIdx = -1 Then
                        LogEvent(LogEventType.Warning, "Structure Production Missing Structure for " & lTmpID)
                        Continue While
                    End If
                End If

                If goFacility(lFacIdx).AddProduction(CInt(oData("ObjectID")), CShort(oData("ObjTypeID")), CByte(oData("OrderNum")), CInt(oData("ProdCount")), CLng(oData("PointsProduced"))) = True Then
                    AddEntityProducing(lFacIdx, ObjectType.eFacility, goFacility(lFacIdx).ObjectID)
                End If 
            End While
			oData.Close()
			oData = Nothing
			oComm = Nothing 

			'Unit Production
			LogEvent(LogEventType.Informational, "Loading Unit Productions...")
			sSQL = "SELECT * FROM tblUnitProduction"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lTmpID As Int32 = CInt(oData("UnitID"))
				Dim oUnit As Unit = Nothing
				Dim lIdx As Int32 = -1
				For X As Int32 = 0 To glUnitUB
					If glUnitIdx(X) = lTmpID Then
						oUnit = goUnit(X)
						lIdx = X
						Exit For
					End If
				Next X
				If oUnit Is Nothing = False Then
					oUnit.bProducing = True
					oUnit.CurrentProduction = New EntityProduction()
					oUnit.RecalcProduction()
					With oUnit.CurrentProduction
						.oParent = oUnit
						.ProductionID = CInt(oData("ObjectID"))
						.ProductionTypeID = CShort(oData("ObjTypeID"))

						If .ProductionTypeID = ProductionType.eFacility Then
							Dim oFacDef As FacilityDef = GetEpicaFacilityDef(.ProductionID)
							If oFacDef Is Nothing = False Then
								Try
									.PointsRequired = oFacDef.ProductionRequirements.PointsRequired
								Catch
								End Try
							End If
						End If

						.OrderNumber = 1
						.PointsProduced = CLng(oData("PointsProduced"))
						.lProdCount = 1
						.lProdX = oUnit.LocX
						.lProdZ = oUnit.LocZ
						.iProdA = CShort(oData("ProdAngle"))
						.lLastUpdateCycle = glCurrentCycle
						Try
							.lFinishCycle = CInt(glCurrentCycle + ((.PointsRequired - .PointsProduced) \ oUnit.mlProdPoints))
						Catch
						End Try
					End With
					AddEntityProducing(lIdx, ObjectType.eUnit, oUnit.ObjectID)

					'for move lock condition
					oUnit.CurrentStatus = oUnit.CurrentStatus Or elUnitStatus.eMoveLock
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Reinforcing facilities
			LogEvent(LogEventType.Informational, "Associating Reinforcing Facilities...")
			For X As Int32 = 0 To glFacilityUB
				If glFacilityIdx(X) <> -1 AndAlso goFacility(X).lReinforcingUnitGroupID > 0 Then
					Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(goFacility(X).lReinforcingUnitGroupID)
					If oUnitGroup Is Nothing = False Then oUnitGroup.AddReinforcer(goFacility(X).ObjectID) Else goFacility(X).lReinforcingUnitGroupID = -1
				End If
			Next X

			''Corporations
			'LogEvent(LogEventType.Informational, "Loading Corporations...")
			'sSQL = "SELECT * FROM tblCorporation"
			'oComm = New OleDb.OleDbCommand(sSQL, goCN)
			'oData = oComm.ExecuteReader(CommandBehavior.Default)
			'While oData.Read
			'    glCorporationUB += 1
			'    ReDim Preserve goCorporation(glCorporationUB)
			'    ReDim Preserve glCorporationIdx(glCorporationUB)
			'    goCorporation(glCorporationUB) = New Corporation()
			'    With goCorporation(glCorporationUB)
			'        .CorporationName = StringToBytes(CStr(oData("CorporationName")))
			'        .CurrentStockValue = CInt(oData("CurrentStockValue"))
			'        .HeadQuarters = GetEpicaFacility(CInt(oData("HQ_ID")))
			'        .MaximumValue = CInt(oData("MaximumValue"))
			'        .MinimumValue = CInt(oData("MinimumValue"))
			'        .ObjectID = CInt(oData("CorporationID"))
			'        .ObjTypeID = ObjectType.eCorporation
			'        glCorporationIdx(glCorporationUB) = .ObjectID
			'    End With
			'End While
			'oData.Close()
			'oData = Nothing
			'oComm = Nothing

			'Player Comm Folder
			LogEvent(LogEventType.Informational, "Loading Player Comm Folders...")
			sSQL = "SELECT * FROM tblPlayerCommFolder WHERE PlayerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
				If oPlayer Is Nothing = False AndAlso oPlayer.InMyDomain = True Then
					With oPlayer
						'since we are loading
						.AddFolder(CInt(oData("PCF_ID")), CStr(oData("FolderName")))
					End With
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Player Comm
			LogEvent(LogEventType.Informational, "Loading Player Comms...")
			sSQL = "SELECT * FROM tblPlayerComm WHERE PlayerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
				If oPlayer Is Nothing = False AndAlso oPlayer.InMyDomain = True Then
                    oPlayer.AddEmailMsg(CInt(oData("PC_ID")), CInt(oData("PCF_ID")), CShort(oData("EncryptLevel")), _
                      CStr(oData("MsgBody")), CStr(oData("MsgTitle")), CInt(oData("SentByID")), CInt(oData("SentOn")), _
                      CByte(oData("ReadFlag")) <> 0, CStr(oData("SentTo")), Nothing)
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Now, player comm attachments
			LogEvent(LogEventType.Informational, "Loading Player Comm Attachments...")
			sSQL = "SELECT tblEmailAttachment.*, tblPlayerComm.PlayerID, tblPlayerComm.PCF_ID FROM tblEmailAttachment " & _
			  "LEFT OUTER JOIN tblPlayerComm ON tblEmailAttachment.PC_ID = tblPlayerComm.PC_ID WHERE PlayerID IN " & sPlayerInStr & " ORDER BY PlayerID"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			If oData Is Nothing = False Then
				Dim oPlayer As Player = Nothing
				While oData.Read
					If oData("PlayerID") Is DBNull.Value Then Continue While
					If oPlayer Is Nothing OrElse oPlayer.ObjectID <> CInt(oData("PlayerID")) Then oPlayer = GetEpicaPlayer(CInt(oData("PlayerID")))
					If oPlayer Is Nothing = False Then
						oPlayer.AddEmailAttachment(CInt(oData("PC_ID")), CInt(oData("PCF_ID")), CByte(oData("AttachNumber")), _
						  CInt(oData("EnvirID")), CShort(oData("EnvirTypeID")), CInt(oData("LocX")), CInt(oData("LocZ")), CStr(oData("WPName")))
					End If
				End While
			End If
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Player to Player Rel
			LogEvent(LogEventType.Informational, "Loading Player to Player Rels")
			sSQL = "SELECT * FROM tblPlayerToPlayerRel"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				oTmpRel = New PlayerRel()

				With oTmpRel
					.oPlayerRegards = GetEpicaPlayer(CInt(oData("Player1ID")))
					.oThisPlayer = GetEpicaPlayer(CInt(oData("Player2ID")))
                    .WithThisScore = CByte(oData("RelTypeID"))

                    .blTotalWarpointsGained = CLng(oData("TotalWPGain"))
                    .lPlayersWPV = CInt(oData("WPV"))

                    Dim lCycles As Int32 = CInt(oData("CyclesToNextUpdate"))
                    If lCycles > 0 Then .lCycleOfNextScoreUpdate = glCurrentCycle + lCycles
                    .TargetScore = CByte(oData("TargetScore"))

                    If .TargetScore >= .WithThisScore Then
                        .lCycleOfNextScoreUpdate = -1
                    ElseIf .lCycleOfNextScoreUpdate > -1 Then
                        AddToQueue(.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oTmpRel.oPlayerRegards.ObjectID, -1, oTmpRel.oThisPlayer.ObjectID, -1, -1, -1, -1, -1)
                    End If
				End With
                oTmpRel.oPlayerRegards.SetPlayerRel(oTmpRel.oThisPlayer.ObjectID, oTmpRel, False)
				oTmpRel = Nothing
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

            ''Stocks
            'LogEvent(LogEventType.Informational, "Loading Stocks...")
            ''TODO: need to get the stocks and put them somewhere...

			'initialize our mineral caches
			LogEvent(LogEventType.Informational, "Initializing Mineral Cache Arrays...")
			sSQL = "SELECT * FROM tblMineralCache"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lParentID As Int32 = CInt(oData("ParentID"))
				Dim iParentTypeID As Int16 = CShort(oData("ParentTypeID"))

                'If iParentTypeID = ObjectType.eSolarSystem Then
                '	Dim oSystem As SolarSystem = GetEpicaSystem(lParentID)
                '	If oSystem Is Nothing OrElse oSystem.InMyDomain = False Then Continue While
                'ElseIf iParentTypeID = ObjectType.ePlanet Then
                '	Dim oPlanet As Planet = GetEpicaPlanet(lParentID)
                '	If oPlanet Is Nothing OrElse oPlanet.InMyDomain = False Then Continue While
                'ElseIf iParentTypeID = ObjectType.eFacility Then
                '	Dim oFac As Facility = GetEpicaFacility(lParentID)
                '	If oFac Is Nothing Then Continue While
                '	If oFac.ParentObject Is Nothing Then Continue While
                '	If CType(oFac.ParentObject, Epica_GUID).InMyDomain = False Then Continue While
                'ElseIf iParentTypeID = ObjectType.eUnit Then
                '	Dim oUnit As Unit = GetEpicaUnit(lParentID)
                '	If oUnit Is Nothing Then Continue While
                '	If oUnit.ParentObject Is Nothing Then Continue While

                '	With CType(oUnit.ParentObject, Epica_GUID)
                '		If .ObjTypeID = ObjectType.ePlanet Then
                '			If .InMyDomain = False Then Continue While
                '		ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                '			If .InMyDomain = False Then Continue While
                '		ElseIf .ObjTypeID = ObjectType.eFacility OrElse .ObjTypeID = ObjectType.eUnit Then
                '			'Try to get to the first geographical object I can find
                '			Dim iTempVal As Int16 = .ObjTypeID
                '			Dim oObj As Object = oUnit.ParentObject
                '			Dim lCnt As Int32 = 0
                '			While iTempVal = ObjectType.eFacility OrElse iTempVal = ObjectType.eUnit
                '				lCnt += 1
                '				With CType(CType(oObj, Epica_Entity).ParentObject, Epica_GUID)
                '					iTempVal = .ObjTypeID
                '					oObj = CType(oObj, Epica_Entity).ParentObject
                '				End With
                '				If lCnt > 100 Then
                '					oObj = Nothing
                '					Exit While
                '				End If
                '			End While
                '                        'If oObj Is Nothing OrElse CType(oObj, Epica_GUID).InMyDomain = False Then Continue While
                '                        'Else
                '                        'LogEvent(LogEventType.CriticalError, "LoadRemaining.MineralCaches Determining InMyDomain was whack!")
                '                        'Continue While	'TODO:what else could there be?
                '		End If
                '	End With
                'ElseIf iParentTypeID = ObjectType.eColony Then
                '	Dim oColony As Colony = GetEpicaColony(lParentID)
                '	If oColony Is Nothing Then Continue While
                '	If oColony.ParentObject Is Nothing Then Continue While
                '	With CType(oColony.ParentObject, Epica_GUID)
                '		If .ObjTypeID = ObjectType.ePlanet Then
                '			If .InMyDomain = False Then Continue While
                '		ElseIf .ObjTypeID = ObjectType.eFacility Then
                '			Dim oTmp As Facility = CType(oColony.ParentObject, Facility)
                '			If oTmp.ParentObject Is Nothing Then Continue While
                '			If CType(oTmp.ParentObject, Epica_GUID).InMyDomain = False Then Continue While
                '		End If
                '	End With
                'End If
                If GetEpicaObject(lParentID, iParentTypeID) Is Nothing = False Then glMineralCacheUB += 1
			End While
			ReDim Preserve goMineralCache(glMineralCacheUB)
			ReDim Preserve glMineralCacheIdx(glMineralCacheUB)
			For X As Int32 = 0 To glMineralCacheUB
				glMineralCacheIdx(X) = -1
			Next X
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'mineral caches
			LogEvent(LogEventType.Informational, "Loading Mineral Caches...")
			sSQL = "SELECT * FROM tblMineralCache"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			Dim lCurrentIdx As Int32 = -1
			While oData.Read

				Dim lParentID As Int32 = CInt(oData("ParentID"))
				Dim iParentTypeID As Int16 = CShort(oData("ParentTypeID"))

                'If iParentTypeID = ObjectType.eSolarSystem Then
                '	Dim oSystem As SolarSystem = GetEpicaSystem(lParentID)
                '	If oSystem Is Nothing OrElse oSystem.InMyDomain = False Then Continue While
                'ElseIf iParentTypeID = ObjectType.ePlanet Then
                '	Dim oPlanet As Planet = GetEpicaPlanet(lParentID)
                '	If oPlanet Is Nothing OrElse oPlanet.InMyDomain = False Then Continue While
                'ElseIf iParentTypeID = ObjectType.eFacility Then
                '	Dim oFac As Facility = GetEpicaFacility(lParentID)
                '	If oFac Is Nothing Then Continue While
                '	If oFac.ParentObject Is Nothing Then Continue While
                '	If CType(oFac.ParentObject, Epica_GUID).InMyDomain = False Then Continue While
                'ElseIf iParentTypeID = ObjectType.eUnit Then
                '	Dim oUnit As Unit = GetEpicaUnit(lParentID)
                '	If oUnit Is Nothing Then Continue While
                '	If oUnit.ParentObject Is Nothing Then Continue While

                '	With CType(oUnit.ParentObject, Epica_GUID)
                '		If .ObjTypeID = ObjectType.ePlanet Then
                '			If .InMyDomain = False Then Continue While
                '		ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                '			If .InMyDomain = False Then Continue While
                '		ElseIf .ObjTypeID = ObjectType.eFacility OrElse .ObjTypeID = ObjectType.eUnit Then
                '			'Try to get to the first geographical object I can find
                '			Dim iTempVal As Int16 = .ObjTypeID
                '			Dim oObj As Object = oUnit.ParentObject
                '			Dim lCnt As Int32 = 0
                '			While iTempVal = ObjectType.eFacility OrElse iTempVal = ObjectType.eUnit
                '				lCnt += 1
                '				With CType(CType(oObj, Epica_Entity).ParentObject, Epica_GUID)
                '					iTempVal = .ObjTypeID
                '					oObj = CType(oObj, Epica_Entity).ParentObject
                '				End With
                '				If lCnt > 100 Then
                '					oObj = Nothing
                '					Exit While
                '				End If
                '			End While
                '                        'If oObj Is Nothing OrElse CType(oObj, Epica_GUID).InMyDomain = False Then Continue While
                '                        'Else
                '                        'LogEvent(LogEventType.CriticalError, "LoadRemaining.MineralCaches Determining InMyDomain was whack!")
                '                        'Continue While	'TODO:what else could there be?
                '		End If
                '	End With
                'ElseIf iParentTypeID = ObjectType.eColony Then
                '	Dim oColony As Colony = GetEpicaColony(lParentID)
                '	If oColony Is Nothing Then Continue While
                '	If oColony.ParentObject Is Nothing Then Continue While
                '	With CType(oColony.ParentObject, Epica_GUID)
                '		If .ObjTypeID = ObjectType.ePlanet Then
                '			If .InMyDomain = False Then Continue While
                '		ElseIf .ObjTypeID = ObjectType.eFacility Then
                '			Dim oTmp As Facility = CType(oColony.ParentObject, Facility)
                '			If oTmp.ParentObject Is Nothing Then Continue While
                '			If CType(oTmp.ParentObject, Epica_GUID).InMyDomain = False Then Continue While
                '		End If
                '	End With
                'End If

                Dim oParentObject As Object = GetEpicaObject(lParentID, iParentTypeID)
                If oParentObject Is Nothing Then Continue While

				lCurrentIdx += 1
				goMineralCache(lCurrentIdx) = New MineralCache()
				With goMineralCache(lCurrentIdx)
					.lServerIndex = lCurrentIdx
					.CacheTypeID = CByte(oData("CacheTypeID"))
					.Concentration = CInt(oData("Concentration"))
					.LocX = CInt(oData("LocX"))
					.LocZ = CInt(oData("LocY"))
					.ObjectID = CInt(oData("CacheID"))
					.ObjTypeID = ObjectType.eMineralCache
					.oMineral = GetEpicaMineral(CInt(oData("MineralID")))
                    .ParentObject = oParentObject 'GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
					.OriginalConcentration = CInt(oData("OriginalConcentration"))
                    .Quantity = CInt(oData("Quantity"))

					'Ensure we register with a containing object's cargo...
					If .ParentObject Is Nothing = False Then
						If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility OrElse CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eUnit Then
							CType(.ParentObject, Epica_Entity).AddCargoRef(CType(goMineralCache(lCurrentIdx), Epica_GUID))
						ElseIf CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eColony Then
							Dim oColony As Colony = CType(.ParentObject, Colony)
							oColony.mlMineralCacheUB += 1
							ReDim Preserve oColony.mlMineralCacheID(oColony.mlMineralCacheUB)
							ReDim Preserve oColony.mlMineralCacheIdx(oColony.mlMineralCacheUB)
							ReDim Preserve oColony.mlMineralCacheMineralID(oColony.mlMineralCacheUB)
							oColony.mlMineralCacheID(oColony.mlMineralCacheUB) = .ObjectID
							oColony.mlMineralCacheIdx(oColony.mlMineralCacheUB) = .lServerIndex
							oColony.mlMineralCacheMineralID(oColony.mlMineralCacheUB) = .oMineral.ObjectID
						End If
					End If
                    glMineralCacheIdx(lCurrentIdx) = .ObjectID
                    .bNeedsAsync = False
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing


			'Component Caches
			LogEvent(LogEventType.Informational, "Loading Component Caches...")
			sSQL = "SELECT * FROM tblComponentCache"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read

				Dim lParentID As Int32 = CInt(oData("ParentID"))
				Dim iParentTypeID As Int16 = CShort(oData("ParentTypeID"))

                'If iParentTypeID = ObjectType.eSolarSystem Then
                '	Dim oSystem As SolarSystem = GetEpicaSystem(lParentID)
                '	If oSystem Is Nothing OrElse oSystem.InMyDomain = False Then Continue While
                'ElseIf iParentTypeID = ObjectType.ePlanet Then
                '	Dim oPlanet As Planet = GetEpicaPlanet(lParentID)
                '	If oPlanet Is Nothing OrElse oPlanet.InMyDomain = False Then Continue While
                'ElseIf iParentTypeID = ObjectType.eFacility Then
                '	Dim oFac As Facility = GetEpicaFacility(lParentID)
                '	If oFac Is Nothing Then Continue While
                '	If oFac.ParentObject Is Nothing Then Continue While
                '	If CType(oFac.ParentObject, Epica_GUID).InMyDomain = False Then Continue While
                'ElseIf iParentTypeID = ObjectType.eUnit Then
                '	Dim oUnit As Unit = GetEpicaUnit(lParentID)
                '	If oUnit Is Nothing Then Continue While
                '	If oUnit.ParentObject Is Nothing Then Continue While

                '	With CType(oUnit.ParentObject, Epica_GUID)
                '		If .ObjTypeID = ObjectType.ePlanet Then
                '			If .InMyDomain = False Then Continue While
                '		ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                '			If .InMyDomain = False Then Continue While
                '		ElseIf .ObjTypeID = ObjectType.eFacility OrElse .ObjTypeID = ObjectType.eUnit Then
                '			'Try to get to the first geographical object I can find
                '			Dim iTempVal As Int16 = .ObjTypeID
                '			Dim oObj As Object = oUnit.ParentObject
                '			Dim lCnt As Int32 = 0
                '			While iTempVal = ObjectType.eFacility OrElse iTempVal = ObjectType.eUnit
                '				lCnt += 1
                '				With CType(CType(oObj, Epica_Entity).ParentObject, Epica_GUID)
                '					iTempVal = .ObjTypeID
                '					oObj = CType(oObj, Epica_Entity).ParentObject
                '				End With
                '				If lCnt > 100 Then
                '					oObj = Nothing
                '					Exit While
                '				End If
                '			End While
                '			If oObj Is Nothing OrElse CType(oObj, Epica_GUID).InMyDomain = False Then Continue While
                '		Else
                '			LogEvent(LogEventType.CriticalError, "LoadRemaining.ComponentCaches Determining InMyDomain was whack!")
                '			Continue While	'TODO:what else could there be?
                '		End If
                '                End With
                '            ElseIf iParentTypeID = ObjectType.eColony Then
                '                Dim oColony As Colony = GetEpicaColony(lParentID)
                '                If oColony Is Nothing Then Continue While
                '                If oColony.ParentObject Is Nothing Then Continue While
                '                With CType(oColony.ParentObject, Epica_GUID)
                '                    If .ObjTypeID = ObjectType.ePlanet Then
                '                        If .InMyDomain = False Then Continue While
                '                    ElseIf .ObjTypeID = ObjectType.eFacility Then
                '                        Dim oTmp As Facility = CType(oColony.ParentObject, Facility)
                '                        If oTmp.ParentObject Is Nothing Then Continue While
                '                        If CType(oTmp.ParentObject, Epica_GUID).InMyDomain = False Then Continue While
                '                    End If
                '                End With
                'End If

				glComponentCacheUB += 1
				ReDim Preserve goComponentCache(glComponentCacheUB)
				ReDim Preserve glComponentCacheIdx(glComponentCacheUB)
				goComponentCache(glComponentCacheUB) = New ComponentCache
				With goComponentCache(glComponentCacheUB)
					.ObjectID = CInt(oData("ComponentCacheID"))
					.ObjTypeID = ObjectType.eComponentCache
					.ComponentOwnerID = CInt(oData("ComponentOwnerID"))
					.ComponentID = CInt(oData("ComponentID"))
					.ComponentTypeID = CShort(oData("ComponentTypeID"))

					.ParentObject = GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
					.LocX = CInt(oData("LocX"))
					.LocZ = CInt(oData("LocZ"))
                    .Quantity = CInt(oData("Quantity"))
                    .bNeedsSaved = False

					'Ensure we register with a containing object's cargo...
					If .ParentObject Is Nothing = False Then
						If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility OrElse CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eUnit Then
                            CType(.ParentObject, Epica_Entity).AddCargoRef(CType(goComponentCache(glComponentCacheUB), Epica_GUID))
                        ElseIf CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eColony Then
                            Dim oColony As Colony = CType(.ParentObject, Colony)
                            oColony.mlComponentCacheUB += 1
                            ReDim Preserve oColony.mlComponentCacheID(oColony.mlComponentCacheUB)
                            ReDim Preserve oColony.mlComponentCacheIdx(oColony.mlComponentCacheUB)
                            ReDim Preserve oColony.mlComponentCacheCompID(oColony.mlComponentCacheUB)
                            oColony.mlComponentCacheID(oColony.mlComponentCacheUB) = .ObjectID
                            oColony.mlComponentCacheIdx(oColony.mlComponentCacheUB) = glComponentCacheUB
                            oColony.mlComponentCacheCompID(oColony.mlComponentCacheUB) = .ComponentID
						End If
					End If
					glComponentCacheIdx(glComponentCacheUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing


			'EntityWeaponDefAmmo
            'LogEvent(LogEventType.Informational, "Loading Ammo Values...")
            'sSQL = "SELECT * FROM tblEntityWeaponDefAmmo"
            'oComm = New OleDb.OleDbCommand(sSQL, goCN)
            'oData = oComm.ExecuteReader(CommandBehavior.Default)
            'Dim lWpnDefID As Int32
            'While oData.Read
            '	'Ok, get our entity
            '	If CShort(oData("ObjTypeID")) = ObjectType.eUnit Then
            '		Dim oUnit As Unit = GetEpicaUnit(CInt(oData("ObjectID")))
            '		If oUnit Is Nothing = False Then
            '			With oUnit
            '				lWpnDefID = CInt(oData("WeaponDefID"))
            '				For lTemp = 0 To .EntityDef.WeaponDefUB
            '					If .EntityDef.WeaponDefs(lTemp).ObjectID = lWpnDefID Then
            '						.lCurrentAmmo(lTemp) = CInt(oData("AmmoCnt"))
            '						Exit For
            '					End If
            '				Next lTemp
            '			End With
            '		End If
            '	Else
            '		Dim oFac As Facility = GetEpicaFacility(CInt(oData("ObjectID")))
            '		If oFac Is Nothing = False Then
            '			With oFac
            '				lWpnDefID = CInt(oData("WeaponDefID"))
            '				For lTemp = 0 To .EntityDef.WeaponDefUB
            '					If .EntityDef.WeaponDefs(lTemp).ObjectID = lWpnDefID Then
            '						.lCurrentAmmo(lTemp) = CInt(oData("AmmoCnt"))
            '						Exit For
            '					End If
            '				Next lTemp
            '			End With
            '		End If
            '	End If

            'End While
            'oData.Close()
            'oData = Nothing
            'oComm = Nothing

			LogEvent(LogEventType.Informational, "Associating Mining facilities to caches...")
			For lIdx As Int32 = 0 To glFacilityUB
				If goFacility(lIdx).yProductionType = ProductionType.eMining Then
					'Ok, its a mining facility, find if there is a mineral cache underneath it
                    'TODO: Make this less generic (hard-coded) and put in code to get the model's size
                    If goFacility(lIdx).LocX = Int32.MinValue OrElse goFacility(lIdx).LocZ = Int32.MinValue Then Continue For
					Dim rcTemp As Rectangle = Rectangle.FromLTRB(goFacility(lIdx).LocX - 50, goFacility(lIdx).LocZ - 50, goFacility(lIdx).LocX + 50, goFacility(lIdx).LocZ + 50)
					'TODO: to speed this up, we could separate the mineral caches into the two different types (mineable and surface)
					For X As Int32 = 0 To glMineralCacheUB
						If glMineralCacheIdx(X) <> -1 Then
							Dim oCache As MineralCache = goMineralCache(X)
                            If oCache Is Nothing = False AndAlso oCache.CacheTypeID = MineralCacheType.eMineable AndAlso oCache.ParentObject Is Nothing = False Then
                                With oCache
                                    If CType(.ParentObject, Epica_GUID).ObjectID = CType(goFacility(lIdx).ParentObject, Epica_GUID).ObjectID Then
                                        If CType(.ParentObject, Epica_GUID).ObjTypeID = CType(goFacility(lIdx).ParentObject, Epica_GUID).ObjTypeID Then
                                            'Ok, check location
                                            If rcTemp.Contains(.LocX, .LocZ) = True Then
                                                'Ok, found one
                                                If .BeingMinedBy Is Nothing Then
                                                    goFacility(lIdx).lCacheIndex = X
                                                    goFacility(lIdx).lCacheID = .ObjectID
                                                    If goFacility(lIdx).Active Then goFacility(lIdx).bMining = True '???
                                                    .BeingMinedBy = goFacility(lIdx)
                                                    AddFacilityMining(lIdx, goFacility(lIdx).ObjectID)
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End With
                            End If
						End If
                    Next X
                ElseIf (goFacility(lIdx).EntityDef.ModelID And 255) = 148 Then        'the OMP model
                    'Ok, get the parent envir
                    If goFacility(lIdx).ParentObject Is Nothing = False AndAlso CType(goFacility(lIdx).ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                        Dim oSystem As SolarSystem = CType(goFacility(lIdx).ParentObject, SolarSystem)
                        If oSystem Is Nothing = False Then
                            Dim oMinedPlanet As Planet = Nothing
                            Dim lDist As Int32 = Int32.MaxValue
                            For X As Int32 = 0 To oSystem.mlPlanetUB
                                Dim lPIdx As Int32 = oSystem.GetPlanetIdx(X)
                                Dim oPlanet As Planet = goPlanet(lPIdx)
                                If oPlanet Is Nothing = False Then
                                    Dim fDX As Single = goFacility(lIdx).LocX - oPlanet.LocX
                                    Dim fDZ As Single = goFacility(lIdx).LocZ - oPlanet.LocZ
                                    fDX *= fDX
                                    fDZ *= fDZ
                                    Dim fTemp As Single = CSng(Math.Sqrt(fDX + fDZ))
                                    If fTemp < lDist Then
                                        lDist = CInt(fTemp)
                                        oMinedPlanet = oPlanet
                                    End If
                                End If
                            Next X

                            If oMinedPlanet Is Nothing = False Then
                                oMinedPlanet.RingMinerID = goFacility(lIdx).ObjectID
                            End If
                        End If
                    End If
                End If
			Next lIdx

			'Trade
			LogEvent(LogEventType.Informational, "Loading Trade Objects...")
			sSQL = "SELECT tblTrade.* FROM tblTrade WHERE TradeID IN (SELECT TradeID FROM tblTradePlayer WHERE PlayerID IN " & sPlayerInStr & ")"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				'glTradeUB += 1
				'ReDim Preserve goTrades(glTradeUB)
				'ReDim Preserve glTradeIdx(glTradeUB)
				'goTrades(glTradeUB) = New Trade
				'With goTrades(glTradeUB)
				Dim oTrade As New DirectTrade()
				With oTrade
					.ObjectID = CInt(oData("TradeID"))
					.ObjTypeID = ObjectType.eTrade
					.CyclesRemaining = CInt(oData("CyclesRemaining"))
					.TradeCycles = CInt(oData("TradeCycles"))
					.TradeStarted = CInt(oData("TradeStarted"))
					.TradeState = CByte(oData("TradeState"))
					.FailureReason = CByte(oData("FailureReason"))

					'glTradeIdx(glTradeUB) = .ObjectID
				End With
				If goGTC Is Nothing Then goGTC = New GalacticTradeSystem()
                goGTC.AddNewDirectTrade(oTrade)

                If (oTrade.TradeState And DirectTrade.eTradeStateValues.InProgress) <> 0 AndAlso oTrade.TradeState <> DirectTrade.eTradeStateValues.TradeRejected AndAlso (oTrade.TradeState And DirectTrade.eTradeStateValues.TradeCompleted) = 0 Then
                    'Create a Queue Item for the full time
                    Try
                        Dim dt As Date = Date.SpecifyKind(GetDateFromNumber(oTrade.TradeStarted), DateTimeKind.Utc).ToLocalTime
                        Dim lCyclesElapsed As Int32 = CInt(Now.Subtract(dt).TotalSeconds * 30)
                        oTrade.CyclesRemaining = oTrade.TradeCycles - lCyclesElapsed
                        AddToQueue(glCurrentCycle + Math.Max(1, oTrade.CyclesRemaining), QueueItemType.eTradeEventFinal, oTrade.ObjectID, oTrade.ObjTypeID, -1, -1, 0, 0, 0, 0)
                    Catch
                    End Try
                End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Trade Players
			LogEvent(LogEventType.Informational, "Loading Trade Players...")
			'TODO: Optimize this to pull only those we need
			sSQL = "SELECT * FROM tblTradePlayer"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oTrade As DirectTrade = goGTC.GetDirectTrade(CInt(oData("TradeID")))
				If oTrade Is Nothing = False Then
					Dim oTP As TradePlayer = Nothing
					If oTrade.oPlayer1TP Is Nothing Then
						oTrade.oPlayer1TP = New TradePlayer()
						oTP = oTrade.oPlayer1TP
					Else
						oTrade.oPlayer2TP = New TradePlayer
						oTP = oTrade.oPlayer2TP
					End If

					If oTP Is Nothing = False Then
						With oTP
							.lTradePostID = CInt(oData("TradePostID"))
							If oData("Notes") Is DBNull.Value = False Then .Notes = StringToBytes(CStr(oData("Notes")))
							.PlayerID = CInt(oData("PlayerID"))
							.TradeID = CInt(oData("TradeID"))
						End With
					End If
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Trade Player Items
			LogEvent(LogEventType.Informational, "Loading Trade Player Items...")
			'TODO: Optimize this to pull only those we need
			sSQL = "SELECT tblTradePlayerItem.PlayerID, tblTradePlayerItem.ObjectID, tblTradePlayerItem.ObjTypeID, tblTradePlayerItem.Quantity, " & _
			  "tblTradePlayerItem.TradeID, tblTradePlayerItem.ExtendedID, tblTradePlayerItem.Extended2ID FROM tblTradePlayerItem"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oTrade As DirectTrade = goGTC.GetDirectTrade(CInt(oData("TradeID")))
				Dim lPlayerID As Int32 = CInt(oData("PlayerID"))
				If oTrade Is Nothing = False Then
					Dim oTP As TradePlayer = Nothing

					If oTrade.oPlayer1TP Is Nothing = False AndAlso oTrade.oPlayer1TP.PlayerID = lPlayerID Then
						oTP = oTrade.oPlayer1TP
					ElseIf oTrade.oPlayer2TP Is Nothing = False AndAlso oTrade.oPlayer2TP.PlayerID = lPlayerID Then
						oTP = oTrade.oPlayer2TP
					End If

					If oTP Is Nothing = False Then
						oTP.AddItemQuantity(CInt(oData("ObjectID")), CShort(oData("ObjTypeID")), CLng(oData("Quantity")), CInt(oData("ExtendedID")), CInt(oData("Extended2ID")))
					End If
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Sell Orders
			LogEvent(LogEventType.Informational, "Loading Sell Orders...")
			sSQL = "SELECT * FROM tblSellOrder"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lTPID As Int32 = CInt(oData("TradePostID"))
				If GetEpicaFacility(lTPID) Is Nothing Then Continue While

				Dim oSellOrder As New SellOrder()
				With oSellOrder
					.lTradeID = CInt(oData("lTradeID"))
					.iTradeTypeID = CShort(oData("iTradeTypeID"))
					.TradePostID = CInt(oData("TradePostID"))
					.blPrice = CLng(oData("Price"))
					.blQuantity = CLng(oData("Quantity"))
					.iExtTypeID = CShort(oData("ExtTypeID"))
                    StringToBytes(CStr(oData("ItemName"))).CopyTo(.yItemName, 0)

                    If .iTradeTypeID = ObjectType.eUnit Then
                        Dim oUnit As Unit = GetEpicaUnit(.lTradeID)
                        If oUnit Is Nothing = False Then oUnit.bUnitInSellOrder = True
                    End If
				End With
				goGTC.LoadSellOrder(oSellOrder)
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Buy Orders
			LogEvent(LogEventType.Informational, "Loading Buy Orders...")
			sSQL = "SELECT * FROM tblBuyOrder WHERE BuyOrderType <> 0"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lTPID As Int32 = CInt(oData("TradePostID"))
				If GetEpicaFacility(lTPID) Is Nothing Then Continue While

				Dim oBuyOrder As New BuyOrder()
				With oBuyOrder
					.blEscrow = CLng(oData("Escrow"))
					.blPaymentAmt = CLng(oData("PaymentAmt"))
					.blQuantity = CLng(oData("Quantity"))
					.BuyOrderID = CInt(oData("BuyOrderID"))
					.iTradeTypeID = CShort(oData("TradeTypeID"))
					.lAcceptedByID = CInt(oData("AcceptedByID"))
					.lAcceptedOn = CInt(oData("AcceptedOn"))
					.lDeadline = CInt(oData("Deadline"))
					.lSpecificID = CInt(oData("SpecificID"))
					.TradePostID = CInt(oData("TradePostID"))
					.yBuyOrderState = CByte(oData("BuyOrderState"))
					.yBuyOrderType = CByte(oData("BuyOrderType"))
				End With
				goGTC.LoadBuyOrder(oBuyOrder)
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Buy Order Properties
			LogEvent(LogEventType.Informational, "Loading Buy Order Properties...")
			sSQL = "SELECT * FROM tblBuyOrderProperty"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oBuyOrder As BuyOrder = goGTC.GetBuyOrderItem(CInt(oData("BuyOrderID")))
				If oBuyOrder Is Nothing = False Then
					oBuyOrder.AddBOProp(CInt(oData("PropertyID")), CInt(oData("PropertyValue")), CByte(oData("CompareType")))
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Trade History
            LogEvent(LogEventType.Informational, "Loading Trade History...")
            Dim lMinDate As Int32 = GetDateAsNumber(Now.AddDays(-14))
            sSQL = "SELECT * FROM tblTradeHistory WHERE PlayerID IN " & sPlayerInStr & " AND TransactionDate > " & lMinDate & " ORDER BY TransactionDate"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oTH As New TradeHistory()
				With oTH
					.blTradeAmt = CLng(oData("TradeAmt"))
					.lOtherPlayerID = CInt(oData("OtherPlayerID"))
					.lPlayerID = CInt(oData("PlayerID"))
					.lTransactionDate = CInt(oData("TransactionDate"))
					.yTradeEventType = CType(oData("TradeEventType"), TradeHistory.TradeEventType)
                    .yTradeResult = CType(oData("TradeResult"), TradeHistory.TradeResult)
                    .lDeliveryTime = CInt(oData("DeliveryTime"))
				End With

				Dim oPlayer As Player = GetEpicaPlayer(oTH.lPlayerID)
				If oPlayer Is Nothing = False Then oPlayer.AddTradeHistory(oTH)
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Trade History Items
			LogEvent(LogEventType.Informational, "Loading Trade History Items...")
            sSQL = "SELECT * FROM tblTradeHistoryItem WHERE PlayerID IN " & sPlayerInStr & " AND TransactionDate > " & lMinDate
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
				If oPlayer Is Nothing = False Then
					oPlayer.LoadTradeHistoryItem(CInt(oData("OtherPlayerID")), CInt(oData("TransactionDate")), CStr(oData("ItemName")), CLng(oData("Quantity")), CType(oData("TradeHistoryItemType"), TradeHistory.TradeHistoryItemType))
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Trade Deliveries
			LogEvent(LogEventType.Informational, "Loading Trade Deliveries...")
			sSQL = "SELECT * FROM tblTradeDelivery"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lToTP As Int32 = CInt(oData("ToTradePostID"))
				Dim oTP As Facility = GetEpicaFacility(lToTP)
				If oTP Is Nothing = False Then
					Dim oTD As New TradeDelivery()
					With oTD
						.lID1 = CInt(oData("ID1"))
                        .iID2 = CShort(oData("ID2"))
                        If .iID2 = ObjectType.eUnit Then
                            'Ok, get the unit
                            Dim oUnit As Unit = GetEpicaUnit(.lID1)
                            If oUnit Is Nothing = False Then
                                'ok, check the unit's parent
                                If oUnit.ParentObject Is Nothing = False Then
                                    Try
                                        Dim oPGuid As Epica_GUID = CType(oUnit.ParentObject, Epica_GUID)
                                        If oPGuid Is Nothing = False Then
                                            If oPGuid.ObjTypeID = ObjectType.eFacility Then
                                                Dim oFac As Facility = CType(oPGuid, Facility)
                                                For X As Int32 = 0 To oFac.lHangarUB
                                                    If oFac.lHangarIdx(X) = .lID1 Then
                                                        If CType(oFac.oHangarContents(X), Epica_GUID).ObjTypeID = ObjectType.eUnit Then
                                                            oFac.lHangarIdx(X) = -1
                                                            Exit For
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Catch
                                    End Try
                                End If
                                oUnit.ParentObject = Nothing
                            End If
                        End If
						.lID3 = CInt(oData("ID3"))
						.lTradePostID = CInt(oData("ToTradePostID"))
						.lSourceTradePostID = CInt(oData("FromTradePostID"))
						.dtDelivery = GetDateFromNumber(CInt(oData("DeliveryOn")))
						.dtStartedOn = GetDateFromNumber(CInt(oData("StartedOn")))
						.blQty = CLng(oData("Quantity"))
					End With
					goGTC.LoadTradeDelivery(oTD)
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'BUGS
			'LogEvent(LogEventType.Informational, "Loading Bug Reports...")
			'sSQL = "SELECT * FROM tblBug"
			'oComm = New OleDb.OleDbCommand(sSQL, goCN)
			'oData = oComm.ExecuteReader(CommandBehavior.Default)
			'While oData.Read
			'	glBugUB = glBugUB + 1
			'	ReDim Preserve goBugs(glBugUB)
			'	goBugs(glBugUB) = New EpicaBug
			'	With goBugs(glBugUB)
			'		.lBugID = CInt(oData("BugID"))
			'		.lCreatedBy = CInt(oData("UserID"))
			'		If oData("DevNotes").Equals(DBNull.Value) = False Then .sDevNotes = CStr(oData("DevNotes")) Else .sDevNotes = ""
			'		If oData("ProblemDesc").Equals(DBNull.Value) = False Then .sProblemDesc = CStr(oData("ProblemDesc")) Else .sProblemDesc = ""
			'		If oData("StepsToProduce").Equals(DBNull.Value) = False Then .sStepsToProduce = CStr(oData("StepsToProduce")) Else .sStepsToProduce = ""
			'		.yCategory = CByte(oData("CategoryID"))
			'		.yOccurs = CByte(oData("OccursID"))
			'		.yPriority = CByte(oData("PriorityID"))
			'		.yStatus = CByte(oData("statusID"))
			'		.ySubCat = CByte(oData("SubCatID"))
			'		.lAssignedTo = CInt(oData("AssignedTo"))
			'	End With
			'End While
			'oData.Close()
			'oData = Nothing
			'oComm = Nothing

			'Agent Effects in progress
			LogEvent(LogEventType.Informational, "Loading Agent Effects...")
			sSQL = "SELECT * FROM tblAgentEffects"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lID As Int32 = CInt(oData("EffectedItemID"))
				Dim iTypeID As Int16 = CShort(oData("EffectedItemTypeID"))
				Dim lRemainingCycle As Int32 = CInt(oData("RemainingCycles"))
				Dim yType As Byte = CByte(oData("EffectType"))
				Dim lAmount As Int32 = CInt(oData("EffectAmount"))
                Dim bAmountAsPerc As Boolean = CInt(oData("yAmountAsPerc")) <> 0
                Dim lCausedByID As Int32 = 0
                If oData("CausedByID") Is DBNull.Value = False Then lCausedByID = CInt(oData("CausedByID"))
                Dim oCausedBy As Player = Nothing
                If lCausedByID > 0 Then oCausedBy = GetEpicaPlayer(lCausedByID)

				If iTypeID = ObjectType.eBudget Then
					'ok, lID is a player
					Dim oPlayer As Player = GetEpicaPlayer(lID)
					If oPlayer Is Nothing = False AndAlso oPlayer.oBudget Is Nothing = False Then
                        Dim oEffect As AgentEffect = oPlayer.oBudget.AddAgentEffect(glCurrentCycle, lRemainingCycle, CType(yType, AgentEffectType), lAmount, bAmountAsPerc, lCausedByID)
                        If oEffect Is Nothing = False AndAlso oCausedBy Is Nothing = False Then oCausedBy.AddImposedEffect(oEffect, oPlayer, "Budget")
					End If
					oPlayer = Nothing
				ElseIf iTypeID = ObjectType.eColony Then
					Dim oColony As Colony = GetEpicaColony(lID)
                    If oColony Is Nothing = False Then
                        Dim oEffect As AgentEffect = oColony.AddAgentEffect(glCurrentCycle, lRemainingCycle, CType(yType, AgentEffectType), lAmount, bAmountAsPerc, lCausedByID)
                        If oEffect Is Nothing = False AndAlso oCausedBy Is Nothing = False Then oCausedBy.AddImposedEffect(oEffect, oColony.Owner, BytesToString(oColony.ColonyName))
                    End If

					oColony = Nothing
				ElseIf iTypeID = ObjectType.eUnit Then
					Dim oUnit As Unit = GetEpicaUnit(lID)
                    If oUnit Is Nothing = False Then
                        Dim oEffect As AgentEffect = oUnit.AddAgentEffect(glCurrentCycle, lRemainingCycle, CType(yType, AgentEffectType), lAmount, bAmountAsPerc, lCausedByID)
                        If oEffect Is Nothing = False AndAlso oCausedBy Is Nothing = False Then oCausedBy.AddImposedEffect(oEffect, oUnit.Owner, BytesToString(oUnit.EntityName))
                    End If
                    oUnit = Nothing
				ElseIf iTypeID = ObjectType.eFacility Then
					Dim oFac As Facility = GetEpicaFacility(lID)
                    If oFac Is Nothing = False Then
                        Dim oEffect As AgentEffect = oFac.AddAgentEffect(glCurrentCycle, lRemainingCycle, CType(yType, AgentEffectType), lAmount, bAmountAsPerc, lCausedByID)
                        If oEffect Is Nothing = False AndAlso oCausedBy Is Nothing = False Then oCausedBy.AddImposedEffect(oEffect, oFac.Owner, BytesToString(oFac.EntityName))
                    End If
				Else
					LogEvent(LogEventType.Warning, "AgentEffect typeID unexpected: " & iTypeID & ". Agent Effect disregarded.")
				End If

			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Player Item Intel
			LogEvent(LogEventType.Informational, "Loading Player Item Intel...")
			sSQL = "SELECT * FROM tblPlayerItemIntel WHERE PlayerID IN " & sPlayerInStr
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
				If oPlayer Is Nothing = False Then
					Dim oPII As New PlayerItemIntel
					With oPII
						.dtTimeStamp = CDate(oData("IntelTimeStamp"))
						.EnvirID = CInt(oData("EnvirID"))
                        .EnvirTypeID = CShort(oData("EnvirTypeID"))
                        .FullKnowledge = Base64.DecodeToByteArray(CStr(oData("FullKnowledge")))
                        .iItemTypeID = CShort(oData("ItemTypeID"))
						.lItemID = CInt(oData("ItemID"))
						.LocX = CInt(oData("LocX"))
						.LocZ = CInt(oData("LocZ"))
						.lOtherPlayerID = CInt(oData("OtherPlayerID"))
						.lPlayerID = CInt(oData("PlayerID"))
						.lValue = CInt(oData("IntelValue"))
						.StatusValue = CInt(oData("StatusValue"))
						.yIntelType = CType(oData("IntelType"), PlayerItemIntel.PlayerItemIntelType)
                        .yArchived = CByte(oData("bArchived"))
					End With
					oPlayer.mlItemIntelUB += 1
					ReDim Preserve oPlayer.moItemIntel(oPlayer.mlItemIntelUB)
					oPlayer.moItemIntel(oPlayer.mlItemIntelUB) = oPII
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Player Alias'
			LogEvent(LogEventType.Informational, "Loading Player Aliasing...")
			sSQL = "SELECT * FROM tblAlias"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
				If oPlayer Is Nothing = False Then
					Dim oOther As Player = GetEpicaPlayer(CInt(oData("OtherPlayerID")))
					If oOther Is Nothing = False Then
						Dim yUserName(19) As Byte
						StringToBytes(CStr(oData("sAliasUserName"))).CopyTo(yUserName, 0)
						Dim yPassword(19) As Byte
						StringToBytes(CStr(oData("sAliasPassword"))).CopyTo(yPassword, 0)
						Dim lRights As Int32 = CInt(oData("lAliasRights"))
						oPlayer.AddPlayerAlias(oOther, yUserName, yPassword, lRights)
						oOther.AddPlayerAliasAllowance(oPlayer, yUserName, yPassword, lRights)
					End If
				End If
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

            ''Load our senate stuff
            'LogEvent(LogEventType.Informational, "Loading Senate Proposals...")
            'sSQL = "SELECT * FROM tblSenateProposal"
            'oComm = New OleDb.OleDbCommand(sSQL, goCN)
            'oData = oComm.ExecuteReader(CommandBehavior.Default)
            'While oData.Read
            '	Dim oProposal As New SenateProposal()
            '	With oProposal
            '		.lProposedBy = CInt(oData("ProposedBy"))
            '		.lVotesAgainst = 0
            '		.lVotesFor = 0
            '		.ObjectID = CInt(oData("ProposalID"))
            '		.ObjTypeID = ObjectType.eSenateLaw
            '		.yDescription = StringToBytes(CStr(oData("ProposalDescription")))
            '		.yProposalState = CType(CByte(oData("ProposalState")), eyProposalState)
            '		.yTitle = StringToBytes(CStr(oData("ProposalTitle")))
            '	End With
            '	Senate.AddNewProposal(oProposal)
            'End While
            'oData.Close()
            'oData = Nothing
            'oComm = Nothing

            'LogEvent(LogEventType.Informational, "Loading Senate Votes...")
            'sSQL = "SELECT * FROM tblSenateProposalVote"
            'oComm = New OleDb.OleDbCommand(sSQL, goCN)
            'oData = oComm.ExecuteReader(CommandBehavior.Default)
            'While oData.Read
            '	Dim oProposal As SenateProposal = Senate.GetProposal(CInt(oData("ProposalID")))
            '	If oProposal Is Nothing = False Then
            '		Dim lVoterID As Int32 = CInt(oData("VoterID"))
            '		Dim yVoteValue As eyVoteValue = CType(oData("VoteValue"), eyVoteValue)
            '		oProposal.HandlePlayerVote(lVoterID, yVoteValue, True)
            '	End If
            'End While
            'oData.Close()
            'oData = Nothing
            'oComm = Nothing

            'LogEvent(LogEventType.Informational, "Recalculating Senate Votes...")
            'Senate.RecalculateAllProposals()

            'LogEvent(LogEventType.Informational, "Loading Senate Messages...")
            'sSQL = "SELECT * FROM tblSenateProposalMsg ORDER BY PostedOn"
            'oComm = New OleDb.OleDbCommand(sSQL, goCN)
            'oData = oComm.ExecuteReader(CommandBehavior.Default)
            'While oData.Read
            '	Dim oProposal As SenateProposal = Senate.GetProposal(CInt(oData("ProposalID")))
            '	If oProposal Is Nothing = False Then
            '		Dim lPosterID As Int32 = CInt(oData("PosterID"))
            '		Dim lPostedOn As Int32 = CInt(oData("PostedOn"))
            '		Dim sMsgData As String = CStr(oData("MsgData"))
            '		oProposal.AddMsg(lPosterID, lPostedOn, sMsgData)
            '	End If
            'End While
            'oData.Close()
            'oData = Nothing
            'oComm = Nothing

            LoadMineralBuyOrders("")
            LoadGuildBillboards("")
            LoadRouteTemplates("")
            If LoadTransports() = False Then Stop

            LogEvent(LogEventType.Informational, "Updating All Values of Colonies...")
			For X As Int32 = 0 To glColonyUB
				goColony(X).UpdateAllValues(-1)
            Next X

            'Ok, time to double check all player's special projects CP count
            LogEvent(LogEventType.Informational, "Verifying all player CP limits...")
            oComm = New OleDb.OleDbCommand("SELECT tblPlayerSpecialTech.PlayerID, Max(tblSpecialTech.NewAttrValue) AS MaxOfNewAttrValue " & _
              "FROM tblSpecialTech LEFT JOIN tblPlayerSpecialTech ON tblSpecialTech.TechID = tblPlayerSpecialTech.SpecialTechID " & _
              "GROUP BY tblPlayerSpecialTech.PlayerID, tblSpecialTech.ProgramControl, tblPlayerSpecialTech.ResPhase " & _
              "HAVING (((tblSpecialTech.ProgramControl)=7) AND ((tblPlayerSpecialTech.ResPhase)=2))", goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                Dim lPlayerID As Int32 = CInt(oData("PlayerID"))
                Dim lMaxVal As Int32 = CInt(oData("MaxOfNewAttrValue"))

                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then
                    If oPlayer.oSpecials.iCPLimit < lMaxVal Then Stop
                Else
                    Stop
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm.Dispose()
            oComm = Nothing

            bResult = True
		Catch
			LogEvent(LogEventType.CriticalError, "LoadRemaining Error: " & Err.Description)
		Finally
			If oData Is Nothing = False Then
				oData.Close()
			End If
			oData = Nothing
			oComm = Nothing
		End Try
		Return bResult
    End Function

    Private Function LoadTransports() As Boolean
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim oComm As OleDb.OleDbCommand
        Dim sSQL As String
        Dim bResult As Boolean = False

        Dim lMinInactiveDate As Int32 = GetDateAsNumber(Now.AddDays(-2))

        LogEvent(LogEventType.Informational, "Loading Transports...")
        Try
            sSQL = "SELECT * FROM tblTransport WHERE OwnerID IN (SELECT PlayerID FROM tblPlayer WHERE (AccountStatus = 1 OR AccountStatus = 101) OR DateAccountWentInactive > " & lMinInactiveDate.ToString & ")"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                Dim oTrans As New Transport
                With oTrans
                    .TransportID = CInt(oData("TransportID"))
                    ReDim .UnitName(19)
                    StringToBytes(CStr(oData("UnitName"))).CopyTo(.UnitName, 0)
                    .OwnerID = CInt(oData("OwnerID"))
                    .LocationID = CInt(oData("LocationID"))
                    .LocationTypeID = CShort(oData("LocationTypeID"))
                    .DestinationID = CInt(oData("DestinationID"))
                    .DestinationTypeID = CShort(oData("DestinationTypeID"))
                    .ETA = Date.SpecifyKind(GetBigDateFromNumber(CLng(oData("ETA"))), DateTimeKind.Utc)
                    .DepartureDate = Date.SpecifyKind(GetBigDateFromNumber(CLng(oData("DepartureDate"))), DateTimeKind.Utc)
                    .UnitDefID = CInt(oData("UnitDefID"))
                    .TransFlags = CByte(oData("RU_Flags"))
                    .CurrentWaypoint = CByte(oData("CurrentWaypoint"))
                    .LocX = CInt(oData("LocX"))
                    .LocZ = CInt(oData("LocZ"))
                    'Test Location and Dest for lost colonies
                    If .LocationTypeID = ObjectType.eColony AndAlso .DestinationTypeID = ObjectType.eColony Then
                        Dim oLocation As Colony = GetEpicaColony(.LocationID)
                        Dim oDestination As Colony = GetEpicaColony(.DestinationID)
                        If oDestination Is Nothing = True AndAlso oLocation Is Nothing = True Then
                            Dim oPlayer As Player = GetEpicaPlayer(.OwnerID)
                            If oPlayer Is Nothing Then Continue While

                            Dim oPlanet As Planet
                            If oPlayer.lIronCurtainPlanet < 1 Then
                                oPlanet = GetEpicaPlanet(oPlayer.lStartedEnvirID)
                            Else
                                oPlanet = GetEpicaPlanet(oPlayer.lIronCurtainPlanet)
                            End If
                            If oPlanet Is Nothing Then Continue While

                            'Inject step to go to HW and unload
                            oTrans.lRouteUB = 0
                            ReDim oTrans.oRoute(oTrans.lRouteUB)
                            .oRoute(.lRouteUB) = New TransportRoute
                            With .oRoute(oTrans.lRouteUB)
                                .OrderNum = 0
                                .oTransport = oTrans
                                .lActionUB = -1
                                ReDim .oActions(-1)
                                .WaypointFlags = 0
                                .DestinationTypeID = ObjectType.ePlanet
                                .DestinationID = oPlanet.ObjectID
                                oTrans.DestinationTypeID = ObjectType.ePlanet
                                oTrans.DestinationID = oPlanet.ObjectID
                                oTrans.CurrentWaypoint = 0
                                oTrans.TransFlags = Transport.elTransportFlags.eEnRoute
                                Dim dtNow As DateTime = DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime
                                oTrans.DepartureDate = DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime
                                oTrans.ETA = dtNow.AddSeconds(60 * 60 * 24) '24 hours
                            End With

                            'Lookup if a player colony exists on this planet.  If so replace that as the destination, and inject an unload route action
                            Dim lCurUB As Int32 = -1
                            If oPlanet.lColonysHereIdx Is Nothing = False Then lCurUB = Math.Min(oPlanet.lColonysHereUB, oPlanet.lColonysHereIdx.GetUpperBound(0))
                            For X As Int32 = 0 To lCurUB
                                Dim lColonyIdx As Int32 = oPlanet.lColonysHereIdx(X)
                                If lColonyIdx > -1 Then
                                    Dim oColony As Colony = goColony(lColonyIdx)
                                    If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = oPlayer.ObjectID Then
                                        If CType(oColony.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet AndAlso CType(oColony.ParentObject, Epica_GUID).ObjectID = oPlanet.ObjectID Then
                                            'ok, valid colony
                                            oTrans.lRouteUB = 0
                                            ReDim oTrans.oRoute(oTrans.lRouteUB)
                                            .oRoute(.lRouteUB) = New TransportRoute
                                            With .oRoute(oTrans.lRouteUB)
                                                .OrderNum = 0
                                                .oTransport = oTrans
                                                .lActionUB = -1
                                                ReDim .oActions(-1)
                                                .WaypointFlags = 0
                                                .DestinationTypeID = ObjectType.eColony
                                                .DestinationID = oColony.ObjectID
                                            End With

                                            oTrans.oRoute(oTrans.lRouteUB).lActionUB += 1
                                            ReDim Preserve oTrans.oRoute(oTrans.lRouteUB).oActions(oTrans.oRoute(oTrans.lRouteUB).lActionUB)
                                            oTrans.oRoute(oTrans.lRouteUB).oActions(oTrans.oRoute(oTrans.lRouteUB).lActionUB) = New TransportRouteAction
                                            With oTrans.oRoute(oTrans.lRouteUB).oActions(oTrans.oRoute(oTrans.lRouteUB).lActionUB)
                                                .ActionOrderNum = 0
                                                .ActionTypeID = TransportRouteAction.TransportRouteActionType.eUnload Or TransportRouteAction.TransportRouteActionType.ePercentage
                                                .Extended1 = -1
                                                .Extended2 = -1
                                                .Extended3 = 100
                                                .oParentRoute = oTrans.oRoute(oTrans.lRouteUB)
                                                Exit For
                                            End With
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        .SaveObject()
                        'Clear out my temp oRoute as it's loaded down below.
                        oTrans.lRouteUB = -1
                        ReDim oTrans.oRoute(-1)
                    End If
                End With
                If oTrans.Owner Is Nothing = False Then oTrans.Owner.AddTransport(oTrans)
            End While
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "LoadTransports.tblTransport: " & ex.Message)
            Stop
        End Try
        If oData Is Nothing = False Then
            oData.Close()
        End If
        oData = Nothing
        oComm = Nothing

        'tblTransportCargo (TransportID, CargoID, CargoTypeID, Quantity, OwnerID)
        LogEvent(LogEventType.Informational, "Loading Transport Cargo...")
        Try
            sSQL = "SELECT * FROM tblTransportCargo"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                Dim lTransOwner As Int32 = CInt(oData("TransOwnerID"))
                Dim oPlayer As Player = GetEpicaPlayer(lTransOwner)
                If oPlayer Is Nothing Then Continue While
                Dim oTrans As Transport = oPlayer.GetTransport(CInt(oData("TransportID")))
                If oTrans Is Nothing Then Continue While
                oTrans.lCargoUB += 1
                ReDim Preserve oTrans.oCargo(oTrans.lCargoUB)
                oTrans.oCargo(oTrans.lCargoUB) = New TransportCargo
                With oTrans.oCargo(oTrans.lCargoUB)
                    .Quantity = CInt(oData("Quantity"))
                    .OwnerID = CInt(oData("OwnerID"))
                    .CargoID = CInt(oData("CargoID"))
                    .CargoTypeID = CShort(oData("CargoTypeID"))
                    .oParentTransport = oTrans
                End With
            End While
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "LoadTransports.tblTransportCargo: " & ex.Message)
            Stop
        End Try
        If oData Is Nothing = False Then
            oData.Close()
        End If
        oData = Nothing
        oComm = Nothing

        'tblTransportRoute (TransportID, OrderNum, DestinationID, DestinationTypeID, WaypointFlags)
        LogEvent(LogEventType.Informational, "Loading Transport Routes...")
        Try
            sSQL = "SELECT * FROM tblTransportRoute ORDER BY OrderNum"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("TransOwnerID")))
                If oPlayer Is Nothing Then Continue While
                Dim oTrans As Transport = oPlayer.GetTransport(CInt(oData("TransportID")))
                If oTrans Is Nothing Then Continue While

                oTrans.lRouteUB += 1
                ReDim Preserve oTrans.oRoute(oTrans.lRouteUB)
                oTrans.oRoute(oTrans.lRouteUB) = New TransportRoute
                With oTrans.oRoute(oTrans.lRouteUB)
                    .OrderNum = CByte(oData("OrderNum"))
                    .oTransport = oTrans
                    .lActionUB = -1
                    ReDim .oActions(-1)
                    .WaypointFlags = 0
                    .DestinationID = CInt(oData("DestinationID"))
                    .DestinationTypeID = CShort(oData("DestinationTypeID"))
                End With
            End While
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "LoadTransports.tblTransportRoute: " & ex.Message)
            Stop
        End Try
        If oData Is Nothing = False Then
            oData.Close()
        End If
        oData = Nothing
        oComm = Nothing

        'tblTransportRouteAction(TransportID, OrderNum, ActionOrderNum, ActionTypeID, Extended1, Extended2, Extended3)
        LogEvent(LogEventType.Informational, "Loading Transport Route Actions")
        Try
            sSQL = "SELECT * FROM tblTransportRouteAction ORDER BY ActionOrderNum"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                Dim oPlayer As Player = GetEpicaPlayer(CInt(oData("TransOwnerID")))
                If oPlayer Is Nothing Then Continue While
                Dim oTrans As Transport = oPlayer.GetTransport(CInt(oData("TransportID")))
                If oTrans Is Nothing Then Continue While
                Dim yRoute As Byte = CByte(oData("OrderNum"))
                If yRoute > oTrans.lRouteUB Then Continue While
                Dim oRoute As TransportRoute = oTrans.oRoute(yRoute)
                If oRoute Is Nothing Then Continue While

                oRoute.lActionUB += 1
                ReDim Preserve oRoute.oActions(oRoute.lActionUB)
                oRoute.oActions(oRoute.lActionUB) = New TransportRouteAction
                With oRoute.oActions(oRoute.lActionUB)
                    .ActionOrderNum = CByte(oData("ActionOrderNum"))
                    .ActionTypeID = CType(oData("ActionTypeID"), TransportRouteAction.TransportRouteActionType)
                    .Extended1 = CInt(oData("Extended1"))
                    .Extended2 = CShort(oData("Extended2"))
                    .Extended3 = CInt(oData("Extended3"))
                    .oParentRoute = oRoute
                End With
            End While
        Catch ex As Exception

        End Try
        If oData Is Nothing = False Then
            oData.Close()
        End If
        oData = Nothing
        oComm = Nothing

        Return True
    End Function

    Private Function LoadRouteTemplates(ByVal sPlayerStr As String) As Boolean
        Dim bResult As Boolean = False
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Try

            Dim sSQL As String
            Dim oTemplates(-1) As RouteTemplate
            Dim lTemplateUB As Int32 = -1

            If sPlayerStr Is Nothing OrElse sPlayerStr = "" Then
                sSQL = "SELECT count(*) FROM tblRouteTemplate"
            Else
                sSQL = "SELECT count(*) FROM tblRouteTemplate WHERE PlayerID IN " & sPlayerStr
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader
            If oData.Read = True Then
                lTemplateUB = CInt(oData(0))
            End If
            oData.Close()
            oData = Nothing
            oComm.Dispose()
            oComm = Nothing

            ReDim oTemplates(lTemplateUB)
            lTemplateUB = -1

            sSQL = sSQL.Replace("count(*)", "*")
            sSQL &= " ORDER BY TemplateName"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                'Load the template here into the oTemplates array
                lTemplateUB += 1
                If lTemplateUB > oTemplates.GetUpperBound(0) Then ReDim Preserve oTemplates(lTemplateUB)
                oTemplates(lTemplateUB) = New RouteTemplate
                With oTemplates(lTemplateUB)
                    .lItemUB = -1
                    .lPlayerID = CInt(oData("PlayerID"))
                    .lTemplateID = CInt(oData("TemplateID"))
                    ReDim .TemplateName(19)
                    StringToBytes(CStr(oData("TemplateName"))).CopyTo(.TemplateName, 0)
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm.Dispose()
            oComm = Nothing

            sSQL = "SELECT * FROM tblRouteTemplateItem ORDER BY TemplateID, OrderNum"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader
            While oData.Read = True
                Dim lTemplateID As Int32 = CInt(oData("TemplateID"))
                For X As Int32 = 0 To lTemplateUB
                    If oTemplates(X).lTemplateID = lTemplateID Then

                        oTemplates(X).lItemUB += 1
                        ReDim Preserve oTemplates(X).uItems(oTemplates(X).lItemUB)
                        With oTemplates(X).uItems(oTemplates(X).lItemUB)
                            .lLocX = CInt(oData("LocX"))
                            .lLocZ = CInt(oData("LocZ"))
                            .lOrderNum = CInt(oData("OrderNum"))

                            .oDest = CType(GetEpicaObject(CInt(oData("DestID")), CShort(oData("DestTypeID"))), Epica_GUID)
                            Dim lTmpVal As Int32 = CInt(oData("LoadItemID"))
                            Dim iTmpVal As Int16 = CShort(oData("LoadItemTypeID"))
                            .oLoadItem = Nothing

                            Dim yFlag As Byte = CByte(oData("ExtraFlags"))
                            .SetLoadItem(lTmpVal, iTmpVal, GetEpicaPlayer(oTemplates(X).lPlayerID), yFlag)
                        End With

                        Exit For
                    End If
                Next X
            End While
            oData.Close()
            oData = Nothing
            oComm.Dispose()
            oComm = Nothing

            'Now, go through our oTemplates array and assign to the players
            For X As Int32 = 0 To lTemplateUB
                If oTemplates(X) Is Nothing = False Then
                    Dim oPlayer As Player = GetEpicaPlayer(oTemplates(X).lPlayerID)
                    If oPlayer Is Nothing = False Then
                        oPlayer.lRouteTemplateUB += 1
                        ReDim Preserve oPlayer.oRouteTemplates(oPlayer.lRouteTemplateUB)
                        oPlayer.oRouteTemplates(oPlayer.lRouteTemplateUB) = oTemplates(X)
                    End If
                End If
            Next X

            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Unable to LoadRouteTemplates: " & ex.Message)
        End Try
        If oData Is Nothing = False Then oData.Close()
        oData = Nothing
        If oComm Is Nothing = False Then oComm.Dispose()
        oComm = Nothing

        Return bResult

    End Function

    Private Function LoadAgentDetails(ByVal sPlayerInStr As String) As Boolean
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim oComm As OleDb.OleDbCommand
        Dim sSQL As String
        Dim bResult As Boolean = False

        Try
            'Agents
            LogEvent(LogEventType.Informational, "Loading Agents...")
            sSQL = "SELECT * FROM tblAgent WHERE OwnerID IN " & sPlayerInStr & " OR TargetID IN " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                glAgentUB = glAgentUB + 1
                ReDim Preserve goAgent(glAgentUB)
                ReDim Preserve glAgentIdx(glAgentUB)
                goAgent(glAgentUB) = New Agent()
                With goAgent(glAgentUB)
                    .LoadFromDataReader(oData, glAgentUB)
                    glAgentIdx(glAgentUB) = .ObjectID
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Agent Skills
            LogEvent(LogEventType.Informational, "Loading Agent Skills...")
            sSQL = "SELECT * FROM tblAgentSkill"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oAgent As Agent = GetEpicaAgent(CInt(oData("AgentID")))
                If oAgent Is Nothing = False Then
                    oAgent.AddSkill(GetEpicaSkill(CInt(oData("SkillID"))), CByte(oData("SkillValue")))
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Update agent engine
            If goAgentEngine Is Nothing Then goAgentEngine = New AgentEngine()
            LogEvent(LogEventType.Informational, "Registering Agents in Agent Engine")
            For X As Int32 = 0 To glAgentUB
                If glAgentIdx(X) <> -1 Then
                    If (goAgent(X).lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
                        'ok, the agent has been captured, all other status is unused
                        If goAgent(X).lCapturedBy > -1 Then
                            Dim oPlayer As Player = GetEpicaPlayer(goAgent(X).lCapturedBy)
                            If oPlayer Is Nothing = False Then
                                oPlayer.oSecurity.AddCapturedAgent(goAgent(X))
                                Continue For
                            End If
                        End If
                        goAgent(X).lAgentStatus = goAgent(X).lAgentStatus Xor AgentStatus.HasBeenCaptured
                    End If

                    'remove Returning Home (assumed to be done between restarts
                    If (goAgent(X).lAgentStatus And AgentStatus.ReturningHome) <> 0 Then
                        goAgent(X).lAgentStatus = goAgent(X).lAgentStatus Xor AgentStatus.ReturningHome
                    End If
                    'if dismissed, add the event
                    If (goAgent(X).lAgentStatus And AgentStatus.Dismissed) <> 0 Then
                        goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentDismissed, goAgent(X), Nothing, glCurrentCycle + 864000 + CInt(Rnd() * 1000))
                    End If
                    'if the agent is infiltrating but is not yet infiltrated
                    If (goAgent(X).lAgentStatus And AgentStatus.Infiltrating) <> 0 AndAlso (goAgent(X).lAgentStatus And AgentStatus.IsInfiltrated) = 0 Then
                        goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentFirstInfiltrate, goAgent(X), Nothing, glCurrentCycle + CInt(Rnd() * 1000))
                    End If
                    'if the agent is infiltrated 
                    If (goAgent(X).lAgentStatus And AgentStatus.IsInfiltrated) <> 0 Then
                        'if it is not at inflevel of 200
                        If goAgent(X).InfiltrationLevel < 200 Then goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentReInfiltrate, goAgent(X), Nothing, glCurrentCycle + Agent.ml_DEINFILTRATE_TIME + CInt(Rnd() * 1000))
                        'add its check in interval
                        goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentCheckIn, goAgent(X), Nothing, glCurrentCycle + Agent.GetCycleDelayByFreq(goAgent(X).yReportFreq))
                    End If
                    'is the agent a new recruit
                    If (goAgent(X).lAgentStatus And AgentStatus.NewRecruit) <> 0 Then
                        goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.RecruitDismiss, goAgent(X), Nothing, glCurrentCycle + 5184000 + CInt(Rnd() * 1000))
                    End If
                End If
            Next X

            'Now, missions currently underway
            LogEvent(LogEventType.Informational, "Loading Player Missions...")
            sSQL = "SELECT * FROM tblPlayerMission WHERE PlayerID IN " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oMission As Mission = GetEpicaMission(CInt(oData("MissionID")))
                Dim oOwner As Player = GetEpicaPlayer(CInt(oData("PlayerID")))
                glPlayerMissionUB += 1
                ReDim Preserve goPlayerMission(glPlayerMissionUB)
                ReDim Preserve glPlayerMissionIdx(glPlayerMissionUB)
                goPlayerMission(glPlayerMissionUB) = New PlayerMission(oMission, oOwner)
                With goPlayerMission(glPlayerMissionUB)
                    .bAlarmThrown = CInt(oData("yAlarmThrown")) <> 0
                    .CasualtyCnt = CInt(oData("CasualtyCnt"))
                    .iTargetTypeID = CShort(oData("TargetTypeID"))
                    .lCurrentPhase = CType(oData("CurrentPhase"), eMissionPhase)
                    .lTargetID = CInt(oData("TargetID"))
                    .lTargetID2 = CInt(oData("TargetID2"))
                    .iTargetTypeID2 = CShort(oData("TargetTypeID2"))
                    .Modifier = CInt(oData("Modifier"))
                    .oTarget = GetEpicaPlayer(CInt(oData("TargetPlayerID")))
                    .PM_ID = CInt(oData("PM_ID"))
                    .ySafeHouseSetting = CByte(oData("SafeHouseSetting"))
                    .lMethodID = CInt(oData("MethodID"))
                    .yArchived = CByte(oData("bArchived"))
                    glPlayerMissionIdx(glPlayerMissionUB) = .PM_ID
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Player Mission Goals
            LogEvent(LogEventType.Informational, "Loading Player Mission Goals...")
            sSQL = "SELECT * FROM tblPlayerMissionGoal"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oPM As PlayerMission = GetEpicaPlayerMission(CInt(oData("PM_ID")))
                If oPM Is Nothing = False Then
                    If oPM.oMissionGoals Is Nothing = False Then
                        Dim lGoalID As Int32 = CInt(oData("GoalID"))
                        Dim lSkillsetID As Int32 = CInt(oData("SkillsetID"))
                        Dim oAgent As Agent = GetEpicaAgent(CInt(oData("AgentID")))
                        Dim oSkill As Skill = GetEpicaSkill(CInt(oData("SkillID")))
                        If oAgent Is Nothing OrElse oSkill Is Nothing Then Continue While

                        If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
                            oPM.oSafeHouseGoal = New PlayerMissionGoal()
                            oPM.oSafeHouseGoal.oGoal = Goal.GetSafehouseGoal()
                            oPM.oSafeHouseGoal.oMission = oPM
                            oPM.oSafeHouseGoal.oSkillSet = oPM.oSafeHouseGoal.oGoal.GetOrAddSkillSet(lSkillsetID, False)
                            Dim oAA As AgentAssignment = oPM.oSafeHouseGoal.AddAgentAssignment(oAgent, oSkill)
                            If oAA Is Nothing = False Then
                                oAA.PointsAccumulated = CShort(oData("PointsAccumulated"))
                                oAA.yStatus = CByte(oData("AAStatus"))
                                oAA.oCoveringAgent = GetEpicaAgent(CInt(oData("CoveringAgentID")))
                            End If
                        Else
                            For X As Int32 = 0 To oPM.oMissionGoals.GetUpperBound(0)
                                If oPM.oMissionGoals(X) Is Nothing = False AndAlso oPM.oMissionGoals(X).oGoal.ObjectID = lGoalID AndAlso oPM.oMission.MethodIDs(X) = oPM.lMethodID Then
                                    If oPM.oMissionGoals(X).oSkillSet Is Nothing OrElse oPM.oMissionGoals(X).oSkillSet.SkillSetID <> lSkillsetID Then
                                        For Y As Int32 = 0 To oPM.oMission.Goals(X).SkillSetUB
                                            If oPM.oMission.Goals(X).SkillSets(Y).SkillSetID = lSkillsetID Then
                                                oPM.oMissionGoals(X).oSkillSet = oPM.oMission.Goals(X).SkillSets(Y)
                                                Exit For
                                            End If
                                        Next Y
                                    End If

                                    Dim oAA As AgentAssignment = oPM.oMissionGoals(X).AddAgentAssignment(oAgent, oSkill)
                                    If oAA Is Nothing = False Then
                                        oAA.PointsAccumulated = CShort(oData("PointsAccumulated"))
                                        oAA.yStatus = CByte(oData("AAStatus"))
                                        oAA.oCoveringAgent = GetEpicaAgent(CInt(oData("CoveringAgentID")))
                                    End If
                                    Exit For
                                End If
                            Next X
                        End If

                    End If
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'Player Mission Phase Cover Agents
            LogEvent(LogEventType.Informational, "Loading Player Mission Phase Cover Agents...")
            sSQL = "SELECT * FROM tblPMPhaseCoverAgent"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oPM As PlayerMission = GetEpicaPlayerMission(CInt(oData("PM_ID")))
                If oPM Is Nothing = False Then
                    Dim lPhaseID As Int32 = CInt(oData("PhaseID"))
                    Dim oAgent As Agent = GetEpicaAgent(CInt(oData("CoverAgentID")))
                    Dim yUsedAsCoverAgent As Byte = CByte(oData("UsedAsCoverAgent"))
                    oPM.AddPhaseCoverAgent(lPhaseID, oAgent, yUsedAsCoverAgent)
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            LogEvent(LogEventType.Informational, "Associating Agents to Missions...")
            sSQL = "SELECT AgentID, MissionID FROM tblAgent WHERE OwnerID IN " & sPlayerInStr & " OR TargetID IN " & sPlayerInStr
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oMission As PlayerMission = GetEpicaPlayerMission(CInt(oData("MissionID")))
                If oMission Is Nothing = False Then
                    Dim oAgent As Agent = GetEpicaAgent(CInt(oData("AgentID")))
                    If oAgent Is Nothing = False Then oAgent.oMission = oMission
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            LogEvent(LogEventType.Informational, "Registering Missions with Agent Engine...")
            For X As Int32 = 0 To glPlayerMissionUB
                If glPlayerMissionIdx(X) <> -1 Then
                    If goPlayerMission(X).lCurrentPhase >= eMissionPhase.eInPlanning AndAlso goPlayerMission(X).lCurrentPhase <= eMissionPhase.eReinfiltrationPhase AndAlso (goPlayerMission(X).lCurrentPhase And (eMissionPhase.eMissionPaused)) = 0 Then ' Or eMissionPhase.eCancelled)) = 0 Then
                        goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.ProcessMission, Nothing, goPlayerMission(X), glCurrentCycle + CInt(Rnd() * 1000))
                    End If
                End If
            Next X

            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "LoadAgentDetails: " & ex.Message)
        Finally
            If oData Is Nothing = False Then
                oData.Close()
            End If
            oData = Nothing
            oComm = Nothing
        End Try

        Return bResult
    End Function

    Public Function LoadMineralBuyOrders(ByVal sSystemList As String) As Boolean

        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim bResult As Boolean = False

        If sSystemList Is Nothing OrElse sSystemList = "" Then
            For X As Int32 = 0 To glSystemUB
                If glSystemIdx(X) <> -1 Then
                    If goSystem(X).InMyDomain = True Then
                        If sSystemList <> "" Then sSystemList &= ", "
                        sSystemList &= goSystem(X).ObjectID
                    End If
                End If
            Next X
        End If

        Try

            Dim lCnt As Int32 = 0
            Dim sSQL As String = "SELECT count(*) FROM tblMineBuyOrder WHERE CacheID IN (SELECT CacheID FROM tblMineralCache WHERE (ParentTypeID = " & CByte(ObjectType.ePlanet)
            sSQL &= " AND ParentID IN (SELECT PlanetID FROM tblPlanet WHERE ParentID IN (" & sSystemList & "))) OR (ParentTypeID = " & CByte(ObjectType.eSolarSystem)
            sSQL &= " AND ParentID IN (" & sSystemList & ")))"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            If oData.Read = True Then
                lCnt = CInt(oData(0))
            End If
            oData.Close()
            oData = Nothing
            oComm = Nothing

            sSQL = "SELECT * FROM tblMineBuyOrder WHERE CacheID IN (SELECT CacheID FROM tblMineralCache WHERE (ParentTypeID = " & CByte(ObjectType.ePlanet)
            sSQL &= " AND ParentID IN (SELECT PlanetID FROM tblPlanet WHERE ParentID IN (" & sSystemList & "))) OR (ParentTypeID = " & CByte(ObjectType.eSolarSystem)
            sSQL &= " AND ParentID IN (" & sSystemList & ")))"

            LogEvent(LogEventType.Informational, "Loading Mineral Buy Orders...")

            Dim oBOs(-1) As MineBuyOrderManager
            Dim lBOIdx(-1) As Int32
            Dim lIdx As Int32 = -1
            Dim lBOUB As Int32 = lCnt - 1
            ReDim oBOs(lBOUB)
            ReDim lBOIdx(lBOUB)

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                lIdx += 1
                If lIdx > lBOUB Then
                    lBOUB += 100
                    ReDim Preserve oBOs(lBOUB)
                    ReDim Preserve lBOIdx(lBOUB)
                End If
                oBOs(lIdx) = New MineBuyOrderManager
                With oBOs(lIdx)
                    Dim oFac As Facility = GetEpicaFacility(CInt(oData("StructureID")))
                    If oFac Is Nothing Then
                        lIdx -= 1
                        Continue While
                    End If
                    Dim oCache As MineralCache = GetEpicaMineralCache(CInt(oData("CacheID")))
                    If oCache Is Nothing Then
                        lIdx -= 1
                        Continue While
                    End If

                    .oParentFacility = oFac
                    .oMineralCache = oCache
                    .lMaxDaysSold = CInt(oData("MaxDaysSold"))
                    .lCurrentConseqDays = CInt(oData("CurrentConseqDays"))
                    .lDaysNotSold = CInt(oData("DaysNotSold"))
                    .bSomethingSold = CByte(oData("SomethingSold")) <> 0

                    .oParentFacility.oMiningBid = oBOs(lIdx)
                    lBOIdx(lIdx) = oFac.ObjectID
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            sSQL = "SELECT * FROM tblMineBuyOrderBid WHERE CacheID IN (SELECT CacheID FROM tblMineralCache WHERE (ParentTypeID = " & CByte(ObjectType.ePlanet)
            sSQL &= " AND ParentID IN (SELECT PlanetID FROM tblPlanet WHERE ParentID IN (" & sSystemList & "))) OR (ParentTypeID = " & CByte(ObjectType.eSolarSystem)
            sSQL &= " AND ParentID IN (" & sSystemList & "))) ORDER BY OrderNum"

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True
                Dim lFacID As Int32 = CInt(oData("StructureID"))
                Dim lCacheID As Int32 = CInt(oData("CacheID"))

                For X As Int32 = 0 To lBOUB
                    If lBOIdx(X) = lFacID Then
                        If oBOs(X).oParentFacility Is Nothing = False AndAlso oBOs(X).oMineralCache Is Nothing = False AndAlso oBOs(X).oParentFacility.ObjectID = lFacID AndAlso oBOs(X).oMineralCache.ObjectID = lCacheID Then
                            oBOs(X).AddBid(CInt(oData("BidAmount")), CInt(oData("MaxQuantity")), GetEpicaPlayer(CInt(oData("PlayerID"))), False)
                            Exit For
                        End If
                    End If
                Next X
            End While

            bResult = True
        Catch
            LogEvent(LogEventType.CriticalError, "LoadMineralBuyOrders Error: " & Err.Description)
        Finally
            If oData Is Nothing = False Then
                oData.Close()
            End If
            oData = Nothing
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Private Function LoadGuildBillboards(ByVal sSystemList As String) As Boolean

        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim bResult As Boolean = False

        If sSystemList Is Nothing OrElse sSystemList = "" Then
            For X As Int32 = 0 To glSystemUB
                If glSystemIdx(X) <> -1 Then
                    If goSystem(X).InMyDomain = True Then
                        If sSystemList <> "" Then sSystemList &= ", "
                        sSystemList &= goSystem(X).ObjectID
                    End If
                End If
            Next X
        End If

        Try
            'tblGuildBillboard (GuildID, PlanetID, SlotID, BidAmount, Duration)
            LogEvent(LogEventType.Informational, "Loading Guild Billboards...")
            Dim sSQL As String = "SELECT * FROM tblGuildBillboard WHERE PlanetID IN (Select PlanetID from tblPlanet WHERE parentID IN (" & sSystemList & "))"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()

            While oData.Read = True
                Dim lPlanetID As Int32 = CInt(oData("PlanetID"))
                Dim lGuildID As Int32 = CInt(oData("GuildID"))
                Dim lSlotID As Int32 = CInt(oData("SlotID"))
                Dim lBid As Int32 = CInt(oData("BidAmount"))
                Dim lDur As Int32 = CInt(oData("Duration"))

                Dim oPlanet As Planet = GetEpicaPlanet(lPlanetID)
                Dim oGuild As Guild = GetEpicaGuild(lGuildID)
                If oPlanet Is Nothing = False AndAlso oGuild Is Nothing = False Then
                    oPlanet.AddBid(lBid, lDur, oGuild, lSlotID, False)
                End If
            End While

            bResult = True
        Catch
            LogEvent(LogEventType.CriticalError, "LoadMineralBuyOrders Error: " & Err.Description)
        Finally
            If oData Is Nothing = False Then
                oData.Close()
            End If
            oData = Nothing
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Function LoadEntityChangingEnvironment(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16) As Epica_Entity
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim oComm As OleDb.OleDbCommand
        Dim sSQL As String
        Dim oEntity As Epica_Entity = Nothing

        Try
            '  Load the entity from the DB
            If iEntityTypeID = ObjectType.eUnit Then
                Dim oUnit As Unit = Nothing
                sSQL = "SELECT * FROM tblUnit WHERE UnitID = " & lEntityID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read = True Then
                    oUnit = New Unit()
                    oUnit.LoadFromDataReader(oData)
                End If
                oData.Close()
                oData = Nothing
                oComm = Nothing
                If oUnit Is Nothing = False Then
                    '  load the route for the entity (if unit)
                    sSQL = "SELECT * FROM tblRouteItem WHERE UnitID = " & lEntityID & " ORDER BY OrderNum"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader(CommandBehavior.Default)
                    While oData.Read = True
                        Dim uNewRouteItem As RouteItem
                        With uNewRouteItem
                            .lLocX = CInt(oData("LocX"))
                            .lLocZ = CInt(oData("LocZ"))
                            .lOrderNum = CInt(oData("OrderNum"))
                            .oDest = CType(GetEpicaObject(CInt(oData("DestID")), CShort(oData("DestTypeID"))), Epica_GUID)
                            Dim lTmpVal As Int32 = CInt(oData("LoadItemID"))
                            Dim iTmpVal As Int16 = CShort(oData("LoadItemTypeID"))
                            .oLoadItem = Nothing

                            Dim yFlag As Byte = CByte(oData("ExtraFlags"))
                            .SetLoadItem(lTmpVal, iTmpVal, oUnit.Owner, yFlag)
                        End With
                        oUnit.AddRouteItem(uNewRouteItem)
                    End While
                    oData.Close()
                    oData = Nothing
                    oComm = Nothing
                End If

                'Now, add the unit to the global array
                AddUnitToGlobalArray(oUnit)

                'and then put the netity as the unit
                oEntity = oUnit
            ElseIf iEntityTypeID = ObjectType.eFacility Then
                Dim oFac As Facility = Nothing
                sSQL = "SELECT * FROM tblStructure WHERE StructureID = " & lEntityID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read = True Then
                    oFac = New Facility()
                    oFac.LoadFromDataReader(oData)
                End If
                oData.Close()
                oData = Nothing
                oComm = Nothing

                'Now, add the facility to the global array
                AddFacilityToGlobalArray(oFac)

                'and then put the entity as the facility
                oEntity = oFac
            Else
                Throw New InvalidCastException("Unable to cast objecttypeid " & iEntityTypeID & " as unit or facility! ObjectID : " & lEntityID)
            End If

            If oEntity Is Nothing Then Throw New NullReferenceException("Failed to load the object! ObjectID: " & lEntityID & ", Type: " & iEntityTypeID)

            '  load mineral cache cargo
            sSQL = "SELECT * FROM tblMineralCache WHERE ParentID = " & lEntityID & " AND ParentTypeID = " & iEntityTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True
                'load the mineral
                'add it to the global array
                'add to the cargo of the entity
                Dim oCache As New MineralCache()
                With oCache
                    .CacheTypeID = CByte(oData("CacheTypeID"))
                    .Concentration = CInt(oData("Concentration"))
                    .LocX = CInt(oData("LocX"))
                    .LocZ = CInt(oData("LocY"))
                    .ObjectID = CInt(oData("CacheID"))
                    .ObjTypeID = ObjectType.eMineralCache
                    .oMineral = GetEpicaMineral(CInt(oData("MineralID")))
                    .ParentObject = oEntity 'GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
                    .OriginalConcentration = CInt(oData("OriginalConcentration"))
                    .Quantity = CInt(oData("Quantity"))
                    .InMyDomain = True

                    .lServerIndex = AddMineralCacheToGlobalArray(oCache)

                    'Ensure we register with a containing object's cargo...
                    If .ParentObject Is Nothing = False Then
                        If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility OrElse CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eUnit Then
                            CType(.ParentObject, Epica_Entity).AddCargoRef(CType(oCache, Epica_GUID))
                        ElseIf CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eColony Then
                            Dim oColony As Colony = CType(.ParentObject, Colony)
                            oColony.mlMineralCacheUB += 1
                            ReDim Preserve oColony.mlMineralCacheID(oColony.mlMineralCacheUB)
                            ReDim Preserve oColony.mlMineralCacheIdx(oColony.mlMineralCacheUB)
                            ReDim Preserve oColony.mlMineralCacheMineralID(oColony.mlMineralCacheUB)
                            oColony.mlMineralCacheID(oColony.mlMineralCacheUB) = .ObjectID
                            oColony.mlMineralCacheIdx(oColony.mlMineralCacheUB) = .lServerIndex
                            oColony.mlMineralCacheMineralID(oColony.mlMineralCacheUB) = .oMineral.ObjectID
                        End If
                    End If

                    .bNeedsAsync = False
                End With

            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            '  load component cache cargo
            sSQL = "SELECT * FROM tblComponentCache WHERE ParentID = " & lEntityID & " AND ParentTypeID = " & iEntityTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True
                'load the component cache
                Dim oCache As New ComponentCache()
                With oCache
                    .ComponentID = CInt(oData("ComponentID"))
                    .ComponentOwnerID = CInt(oData("ComponentOwnerID"))
                    .ComponentTypeID = CShort(oData("ComponentTypeID"))
                    .InMyDomain = True
                    .LocX = CInt(oData("LocX"))
                    .LocZ = CInt(oData("LocZ"))
                    .ObjectID = CInt(oData("ComponentCacheID"))
                    .ObjTypeID = ObjectType.eComponentCache
                    .ParentObject = oEntity 'GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
                    .Quantity = CInt(oData("Quantity"))
                    .yCacheTypeID = 0

                    .bNeedsSaved = False

                    'TODO: ensure the component cache's component is loaded

                    'add it to the global array
                    Dim lIdx As Int32 = AddComponentCacheToGlobalArray(oCache)
                    'add to the cargo of the entity
                    If .ParentObject Is Nothing = False Then
                        If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility OrElse CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eUnit Then
                            CType(.ParentObject, Epica_Entity).AddCargoRef(CType(oCache, Epica_GUID))
                        ElseIf CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eColony Then
                            Dim oColony As Colony = CType(.ParentObject, Colony)
                            oColony.mlComponentCacheUB += 1
                            ReDim Preserve oColony.mlComponentCacheID(oColony.mlComponentCacheUB)
                            ReDim Preserve oColony.mlComponentCacheIdx(oColony.mlComponentCacheUB)
                            ReDim Preserve oColony.mlComponentCacheCompID(oColony.mlComponentCacheUB)
                            oColony.mlComponentCacheID(oColony.mlComponentCacheUB) = .ObjectID
                            oColony.mlComponentCacheIdx(oColony.mlComponentCacheUB) = lIdx
                            oColony.mlComponentCacheCompID(oColony.mlComponentCacheUB) = .ComponentID
                        End If
                    End If
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            '  load hangar contents
            sSQL = "SELECT * FROM tblUnit WHERE ParentID = " & lEntityID & " AND ParentTypeID = " & iEntityTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True
                'load the unit
                Dim oUnit As New Unit()
                oUnit.LoadFromDataReader(oData)

                'TODO: Ensure the unit's Owner is "Interested" in me...

                'add it to the global array
                Dim lIdx As Int32 = AddUnitToGlobalArray(oUnit)
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "LoadEntityChangingEnvironment Error: " & ex.Message)
        Finally
            If oData Is Nothing = False Then
                oData.Close()
            End If
            oData = Nothing
            oComm = Nothing
        End Try

        Return oEntity
    End Function

    Public Function LoadSharedPlayerData(ByVal lPlayerID As Int32) As Boolean
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim oComm As OleDb.OleDbCommand
        Dim sSQL As String
        Dim bResult As Boolean = False

        Try
            'Ok, player now has an interest in this primary... Load the following data specific to the player:
            Dim sFormationDefIDList As String = ""
            Dim oFormationDefs() As FormationDef = Nothing
            Dim lFormationDefUB As Int32 = -1
            '   Formations
            sSQL = "SELECT * FROM tblFormationDef WHERE OwnerID = " & lPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True

                lFormationDefUB += 1
                ReDim Preserve oFormationDefs(lFormationDefUB)

                If sFormationDefIDList <> "" Then sFormationDefIDList &= ", "
                Dim lFormationDefID As Int32 = CInt(oData("FormationDefID"))
                sFormationDefIDList &= lFormationDefID.ToString

                'Load the Formation def into a temporary formation defs
                oFormationDefs(lFormationDefUB) = New FormationDef()
                With oFormationDefs(lFormationDefUB)
                    .FormationID = lFormationDefID
                    .lCellSize = CInt(oData("CellSize"))
                    .lOwnerID = CInt(oData("OwnerID"))
                    .yCriteria = CType(CByte(oData("yCriteria")), CriteriaType)
                    ReDim .yName(19)
                    StringToBytes(CStr(oData("FormationName"))).CopyTo(.yName, 0)
                    .yDefault = CByte(oData("yDefault"))
                    ReDim .mlSlots(FormationDef.ml_GRID_SIZE_WH - 1, FormationDef.ml_GRID_SIZE_WH - 1)
                End With
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            '   FormationDefs
            If sFormationDefIDList <> "" Then
                sSQL = "SELECT * FROM tblFormationDefSlot WHERE FormationDefID IN (" & sFormationDefIDList & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                While oData.Read = True
                    Dim lFormationDefID As Int32 = CInt(oData("FormationDefID"))
                    For X As Int32 = 0 To lFormationDefUB
                        If oFormationDefs(X).FormationID = lFormationDefID Then
                            Dim lSlotX As Int32 = CInt(oData("SlotX"))
                            Dim lSlotY As Int32 = CInt(oData("SlotY"))
                            Dim lSlotVal As Int32 = CInt(oData("SlotValue"))
                            oFormationDefs(X).mlSlots(lSlotX, lSlotY) = lSlotVal
                        End If
                    Next X
                End While
                oData.Close()
                oData = Nothing
                oComm = Nothing

                For X As Int32 = 0 To lFormationDefUB
                    Dim oFormation As FormationDef = GetEpicaFormation(oFormationDefs(X).FormationID)
                    If oFormation Is Nothing = True Then
                        AddFormationToGlobalArray(oFormationDefs(X))
                    End If
                    Dim yForward() As Byte = oFormationDefs(X).GetAsAddMsg()

                    'ok, we need to send this to all domains and pathfinding 
                    goMsgSys.BroadcastToDomains(yForward)
                    goMsgSys.SendToPathfinding(yForward)
                Next X
            End If

            '   Techs/Weapons
            LoadTechs("OwnerID = " & lPlayerID, False)

            '   Entity Defs
            sSQL = "SELECT * FROM tblUnitDef WHERE OwnerID = " & lPlayerID
            Dim sUnitDefList As String = ""
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True
                Dim ldefID As Int32 = CInt(oData("UnitDefID"))
                Dim oDef As Epica_Entity_Def = GetEpicaUnitDef(ldefID)
                If oDef Is Nothing = True Then
                    oDef = New Epica_Entity_Def
                    With oDef
                        .FillFromDataReader(oData, ObjectType.eUnitDef)
                        Dim lIdx As Int32 = AddUnitDefToGlobalArray(oDef)
                        If sUnitDefList <> "" Then sUnitDefList &= ", "
                        sUnitDefList &= .ObjectID
                        If .oPrototype Is Nothing = False AndAlso .oPrototype.oHullTech Is Nothing = False Then
                            .oPrototype.oHullTech.GetSlotCriticalChances(oDef)
                        End If
                    End With
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            sSQL = "SELECT * FROM tblStructureDef WHERE OwnerID = " & lPlayerID
            Dim sStructureDefList As String = ""
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read = True
                Dim lDefID As Int32 = CInt(oData("FacilityDefID"))
                Dim oDef As FacilityDef = GetEpicaFacilityDef(lDefID)
                If oDef Is Nothing = True Then
                    oDef = New FacilityDef
                    With oDef
                        .FillFromDataReader(oData, ObjectType.eFacilityDef)
                        Dim lIdx As Int32 = AddFacilityDefToGlobalArray(oDef)
                        If sStructureDefList <> "" Then sStructureDefList &= ", "
                        sStructureDefList &= .ObjectID
                        If .oPrototype Is Nothing = False AndAlso .oPrototype.oHullTech Is Nothing = False Then
                            .oPrototype.oHullTech.GetSlotCriticalChances(CType(oDef, Epica_Entity_Def))
                        End If
                    End With
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            '     WeaponDefs
            'Weapon Definitions....
            'LogEvent(LogEventType.Informational, "Loading Weapon Defs...")
            sSQL = "SELECT tblWeaponDef.*, tblWeapon.OwnerID FROM tblWeaponDef LEFT OUTER JOIN tblWeapon ON tblWeaponDef.WeaponID = tblWeapon.WeaponID WHERE tblWeapon.OwnerID = " & lPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim lID As Int32 = CInt(oData("WeaponDefID"))
                Dim oWpnDef As WeaponDef = GetEpicaWeaponDef(lID)
                If oWpnDef Is Nothing = False Then
                    With oWpnDef
                        .FillFromDataReader(oData)

                        'Now, place it
                        Dim lIdx As Int32 = AddWeaponDefToGlobalArray(oWpnDef)
                    End With
                End If
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            'LogEvent(LogEventType.Informational, "Loading Unit Def Weapons...")
            If sUnitDefList <> "" Then
                sSQL = "SELECT * FROM tblUnitDefWeapon WHERE UnitDefID IN (" & sUnitDefList & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                While oData.Read
                    Dim oDef As Epica_Entity_Def = GetEpicaUnitDef(CInt(oData("UnitDefID")))
                    If oDef Is Nothing = False Then
                        Dim oTmpWeaponDef As WeaponDef = GetEpicaWeaponDef(CInt(oData("WeaponDefID")))
                        If oTmpWeaponDef Is Nothing = False Then
                            oDef.AddWeaponDef(CInt(oData("UDW_ID")), oTmpWeaponDef, CType(oData("Quadrant"), UnitArcs), CInt(oData("EntityStatusGroup")), CInt(oData("iAmmoCap")))
                        End If
                    End If
                End While
                oData.Close()
                oData = Nothing
                oComm = Nothing
            End If

            'LogEvent(LogEventType.Informational, "Loading Structure Def Weapons...")
            If sStructureDefList <> "" Then
                sSQL = "SELECT * FROM tblStructureDefWeapon WHERE StructureDefID IN (" & sStructureDefList & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                While oData.Read
                    Dim oTmpFacilityDef As FacilityDef = GetEpicaFacilityDef(CInt(oData("StructureDefID")))
                    If oTmpFacilityDef Is Nothing = False Then
                        Dim oTmpWeaponDef As WeaponDef = GetEpicaWeaponDef(CInt(oData("WeaponDefID")))
                        If oTmpWeaponDef Is Nothing = False Then
                            oTmpFacilityDef.AddWeaponDef(CInt(oData("SDW_ID")), oTmpWeaponDef, CType(oData("Quadrant"), UnitArcs), CInt(oData("EntityStatusGroup")), CInt(oData("iAmmoCap")))
                        End If
                    End If
                End While
                oData.Close()
                oData = Nothing
                oComm = Nothing
            End If

            '     HangarDoors
            'LogEvent(LogEventType.Informational, "Loading Entity Def Hangar Doors...")
            If sStructureDefList <> "" AndAlso sUnitDefList <> "" Then
                sSQL = "SELECT * FROM tblEntityDefHangarDoor WHERE "
                If sStructureDefList <> "" Then
                    sSQL &= "(EntityDefID IN (" & sStructureDefList & ") AND EntityDefTypeID = " & ObjectType.eFacilityDef & ")"
                End If
                If sUnitDefList <> "" Then
                    sSQL &= "(EntityDefID IN (" & sUnitDefList & ") AND EntityDefTypeID = " & ObjectType.eUnitDef & ")"
                End If
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                While oData.Read
                    Dim oTmpDef As Epica_Entity_Def = CType(GetEpicaObject(CInt(oData("EntityDefID")), CShort(oData("EntityDefTypeID"))), Epica_Entity_Def)
                    If oTmpDef Is Nothing = False Then
                        oTmpDef.AddHangarDoor(CInt(oData("ED_HD_ID")), CInt(oData("DoorSize")), CByte(oData("SideArc")))
                    End If
                    oTmpDef = Nothing
                End While
                oData.Close()
                oData = Nothing
                oComm = Nothing


                'LogEvent(LogEventType.Informational, "Loading Production Costs...")
                sSQL = "SELECT * FROM tblProductionCost WHERE "
                If sStructureDefList <> "" Then
                    sSQL &= "(ObjectID IN (" & sStructureDefList & ") AND ObjTypeID = " & ObjectType.eFacilityDef & ")"
                End If
                If sUnitDefList <> "" Then
                    sSQL &= "(ObjectID IN (" & sUnitDefList & ") AND ObjTypeID = " & ObjectType.eUnitDef & ")"
                End If

                Dim sPC_ID As String = ""
                Dim oPCs() As ProductionCost = Nothing
                Dim lPCUB As Int32 = -1

                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                While oData.Read
                    'Produceable items... units, facilities
                    If sPC_ID <> "" Then sPC_ID &= ", "
                    sPC_ID &= CInt(oData("PC_ID"))

                    Select Case Val(oData("ObjTypeID"))
                        Case ObjectType.eFacilityDef
                            Dim oDef As FacilityDef = GetEpicaFacilityDef(CInt(oData("ObjectID")))
                            If oDef Is Nothing = False Then
                                oDef.ProductionRequirements = New ProductionCost()
                                With oDef.ProductionRequirements
                                    .ColonistCost = CInt(oData("Colonists"))
                                    .CreditCost = CLng(oData("Credits"))
                                    .EnlistedCost = CInt(oData("Enlisted"))
                                    .ObjectID = CInt(oData("ObjectID"))
                                    .ObjTypeID = CShort(oData("ObjTypeID"))
                                    .OfficerCost = CInt(oData("Officers"))
                                    .PointsRequired = CLng(oData("PointsRequired"))
                                    .ProductionCostType = CByte(oData("ProductionCostType"))
                                    .PC_ID = CInt(oData("PC_ID"))
                                End With
                                lPCUB += 1
                                ReDim Preserve oPCs(lPCUB)
                                oPCs(lPCUB) = oDef.ProductionRequirements
                            End If
                        Case ObjectType.eUnitDef
                            Dim oDef As Epica_Entity_Def = GetEpicaUnitDef(CInt(oData("ObjectID")))
                            If oDef Is Nothing = False Then
                                oDef.ProductionRequirements = New ProductionCost()
                                With oDef.ProductionRequirements
                                    .ColonistCost = CInt(oData("Colonists"))
                                    .CreditCost = CLng(oData("Credits"))
                                    .EnlistedCost = CInt(oData("Enlisted"))
                                    .ObjectID = CInt(oData("ObjectID"))
                                    .ObjTypeID = CShort(oData("ObjTypeID"))
                                    .OfficerCost = CInt(oData("Officers"))
                                    .PointsRequired = CLng(oData("PointsRequired"))
                                    .ProductionCostType = CByte(oData("ProductionCostType"))
                                    .PC_ID = CInt(oData("PC_ID"))
                                End With
                                lPCUB += 1
                                ReDim Preserve oPCs(lPCUB)
                                oPCs(lPCUB) = oDef.ProductionRequirements
                            End If
                    End Select
                End While
                oData.Close()
                oData = Nothing
                oComm = Nothing

                'Now, go back thru and add our minerals
                'LogEvent(LogEventType.Informational, "Loading Production Item Costs...")
                If sPC_ID <> "" Then
                    sSQL = "SELECT * FROM tblProductionCostItem where PC_ID IN (" & sPC_ID & ")"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader(CommandBehavior.Default)
                    While oData.Read
                        Dim lPC_ID As Int32 = CInt(oData("PC_ID"))
                        For X As Int32 = 0 To lPCUB
                            If oPCs(X).ObjectID = lPC_ID Then
                                With oPCs(X)
                                    .ItemCostUB += 1
                                    ReDim Preserve .ItemCosts(.ItemCostUB)
                                End With
                                With oPCs(X).ItemCosts(oPCs(X).ItemCostUB)
                                    .ItemID = CInt(oData("ItemID"))
                                    .ItemTypeID = CShort(oData("ItemTypeID"))
                                    .oProdCost = oPCs(X)
                                    .PC_ID = CInt(oData("PC_ID"))
                                    .PCM_ID = CInt(oData("PCM_ID"))
                                    .QuantityNeeded = CInt(oData("Quantity"))
                                End With

                            End If
                        Next X

                    End While
                    oData.Close()
                    oData = Nothing
                    oComm = Nothing
                End If

                'Now, load our EntityDef.EntityMineralDefs :)
                'LogEvent(LogEventType.Informational, "Loading Entity Def Minerals...")
                sSQL = "SELECT * FROM tblEntityDefMineral WHERE "
                If sStructureDefList <> "" Then
                    sSQL &= "(EntityDefID IN (" & sStructureDefList & ") AND EntityDefTypeID = " & ObjectType.eFacilityDef & ")"
                End If
                If sUnitDefList <> "" Then
                    sSQL &= "(EntityDefID IN (" & sUnitDefList & ") AND EntityDefTypeID = " & ObjectType.eUnitDef & ")"
                End If
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                While oData.Read
                    Select Case Val(oData("EntityDefTypeID"))
                        Case ObjectType.eFacilityDef
                            Dim oTmpFacilityDef As FacilityDef = GetEpicaFacilityDef(CInt(oData("EntityDefID")))
                            If oTmpFacilityDef Is Nothing = False Then

                                oTmpFacilityDef.lEntityDefMineralUB += 1
                                ReDim Preserve oTmpFacilityDef.EntityDefMinerals(oTmpFacilityDef.lEntityDefMineralUB)
                                oTmpFacilityDef.EntityDefMinerals(oTmpFacilityDef.lEntityDefMineralUB) = New EntityDefMineral()
                                With oTmpFacilityDef.EntityDefMinerals(oTmpFacilityDef.lEntityDefMineralUB)
                                    .iEntityDefTypeID = ObjectType.eFacilityDef
                                    .lQuantity = CInt(oData("Quantity"))
                                    .oEntityDef = oTmpFacilityDef
                                    .oMineral = GetEpicaMineral(CInt(oData("MineralID")))
                                End With
                            End If
                        Case ObjectType.eUnitDef
                            Dim oTmpUnitDef As Epica_Entity_Def = GetEpicaUnitDef(CInt(oData("EntityDefID")))
                            If oTmpUnitDef Is Nothing = False Then

                                oTmpUnitDef.lEntityDefMineralUB += 1
                                ReDim Preserve oTmpUnitDef.EntityDefMinerals(oTmpUnitDef.lEntityDefMineralUB)
                                oTmpUnitDef.EntityDefMinerals(oTmpUnitDef.lEntityDefMineralUB) = New EntityDefMineral()
                                With oTmpUnitDef.EntityDefMinerals(oTmpUnitDef.lEntityDefMineralUB)
                                    .iEntityDefTypeID = ObjectType.eUnitDef
                                    .lQuantity = CInt(oData("Quantity"))
                                    .oEntityDef = oTmpUnitDef
                                    .oMineral = GetEpicaMineral(CInt(oData("MineralID")))
                                End With
                            End If
                        Case Else
                            LogEvent(LogEventType.Warning, "Entity Def Mineral.EntityTypeID of " & CInt(oData("EntityDefTypeID")) & " unexpected. Record skipped.")
                    End Select
                End While
                oData.Close()
                oData = Nothing
                oComm = Nothing

            End If


            '   Guild ???

            '   Player Rels
            'LogEvent(LogEventType.Informational, "Loading Player to Player Rels")
            sSQL = "SELECT * FROM tblPlayerToPlayerRel"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim oTmpRel As New PlayerRel()

                With oTmpRel
                    .oPlayerRegards = GetEpicaPlayer(CInt(oData("Player1ID")))
                    .oThisPlayer = GetEpicaPlayer(CInt(oData("Player2ID")))
                    .WithThisScore = CByte(oData("RelTypeID"))

                    .blTotalWarpointsGained = CLng(oData("TotalWPGain"))
                    .lPlayersWPV = CInt(oData("WPV"))

                    Dim lCycles As Int32 = CInt(oData("CyclesToNextUpdate"))
                    If lCycles > 0 Then .lCycleOfNextScoreUpdate = glCurrentCycle + lCycles
                    .TargetScore = CByte(oData("TargetScore"))

                    If .TargetScore >= .WithThisScore Then
                        .lCycleOfNextScoreUpdate = -1
                    ElseIf .lCycleOfNextScoreUpdate > -1 Then
                        AddToQueue(.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oTmpRel.oPlayerRegards.ObjectID, -1, oTmpRel.oThisPlayer.ObjectID, -1, -1, -1, -1, -1)
                    End If
                End With
                oTmpRel.oPlayerRegards.SetPlayerRel(oTmpRel.oThisPlayer.ObjectID, oTmpRel, False)
                oTmpRel = Nothing
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "LoadEntityChangingEnvironment Error: " & ex.Message)
        Finally
            If oData Is Nothing = False Then
                oData.Close()
            End If
            oData = Nothing
            oComm = Nothing
        End Try

        Return bResult
    End Function

	Public Function GetSkillListResponse() As Byte()
		If mySkillListResp Is Nothing Then
			Dim iCnt As Int16 = CShort(glSkillUB + 1)
			Dim X As Int32
			Dim lDestPos As Int32

			ReDim mySkillListResp((iCnt * Skill.SKILL_MSG_LEN) + 3)

			'Now, do it
			System.BitConverter.GetBytes(GlobalMessageCode.eRequestSkillList).CopyTo(mySkillListResp, 0)
			System.BitConverter.GetBytes(iCnt).CopyTo(mySkillListResp, 2)

			lDestPos = 4

			For X = 0 To glSkillUB
				goSkill(X).GetObjAsString.CopyTo(mySkillListResp, lDestPos) : lDestPos += Skill.SKILL_MSG_LEN
			Next X
		End If
		Return mySkillListResp
	End Function
 
	Public Sub BeginPrimarySave()
		'Only call this once all domain servers have turned in...
		gfrmDisplayForm.AddEventLine("Domain Servers Turned In... Beginning Save Process")
		Dim oThread As Threading.Thread = New Threading.Thread(AddressOf PrimarySaveData)
		'Dim oThread As Threading.Thread = New Threading.Thread(AddressOf NewPrimarySaveData)
		oThread.Start()
	End Sub

    'Public Sub SaveSixMinuteSnapshot(ByVal lCensus As Int32)
    '    Dim lDT As Int32 = GetDateAsNumber(Now)
    '    Dim oComm As OleDb.OleDbCommand = Nothing

    '    Dim lCnt As Int32 = 0
    '    For X As Int32 = 0 To glPlayerUB
    '        If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).oSocket Is Nothing = False Then
    '            lCnt += 1
    '            Try
    '                oComm = New OleDb.OleDbCommand("INSERT INTO tblSnapShotPlayer (DateTimeStamp, PlayerID) VALUES (" & lDT & ", " & glPlayerIdx(X) & ")", goCN)
    '                oComm.ExecuteNonQuery()
    '            Catch ex As Exception
    '                LogEvent(LogEventType.CriticalError, "Error saving Six Minute Snapshot person: " & ex.Message)
    '            Finally
    '                If oComm Is Nothing = False Then oComm.Dispose()
    '                oComm = Nothing
    '            End Try

    '        End If
    '    Next X


    '    Dim lPirateUnit As Int32 = 0
    '    Dim lPirateFac As Int32 = 0

    '    For X As Int32 = 0 To glUnitUB
    '        Try
    '            If glUnitIdx(X) <> -1 Then
    '                Dim oUnit As Unit = goUnit(X)
    '                If oUnit Is Nothing = False AndAlso oUnit.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then lPirateUnit += 1
    '            End If
    '        Catch
    '        End Try
    '    Next X
    '    For X As Int32 = 0 To glFacilityUB
    '        Try
    '            If glFacilityIdx(X) <> -1 Then
    '                Dim oFac As Facility = goFacility(X)
    '                If oFac Is Nothing = False Then
    '                    If oFac.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then lPirateFac += 1
    '                End If
    '            End If
    '        Catch
    '        End Try
    '    Next X


    '    Dim sSQL As String = "INSERT INTO tblSnapShot (DateTimeStamp, NumberPlayersOnline, GalacticCensus, PirateUnitCnt, PirateFacilityCnt, ErrorCount) VALUES (" & _
    '     lDT & ", " & lCnt & ", " & lCensus & ", " & lPirateUnit & ", " & lPirateFac & ", 0)"
    '    Try
    '        oComm = New OleDb.OleDbCommand(sSQL, goCN)
    '        oComm.ExecuteNonQuery()
    '    Catch ex As Exception
    '        LogEvent(LogEventType.CriticalError, "SaveSnapshot: " & ex.Message)
    '    Finally
    '        If oComm Is Nothing = False Then oComm.Dispose()
    '        oComm = Nothing
    '    End Try

    'End Sub

	Private Sub NewPrimarySaveData()
		'Sleep to give everyone time to ponder the absolution that is about to become the Universe... all stored to a database file that
		'  sits on some machine... makes you wonder... are we in a similar universe? Does someone's HDD fill up to save the entity that
		'  we refer to ourselves as? What happens at night? Does a backup occur where our entire universe... tiny compared to the beings
		'  that possess us... is cloned and stored for safe keeping.... does that clone represent another dimension??????
		Threading.Thread.Sleep(10)

		'Ok, here, we save all data...
		'TODO: Add to this as the game gets closer to finished!!!
		Dim X As Int32
		Dim lXStart As Int32

		gfrmDisplayForm.AddEventLine("Saving Msg Monitor Log...")
		goMsgSys.moMonitor.SaveAll()

		Dim oAllDataSaver As New AllDataSaver()
		oAllDataSaver.InitializeCommand()

		gfrmDisplayForm.AddEventLine("Saving Agents...")
		lXStart = 0
AGENT_SAVE_BEGIN:
		Try
			For X = lXStart To glAgentUB
				If glAgentIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goAgent(X).GetSaveObjectText()) = False Then
						LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData: " & goAgent(X).ObjectID)
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Agent error: " & ex.Message)
		End Try
		If X < glAgentUB Then
			lXStart = X + 1
			GoTo AGENT_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Colonies...")
		lXStart = 0
COLONY_SAVE_BEGIN:
		Try
			For X = lXStart To glColonyUB
				If glColonyIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goColony(X).GetSaveObjectText()) = False Then
						LogEvent(LogEventType.CriticalError, "Unable to save Colony in SaveAllData: " & goColony(X).ObjectID)
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Colony error: " & ex.Message)
		End Try
		If X < glColonyUB Then
			lXStart = X + 1
			GoTo COLONY_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Component Caches...")
		lXStart = 0
COMPONENT_CACHE_SAVE_BEGIN:
		Try
			For X = lXStart To glComponentCacheUB
				If glComponentCacheIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goComponentCache(X).GetSaveObjectText()) = False Then
						LogEvent(LogEventType.CriticalError, "Unable to save ComponentCache in SaveAllData: " & goComponentCache(X).ObjectID)
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Component Cache error: " & ex.Message)
		End Try
		If X < glComponentCacheUB Then
			lXStart = X + 1
			GoTo COMPONENT_CACHE_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()


		'		gfrmDisplayForm.AddEventLine("Saving Corporations...")
		'		lXStart = 0

		'CORPORATION_SAVE_BEGIN:
		'		Try
		'			For X = lXStart To glCorporationUB
		'				If glCorporationIdx(X) <> -1 Then goCorporation(X).SaveObject()
		'			Next X
		'		Catch ex As Exception
		'			gfrmDisplayForm.AddEventLine("Save Corporation error: " & ex.Message)
		'		End Try
		'		If X < glCorporationUB Then
		'			lXStart = X + 1
		'			GoTo CORPORATION_SAVE_BEGIN
		'		End If

		'        gfrmDisplayForm.AddEventLine("Saving Engine Techs...")
		'        lXStart = 0
		'ENGINE_SAVE_BEGIN:
		'        Try
		'            For X = lXStart To glEngineUB
		'                If glEngineIdx(X) <> -1 Then goEngine(X).SaveObject()
		'            Next X
		'        Catch ex As Exception
		'            gfrmDisplayForm.AddEventLine("Save Engine error: " & ex.Message)
		'        End Try
		'        If X < glEngineUB Then
		'            lXStart = X + 1
		'            GoTo ENGINE_SAVE_BEGIN
		'        End If

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Galaxies...")
		lXStart = 0
GALAXY_SAVE_BEGIN:
		Try
			For X = lXStart To glGalaxyUB
				If glGalaxyIdx(X) <> -1 AndAlso goGalaxy(X).InMyDomain = True Then
					If oAllDataSaver.AddCommandText(goGalaxy(X).GetSaveObjectText()) = False Then
						LogEvent(LogEventType.CriticalError, "Unable to save Galaxy: " & goGalaxy(X).ObjectID)
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Galaxy error: " & ex.Message)
		End Try
		If X < glGalaxyUB Then
			lXStart = X + 1
			GoTo GALAXY_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		gfrmDisplayForm.AddEventLine("Saving Guilds...")
		lXStart = 0
GUILD_SAVE_BEGIN:
		Try
			For X = lXStart To glGuildUB
				If glGuildIdx(X) <> -1 AndAlso goGuild(X) Is Nothing = False Then
					goGuild(X).SaveObject()
				End If
			Next
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Guild Error: " & ex.Message)
		End Try
		If X < glGuildUB Then
			lXStart = X + 1
			GoTo GUILD_SAVE_BEGIN
		End If 
 
		gfrmDisplayForm.AddEventLine("Saving Player Missions...")
		lXStart = 0
PLAYER_MISSION_SAVE_BEGIN:
		Try
			For X = lXStart To glPlayerMissionUB
				If glPlayerMissionIdx(X) <> -1 Then goPlayerMission(X).SaveObject()
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Player Mission error: " & ex.Message)
		End Try
		If X < glPlayerMissionUB Then
			lXStart = X + 1
			GoTo PLAYER_MISSION_SAVE_BEGIN
		End If 
 
		gfrmDisplayForm.AddEventLine("Saving Minerals...")
		lXStart = 0
MINERAL_SAVE_BEGIN:
		Try
			For X = lXStart To glMineralUB
				If glMineralIdx(X) <> -1 Then goMineral(X).SaveObject()
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Mineral error: " & ex.Message)
		End Try
		If X < glMineralUB Then
			lXStart = X + 1
			GoTo MINERAL_SAVE_BEGIN
		End If 

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Mineral Caches...")
		lXStart = 0
MINERAL_CACHE_SAVE_BEGIN:
		Try
			For X = lXStart To glMineralCacheUB
				If glMineralCacheIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goMineralCache(X).GetSaveObjectText()) = False Then
						LogEvent(LogEventType.CriticalError, "Unable to save MineralCache: " & goMineralCache(X).ObjectID)
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Mineral Cache error: " & ex.Message)
		End Try
		If X < glMineralCacheUB Then
			lXStart = X + 1
			GoTo MINERAL_CACHE_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		gfrmDisplayForm.AddEventLine("Saving Nebulae...")
		lXStart = 0
NEBULA_SAVE_BEGIN:
		Try
			For X = lXStart To glNebulaUB
				If glNebulaIdx(X) <> -1 AndAlso goNebula(X).InMyDomain = True Then goNebula(X).SaveObject()
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Nebulae error: " & ex.Message)
		End Try
		If X < glNebulaUB Then
			lXStart = X + 1
			GoTo NEBULA_SAVE_BEGIN
		End If

		gfrmDisplayForm.AddEventLine("Saving Planets...")
		lXStart = 0
PLANET_SAVE_BEGIN:
		Try
			For X = lXStart To glPlanetUB
				If glPlanetIdx(X) <> -1 AndAlso goPlanet(X).InMyDomain = True Then goPlanet(X).SaveObject()
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Planet error: " & ex.Message)
		End Try
		If X < glPlanetUB Then
			lXStart = X + 1
			GoTo PLANET_SAVE_BEGIN
		End If

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Players...")
		lXStart = 0
PLAYER_SAVE_BEGIN:
		Try
			For X = lXStart To glPlayerUB
				If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True Then 'goPlayer(X).SaveObject(False)
					If goPlayer(X).GetSaveObjectText(oAllDataSaver) = False Then
						If goPlayer(X).SaveObject(False) = False Then LogEvent(LogEventType.CriticalError, "Unable to save Player: " & goPlayer(X).ObjectID)
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Player error: " & ex.Message)
		End Try
		If X < glPlayerUB Then
			lXStart = X + 1
			GoTo PLAYER_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

        'oAllDataSaver.InitializeCommand()
        'Try
        '	If oAllDataSaver.AddCommandText(Senate.GetSaveObjectText) = False Then
        '		LogEvent(LogEventType.CriticalError, "Save Senate could not be saved!")
        '	End If
        'Catch ex As Exception
        '	LogEvent(LogEventType.CriticalError, "Save Senate: " & ex.Message)
        'End Try
        'If oAllDataSaver.ExecuteFinalCommand() = False Then
        '	LogEvent(LogEventType.CriticalError, "Unable to save Senate in SaveAllData!")
        'End If
        'oAllDataSaver.DisposeMe()

		gfrmDisplayForm.AddEventLine("Saving Solar Systems...")
		lXStart = 0
SOLAR_SYSTEM_SAVE_BEGIN:
		Try
			For X = lXStart To glSystemUB
				If glSystemIdx(X) <> -1 AndAlso goSystem(X).InMyDomain = True Then goSystem(X).SaveObject()
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save System error: " & ex.Message)
		End Try
		If X < glSystemUB Then
			lXStart = X + 1
			GoTo SOLAR_SYSTEM_SAVE_BEGIN
		End If

		gfrmDisplayForm.AddEventLine("Saving GTC...")
		lXStart = 0
TRADE_SAVE_BEGIN:
		Try
			goGTC.SaveGTC()
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save GTC error: " & ex.Message)
		End Try

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Facilities...")
		lXStart = 0
FACILITY_SAVE_BEGIN:
		Try
			For X = lXStart To glFacilityUB
				If glFacilityIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goFacility(X).GetSaveObjectText()) = False Then
						goFacility(X).SaveObject()
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Facility error: " & ex.Message)
		End Try
		If X < glFacilityUB Then
			lXStart = X + 1
			GoTo FACILITY_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Facility Defs...")
		lXStart = 0
FACILITYDEF_SAVE_BEGIN:
		Try
			For X = lXStart To glFacilityDefUB
				If glFacilityDefIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goFacilityDef(X).GetSaveObjectText()) = False Then
						goFacilityDef(X).SaveObject()
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Facility Def error: " & ex.Message)
		End Try
		If X < glFacilityDefUB Then
			lXStart = X + 1
			GoTo FACILITYDEF_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Units...")
		lXStart = 0
UNIT_SAVE_BEGIN:
		Try
			For X = lXStart To glUnitUB
				If glUnitIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goUnit(X).GetSaveObjectText) = False Then
						goUnit(X).SaveObject()
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save unit error: " & ex.Message)
		End Try
		If X < glUnitUB Then
			lXStart = X + 1
			GoTo UNIT_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving UnitDefs...")
		lXStart = 0
UNITDEF_SAVE_BEGIN:
		Try
			For X = lXStart To glUnitDefUB
				If glUnitDefIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goUnitDef(X).GetSaveObjectText()) = False Then
						goUnitDef(X).SaveObject()
					End If
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Unit Def error: " & ex.Message)
		End Try
		If X < glUnitDefUB Then
			lXStart = X + 1
			GoTo UNITDEF_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Unit Groups...")
		lXStart = 0
UNIT_GROUP_SAVE_BEGIN:
		Try
			For X = lXStart To glUnitGroupUB
				If glUnitGroupIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goUnitGroup(X).GetSaveObjectText()) = False Then goUnitGroup(X).SaveObject()
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Unit Group error: " & ex.Message)
		End Try
		If X < glUnitGroupUB Then
			lXStart = X + 1
			GoTo UNIT_GROUP_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		oAllDataSaver.InitializeCommand()
		gfrmDisplayForm.AddEventLine("Saving Weapon Defs...")
		lXStart = 0
WEAPON_DEF_SAVE_BEGIN:
		Try
			For X = lXStart To glWeaponDefUB
				If glWeaponDefIdx(X) <> -1 Then
					If oAllDataSaver.AddCommandText(goWeaponDefs(X).GetSaveObjectText()) = False Then goWeaponDefs(X).SaveObject()
				End If
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Weapon Def error: " & ex.Message)
		End Try
		If X < glWeaponDefUB Then
			lXStart = X + 1
			GoTo WEAPON_DEF_SAVE_BEGIN
		End If
		If oAllDataSaver.ExecuteFinalCommand() = False Then
			LogEvent(LogEventType.CriticalError, "Unable to save Agent in SaveAllData!")
		End If
		oAllDataSaver.DisposeMe()

		gfrmDisplayForm.AddEventLine("Saving Wormholes...")
		lXStart = 0
WORMHOLE_SAVE_BEGIN:
		Try
			For X = lXStart To glWormholeUB
				If glWormholeIdx(X) <> -1 AndAlso goWormhole(X).InMyDomain = True Then goWormhole(X).SaveObject()
			Next X
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("Save Wormhole error: " & ex.Message)
		End Try
		If X < glWormholeUB Then
			lXStart = X + 1
			GoTo WORMHOLE_SAVE_BEGIN
		End If

		'gfrmDisplayForm.AddEventLine("Flushing Save all Data... This may take a few moments...")
		'If oAllDataSaver.ExecuteFinalCommand() = False Then
		'	LogEvent(LogEventType.CriticalError, "ExecuteFinalCommand Failed. Resaving the old way.")
		'	PrimarySaveData()
		'	Return
		'End If
		'oAllDataSaver.DisposeMe()

		gfrmDisplayForm.AddEventLine("All objects saved... forcing disconnect on remaining connections.")
		goMsgSys.ForceDisconnectAll()

		gfrmDisplayForm.AddEventLine("Finalizing Globals...")
		FinalizeGlobals()

		gfrmDisplayForm.AddEventLine("Program Ready to Terminate...")

	End Sub

    Public Sub PrimarySaveData()
        'Sleep to give everyone time to ponder the absolution that is about to become the Universe... all stored to a database file that
        '  sits on some machine... makes you wonder... are we in a similar universe? Does someone's HDD fill up to save the entity that
        '  we refer to ourselves as? What happens at night? Does a backup occur where our entire universe... tiny compared to the beings
        '  that possess us... is cloned and stored for safe keeping.... does that clone represent another dimension??????
        Threading.Thread.Sleep(10)

        'Ok, here, we save all data...
        Dim X As Int32
        Dim lXStart As Int32

        'Before beginning, set all system's unit and facility counts
        For X = 0 To glSystemUB
            If goSystem(X) Is Nothing = False Then
                goSystem(X).lUnitCount = 0
                goSystem(X).lFacilityCount = 0
            End If
        Next X

        'Any factions that are invalid, remove
        For X = 0 To glPlayerUB
            If glPlayerIdx(X) > -1 Then
                Dim oPlayer As Player = goPlayer(X)
                If oPlayer Is Nothing = False Then
                    For Y As Int32 = 0 To 4
                        If oPlayer.ySlotState(Y) <> eySlotState.Accepted Then
                            Dim lOther As Int32 = oPlayer.lSlotID(Y)
                            Dim oOther As Player = GetEpicaPlayer(lOther)
                            If oOther Is Nothing = False Then
                                For Z As Int32 = 0 To 2
                                    If oOther.lFactionID(Z) = oPlayer.ObjectID Then
                                        oOther.lFactionID(Z) = -1
                                    End If
                                Next Z
                            End If
                            oPlayer.lSlotID(Y) = -1
                            oPlayer.ySlotState(Y) = 0
                        End If
                    Next 
                End If
            End If
        Next X

        'Ok, before proceeding... let's go through and store all facility production in an array...
        Dim oFacProd() As EntityProduction = Nothing
        Dim lFacProdStructID() As Int32 = Nothing
        Dim bVerified() As Boolean = Nothing
        Dim lFacProdUB As Int32 = -1

        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"
        Dim oFS As New System.IO.FileStream(sPath & "SP_" & Now.ToString("MMddyyhhmmss") & ".txt", IO.FileMode.Create)
        Dim oWriter As New System.IO.StreamWriter(oFS)

        For X = 0 To glFacilityUB
            If glFacilityIdx(X) > -1 Then
                Dim oFac As Facility = goFacility(X)
                If oFac Is Nothing = False AndAlso oFac.Active = True AndAlso oFac.CurrentProduction Is Nothing = False Then
                    lFacProdUB += 1
                    ReDim Preserve oFacProd(lFacProdUB)
                    ReDim Preserve bVerified(lFacProdUB)
                    ReDim Preserve lFacProdStructID(lFacProdUB)
                    bVerified(lFacProdUB) = False

                    Dim sLine As String = ""

                    oFacProd(lFacProdUB) = New EntityProduction
                    With oFacProd(lFacProdUB)
                        .bPaidFor = oFac.CurrentProduction.bPaidFor
                        .fAmmoSize = oFac.CurrentProduction.fAmmoSize
                        .iExtendedTypeID = oFac.CurrentProduction.iExtendedTypeID
                        .iProdA = oFac.CurrentProduction.iProdA
                        .lAmmoQty = oFac.CurrentProduction.lAmmoQty
                        .lAmmoWpnTechID = oFac.CurrentProduction.lAmmoWpnTechID
                        .lExtendedID = oFac.CurrentProduction.lExtendedID
                        .lFinishCycle = oFac.CurrentProduction.lFinishCycle
                        .lLastUpdateCycle = oFac.CurrentProduction.lLastUpdateCycle
                        .lProdCount = oFac.CurrentProduction.lProdCount
                        .lProdX = oFac.CurrentProduction.lProdX
                        .lProdZ = oFac.CurrentProduction.lProdZ
                        .oParent = oFac.CurrentProduction.oParent
                        .OrderNumber = 0
                        .PointsProduced = oFac.CurrentProduction.PointsProduced
                        .PointsRequired = oFac.CurrentProduction.PointsRequired
                        .ProdCost = oFac.CurrentProduction.ProdCost
                        .ProductionID = oFac.CurrentProduction.ProductionID
                        .ProductionTypeID = oFac.CurrentProduction.ProductionTypeID
                        .yExtendedType = oFac.CurrentProduction.yExtendedType

                        sLine = CInt(.bPaidFor).ToString & "|" & oFac.ObjectID.ToString & "|" & .OrderNumber & "|" & .PointsProduced.ToString & "|" & .ProductionID.ToString & "|" & .ProductionTypeID.ToString & "|" & .lProdCount.ToString
                        oWriter.WriteLine(sLine)
                    End With
                    lFacProdStructID(lFacProdUB) = oFac.ObjectID
                End If
            End If
        Next X
        oWriter.Close()
        oFS.Close()
        oWriter.Dispose()
        oFS.Dispose()

        Form1.bHideEventLines = False
        LogEvent(LogEventType.Informational, "Saving Msg Monitor Log...")
        Form1.bHideEventLines = True
        goMsgSys.moMonitor.SaveAll()

        Try
            'TODO: This breaks when multiple primaries are used
            Dim oComm As New OleDb.OleDbCommand("DELETE FROM tblStructureProduction", goCN)
            oComm.ExecuteNonQuery()
            oComm = Nothing
        Catch
            LogEvent(LogEventType.Informational, "Unable to clear Structure Production Table!")
        End Try

        LogEvent(LogEventType.Informational, "Saving Agents...")
        lXStart = 0
AGENT_SAVE_BEGIN:
        Try
            For X = lXStart To glAgentUB
                If glAgentIdx(X) <> -1 Then goAgent(X).SaveObject()
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Agent error: " & ex.Message)
        End Try
        If X < glAgentUB Then
            lXStart = X + 1
            GoTo AGENT_SAVE_BEGIN
        End If

        LogEvent(LogEventType.Informational, "Saving Colonies...")
        lXStart = 0
COLONY_SAVE_BEGIN:
        Try
            For X = lXStart To glColonyUB
                If glColonyIdx(X) <> -1 Then goColony(X).SaveObject()
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Colony error: " & ex.Message)
        End Try
        If X < glColonyUB Then
            lXStart = X + 1
            GoTo COLONY_SAVE_BEGIN
        End If

        LogEvent(LogEventType.Informational, "Saving Component Caches...")
        lXStart = 0

COMPONENT_CACHE_SAVE_BEGIN:
        Try
            For X = lXStart To glComponentCacheUB
                If glComponentCacheIdx(X) <> -1 Then goComponentCache(X).SaveObject()
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Component Cache error: " & ex.Message)
        End Try
        If X < glComponentCacheUB Then
            lXStart = X + 1
            GoTo COMPONENT_CACHE_SAVE_BEGIN
        End If

        'MSC - 10/02/08 - obsolete
        '        logevent(LogEventType.Informational, "Saving Corporations...")
        '        lXStart = 0

        'CORPORATION_SAVE_BEGIN:
        '        Try
        '            For X = lXStart To glCorporationUB
        '                If glCorporationIdx(X) <> -1 Then goCorporation(X).SaveObject()
        '            Next X
        '        Catch ex As Exception
        '            logevent(LogEventType.Informational, "Save Corporation error: " & ex.Message)
        '        End Try
        '        If X < glCorporationUB Then
        '            lXStart = X + 1
        '            GoTo CORPORATION_SAVE_BEGIN
        '        End If

        'MSC - 10/02/08 - never saved on primary
        '        logevent(LogEventType.Informational, "Saving Galaxies...")
        '        lXStart = 0
        'GALAXY_SAVE_BEGIN:
        '        Try
        '            For X = lXStart To glGalaxyUB
        '                If glGalaxyIdx(X) <> -1 AndAlso goGalaxy(X).InMyDomain = True Then goGalaxy(X).SaveObject()
        '            Next X
        '        Catch ex As Exception
        '            logevent(LogEventType.Informational, "Save Galaxy error: " & ex.Message)
        '        End Try
        '        If X < glGalaxyUB Then
        '            lXStart = X + 1
        '            GoTo GALAXY_SAVE_BEGIN
        '        End If

        LogEvent(LogEventType.Informational, "Saving Guilds...")
        lXStart = 0
GUILD_SAVE_BEGIN:
        Try
            For X = lXStart To glGuildUB
                If glGuildIdx(X) <> -1 AndAlso goGuild(X) Is Nothing = False Then
                    'If goGuild(X).InMyDomain = True Then
                    goGuild(X).SaveObject()
                    'End If
                End If
            Next
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Guild Error: " & ex.Message)
        End Try
        If X < glGuildUB Then
            lXStart = X + 1
            GoTo GUILD_SAVE_BEGIN
        End If

        'MSC - 10/02/08 - obsolete
        '        logevent(LogEventType.Informational, "Saving GNS...")
        '        lXStart = 0
        'GNS_SAVE_BEGIN:
        '        Try
        '            For X = lXStart To glGNSUB
        '                If glGNSIdx(X) <> -1 Then goGNS(X).SaveObject()
        '            Next X
        '        Catch ex As Exception
        '            logevent(LogEventType.Informational, "Save GNS error: " & ex.Message)
        '        End Try
        '        If X < glGNSUB Then
        '            lXStart = X + 1
        '            GoTo GNS_SAVE_BEGIN
        '        End If

        LogEvent(LogEventType.Informational, "Saving Player Missions...")
        lXStart = 0
PLAYER_MISSION_SAVE_BEGIN:
        Try
            For X = lXStart To glPlayerMissionUB
                If glPlayerMissionIdx(X) <> -1 Then goPlayerMission(X).SaveObject()
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Player Mission error: " & ex.Message)
        End Try
        If X < glPlayerMissionUB Then
            lXStart = X + 1
            GoTo PLAYER_MISSION_SAVE_BEGIN
        End If

        'MSC - 10/02/08 - saved only on creation now
        '        logevent(LogEventType.Informational, "Saving Minerals...")
        '        lXStart = 0
        'MINERAL_SAVE_BEGIN:
        '        Try
        '            For X = lXStart To glMineralUB
        '                If glMineralIdx(X) <> -1 Then goMineral(X).SaveObject()
        '            Next X
        '        Catch ex As Exception
        '            logevent(LogEventType.Informational, "Save Mineral error: " & ex.Message)
        '        End Try
        '        If X < glMineralUB Then
        '            lXStart = X + 1
        '            GoTo MINERAL_SAVE_BEGIN
        '        End If

        LogEvent(LogEventType.Informational, "Saving Mineral Caches...")
        lXStart = 0
MINERAL_CACHE_SAVE_BEGIN:
        Try
            For X = lXStart To glMineralCacheUB
                If glMineralCacheIdx(X) <> -1 Then goMineralCache(X).SaveObject()
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Mineral Cache error: " & ex.Message)
        End Try
        If X < glMineralCacheUB Then
            lXStart = X + 1
            GoTo MINERAL_CACHE_SAVE_BEGIN
        End If

        'MSC - 10/02/08 - never saved on primary ever
        '        logevent(LogEventType.Informational, "Saving Nebulae...")
        '        lXStart = 0
        'NEBULA_SAVE_BEGIN:
        '        Try
        '            For X = lXStart To glNebulaUB
        '                If glNebulaIdx(X) <> -1 AndAlso goNebula(X).InMyDomain = True Then goNebula(X).SaveObject()
        '            Next X
        '        Catch ex As Exception
        '            logevent(LogEventType.Informational, "Save Nebulae error: " & ex.Message)
        '        End Try
        '        If X < glNebulaUB Then
        '            lXStart = X + 1
        '            GoTo NEBULA_SAVE_BEGIN
        '        End If

        LogEvent(LogEventType.Informational, "Saving Planets...")
        lXStart = 0
PLANET_SAVE_BEGIN:
        Try
            For X = lXStart To glPlanetUB
                If glPlanetIdx(X) <> -1 AndAlso goPlanet(X).InMyDomain = True Then goPlanet(X).SaveObject()
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Planet error: " & ex.Message)
        End Try
        If X < glPlanetUB Then
            lXStart = X + 1
            GoTo PLANET_SAVE_BEGIN
        End If

        LogEvent(LogEventType.Informational, "Saving Players...")
        lXStart = 0
PLAYER_SAVE_BEGIN:
        Try
            For X = lXStart To glPlayerUB
                If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True Then
                    'goPlayer(X).RegenerateWPVValues()
                    goPlayer(X).SaveObject(False)
                End If
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Player error: " & ex.Message)
        End Try
        If X < glPlayerUB Then
            lXStart = X + 1
            GoTo PLAYER_SAVE_BEGIN
        End If

        'MSC - 10/02/08 - no longer a saved item
        '        logevent(LogEventType.Informational, "Saving Skills...")
        '        lXStart = 0
        'SKILL_SAVE_BEGIN:
        '        Try
        '            For X = lXStart To glSkillUB
        '                If glSkillIdx(X) <> -1 Then goSkill(X).SaveObject()
        '            Next X
        '        Catch ex As Exception
        '            logevent(LogEventType.Informational, "Save Skill error: " & ex.Message)
        '        End Try
        '        If X < glSkillUB Then
        '            lXStart = X + 1
        '            GoTo SKILL_SAVE_BEGIN
        '        End If

        LogEvent(LogEventType.Informational, "Saving Solar Systems...")
        lXStart = 0
SOLAR_SYSTEM_SAVE_BEGIN:
        Try
            For X = lXStart To glSystemUB
                If glSystemIdx(X) <> -1 AndAlso goSystem(X).InMyDomain = True Then goSystem(X).SaveObject()
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save System error: " & ex.Message)
        End Try
        If X < glSystemUB Then
            lXStart = X + 1
            GoTo SOLAR_SYSTEM_SAVE_BEGIN
        End If

        LogEvent(LogEventType.Informational, "Saving GTC...")
        lXStart = 0
TRADE_SAVE_BEGIN:
        Try
            goGTC.SaveGTC()
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save GTC error: " & ex.Message)
        End Try

        LogEvent(LogEventType.Informational, "Preparing to Save Facilities...")
        Try
            Dim sSQL As String = "DELETE FROM tblStructureProduction"
            sSQL &= vbCrLf & "DELETE FROM tblAgentEffects WHERE EffectedItemTypeID = " & CInt(ObjectType.eFacility).ToString
            sSQL &= vbCrLf & "DELETE FROM tblMineBuyOrder"
            sSQL &= vbCrLf & "DELETE FROM tblMineBuyOrderBid"

            Dim oComm As New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Preparing to Save Facilities Error: " & ex.Message)
            Stop
        End Try

        Dim oAllDataSaver As New AllDataSaver()
        oAllDataSaver.InitializeCommand()
        LogEvent(LogEventType.Informational, "Saving Facilities...")
        lXStart = 0
        Dim lCnt As Int32 = 0
        Dim bGood As Boolean = True

FACILITY_SAVE_BEGIN:
        Try
            For X = lXStart To glFacilityUB
                If glFacilityIdx(X) <> -1 Then
                    If goFacility(X).ParentObject Is Nothing = False Then
                        Dim iTypeID As Int16 = CType(goFacility(X).ParentObject, Epica_GUID).ObjTypeID
                        If iTypeID = ObjectType.eSolarSystem Then
                            CType(goFacility(X).ParentObject, SolarSystem).lFacilityCount += 1
                        ElseIf iTypeID = ObjectType.ePlanet Then
                            Dim oPlanet As Planet = CType(goFacility(X).ParentObject, Planet)
                            If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then oPlanet.ParentSystem.lFacilityCount += 1
                        End If
                    End If
                    If oAllDataSaver.AddCommandText(goFacility(X).GetSaveObjectText()) = False Then
                        goFacility(X).SaveObject()
                    Else
                        lCnt += 1
                    End If
                End If

                If lCnt > 1500 Then
                    If oAllDataSaver.ExecuteFinalCommand() = False Then
                        LogEvent(LogEventType.CriticalError, "Unable to save Facility in SaveAllData! Doing it the SLOW way...")
                        bGood = False
                        X = glFacilityUB
                        Exit For
                    Else
                        oAllDataSaver.DisposeMe()
                        System.GC.Collect()
                        oAllDataSaver.InitializeCommand()
                    End If
                    lCnt = 0
                End If
            Next X
        Catch ex_oom As OutOfMemoryException
            X -= 1001
            oAllDataSaver.DisposeMe()
            oAllDataSaver.InitializeCommand()
            LogEvent(LogEventType.Warning, "Save Facility OOM: Resetting X to " & X)
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Facility error: " & ex.Message)
        End Try
        If X < glFacilityUB Then
            lXStart = X + 1
            GoTo FACILITY_SAVE_BEGIN
        End If
        If bGood = True Then
            If oAllDataSaver.ExecuteFinalCommand() = False Then
                LogEvent(LogEventType.CriticalError, "Unable to save Facility in SaveAllData! Doing it the SLOW way...")
                bGood = False
            End If
        End If
        oAllDataSaver.DisposeMe()

        If bGood = False Then
            'need to clear out the facility counts
            For X = 0 To glSystemUB
                If goSystem(X) Is Nothing = False Then goSystem(X).lFacilityCount = 0
            Next X

            LogEvent(LogEventType.Informational, "Saving Facilities (SLOW)...")
            lXStart = 0
FACILITY_SAVE_BEGIN_SLOW:
            Try
                For X = lXStart To glFacilityUB
                    If glFacilityIdx(X) <> -1 Then
                        If goFacility(X).ParentObject Is Nothing = False Then
                            Dim iTypeID As Int16 = CType(goFacility(X).ParentObject, Epica_GUID).ObjTypeID
                            If iTypeID = ObjectType.eSolarSystem Then
                                CType(goFacility(X).ParentObject, SolarSystem).lFacilityCount += 1
                            ElseIf iTypeID = ObjectType.ePlanet Then
                                Dim oPlanet As Planet = CType(goFacility(X).ParentObject, Planet)
                                If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then oPlanet.ParentSystem.lFacilityCount += 1
                            End If
                        End If

                        goFacility(X).SaveObject()
                    End If
                Next X
            Catch ex As Exception
                LogEvent(LogEventType.Informational, "Save Facility error: " & ex.Message)
            End Try
            If X < glFacilityUB Then
                lXStart = X + 1
                GoTo FACILITY_SAVE_BEGIN_SLOW
            End If
        End If

        LogEvent(LogEventType.Informational, "Preparing to Save Units...")
        Try
            Dim sSQL As String = "DELETE FROM tblAgentEffects WHERE EffectedItemTypeID = " & CInt(ObjectType.eUnit).ToString
            sSQL &= vbCrLf & "DELETE FROM tblUnitProduction"
            sSQL &= vbCrLf & "DELETE FROM tblRouteItem"

            Dim oComm As New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Preparing to Save Units Error: " & ex.Message)
            Stop
        End Try


        LogEvent(LogEventType.Informational, "Saving Units...")
        lXStart = 0

        oAllDataSaver.InitializeCommand()
        lCnt = 0
        bGood = True

UNIT_SAVE_BEGIN:
        Try
            For X = lXStart To glUnitUB
                If glUnitIdx(X) <> -1 Then

                    If goUnit(X).ParentObject Is Nothing = False Then
                        Dim iTypeID As Int16 = CType(goUnit(X).ParentObject, Epica_GUID).ObjTypeID
                        If iTypeID = ObjectType.eSolarSystem Then
                            CType(goUnit(X).ParentObject, SolarSystem).lUnitCount += 1
                        ElseIf iTypeID = ObjectType.ePlanet Then
                            Dim oPlanet As Planet = CType(goUnit(X).ParentObject, Planet)
                            If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then oPlanet.ParentSystem.lUnitCount += 1
                        End If
                    End If

                    If oAllDataSaver.AddCommandText(goUnit(X).GetSaveObjectText()) = False Then
                        goUnit(X).SaveObject()
                    End If
                    lCnt += 1
                End If

                If lCnt > 750 Then
                    If oAllDataSaver.ExecuteFinalCommand() = False Then
                        LogEvent(LogEventType.CriticalError, "Unable to save Unit in SaveAllData! Doing it the SLOW way...")
                        bGood = False
                        X = glUnitUB
                        Exit For
                    Else
                        oAllDataSaver.DisposeMe()
                        System.GC.Collect()
                        oAllDataSaver.InitializeCommand()
                    End If
                    lCnt = 0
                End If
            Next X
        Catch ex_oom As OutOfMemoryException
            X -= 1001
            oAllDataSaver.DisposeMe()
            oAllDataSaver.InitializeCommand()
            LogEvent(LogEventType.Warning, "Save Unit OOM: Resetting X to " & X)
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Unit error: " & ex.Message)
        End Try
        If X < glUnitUB Then
            lXStart = X + 1
            GoTo UNIT_SAVE_BEGIN
        End If
        If bGood = True Then
            If oAllDataSaver.ExecuteFinalCommand() = False Then
                LogEvent(LogEventType.CriticalError, "Unable to save Unit in SaveAllData! Doing it the SLOW way...")
                bGood = False
            End If
        End If
        oAllDataSaver.DisposeMe()

        If bGood = False Then
            'need to clear out the facility counts
            For X = 0 To glSystemUB
                If goSystem(X) Is Nothing = False Then goSystem(X).lUnitCount = 0
            Next X

            LogEvent(LogEventType.Informational, "Saving Units (SLOW)...")
            lXStart = 0
UNIT_SAVE_BEGIN_SLOW:
            Try
                For X = lXStart To glUnitUB
                    If glUnitIdx(X) <> -1 Then

                        If goUnit(X).ParentObject Is Nothing = False Then
                            Dim iTypeID As Int16 = CType(goUnit(X).ParentObject, Epica_GUID).ObjTypeID
                            If iTypeID = ObjectType.eSolarSystem Then
                                CType(goUnit(X).ParentObject, SolarSystem).lUnitCount += 1
                            ElseIf iTypeID = ObjectType.ePlanet Then
                                Dim oPlanet As Planet = CType(goUnit(X).ParentObject, Planet)
                                If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then oPlanet.ParentSystem.lUnitCount += 1
                            End If
                        End If

                        goUnit(X).SaveObject()
                    End If
                Next X
            Catch ex As Exception
                LogEvent(LogEventType.Informational, "Save unit error: " & ex.Message)
            End Try
            If X < glUnitUB Then
                lXStart = X + 1
                GoTo UNIT_SAVE_BEGIN_SLOW
            End If
        End If


        LogEvent(LogEventType.Informational, "Saving Unit Groups...")
        lXStart = 0
UNIT_GROUP_SAVE_BEGIN:
        Try
            For X = lXStart To glUnitGroupUB
                If glUnitGroupIdx(X) <> -1 Then goUnitGroup(X).SaveObject()
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Unit Group error: " & ex.Message)
        End Try
        If X < glUnitGroupUB Then
            lXStart = X + 1
            GoTo UNIT_GROUP_SAVE_BEGIN
        End If

        '        logevent(LogEventType.Informational, "Saving Weapon Defs...")
        '        lXStart = 0
        'WEAPON_DEF_SAVE_BEGIN:
        '        Try
        '            For X = lXStart To glWeaponDefUB
        '                If glWeaponDefIdx(X) <> -1 Then goWeaponDefs(X).SaveObject()
        '            Next X
        '        Catch ex As Exception
        '            logevent(LogEventType.Informational, "Save Weapon Def error: " & ex.Message)
        '        End Try
        '        If X < glWeaponDefUB Then
        '            lXStart = X + 1
        '            GoTo WEAPON_DEF_SAVE_BEGIN
        '        End If

        LogEvent(LogEventType.Informational, "Saving Wormholes...")
        lXStart = 0
WORMHOLE_SAVE_BEGIN:
        Try
            For X = lXStart To glWormholeUB
                If glWormholeIdx(X) <> -1 AndAlso goWormhole(X).InMyDomain = True Then goWormhole(X).SaveObject()
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Informational, "Save Wormhole error: " & ex.Message)
        End Try
        If X < glWormholeUB Then
            lXStart = X + 1
            GoTo WORMHOLE_SAVE_BEGIN
        End If

        LogEvent(LogEventType.Informational, "All objects saved... forcing disconnect on remaining connections.")
        goMsgSys.ForceDisconnectAll()

        LogEvent(LogEventType.Informational, "Verifying the Structure Production Items...")
        Try
            Dim sSQL As String = "select * from tblstructureproduction spm where ordernum = (select top 1 ordernum from tblstructureproduction sp where spm.structureid = sp.structureid order by ordernum)"
            Dim oComm As OleDb.OleDbCommand = Nothing
            Dim oData As OleDb.OleDbDataReader = Nothing
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            While oData.Read = True
                Try
                    Dim lStructID As Int32 = CInt(oData("StructureID"))
                    Dim yOrderNum As Byte = CByte(oData("OrderNum"))
                    Dim blPoints As Int64 = CLng(oData("PointsProduced"))
                    Dim lObjectID As Int32 = CInt(oData("ObjectID"))
                    Dim iObjTypeID As Int16 = CShort(oData("ObjTypeID"))
                    Dim lProdCount As Int32 = CInt(oData("ProdCount"))

                    For X = 0 To lFacProdUB
                        If lFacProdStructID(X) = lStructID Then
                            If bVerified(X) = False Then
                                If oFacProd(X).ProductionID = lObjectID AndAlso oFacProd(X).ProductionTypeID = iObjTypeID AndAlso blPoints >= oFacProd(X).PointsProduced Then
                                    bVerified(X) = True
                                Else
                                    'Stop
                                    Dim oTmpComm As OleDb.OleDbCommand = New OleDb.OleDbCommand("UPDATE tblstructureproduction SET PointsProduced = " & oFacProd(X).PointsProduced.ToString & " WHERE StructureID = " & lStructID & " and OrderNum = " & yOrderNum & " AND ObjectID = " & lObjectID.ToString & " AND ObjTypeID = " & iObjTypeID, goCN)
                                    oTmpComm.ExecuteNonQuery()
                                    oTmpComm.Dispose()
                                End If
                            End If
                            Exit For
                        End If
                    Next X
                Catch ex As Exception
                    Stop
                End Try
            End While
            oData.Close()
            oData = Nothing
            oComm.Dispose()
            oComm = Nothing
        Catch ex As Exception
            Stop
        End Try

        'Now, save the unit/facility counts of the systems
        Try
            Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand("DELETE FROM tblSystemEntityCount", goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing
        Catch
        End Try
        Try
            For X = 0 To glSystemUB
                If goSystem(X) Is Nothing = False Then

                    Dim oComm As OleDb.OleDbCommand = Nothing
                    Try
                        oComm = New OleDb.OleDbCommand("INSERT INTO tblSystemEntityCount (SystemID, UnitCount, FacilityCount) VALUES (" & goSystem(X).ObjectID.ToString & ", " & goSystem(X).lUnitCount.ToString & ", " & goSystem(X).lFacilityCount.ToString & ")", goCN)
                        oComm.ExecuteNonQuery()
                    Catch
                    End Try
                    If oComm Is Nothing = False Then oComm.Dispose()
                    oComm = Nothing
                End If
            Next X
        Catch
        End Try

        For X = 0 To lFacProdUB
            If bVerified(X) = False Then
                'Stop
                With oFacProd(X)
                    Dim oTmpComm As OleDb.OleDbCommand = New OleDb.OleDbCommand("INSERT INTO tblstructureproduction (StructureID, OrderNum, PointsProduced, ObjectID, ObjTypeID, ProdCount) VALUES (" & lFacProdStructID(X).ToString & ", 0, " & .PointsProduced.ToString & ", " & .ProductionID.ToString & ", " & .ProductionTypeID.ToString & ", 1)", goCN)
                    oTmpComm.ExecuteNonQuery()
                    oTmpComm.Dispose()
                End With
            End If
        Next X


        LogEvent(LogEventType.Informational, "Finalizing Globals...")
        FinalizeGlobals()

        Form1.bHideEventLines = False
        gfrmDisplayForm.AddEventLine("Program Ready to Terminate...")

        gfrmDisplayForm.ReEnableButton1()
        'Threading.Thread.Sleep(5000)
        'Application.Exit()

    End Sub

	Public Sub LoadInstancedPlanet(ByVal lInstanceID As Int32)
		Dim oPlanet As Planet = GetEpicaPlanet(lInstanceID)
		If oPlanet Is Nothing Then Return
		If oPlanet.oDomain Is Nothing Then Return
		If oPlanet.oDomain.DomainSocket Is Nothing Then Return

		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand = Nothing
		Dim oData As OleDb.OleDbDataReader = Nothing

		Try

			Dim lPlayerID As Int32 = lInstanceID - 500000000
			Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
			If oPlayer Is Nothing Then Return

            oPlayer.InMyDomain = True

            Dim sColonyList As String = "" 

			'Ok, an instanced planet can have:
			'	Colonies
			sSQL = "SELECT * FROM tblColony WHERE ParentID = " & lInstanceID & " AND ParentTypeID = " & CInt(ObjectType.ePlanet)
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oColony As Colony = New Colony
				With oColony
					.LoadFromDataReader(oData)

					Dim lCurUB As Int32 = -1
					If glColonyIdx Is Nothing = False Then lCurUB = Math.Min(glColonyIdx.GetUpperBound(0), glColonyUB)
					Dim lIdx As Int32 = -1
					For X As Int32 = 0 To lCurUB
						If glColonyIdx(X) = .ObjectID Then
							'already loaded
							Continue While
						ElseIf glColonyIdx(X) = -1 AndAlso lIdx = -1 Then
							lIdx = X
						End If
					Next X
					If lIdx = -1 Then
						lIdx = glColonyUB + 1
						ReDim Preserve goColony(lIdx)
						ReDim Preserve glColonyIdx(lIdx)
						glColonyIdx(lIdx) = -1
						glColonyUB += 1
					End If
					goColony(lIdx) = oColony
                    glColonyIdx(lIdx) = .ObjectID

					'Now, add this colony index to the Owner's Fast Colony Lookup
                    .Owner.AddColonyIndex(lIdx)

                    If sColonyList <> "" Then sColonyList &= ", "
                    sColonyList &= .ObjectID

					'We should have our PLANETS by now
					If .ParentObject Is Nothing = False Then
						If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
							CType(.ParentObject, Planet).AddColonyReference(lIdx)
						End If
					End If
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'	Facilities
			Dim sFacList As String = ""
			sSQL = "SELECT * FROM tblStructure WHERE ParentID = " & lInstanceID & " AND ParentTypeID = " & CInt(ObjectType.ePlanet)
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim lFacID As Int32 = CInt(oData("StructureID"))

				Dim lCurUB As Int32 = -1
				If glFacilityIdx Is Nothing = False Then lCurUB = Math.Min(glFacilityIdx.GetUpperBound(0), glFacilityUB)
				Dim lIdx As Int32 = -1
				For X As Int32 = 0 To lCurUB
					If glFacilityIdx(X) = lFacID Then
						'already loaded
						lIdx = X
						Exit For
						'Continue While
					ElseIf glFacilityIdx(X) = -1 AndAlso lIdx = -1 Then
						lIdx = X
					End If
				Next X
				If lIdx = -1 Then
					lIdx = glFacilityUB + 1
					ReDim Preserve goFacility(lIdx)
					ReDim Preserve glFacilityIdx(lIdx)
					glFacilityIdx(lIdx) = -1
					glFacilityUB += 1
				End If

				Dim oFac As New Facility()
				With oFac
					.ServerIndex = lIdx
					.LoadFromDataReader(oData)
					.AutoLaunch = True
					goFacility(lIdx) = oFac
                    glFacilityIdx(lIdx) = .ObjectID
                    If .ParentObject Is Nothing Then
                        glFacilityIdx(lIdx) = -1
                    Else
                        AddLookupFacility(.ObjectID, .ObjTypeID, .ServerIndex)
                    End If
                End With

				If sFacList <> "" Then sFacList &= ", "
                sFacList &= lFacID

                'ok, now, send this facility down
				oPlanet.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand_CE))
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

            '	Units
            Dim sUnitList As String = ""
			sSQL = "SELECT * FROM tblUnit WHERE (ParentID = " & lInstanceID & " AND ParentTypeID = " & CInt(ObjectType.ePlanet) & ")"
			If sFacList <> "" Then sSQL &= " OR (ParentTypeID = " & CInt(ObjectType.eFacility) & " AND ParentID IN (" & sFacList & "))"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				Dim oUnit As New Unit
				With oUnit
					.LoadFromDataReader(oData)

                    Dim lIdx As Int32 = AddUnitToGlobalArray(oUnit)

                    'Dim lCurUB As Int32 = -1
                    'If glUnitIdx Is Nothing = False Then lCurUB = Math.Min(glUnitIdx.GetUpperBound(0), glUnitUB)
                    'Dim lIdx As Int32 = -1
                    'For X As Int32 = 0 To lCurUB
                    '	If glUnitIdx(X) = .ObjectID Then
                    '		'already loaded
                    '		'Continue While
                    '		lIdx = X
                    '		Exit For
                    '	ElseIf glUnitIdx(X) = -1 AndAlso lIdx = -1 Then
                    '		lIdx = X
                    '	End If
                    'Next X
                    'If lIdx = -1 Then
                    '	lIdx = glUnitUB + 1
                    '	ReDim Preserve goUnit(lIdx)
                    '	ReDim Preserve glUnitIdx(lIdx)
                    '	glUnitIdx(lIdx) = -1
                    '	glUnitUB += 1
                    'End If
                    'goUnit(lIdx) = oUnit
                    'glUnitIdx(lIdx) = .ObjectID
                    If sUnitList <> "" Then sUnitList &= ", "
                    sUnitList &= .ObjectID.ToString

                    Dim lTemp As Int32 = CInt(oData("UnitGroupID"))
                    Dim lParentTypeID As Int32 = CInt(oData("ParentTypeID"))
                    If lParentTypeID = ObjectType.eGalaxy Then
                        Dim lTestFleetID As Int32 = CInt(oData("TestFleetID"))
                        If lTemp < 1 Then lTemp = lTestFleetID
                    End If

                    If lTemp > 0 Then
                        For X As Int32 = 0 To glUnitGroupUB
                            If glUnitGroupIdx(X) = lTemp Then
                                goUnitGroup(X).AddUnit(lIdx, False)
                                Exit For
                            End If
                        Next X
                    End If

                    If .ParentObject Is Nothing Then glUnitIdx(lIdx) = -1

                    'ok, now, send this unit down
                    If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet AndAlso CType(.ParentObject, Epica_GUID).ObjectID = oPlanet.ObjectID Then
                        oPlanet.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand_CE))
                    End If
                End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'	Mineralcaches
			sSQL = "SELECT * FROM tblMineralCache WHERE (ParentID = " & lInstanceID & " AND ParentTypeID = " & CInt(ObjectType.ePlanet) & ")"
            If sFacList <> "" Then sSQL &= " OR (ParentTypeID = " & CInt(ObjectType.eFacility) & " AND ParentID IN (" & sFacList & "))"
            If sColonyList <> "" Then sSQL &= " OR (ParentTypeID = " & CInt(ObjectType.eColony) & " AND ParentID IN (" & sColonyList & "))"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read

				Dim oCache As New MineralCache
				With oCache
					.lServerIndex = -1
					.CacheTypeID = CByte(oData("CacheTypeID"))
					.Concentration = CInt(oData("Concentration"))
					.LocX = CInt(oData("LocX"))
					.LocZ = CInt(oData("LocY"))
					.ObjectID = CInt(oData("CacheID"))
					.ObjTypeID = ObjectType.eMineralCache
					.oMineral = GetEpicaMineral(CInt(oData("MineralID")))
					.ParentObject = GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
					.OriginalConcentration = CInt(oData("OriginalConcentration"))
					.Quantity = CInt(oData("Quantity"))


					Dim lCacheID As Int32 = CInt(oData("CacheID"))
					Dim lCurUB As Int32 = -1
					If glMineralCacheIdx Is Nothing = False Then lCurUB = Math.Min(glMineralCacheIdx.GetUpperBound(0), glMineralCacheUB)
					Dim lIdx As Int32 = -1
					For X As Int32 = 0 To lCurUB
						If glMineralCacheIdx(X) = lCacheID Then
							'already loaded
							'Continue While
							lIdx = X
							Exit For
						ElseIf glMineralCacheIdx(X) = -1 AndAlso lIdx = -1 Then
							lIdx = X
						End If
					Next X
					If lIdx = -1 Then
						lIdx = glMineralCacheUB + 1
						ReDim Preserve goMineralCache(lIdx)
						ReDim Preserve glMineralCacheIdx(lIdx)
						glMineralCacheIdx(lIdx) = -1
						glMineralCacheUB += 1
					End If
					goMineralCache(lIdx) = oCache
					glMineralCacheIdx(lIdx) = oCache.ObjectID
					oCache.lServerIndex = lIdx

					'Ensure we register with a containing object's cargo...
					If .ParentObject Is Nothing = False Then
						If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility OrElse CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eUnit Then
							CType(.ParentObject, Epica_Entity).AddCargoRef(CType(oCache, Epica_GUID))
						ElseIf CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet AndAlso CType(.ParentObject, Epica_GUID).ObjectID = oPlanet.ObjectID Then
							'ok, now, send this cache down if it is in the planet
                            oPlanet.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oCache, GlobalMessageCode.eAddObjectCommand_CE))
                        ElseIf CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eColony Then
                            CType(.ParentObject, Colony).AdjustColonyMineralCache(.oMineral.ObjectID, .Quantity)
                        End If
					End If

                    .bNeedsAsync = False
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing




			Dim lColonyIdx As Int32 = oPlayer.GetColonyFromParent(lInstanceID, ObjectType.ePlanet)
            If lColonyIdx > -1 AndAlso glColonyIdx(lColonyIdx) > -1 Then
                Dim oColony As Colony = goColony(lColonyIdx)
                If oColony Is Nothing = False Then
                    For X As Int32 = 0 To oColony.ChildrenUB
                        If oColony.lChildrenIdx(X) > -1 Then
                            Dim oFac As Facility = oColony.oChildren(X)
                            If oFac Is Nothing = False AndAlso oFac.yProductionType = ProductionType.eMining Then

                                Dim rcTemp As Rectangle = Rectangle.FromLTRB(oFac.LocX - 50, oFac.LocZ - 50, oFac.LocX + 50, oFac.LocZ + 50)

                                For Y As Int32 = 0 To glMineralCacheUB
                                    If glMineralCacheIdx(Y) <> -1 Then
                                        Dim oCache As MineralCache = goMineralCache(Y)
                                        If oCache Is Nothing = False AndAlso oCache.CacheTypeID = MineralCacheType.eMineable Then
                                            With oCache
                                                If CType(.ParentObject, Epica_GUID).ObjectID = CType(oFac.ParentObject, Epica_GUID).ObjectID Then
                                                    If CType(.ParentObject, Epica_GUID).ObjTypeID = CType(oFac.ParentObject, Epica_GUID).ObjTypeID Then
                                                        'Ok, check location
                                                        If rcTemp.Contains(.LocX, .LocZ) = True Then
                                                            'Ok, found one
                                                            If .BeingMinedBy Is Nothing Then
                                                                oFac.lCacheIndex = Y
                                                                oFac.lCacheID = .ObjectID
                                                                If oFac.Active Then oFac.bMining = True '???
                                                                .BeingMinedBy = oFac
                                                                AddFacilityMining(oFac.ServerIndex, oFac.ObjectID)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End With
                                        End If
                                    End If
                                Next Y
                            End If

                        End If
                    Next X
                    oColony.UpdateAllValues(-1)
                End If

            End If

            LogEvent(LogEventType.Informational, "Loading Unit Routes...")
            If sUnitList <> "" Then
                sSQL = "SELECT * FROM tblRouteItem WHERE UnitID IN (" & sUnitList & ") ORDER BY OrderNum"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                While oData.Read
                    Dim lUnitID As Int32 = CInt(oData("UnitID"))
                    Dim oUnit As Unit = GetEpicaUnit(lUnitID)
                    If oUnit Is Nothing = False Then
                        Dim uNewRouteItem As RouteItem
                        With uNewRouteItem
                            .lLocX = CInt(oData("LocX"))
                            .lLocZ = CInt(oData("LocZ"))
                            .lOrderNum = CInt(oData("OrderNum"))
                            .oDest = CType(GetEpicaObject(CInt(oData("DestID")), CShort(oData("DestTypeID"))), Epica_GUID)
                            Dim lTmpVal As Int32 = CInt(oData("LoadItemID"))
                            Dim iTmpVal As Int16 = CShort(oData("LoadItemTypeID"))
                            .oLoadItem = Nothing

                            Dim yFlag As Byte = CByte(oData("ExtraFlags"))
                            .SetLoadItem(lTmpVal, iTmpVal, oUnit.Owner, yFlag)
                        End With
                        oUnit.AddRouteItem(uNewRouteItem)
                    End If
                End While
                oData.Close()
                oData = Nothing
                oComm = Nothing
            End If

            LoadAgentDetails("(" & oPlayer.ObjectID.ToString & ")")

            If oPlayer.lTutorialStep >= 297 Then
                AureliusAI.SpawnNextWave(oPlayer.ObjectID, lInstanceID)
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "LoadInstancedPlanet: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            oData = Nothing
            oComm = Nothing
        End Try

	End Sub

	Public Sub SaveAndUnloadInstancedPlanet(ByVal lTempID As Int32)
		'Forcefully finish all production?
		FastForwardProduction(lTempID, -1, -1)

		lTempID += 500000000
		Dim oPlanet As Planet = GetEpicaPlanet(lTempID)

		If oPlanet Is Nothing Then Return
		If oPlanet.oDomain Is Nothing Then Return
		If oPlanet.oDomain.DomainSocket Is Nothing Then Return

		'Send to Region the SaveAndUnloadInstancedPlanet command
		Dim yMsg(5) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eSaveAndUnloadInstance).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(lTempID).CopyTo(yMsg, lPos) : lPos += 4
		oPlanet.oDomain.DomainSocket.SendData(yMsg)
	End Sub

    Public Sub SpawnPirateFactory(ByVal lPlanetID As Int32)
        'spawn the player's pirate medium facility
        Dim oPlanet As Planet = GetEpicaPlanet(lPlanetID)
        If oPlanet Is Nothing Then Return

        Dim oFacDef As FacilityDef = GetEpicaFacilityDef(49)
        Dim oFac As Facility = New Facility
        With oFac
            .bProducing = False
            .CurrentProduction = Nothing
            .CurrentSpeed = 0
            .EntityDef = oFacDef
            .CurrentStatus = 0
            For Y As Int32 = 0 To oFacDef.lSideCrits.Length - 1
                .CurrentStatus = .CurrentStatus Or oFacDef.lSideCrits(Y)
            Next Y

            .lPirate_For_PlayerID = lPlanetID - 500000000

            oFacDef.DefName.CopyTo(.EntityName, 0)
            .ExpLevel = 0
            .Fuel_Cap = oFacDef.Fuel_Cap
            .iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage
            .iTargetingTactics = 0
            .LocAngle = 0

            .LocX = 14000 : .LocZ = -14500

            .Owner = GetEpicaPlayer(gl_HARDCODE_PIRATE_PLAYER_ID)
            .ParentObject = oPlanet
            .ObjTypeID = ObjectType.eFacility

            .Q1_HP = oFacDef.Q1_MaxHP
            .Q2_HP = oFacDef.Q2_MaxHP
            .Q3_HP = oFacDef.Q3_MaxHP
            .Q4_HP = oFacDef.Q4_MaxHP
            .Shield_HP = oFacDef.Shield_MaxHP
            .Structure_HP = Math.Min(6000, oFacDef.Structure_MaxHP)
            .yProductionType = oFacDef.ProductionTypeID

            .DataChanged()

            If .SaveObject() = True Then
                'Now, find a suitable place...
                Dim lIdx As Int32 = -1
                'SyncLock goFacility
                '    For Y As Int32 = 0 To glFacilityUB
                '        If glFacilityIdx(Y) = -1 Then
                '            lIdx = Y
                '            Exit For
                '        End If
                '    Next Y

                '    If lIdx = -1 Then
                '        ReDim Preserve glFacilityIdx(glFacilityUB + 1)
                '        ReDim Preserve goFacility(glFacilityUB + 1)
                '        glFacilityUB += 1
                '        lIdx = glFacilityUB
                '    End If

                '    goFacility(lIdx) = oFac
                '    glFacilityIdx(lIdx) = oFac.ObjectID
                'End SyncLock
                lIdx = AddFacilityToGlobalArray(oFac)


                .CurrentStatus = goFacility(lIdx).CurrentStatus Or elUnitStatus.eFacilityPowered
                .ServerIndex = lIdx
                .DataChanged()
                .SaveObject()
            Else
                LogEvent(LogEventType.CriticalError, "SpawnPirateFactory failed.")
            End If
            CType(oFac.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
        End With
        AddPirateProductionItem(oFac)

        oFac = Nothing
        oFacDef = GetEpicaFacilityDef(46)
        For lTmpIdx As Int32 = 0 To 3
            oFac = New Facility()
            With oFac
                .bProducing = False
                .CurrentProduction = Nothing
                .CurrentSpeed = 0
                .EntityDef = oFacDef
                .CurrentStatus = 0
                For Y As Int32 = 0 To oFacDef.lSideCrits.Length - 1
                    .CurrentStatus = .CurrentStatus Or oFacDef.lSideCrits(Y)
                Next Y

                oFacDef.DefName.CopyTo(.EntityName, 0)
                .ExpLevel = 0
                .Fuel_Cap = oFacDef.Fuel_Cap
                .iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage
                .iTargetingTactics = 0
                .LocAngle = 0

                Select Case lTmpIdx
                    Case 0
                        .LocX = 13000 : .LocZ = -13500
                    Case 1
                        .LocX = 15000 : .LocZ = -13500
                    Case 2
                        .LocX = 13000 : .LocZ = -15500
                    Case Else
                        .LocX = 15000 : .LocZ = -15500
                End Select


                .Owner = GetEpicaPlayer(gl_HARDCODE_PIRATE_PLAYER_ID)
                .ParentObject = oPlanet
                .ObjTypeID = ObjectType.eFacility

                .Q1_HP = oFacDef.Q1_MaxHP
                .Q2_HP = oFacDef.Q2_MaxHP
                .Q3_HP = oFacDef.Q3_MaxHP
                .Q4_HP = oFacDef.Q4_MaxHP
                .Shield_HP = oFacDef.Shield_MaxHP
                .Structure_HP = oFacDef.Structure_MaxHP
                .yProductionType = oFacDef.ProductionTypeID

                .DataChanged()

                If .SaveObject() = True Then
                    'Now, find a suitable place...
                    Dim lIdx As Int32 = -1
                    'SyncLock goFacility
                    '    For Y As Int32 = 0 To glFacilityUB
                    '        If glFacilityIdx(Y) = -1 Then
                    '            lIdx = Y
                    '            Exit For
                    '        End If
                    '    Next Y

                    '    If lIdx = -1 Then
                    '        ReDim Preserve glFacilityIdx(glFacilityUB + 1)
                    '        ReDim Preserve goFacility(glFacilityUB + 1)
                    '        glFacilityUB += 1
                    '        lIdx = glFacilityUB
                    '    End If

                    '    goFacility(lIdx) = oFac
                    '    glFacilityIdx(lIdx) = oFac.ObjectID
                    'End SyncLock
                    lIdx = AddFacilityToGlobalArray(oFac)

                    .CurrentStatus = goFacility(lIdx).CurrentStatus Or elUnitStatus.eFacilityPowered
                    .ServerIndex = lIdx
                    .DataChanged()
                    .SaveObject()
                Else
                    LogEvent(LogEventType.CriticalError, "SpawnPirateFactoryTurrets failed.")
                End If
                CType(oFac.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
            End With
        Next lTmpIdx
	End Sub

	Public Function LoadSingleSystem(ByVal lSystemID As Int32) As Boolean

		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand = Nothing
		Dim oData As OleDb.OleDbDataReader = Nothing
		Dim bRes As Boolean = False

		Try
			'System Objects
			LogEvent(LogEventType.Informational, "Operator assigning new system: SystemID = " & lSystemID & "...")
			LogEvent(LogEventType.Informational, "Loading system...")

			Dim oSystem As SolarSystem = GetEpicaSystem(lSystemID)
			If oSystem Is Nothing = False Then
				LogEvent(LogEventType.Informational, "Already has system " & lSystemID)
				Return False
			End If

			sSQL = "SELECT * FROM tblSolarSystem WHERE SystemID = " & lSystemID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read

				oSystem = New SolarSystem()
				With oSystem
					.LocX = CInt(oData("LocX"))
					.LocY = CInt(oData("LocY"))
					.LocZ = CInt(oData("LocZ"))
					.ObjectID = CInt(oData("SystemID"))
					.ObjTypeID = ObjectType.eSolarSystem
					.ParentGalaxy = GetEpicaGalaxy(CInt(oData("galaxyid")))
					.StarType1ID = CByte(oData("StarType1ID"))
					.StarType2ID = CByte(oData("StarType2ID"))
					.StarType3ID = CByte(oData("StarType3ID"))
					.SystemName = StringToBytes(CStr(oData("SystemName")))
					.SystemType = CByte(oData("SystemType"))
                    .OwnerID = CInt(oData("OwnerID"))
                    .FleetJumpPointX = CInt(oData("FleetJumpPointX"))
                    .FleetJumpPointZ = CInt(oData("FleetJumpPointZ"))
                End With

				glSystemUB = glSystemUB + 1
				ReDim Preserve goSystem(glSystemUB)
				ReDim Preserve glSystemIdx(glSystemUB)
				goSystem(glSystemUB) = oSystem
				glSystemIdx(glSystemUB) = lSystemID
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Planet Objects
			LogEvent(LogEventType.Informational, "Loading Planets for system " & lSystemID & "...")
			sSQL = "SELECT * FROM tblPlanet WHERE ParentID = " & lSystemID & " ORDER BY PlanetID"
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glPlanetUB = glPlanetUB + 1
				ReDim Preserve goPlanet(glPlanetUB)
				ReDim Preserve glPlanetIdx(glPlanetUB)
				goPlanet(glPlanetUB) = New Planet()
				With goPlanet(glPlanetUB)
					.Atmosphere = CByte(oData("Atmosphere"))
					.AxisAngle = CInt(oData("AxisAngle"))
					.Gravity = CByte(oData("Gravity"))
					.Hydrosphere = CByte(oData("Hydrosphere"))
					.InnerRingRadius = CInt(oData("RingInnerRadius"))
					.LocX = CInt(oData("LocX"))
					.LocY = CInt(oData("LocY"))
					.LocZ = CInt(oData("LocZ"))
					.ObjectID = CInt(oData("PlanetID"))
					.ObjTypeID = ObjectType.ePlanet
					.OuterRingRadius = CInt(oData("RingOuterRadius"))
					.ParentSystem = oSystem
					'Set the parent system's Planet Index
					.ParentSystem.AddPlanetIndex(glPlanetUB)
					.PlanetName = StringToBytes(CStr(oData("PlanetName")))
					.PlanetRadius = CShort(oData("PlanetRadius"))
					.PlanetSizeID = CByte(oData("PlanetSizeID"))
					.PlanetTypeID = CByte(oData("PlanetTypeID"))
					.RingDiffuse = CInt(oData("RingDiffuse"))
					.RotationDelay = CShort(oData("RotationDelay"))
					.SurfaceTemperature = CByte(oData("SurfaceTemperature"))
					.Vegetation = CByte(oData("Vegetation"))
					.blOriginalMineralQuantity = CLng(oData("OrigMinQty"))
					.ySentGNSLowRes = CByte(oData("SentGNSLowRes"))
                    .lPrimaryComposition = CInt(oData("PrimaryComposition"))
                    .RingMineralConcentration = CInt(oData("RingMineralConcentration"))
                    .RingMineralID = CInt(oData("RingMineralID"))

					.PlayerSpawns = CInt(oData("PlayerSpawns"))
					.SpawnLocked = CByte(oData("SpawnLocked")) <> 0
					.OwnerID = CInt(oData("OwnerID"))

					glPlanetIdx(glPlanetUB) = .ObjectID
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

			'Wormhole objects
			LogEvent(LogEventType.Informational, "Loading Wormholes for system " & lSystemID & "...")
			sSQL = "SELECT * FROM tblWormhole WHERE System1ID = " & lSystemID & " OR System2ID = " & lSystemID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oData = oComm.ExecuteReader(CommandBehavior.Default)
			While oData.Read
				glWormholeUB = glWormholeUB + 1
				ReDim Preserve goWormhole(glWormholeUB)
				ReDim Preserve glWormholeIdx(glWormholeUB)
				goWormhole(glWormholeUB) = New Wormhole()
				With goWormhole(glWormholeUB)
                    .WormholeFlags = CInt(oData("WormholeFlag"))
					.LocX1 = CInt(oData("LocX1"))
					.LocY1 = CInt(oData("LocY1"))
					.LocX2 = CInt(oData("LocX2"))
					.LocY2 = CInt(oData("LocY2"))
					.ObjectID = CInt(oData("WormholeID"))
					.ObjTypeID = ObjectType.eWormhole
					.StartCycle = CInt(oData("StartCycle"))
					.System1 = GetEpicaSystem(CInt(oData("System1ID")))
					.System2 = GetEpicaSystem(CInt(oData("System2ID")))
					glWormholeIdx(glWormholeUB) = .ObjectID

					If .System1 Is Nothing = False Then .System1.AddWormholeReference(goWormhole(glWormholeUB))
					If .System2 Is Nothing = False Then .System2.AddWormholeReference(goWormhole(glWormholeUB))
				End With
			End While
			oData.Close()
			oData = Nothing
			oComm = Nothing

            LogEvent(LogEventType.Informational, "Creating room for mineral caches...")
            sSQL = "SELECT count(*) FROM tblMineralCache WHERE ParentTypeID = 3 AND ParentID IN (SELECT PlanetID FROM tblPlanet WHERE ParentID = " & lSystemID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            Dim lMinCacheCnt As Int32 = 0
            If oData.Read = True Then
                lMinCacheCnt = CInt(oData(0))
            End If
            oData.Close()
            oData = Nothing
            oComm = Nothing

            Dim lNewUB As Int32 = glMineralCacheUB + lMinCacheCnt
            Dim lOldUB As Int32 = glMineralCacheUB
            ReDim Preserve goMineralCache(lNewUB)
            ReDim Preserve glMineralCacheIdx(lNewUB)
            glMineralCacheUB = lNewUB
            For X As Int32 = lOldUB + 1 To lNewUB
                glMineralCacheIdx(X) = -1
            Next X

            LogEvent(LogEventType.Informational, "Loading Mineral Caches...")
            sSQL = "SELECT * FROM tblMineralCache WHERE ParentTypeID = 3 AND ParentID IN (SELECT PlanetID FROM tblPlanet WHERE ParentID = " & lSystemID & ")"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader(CommandBehavior.Default)
            While oData.Read
                Dim lParentID As Int32 = CInt(oData("ParentID"))
                Dim iParentTypeID As Int16 = CShort(oData("ParentTypeID"))

                Dim oParentObject As Object = GetEpicaObject(lParentID, iParentTypeID)
                If oParentObject Is Nothing Then Continue While

                Dim oCache As New MineralCache
                With oCache
                    .CacheTypeID = CByte(oData("CacheTypeID"))
                    .Concentration = CInt(oData("Concentration"))
                    .LocX = CInt(oData("LocX"))
                    .LocZ = CInt(oData("LocY"))
                    .ObjectID = CInt(oData("CacheID"))
                    .ObjTypeID = ObjectType.eMineralCache
                    .oMineral = GetEpicaMineral(CInt(oData("MineralID")))
                    .ParentObject = oParentObject 'GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
                    .OriginalConcentration = CInt(oData("OriginalConcentration"))
                    .Quantity = CInt(oData("Quantity"))
                    .bNeedsAsync = False
                End With
                oCache.lServerIndex = AddMineralCacheToGlobalArray(oCache)
            End While
            oData.Close()
            oData = Nothing
            oComm = Nothing

            bRes = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "LoadSingleSystem: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            oData = Nothing
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try
		Return bRes
		
    End Function

    Public Sub PeriodicProdSave()
        Dim oSB As New System.Text.StringBuilder
        Dim lCnt As Int32 = 0

        oSB.AppendLine("DROP TABLE tblStructureProductionPeriodic_Prev")
        oSB.AppendLine("SELECT * INTO tblStructureProductionPeriodic_Prev FROM tblStructureProductionPeriodic")

        oSB.AppendLine("DELETE FROM tblStructureProductionPeriodic")

        For X As Int32 = 0 To glFacilityUB
            If glFacilityIdx(X) > -1 Then
                Dim oFac As Facility = goFacility(X)
                If oFac Is Nothing = False AndAlso oFac.Active = True AndAlso oFac.CurrentProduction Is Nothing = False Then

                    With oFac.CurrentProduction
                        Dim sSQL As String = "INSERT INTO tblstructureproductionperiodic (StructureID, OrderNum, PointsProduced, ObjectID, ObjTypeID, ProdCount) VALUES (" & oFac.ObjectID.ToString & ", 0, " & .PointsProduced.ToString & ", " & .ProductionID.ToString & ", " & .ProductionTypeID.ToString & ", 1)"
                        oSB.AppendLine(sSQL)
                    End With
                    lCnt += 1
                    If lCnt > 499 Then
                        Try
                            Dim oTmpComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(oSB.ToString, goCN)
                            oTmpComm.ExecuteNonQuery()
                            oTmpComm.Dispose()
                        Catch
                        End Try

                        oSB = Nothing
                        oSB = New System.Text.StringBuilder
                        lCnt = 0
                    End If
                End If
            End If
        Next X

        If oSB.Length > 0 Then
            Try
                Dim oTmpComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(oSB.ToString, goCN)
                oTmpComm.ExecuteNonQuery()
                oTmpComm.Dispose()
            Catch
            End Try
        End If
    End Sub
#End Region

#Region "  Get Object Functions  "
    Public Function GetEpicaAgent(ByVal lObjectID As Int32) As Agent
        Dim X As Int32
        For X = 0 To glAgentUB
            If glAgentIdx(X) = lObjectID Then
                Return goAgent(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaColony(ByVal lObjectID As Int32) As Colony
        Dim X As Int32

        'NOTE: this is one of the SLOWEST ways to get a colony

        For X = 0 To glColonyUB
            If glColonyIdx(X) = lObjectID Then
                Return goColony(X)
            End If
        Next X
        Return Nothing
    End Function

	Public Function GetEpicaComponentCache(ByVal lCacheID As Int32) As ComponentCache
		Dim lCurUB As Int32 = Math.Min(glComponentCacheUB, glComponentCacheIdx.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB
			If glComponentCacheIdx(X) = lCacheID Then Return goComponentCache(X)
		Next X
		Return Nothing
	End Function

    Public Function GetEpicaCorporation(ByVal lObjectID As Int32) As Corporation
        Dim X As Int32
        For X = 0 To glCorporationUB
            If glCorporationIdx(X) = lObjectID Then
                Return goCorporation(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaEntity(ByVal lObjID As Int32, ByVal iObjTypeID As Int16) As Epica_Entity
        If iObjTypeID = ObjectType.eUnit Then Return GetEpicaUnit(lObjID)
        If iObjTypeID = ObjectType.eFacility Then Return GetEpicaFacility(lObjID)
        Return Nothing
    End Function

    Public Function GetEpicaFacility(ByVal lObjectID As Int32) As Facility
        Dim X As Int32 = LookupFacility(lObjectID, ObjectType.eFacility)
        If X < 0 OrElse X > glFacilityUB Then
            For X = 0 To glFacilityUB
                If glFacilityIdx(X) = lObjectID Then
                    Dim oFac As Facility = goFacility(X)
                    If oFac Is Nothing = False AndAlso oFac.ObjectID = lObjectID Then Return oFac Else Return Nothing
                End If
            Next X
            Return Nothing
        Else
            Return goFacility(X)
        End If
    End Function

    Public Function GetEpicaFacilityDef(ByVal lObjectID As Int32) As FacilityDef
        Dim X As Int32
        For X = 0 To glFacilityDefUB
            If glFacilityDefIdx(X) = lObjectID Then
                Return goFacilityDef(X)
            End If
        Next X
        Return Nothing
	End Function

	Public Function GetEpicaFormation(ByVal lFormationID As Int32) As FormationDef
		Dim lCurUB As Int32 = Math.Min(glFormationDefUB, glFormationDefIdx.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB
			If glFormationDefIdx(X) = lFormationID Then Return goFormationDefs(X)
		Next X
		Return Nothing
	End Function

    Public Function GetEpicaGalaxy(ByVal lObjectID As Int32) As Galaxy
        Dim X As Int32
        For X = 0 To glGalaxyUB
            If glGalaxyIdx(X) = lObjectID Then
                Return goGalaxy(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaGNS(ByVal lObjectID As Int32) As GNS
        Dim X As Int32
        For X = 0 To glGNSUB
            If glGNSIdx(X) = lObjectID Then
                Return goGNS(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaGoal(ByVal lObjectID As Int32) As Goal
        Dim X As Int32
        For X = 0 To glGoalUB
            If glGoalIdx(X) = lObjectID Then
                Return goGoal(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaGuild(ByVal lObjectID As Int32) As Guild
        For X As Int32 = 0 To glGuildUB
            If glGuildIdx(X) = lObjectID Then Return goGuild(X)
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaMineral(ByVal lObjectID As Int32) As Mineral
        Dim X As Int32
        For X = 0 To glMineralUB
            If glMineralIdx(X) = lObjectID Then
                Return goMineral(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaMineralCache(ByVal lObjectID As Int32) As MineralCache
        Dim X As Int32
        For X = 0 To glMineralCacheUB
            If glMineralCacheIdx(X) = lObjectID Then
                Return goMineralCache(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaMission(ByVal lObjectID As Int32) As Mission
        Dim X As Int32
        For X = 0 To glMissionUB
            If glMissionIdx(X) = lObjectID Then
                Return goMission(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaNebula(ByVal lObjectID As Int32) As Nebula
        Dim X As Int32
        For X = 0 To glNebulaUB
            If glNebulaIdx(X) = lObjectID Then
                Return goNebula(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaPlanet(ByVal lObjectID As Int32) As Planet
        Dim X As Int32
        For X = 0 To glPlanetUB
            If glPlanetIdx(X) = lObjectID Then
                Return goPlanet(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaPlayer(ByVal lObjectID As Int32) As Player
        'Dim X As Int32
        'For X = 0 To glPlayerUB
        '    If glPlayerIdx(X) = lObjectID Then
        '        Return goPlayer(X)
        '    End If
        'Next X
        'Return Nothing
        If lObjectID = -1 Then Return Nothing
        If lObjectID > glPlayerUB Then
            LogEvent(LogEventType.PossibleCheat, "GetEpicaPlayer is trying to reference player ID lObjectID:" & CStr(lObjectID))
            Return Nothing

        Else
            If glPlayerIdx(lObjectID) <> lObjectID AndAlso lObjectID <> 0 Then
                LogEvent(LogEventType.CriticalError, "GetEpicaPlayer is trying to reference player ID lObjectID which is not in the playeridx:" & CStr(lObjectID))
                'some serious issiues
            End If
            Return goPlayer(lObjectID)


        End If
    End Function

    Public Function GetEpicaPlayerMission(ByVal lObjectID As Int32) As PlayerMission
        For X As Int32 = 0 To glPlayerMissionUB
            If glPlayerMissionIdx(X) = lObjectID Then Return goPlayerMission(X)
        Next X
        Return Nothing
    End Function

	'Public Function GetEpicaSenate(ByVal lObjectID As Int32) As Senate
	'    Dim X As Int32
	'    For X = 0 To glSenateUB
	'        If glSenateIdx(X) = lObjectID Then
	'            Return goSenate(X)
	'        End If
	'    Next X
	'    Return Nothing
	'End Function

    Public Function GetEpicaSkill(ByVal lObjectID As Int32) As Skill
        Dim X As Int32
        For X = 0 To glSkillUB
            If glSkillIdx(X) = lObjectID Then
                Return goSkill(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaSystem(ByVal lObjectID As Int32) As SolarSystem
        Dim X As Int32
        For X = 0 To glSystemUB
            If glSystemIdx(X) = lObjectID Then
                Return goSystem(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaSpecialTech(ByVal lObjectID As Int32) As SpecialTech
        'DO NOT INCLUDE THIS IN GETEPICAOBJECT!!!!!
        For X As Int32 = 0 To glSpecialTechUB
            If glSpecialTechIdx(X) = lObjectID Then Return goSpecialTechs(X)
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaUnit(ByVal lObjectID As Int32) As Unit
        Dim X As Int32
        For X = 0 To glUnitUB
            If glUnitIdx(X) = lObjectID Then
                Return goUnit(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaUnitDef(ByVal lObjectID As Int32) As Epica_Entity_Def
        Dim X As Int32
        For X = 0 To glUnitDefUB
            If glUnitDefIdx(X) = lObjectID Then
                Return goUnitDef(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaUnitGroup(ByVal lObjectID As Int32) As UnitGroup
        For X As Int32 = 0 To glUnitGroupUB
            If glUnitGroupIdx(X) = lObjectID Then Return goUnitGroup(X)
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaWeaponDef(ByVal lObjectID As Int32) As WeaponDef
        Dim X As Int32
        For X = 0 To glWeaponDefUB
            If glWeaponDefIdx(X) = lObjectID Then
                Return goWeaponDefs(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaWormhole(ByVal lObjectID As Int32) As Wormhole
        Dim X As Int32
        For X = 0 To glWormholeUB
            If glWormholeIdx(X) = lObjectID Then
                Return goWormhole(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetEpicaMineralProperty(ByVal lObjectID As Int32) As MineralProperty
        Dim X As Int32
        For X = 0 To glMineralPropertyUB
            If glMineralPropertyIdx(X) = lObjectID Then
                Return goMineralProperty(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetAndSaveEpicaObject(ByVal lObjectID As Int32, ByVal iTypeID As Int16) As Boolean
        'Ok, get the object first
        Dim oObj As Epica_GUID = CType(GetEpicaObject(lObjectID, iTypeID), Epica_GUID)
        If oObj Is Nothing = False Then

            'These are the saveable ones
            Select Case oObj.ObjTypeID
                Case ObjectType.eAgent
                    Return CType(oObj, Agent).SaveObject()
                Case ObjectType.eColony
                    Return CType(oObj, Colony).SaveObject()
                Case ObjectType.eComponentCache
                    Return CType(oObj, ComponentCache).SaveObject()
                Case ObjectType.eFacility
                    Return CType(oObj, Facility).SaveObject()
                Case ObjectType.eFacilityDef
                    Return CType(oObj, FacilityDef).SaveObject()
                Case ObjectType.eGalacticTrade
                    Return goGTC.SaveGTC()
                Case ObjectType.eGalaxy
                    Return CType(oObj, Galaxy).SaveObject()
                Case ObjectType.eGuild
                    Return CType(oObj, Guild).SaveObject()
                Case ObjectType.eMineral
                    Return CType(oObj, Mineral).SaveObject()
                Case ObjectType.eMineralCache
                    Return CType(oObj, MineralCache).SaveObject()
                Case ObjectType.eNebula
                    Return CType(oObj, Nebula).SaveObject()
                Case ObjectType.ePlanet
                    Return CType(oObj, Planet).SaveObject()
                Case ObjectType.ePlayer
                    Return CType(oObj, Player).SaveObject(False)
                Case ObjectType.ePlayerComm
                    Return CType(oObj, PlayerComm).SaveObject(-1)
                Case ObjectType.ePlayerIntel
                    Return CType(oObj, PlayerIntel).SaveObject()                 
                Case ObjectType.eSkill
                    Return CType(oObj, Skill).SaveObject()
                Case ObjectType.eSolarSystem
                    Return CType(oObj, SolarSystem).SaveObject()
                Case ObjectType.eTrade
                    Return CType(oObj, DirectTrade).SaveObject()
                Case ObjectType.eUnit
                    Return CType(oObj, Unit).SaveObject()
                Case ObjectType.eUnitDef
                    Return CType(oObj, Epica_Entity_Def).SaveObject()
                Case ObjectType.eUnitGroup
                    Return CType(oObj, UnitGroup).SaveObject()
                Case ObjectType.eUnitWeaponDef
                    Return CType(oObj, UnitWeaponDef).SaveObject()
                Case ObjectType.eWeaponDef
                    Return CType(oObj, WeaponDef).SaveObject()
                Case ObjectType.eWormhole
                    Return CType(oObj, Wormhole).SaveObject()


                    'MSC - 10/01/08 - these should be saved only upon submission or research state change
                    'Case ObjectType.eAlloyTech
                    '    Return CType(oObj, AlloyTech).SaveObject()
                    'Case ObjectType.eArmorTech
                    '    Return CType(oObj, ArmorTech).SaveObject()
                    'Case ObjectType.eEngineTech
                    '    Return CType(oObj, EngineTech).SaveObject()
                    'Case ObjectType.eHullTech
                    '    Return CType(oObj, HullTech).SaveObject()
                    'Case ObjectType.ePrototype
                    '    Return CType(oObj, Prototype).SaveObject()
                    'Case ObjectType.eRadarTech
                    '    Return CType(oObj, RadarTech).SaveObject()
                    'Case ObjectType.eShieldTech
                    '    Return CType(oObj, ShieldTech).SaveObject()
                    'Case ObjectType.eWeaponTech
                    '    Return CType(oObj, BaseWeaponTech).SaveObject()

            End Select
        End If

    End Function

    Public Function GetEpicaObject(ByVal lObjectID As Int32, ByVal iTypeID As Int16) As Object
        Select Case iTypeID
            Case ObjectType.eUnit
                Return GetEpicaUnit(lObjectID)
            Case ObjectType.eUnitDef
                Return GetEpicaUnitDef(lObjectID)
            Case ObjectType.eFacility
                Return GetEpicaFacility(lObjectID)
            Case ObjectType.eFacilityDef
                Return GetEpicaFacilityDef(lObjectID)
            Case ObjectType.eGuild
                Return GetEpicaGuild(lObjectID)
            Case ObjectType.eAgent
                Return GetEpicaAgent(lObjectID)
            Case ObjectType.eColony
                Return GetEpicaColony(lObjectID)
            Case ObjectType.eCorporation
                Return GetEpicaCorporation(lObjectID)
            Case ObjectType.eGalaxy
                Return GetEpicaGalaxy(lObjectID)
            Case ObjectType.eGNS
                Return GetEpicaGNS(lObjectID)
            Case ObjectType.eGoal
                Return GetEpicaGoal(lObjectID)
            Case ObjectType.eMineral
                Return GetEpicaMineral(lObjectID)
            Case ObjectType.eMineralCache
                Return GetEpicaMineralCache(lObjectID)
            Case ObjectType.eMission
                Return GetEpicaMission(lObjectID)
            Case ObjectType.eNebula
                Return GetEpicaNebula(lObjectID)
            Case ObjectType.eObjectOrder
                'No global object for this
            Case ObjectType.ePlanet
                Return GetEpicaPlanet(lObjectID)
            Case ObjectType.ePlayer
                Return GetEpicaPlayer(lObjectID)
            Case ObjectType.ePlayerComm
                'No global object for this
				'Case ObjectType.eSenate
				'Return GetEpicaSenate(lObjectID)
            Case ObjectType.eSenateLaw
                'No global object for this
            Case ObjectType.eSkill
                Return GetEpicaSkill(lObjectID)
            Case ObjectType.eSolarSystem
                Return GetEpicaSystem(lObjectID)
            Case ObjectType.eUnitGroup
                Return GetEpicaUnitGroup(lObjectID)
            Case ObjectType.eUniverse
                'No global object for this
            Case ObjectType.eWeaponDef
                Return GetEpicaWeaponDef(lObjectID)
            Case ObjectType.eWormhole
                Return GetEpicaWormhole(lObjectID)
            Case ObjectType.eMineralProperty
                Return GetEpicaMineralProperty(lObjectID)
            Case ObjectType.eTrade
                Return goGTC.GetDirectTrade(lObjectID)
            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.ePrototype, ObjectType.eSpecialTech
                Return QuickLookupTechnology(lObjectID, iTypeID)
            Case ObjectType.eComponentCache
                For X As Int32 = 0 To glComponentCacheUB
                    If glComponentCacheIdx(X) = lObjectID Then
                        Return QuickLookupTechnology(goComponentCache(X).ComponentID, goComponentCache(X).ComponentTypeID)
                    End If
                Next X
            Case Else
                If iTypeID < 0 Then
                    'could be a component cache
                    Select Case Math.Abs(iTypeID)
                        Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.ePrototype, ObjectType.eSpecialTech
                            For X As Int32 = 0 To glComponentCacheUB
                                If glComponentCacheIdx(X) = lObjectID Then
                                    Return QuickLookupTechnology(goComponentCache(X).ComponentID, goComponentCache(X).ComponentTypeID)
                                End If
                            Next X
                    End Select
                End If
                'No object to return
        End Select

        Return Nothing
    End Function
#End Region

#Region "  Name Generation  "
    Private msMaleFirst() As String
    Private msFemaleFirst() As String
    Private msLastName() As String
    Private mbNameGenerationLoaded As Boolean = False

    Public Sub GenerateName(ByVal bMale As Boolean, ByRef sFirstName As String, ByRef sLastName As String)
        If mbNameGenerationLoaded = False Then LoadNameGeneration()

        If bMale = True Then
            sFirstName = msMaleFirst(CInt(Int(Rnd() * msMaleFirst.Length)))
        Else
			sFirstName = msFemaleFirst(CInt(Int(Rnd() * msFemaleFirst.Length)))
        End If
        sLastName = msLastName(CInt(Int(Rnd() * msLastName.Length)))
    End Sub

    Public Function GetFullGeneratedName(ByVal bMale As Boolean) As String
        Dim sFirstName As String = ""
        Dim sLastName As String = ""
        GenerateName(bMale, sFirstName, sLastName)
        Return sFirstName & " " & sLastName
    End Function

    Private Sub LoadNameGeneration()
        mbNameGenerationLoaded = True

        Dim oFile As IO.FileStream
        Dim oReader As IO.StreamReader
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"

        oFile = New IO.FileStream(sPath & "Male.txt", IO.FileMode.Open)
        oReader = New IO.StreamReader(oFile)

        ReDim msMaleFirst(-1)
        While oReader.EndOfStream = False
            ReDim Preserve msMaleFirst(msMaleFirst.GetUpperBound(0) + 1)
            msMaleFirst(msMaleFirst.GetUpperBound(0)) = oReader.ReadLine
        End While
        oReader.Dispose()
        oReader = Nothing
        oFile.Dispose()
        oFile = Nothing

        oFile = New IO.FileStream(sPath & "Female.txt", IO.FileMode.Open)
        oReader = New IO.StreamReader(oFile)

        ReDim msFemaleFirst(-1)
        While oReader.EndOfStream = False
            ReDim Preserve msFemaleFirst(msFemaleFirst.GetUpperBound(0) + 1)
            msFemaleFirst(msFemaleFirst.GetUpperBound(0)) = oReader.ReadLine
        End While
        oReader.Dispose()
        oReader = Nothing
        oFile.Dispose()
        oFile = Nothing

        oFile = New IO.FileStream(sPath & "LastName.txt", IO.FileMode.Open)
        oReader = New IO.StreamReader(oFile)

        ReDim msLastName(-1)
        While oReader.EndOfStream = False
            ReDim Preserve msLastName(msLastName.GetUpperBound(0) + 1)
            msLastName(msLastName.GetUpperBound(0)) = oReader.ReadLine
        End While
        oReader.Dispose()
        oReader = Nothing
        oFile.Dispose()
        oFile = Nothing
    End Sub

#End Region

#Region "  Bad Word Filter  "
    Private msWordChecks() As String = Nothing
    Private mlReplacementLen() As Int32 = Nothing
    Private mlWordCheckUB As Int32 = -1
    Private Sub InitializeWordChecks()
        AddWordCheck("fuck")
        AddWordCheck("bitch")
        AddWordCheck("shit")
        AddWordCheck("cunt")
        AddWordCheck("dick")
        AddWordCheck("fagit")
        AddWordCheck("faggot")
        AddWordCheck("nigger")
        AddWordCheck("pussy")
        AddWordCheck("puss")
        AddWordCheck("cock")
        AddWordCheck("fag")
        AddWordCheck("whore")
        AddWordCheck("dike")
        AddWordCheck("kike")
        AddWordCheck("jew")
    End Sub

    Private Sub AddWordCheck(ByVal sWord As String)
        ReDim Preserve msWordChecks(mlWordCheckUB + 2)
        ReDim Preserve mlReplacementLen(mlWordCheckUB + 2)

        Dim lNewIdx As Int32 = mlWordCheckUB + 1
        Dim lOtherIdx As Int32 = mlWordCheckUB + 2
        mlReplacementLen(lNewIdx) = sWord.Length
        mlReplacementLen(lOtherIdx) = sWord.Length
        msWordChecks(lNewIdx) = "\b"
        msWordChecks(lOtherIdx) = ""

        For X As Int32 = 0 To sWord.Length - 1
            Dim sChr As String = sWord(X).ToString
            Dim sCheckVal As String = "[" & sChr.ToLower & "|" & sChr.ToUpper
            If sChr.ToUpper = "S" Then sCheckVal &= "|$"
            If sChr.ToUpper = "I" Then sCheckVal &= "|!|1||"
            If sChr.ToUpper = "A" Then sCheckVal &= "|@"
            If sChr.ToUpper = "O" Then sCheckVal &= "|0"
            If sChr.ToUpper = "H" Then sCheckVal &= "|4"
            sCheckVal &= "]"
            msWordChecks(lNewIdx) &= sCheckVal & "[\W]?"
            msWordChecks(lOtherIdx) &= sCheckVal
        Next X

        msWordChecks(lOtherIdx) &= "\b"

        mlWordCheckUB += 2
    End Sub

    Public Function FilterBadWords(ByVal sInitial As String) As String

        If mlWordCheckUB = -1 Then InitializeWordChecks()

        For X As Int32 = 0 To msWordChecks.GetUpperBound(0)
            sInitial = System.Text.RegularExpressions.Regex.Replace(sInitial, msWordChecks(X), StrDup(mlReplacementLen(X), "*"), System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace)
        Next X

        Return sInitial
    End Function
#End Region
 
    Public Function GetDateFromNumber(ByVal lVal As Int32) As Date
        If lVal = 0 Then Return Date.MinValue

        Dim sVal As String = lVal.ToString

        'Work from right to left
        Dim lUB As Int32 = sVal.Length ' - 1

        'Ok, the bare minimum for this to work is 8
        If lUB < 8 Then Return Date.MinValue

        'Minute, last two values
        Dim sMin As String = sVal.Substring(lUB - 2)
        'Hour, two less from minute
        Dim sHr As String = sVal.Substring(lUB - 4, 2)
        'etc...
        Dim sDay As String = sVal.Substring(lUB - 6, 2)
        Dim sMon As String = sVal.Substring(lUB - 8, 2)

        Dim sYr As String = ""
        If lUB = 9 Then
            sYr = "0" & sVal.Substring(lUB - 9, 1)
        Else : sYr = sVal.Substring(lUB - 10, 2)
        End If

        Return CDate(sMon & "/" & sDay & "/20" & sYr & " " & sHr & ":" & sMin)
    End Function

    Public Function GetBigDateFromNumber(ByVal blVal As Int64) As DateTime
        If blVal = 0 Then Return Date.MinValue

        Dim sVal As String = blVal.ToString

        'Work from right to left
        Dim lUB As Int32 = sVal.Length ' - 1

        'Ok, the bare minimum for this to work is 8
        If lUB < 10 Then Return Date.MinValue

        Dim sSec As String = sVal.Substring(lUB - 2)
        'Minute, last two values
        Dim sMin As String = sVal.Substring(lUB - 4, 2)
        'Hour, two less from minute
        Dim sHr As String = sVal.Substring(lUB - 6, 2)
        'etc...
        Dim sDay As String = sVal.Substring(lUB - 8, 2)
        Dim sMon As String = sVal.Substring(lUB - 10, 2)

        Dim sYr As String = ""
        If lUB = 11 Then
            sYr = "0" & sVal.Substring(lUB - 11, 1)
        Else : sYr = sVal.Substring(lUB - 12, 2)
        End If

        Return CDate(sMon & "/" & sDay & "/20" & sYr & " " & sHr & ":" & sMin)

    End Function

    Public Function GetDateAsNumber(ByVal dtDate As Date) As Int32
		Return CInt((dtDate.ToString("yyMMddHHmm")))
    End Function

    Public Function GetStringFromBytes(ByRef yData() As Byte, ByVal lPos As Int32, ByVal lDataLen As Int32) As String
        Dim yTemp(lDataLen - 1) As Byte
        Array.Copy(yData, lPos, yTemp, 0, lDataLen)
        Dim lLen As Int32 = yTemp.Length
        For Y As Int32 = 0 To yTemp.Length - 1
            If yTemp(Y) = 0 Then
                lLen = Y
                Exit For
            End If
        Next Y
        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)
	End Function

    Public Function GetEntireErrorMsg(ByVal ex As Exception) As String
        Dim oSB As New System.Text.StringBuilder
        oSB.AppendLine(ex.Message)
        While ex.InnerException Is Nothing = False
            oSB.AppendLine(ex.InnerException.Message)
            ex = ex.InnerException
        End While
        Return oSB.ToString
    End Function

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef lEndX As Single, ByRef lEndY As Single, ByVal dDegree As Single)
        Dim dDX As Single
        Dim dDY As Single
        Dim dRads As Single

        dRads = dDegree * CSng(Math.PI / 180.0F)
        dDX = lEndX - lAxisX
        dDY = lEndY - lAxisY
        lEndX = lAxisX + CSng((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        lEndY = lAxisY + -CSng((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

#Region "  IP Parsing  "
    Public Function ParseForIpAddress(ByVal sIpAddress As String, Optional ByRef bRoutable As Boolean = True) As String
        'bRoutable: Only return routable ip's, not 10.x.x.x or 192.168.x.x or 172.16.x.x ~ 172.31.x.x or 5.0.x.x ~ 5.63.0.0
        'Will return Routable as a first priority
        If sIpAddress.ToLower = sIpAddress.ToUpper Then Return sIpAddress
        Dim sTempIP As String = ""
        Dim lPriority As Integer = 0
        Dim lTempPriority As Integer = -1

        If sIpAddress.ToUpper <> sIpAddress.ToLower Then 'Hostname -> Resolve it
            Dim IPList As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(sIpAddress)
            Dim IPaddr As System.Net.IPAddress
            For Each IPaddr In IPList.AddressList
                If (IPaddr.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork) Then
                    If IsPrivateIP(IPaddr.ToString) = False Then
                        Return IPaddr.ToString
                    Else
                        If lPriority > lTempPriority Then
                            sTempIP = IPaddr.ToString
                            lTempPriority = lPriority
                        End If
                    End If
                End If
            Next
        End If
        If bRoutable = False Then Return sTempIP
        Return ""

    End Function

    Private Function IsPrivateIP(ByVal CheckIP As String, Optional ByRef lPriority As Integer = -1) As Boolean
        Dim Quad() As Byte = CType(System.Net.IPAddress.Parse(CheckIP), System.Net.IPAddress).GetAddressBytes
        Select Case Quad(0)
            Case 10
                lPriority = 30
                Return True
            Case 172
                lPriority = 20
                If Quad(1) >= 16 And Quad(1) <= 31 Then Return True
            Case 5
                lPriority = 10
                If Quad(1) >= 0 And Quad(1) <= 63 Then Return True
            Case 192
                lPriority = 30
                If Quad(1) = 168 Then Return True
        End Select
        Return False
    End Function
#End Region

#Region "  Respawn List System  "
    Private Class RespawnListItem
        Public oSys As SolarSystem
        Public lActiveWars As Int32 = 0
        Public lMinCacheCount As Int32 = 0
        Public blTotalCacheQty As Int64 = 0
        Public lTotalCacheConcentration As Int32 = 0

        Public lPlayerCountByRank(7) As Int32

        Public colPlayersHere As New Collection

        Public Function GetActiveWHLinks() As Byte
            Dim lCnt As Int32 = 0
            Dim lFlagCheck As Int32 = (elWormholeFlag.eSystem1Detectable Or elWormholeFlag.eSystem2Detectable Or elWormholeFlag.eSystem1Jumpable Or elWormholeFlag.eSystem2Jumpable)
            If oSys Is Nothing = False Then
                For X As Int32 = 0 To oSys.mlWormholeUB
                    If oSys.moWormholes(X) Is Nothing = False Then
                        If (oSys.moWormholes(X).WormholeFlags And lFlagCheck) = lFlagCheck Then
                            lCnt += 1
                        End If
                    End If
                Next X
            End If
            If lCnt < 0 Then lCnt = 0
            If lCnt > 255 Then lCnt = 255
            Return CByte(lCnt)
        End Function
    End Class

    Private mlLastRespawnCompileCycle As Int32
    Private myCompiledRespawnList() As Byte
    Private mlRespawnCompileSyncLock(-1) As Int32
    Public Function CompileRespawnList() As Byte()

        SyncLock mlRespawnCompileSyncLock
            Try
                'if not ready or 1 longer than an hour - then regenerate it
                If myCompiledRespawnList Is Nothing OrElse glCurrentCycle - mlLastRespawnCompileCycle > 108000 Then
                    'Compile our respawn list

                    'Give everyone a chance to rest...
                    Threading.Thread.Sleep(10)

                    Dim colList As New Collection

                    For X As Int32 = 0 To glSystemUB
                        If glSystemIdx(X) > 0 Then
                            Dim oSys As SolarSystem = goSystem(X)
                            If oSys Is Nothing = False Then
                                If oSys.SystemType = SolarSystem.elSystemType.RespawnSystem OrElse oSys.SystemType = SolarSystem.elSystemType.UnlockedSystem Then
                                    Dim oNew As New RespawnListItem
                                    oNew.oSys = oSys
                                    For Y As Int32 = 0 To oNew.lPlayerCountByRank.GetUpperBound(0)
                                        oNew.lPlayerCountByRank(Y) = 0
                                    Next Y

                                    colList.Add(oNew, "Sys" & oSys.ObjectID.ToString)
                                End If
                            End If
                        End If
                    Next X

                    'To get active wars, we go thru the colony list
                    For X As Int32 = 0 To glColonyUB
                        If glColonyIdx(X) > 0 Then
                            Dim oCol As Colony = goColony(X)
                            If oCol Is Nothing = False AndAlso oCol.ParentObject Is Nothing = False Then
                                If CType(oCol.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                                    Dim oParentPlanet As Planet = CType(oCol.ParentObject, Planet)
                                    If oParentPlanet Is Nothing = False AndAlso oParentPlanet.ParentSystem Is Nothing = False Then
                                        Dim sKey As String = "Sys" & oParentPlanet.ParentSystem.ObjectID.ToString
                                        If colList.Contains(sKey) = True Then
                                            Dim oItem As RespawnListItem = CType(colList(sKey), RespawnListItem)

                                            If oItem Is Nothing = False Then
                                                'Ok, check our players here...
                                                Dim sPKey As String = "P" & oCol.Owner.ObjectID.ToString
                                                If oItem.colPlayersHere.Contains(sPKey) = False Then
                                                    oItem.colPlayersHere.Add(oCol.Owner, sPKey)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next X

                    'give everyone a chance to breathe
                    Threading.Thread.Sleep(10)

                    'Now, our colList has a list of all systems, and each system has a list of players located there
                    For Each oItem As RespawnListItem In colList
                        For Each oPlayer As Player In oItem.colPlayersHere
                            'Ok, 
                            'get the number of wars for that player and increment oItem.lActiveWars
                            oItem.lActiveWars += oPlayer.GetPlayerWarCount()

                            'Increment our rank value
                            Dim yPTitle As Byte = oPlayer.yPlayerTitle
                            If (yPTitle And Player.PlayerRank.ExRankShift) <> 0 Then yPTitle = yPTitle Xor Player.PlayerRank.ExRankShift
                            oItem.lPlayerCountByRank(yPTitle) += 1
                        Next
                    Next

                    'give everyone a chance to breathe
                    Threading.Thread.Sleep(10)

                    'Now, do our mineral caches
                    For X As Int32 = 0 To glMineralCacheUB
                        Dim oCache As MineralCache = goMineralCache(X)
                        If oCache Is Nothing = False Then
                            If oCache.ParentObject Is Nothing = False AndAlso CType(oCache.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                                Dim oParentPlanet As Planet = CType(oCache.ParentObject, Planet)
                                If oParentPlanet Is Nothing OrElse oParentPlanet.ParentSystem Is Nothing Then Continue For
                                Dim sKey As String = "Sys" & oParentPlanet.ParentSystem.ObjectID.ToString
                                If colList.Contains(sKey) = True Then
                                    Dim oItem As RespawnListItem = CType(colList(sKey), RespawnListItem)
                                    If oItem Is Nothing = False Then
                                        'Now, update our values
                                        oItem.blTotalCacheQty += oCache.Quantity
                                        oItem.lTotalCacheConcentration += oCache.Concentration
                                        oItem.lMinCacheCount += 1
                                    End If
                                End If
                            End If
                        End If
                    Next X

                    'give everyone a chance to breathe
                    Threading.Thread.Sleep(10)

                    Dim yFinal(1000) As Byte
                    Dim lSingleMsgLen As Int32
                    Dim lFinalPos As Int32 = 0

                    'Now, let's compile our data and send it
                    For Each oItem As RespawnListItem In colList
                        If oItem Is Nothing = False Then

                            'Ok, get our planet owner data
                            Dim lPlayerID(-1) As Int32
                            Dim yOwnedCnt(-1) As Byte

                            For X As Int32 = 0 To oItem.oSys.mlPlanetUB
                                Dim lPIdx As Int32 = oItem.oSys.GetPlanetIdx(X)
                                If lPIdx > -1 AndAlso lPIdx <= glPlanetUB Then
                                    Dim oPlanet As Planet = goPlanet(lPIdx)
                                    If oPlanet Is Nothing = False AndAlso oPlanet.OwnerID > -1 Then
                                        Dim bFound As Boolean = False
                                        For Y As Int32 = 0 To lPlayerID.GetUpperBound(0)
                                            If lPlayerID(Y) = oPlanet.OwnerID Then
                                                bFound = True
                                                yOwnedCnt(Y) += CByte(1)
                                                Exit For
                                            End If
                                        Next Y
                                        If bFound = False Then
                                            ReDim Preserve lPlayerID(lPlayerID.GetUpperBound(0) + 1)
                                            ReDim Preserve yOwnedCnt(lPlayerID.GetUpperBound(0))
                                            lPlayerID(lPlayerID.GetUpperBound(0)) = oPlanet.OwnerID
                                            yOwnedCnt(yOwnedCnt.GetUpperBound(0)) = 1
                                        End If
                                    End If
                                End If
                            Next X

                            Dim lPlanetOwnedCnt As Int32 = lPlayerID.GetUpperBound(0) + 1

                            Dim lCurPos As Int32 = 0
                            Dim yCur(50 + (lPlanetOwnedCnt * 5)) As Byte

                            System.BitConverter.GetBytes(GlobalMessageCode.eRespawnSelection).CopyTo(yCur, lCurPos) : lCurPos += 2
                            System.BitConverter.GetBytes(oItem.oSys.ObjectID).CopyTo(yCur, lCurPos) : lCurPos += 4
                            oItem.oSys.SystemName.CopyTo(yCur, lCurPos) : lCurPos += 20

                            Dim lPCnt As Int32 = oItem.oSys.mlPlanetUB + 1
                            If lPCnt > 255 Then lPCnt = 255
                            If lPCnt < 0 Then lPCnt = 0
                            yCur(lCurPos) = CByte(lPCnt) : lCurPos += 1

                            If oItem.lActiveWars > Int16.MaxValue Then oItem.lActiveWars = Int16.MaxValue
                            Dim iWars As Int16 = CShort(oItem.lActiveWars)
                            System.BitConverter.GetBytes(iWars).CopyTo(yCur, lCurPos) : lCurPos += 2    '28

                            System.BitConverter.GetBytes(oItem.lMinCacheCount).CopyTo(yCur, lCurPos) : lCurPos += 4 '32
                            System.BitConverter.GetBytes(oItem.blTotalCacheQty).CopyTo(yCur, lCurPos) : lCurPos += 8 '40
                            Dim lAvg As Int32 = oItem.lTotalCacheConcentration \ oItem.lMinCacheCount
                            If lAvg > 255 Then lAvg = 255
                            If lAvg < 0 Then lAvg = 0
                            yCur(lCurPos) = CByte(lAvg) : lCurPos += 1 '41

                            For Y As Int32 = 0 To 6
                                Dim lTemp As Int32 = oItem.lPlayerCountByRank(Y)
                                If lTemp > 255 Then lTemp = 255
                                If lTemp < 0 Then lTemp = 0
                                yCur(lCurPos) = CByte(lTemp) : lCurPos += 1
                            Next Y
                            '48

                            yCur(lCurPos) = oItem.GetActiveWHLinks() : lCurPos += 1 '49

                            If lPlanetOwnedCnt > 255 Then lPlanetOwnedCnt = 255
                            If lPlanetOwnedCnt < 0 Then lPlanetOwnedCnt = 0
                            yCur(lCurPos) = CByte(lPlanetOwnedCnt) : lCurPos += 1 '50

                            For X As Int32 = 0 To lPlanetOwnedCnt - 1
                                System.BitConverter.GetBytes(lPlayerID(X)).CopyTo(yCur, lCurPos) : lCurPos += 4
                                yCur(lCurPos) = yOwnedCnt(X) : lCurPos += 1
                            Next X

                            'Append this msg to the final
                            lSingleMsgLen = yCur.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lFinalPos + lSingleMsgLen + 2 > yFinal.Length Then
                                ReDim Preserve yFinal(yFinal.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yFinal, lFinalPos)
                            lFinalPos += 2
                            yCur.CopyTo(yFinal, lFinalPos)
                            lFinalPos += lSingleMsgLen

                        End If
                    Next

                    'Now, store this result...
                    If lFinalPos <> 0 Then
                        ReDim myCompiledRespawnList(lFinalPos - 1)
                        Array.Copy(yFinal, 0, myCompiledRespawnList, 0, lFinalPos - 1)
                    End If

                    mlLastRespawnCompileCycle = glCurrentCycle
                End If
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "CompileRespawnList: " & ex.Message)
                mlLastRespawnCompileCycle = 0
                Return Nothing
            End Try
        End SyncLock

        'If myCompiledRespawnList Is Nothing = False Then
        '    'And then send it
        '    Dim col As Collection = gcolRespawnListPlayers
        '    gcolRespawnListPlayers = New Collection
        '    For Each oPlayer As Player In col
        '        If oPlayer Is Nothing = False Then
        '            If oPlayer.oSocket Is Nothing = False Then oPlayer.oSocket.SendLenAppendedData(myCompiledRespawnList)
        '        End If
        '    Next
        'End If
        Return myCompiledRespawnList
    End Function
#End Region


End Module
