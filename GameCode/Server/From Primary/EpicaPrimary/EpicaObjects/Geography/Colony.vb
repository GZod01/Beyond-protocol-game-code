Option Strict On

Public Enum eResourcesResult As Byte
    Insufficient_Clear = 0
    Insufficient_Wait = 1
    Sufficient = 255
End Enum
Public Enum eyHasRequiredResourcesFlags As Byte
    NoFlags = 0
    NoReduction = 1
    DoNotWait = 2
    DoNotAlertLowRes = 4
End Enum
Public Class Colony
    Inherits Epica_GUID

#Region " Constants for Expense Calculations "
    Private Const mlBudgetUpdateTime As Int32 = 17280

    Private Const mfPopUpkeepMult As Single = (16.6F / mlBudgetUpdateTime) * 5
    Private Const mfResearchFacMult As Single = ((70000 * 15) / mlBudgetUpdateTime) * 5
    Private Const mfFactoryMult As Single = ((30000 * 2) / mlBudgetUpdateTime) * 5
    Private Const mfSpaceportMult As Single = ((2300000 * 5) / mlBudgetUpdateTime) * 5
    Private Const mfOtherMult As Single = ((100000 * 1) / mlBudgetUpdateTime) * 5
    Private Const mfTurretMult As Single = (10.0F) * 5

    Private Const mfUnemploymentMult As Single = 0.2F
#End Region

    Private Const mf_SCIENTIST_ANNUAL_PAY As Single = 10.0F
    Private Const mf_FACTORY_ANNUAL_PAY As Single = 4.5F
    Private Const mf_OTHER_ANNUAL_PAY As Single = 2.5F

    Public ColonyName(19) As Byte
    Public ParentObject As Object           'can be either a planet or system or facility
    Public Owner As Player
    Public LocX As Int32
    Public LocY As Int32

    Private mfGrowthRateCalc As Single
    Private mlLastPositiveCycle As Int32
    Public mbSentGNSNews As Boolean = False
    Public ColonyStart As Int32
    Public MaxPopulation As Int32

    Public PowerGeneration As Int32         'was short
    Public PowerConsumption As Int32        'was short
    Public TotalPowerNeeded As Int32
    Public TaxRate As Byte
    Public AverageIncome As Short           'in thousands?
    Public CostOfLiving As Short            'in thousands?
    Public DesireForWar As Byte             '
    Public GovScore As Byte

    Public Population As Int32              'number of colonists here (Total)

    'Based on the children facilities
    Public ColonyEnlisted As Int32          'total number of enlisted personnel available for production at this colony
    Public ColonyOfficers As Int32          'total number of officers available for production at this colony

    'TechLevels
    Public PowerLevel As Byte
    Public ElectronicsLevel As Byte
    Public MaterialsLevel As Byte

    Public oChildren() As Facility          'pointers, NOT actuals
    Public ChildrenUB As Int32 = -1
    Public lChildrenIdx() As Int32

    Private mbInUpdatePowerNeeds As Boolean = False

    Public WorkforceEfficiency As Single = 0.0F
    Public NumberOfJobs As Int32 = 0
    Public PoweredHousing As Int32 = 0
    Public UnpoweredHousing As Int32 = 0
    Public TaxableWorkforce As Int32 = 0
    Public MoraleMultiplier As Single = 1.0F
    Public UnemploymentRate As Byte = 0
    Public ColonyGrowthRate As Int32 = 0

    Public ScientistJobs As Int32 = 0
    Public FactoryJobs As Int32 = 0
    Public OtherJobs As Int32 = 0

    Public iControlledMorale As Int16 = 0
    Public iControlledGrowth As Int16 = 0

    Private mySendString() As Byte

    Public bCCInProduction As Boolean = False           'indicates that a command center is in production
    Public bTradepostInProduction As Boolean = False    'indicates that a tradepost is in production

    Private mfLastMoraleMultiplier As Single = 0.0F

#Region "  Colony Cache Management  "
    Public mlMineralCacheIdx() As Int32         'server index to the global glMineralCacheIdx()
    Public mlMineralCacheID() As Int32          'for verification that idx matches to the right id
    Public mlMineralCacheMineralID() As Int32   'the mineral ID for the mineral cache
    Public mlMineralCacheUB As Int32 = -1

    Public mlComponentCacheIdx() As Int32       'server index to the global glcomponentcacheidx()
    Public mlComponentCacheID() As Int32        'for verification that IDX matches to the right ID
    Public mlComponentCacheCompID() As Int32    'the component ID for the component cache
    Public mlComponentCacheUB As Int32 = -1

    Public ReadOnly Property TotalCargoCapAvailable() As Int32
        Get
            Dim lFacilityAvail As Int32 = 0

            Try
                For X As Int32 = 0 To ChildrenUB
                    If oChildren(X) Is Nothing = False AndAlso ProductionTypeSharesColonyCargo(oChildren(X).yProductionType) = True AndAlso oChildren(X).Active = True Then
                        If (oChildren(X).CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then lFacilityAvail += oChildren(X).Cargo_Cap
                    End If
                Next X

                For X As Int32 = 0 To mlAgentEffectUB
                    If EffectValid(X) = True Then
                        If moAgentEffects(X).yType = AgentEffectType.eCargoBay Then
                            lFacilityAvail = CInt(lFacilityAvail * (moAgentEffects(X).lAmount / 100.0F))
                        End If
                    End If
                Next X
            Catch exOverflow As OverflowException
                lFacilityAvail = Int32.MaxValue
            Catch ex As Exception

            End Try

            Return lFacilityAvail
        End Get
    End Property

    Public ReadOnly Property TotalCargoCapUsed() As Int32
        Get
            Dim lMineralCacheUsed As Int32 = 0
            Dim lComponentCacheUsed As Int32 = 0

            Try
                For X As Int32 = 0 To mlMineralCacheUB
                    Dim lCacheIdx As Int32 = mlMineralCacheIdx(X)
                    If lCacheIdx > -1 Then
                        If glMineralCacheIdx(lCacheIdx) = mlMineralCacheID(X) Then
                            If goMineralCache(lCacheIdx) Is Nothing = False Then lMineralCacheUsed += goMineralCache(lCacheIdx).Quantity
                        Else
                            mlMineralCacheIdx(X) = -1
                        End If
                    End If
                Next X

                For X As Int32 = 0 To mlComponentCacheUB
                    Dim lCacheIdx As Int32 = mlComponentCacheIdx(X)
                    If lCacheIdx > -1 Then
                        If glComponentCacheIdx(lCacheIdx) = mlComponentCacheID(X) Then
                            If goComponentCache(lCacheIdx) Is Nothing = False Then lComponentCacheUsed += goComponentCache(lCacheIdx).Quantity
                        Else
                            mlComponentCacheIdx(X) = -1
                        End If
                    End If
                Next X
            Catch exOverflow As OverflowException
                Return Int32.MaxValue
            Catch ex As Exception

            End Try

            Return lMineralCacheUsed + lComponentCacheUsed
        End Get
    End Property

    Public ReadOnly Property Cargo_Cap() As Int32
        Get
            Dim lFacilityAvail As Int32 = TotalCargoCapAvailable
            Dim lCacheUsed As Int32 = TotalCargoCapUsed
            Return lFacilityAvail - lCacheUsed
        End Get
    End Property

    Public Function AdjustColonyMineralCache(ByVal lMineralID As Int32, ByVal lQty As Int32) As MineralCache
        Dim lCurUB As Int32 = -1
        If mlMineralCacheIdx Is Nothing = False Then lCurUB = Math.Min(mlMineralCacheUB, mlMineralCacheIdx.GetUpperBound(0))
        If mlMineralCacheID Is Nothing = False Then lCurUB = Math.Min(lCurUB, mlMineralCacheID.GetUpperBound(0))

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lCurUB
            If mlMineralCacheIdx(X) > -1 Then
                If glMineralCacheIdx(mlMineralCacheIdx(X)) = mlMineralCacheID(X) Then
                    If mlMineralCacheMineralID(X) = lMineralID Then
                        Dim oCache As MineralCache = goMineralCache(mlMineralCacheIdx(X))
                        If oCache Is Nothing = False Then
                            oCache.Quantity += lQty

                            If lMineralID = 41991 AndAlso oCache.Quantity > 487 Then
                                If Me.ParentObject Is Nothing = False AndAlso CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                                    CType(Me.ParentObject, Planet).MicroExplosion()
                                End If
                            End If

                            Return oCache
                        End If
                    End If
                Else
                    mlMineralCacheIdx(X) = -1
                End If
            ElseIf lIdx = -1 Then
                lIdx = X
            End If
        Next X

        'Ok, we're here, we need to add the mineral cache
        Dim lCacheIdx As Int32 = AddMineralCache(Me.ObjectID, Me.ObjTypeID, MineralCacheType.eGround, lQty, lQty, 0, 0, GetEpicaMineral(lMineralID))
        If lCacheIdx > -1 Then
            If lIdx = -1 Then
                lIdx = mlMineralCacheUB + 1
                ReDim Preserve mlMineralCacheID(lIdx)
                ReDim Preserve mlMineralCacheIdx(lIdx)
                ReDim Preserve mlMineralCacheMineralID(lIdx)
                mlMineralCacheUB = lIdx
            End If
            mlMineralCacheIdx(lIdx) = lCacheIdx
            mlMineralCacheID(lIdx) = glMineralCacheIdx(lCacheIdx)
            mlMineralCacheMineralID(lIdx) = lMineralID
            Return goMineralCache(lCacheIdx)
        End If
        Return Nothing
    End Function
    Public Function AdjustColonyComponentCache(ByVal lComponentID As Int32, ByVal iComponentTypeID As Int16, ByVal lComponentOwnerID As Int32, ByVal lQty As Int32) As ComponentCache
        Dim lCurUB As Int32 = -1
        If mlComponentCacheIdx Is Nothing = False Then lCurUB = Math.Min(mlComponentCacheUB, mlComponentCacheIdx.GetUpperBound(0))
        If mlComponentCacheID Is Nothing = False Then lCurUB = Math.Min(lCurUB, mlComponentCacheID.GetUpperBound(0))

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lCurUB
            If mlComponentCacheIdx(X) > -1 Then
                If glComponentCacheIdx(mlComponentCacheIdx(X)) = mlComponentCacheID(X) Then
                    If mlComponentCacheCompID(X) = lComponentID Then
                        Dim oCache As ComponentCache = goComponentCache(mlComponentCacheIdx(X))
                        If oCache Is Nothing = False AndAlso oCache.ComponentID = lComponentID AndAlso oCache.ComponentTypeID = iComponentTypeID Then
                            oCache.Quantity += lQty
                            Return oCache
                        End If
                    End If
                Else
                    mlComponentCacheIdx(X) = -1
                End If
            ElseIf lIdx = -1 Then
                lIdx = X
            End If
        Next X

        'Ok, we're here, we need to add the mineral cache
        If lComponentOwnerID <> Me.Owner.ObjectID AndAlso lComponentOwnerID > 0 AndAlso Me.Owner.HasTechKnowledge(lComponentID, iComponentTypeID, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2) = False Then
            Dim oP As Player = GetEpicaPlayer(lComponentOwnerID)
            If oP Is Nothing = False Then
                Dim oTech As Epica_Tech = oP.GetTech(lComponentID, iComponentTypeID)
                If oTech Is Nothing Then Return Nothing
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(Me.Owner, oTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2, False)
                Me.Owner.SendPlayerMessage(oPTK.GetAddMsg(), False, 0)
            End If
        End If

        Dim lCacheIdx As Int32 = AddComponentCache(Me.ObjectID, Me.ObjTypeID, lQty, 0, 0, lComponentID, iComponentTypeID, lComponentOwnerID, 0)
        If lCacheIdx > -1 Then
            If lIdx = -1 Then
                lIdx = mlComponentCacheUB + 1
                ReDim Preserve mlComponentCacheID(lIdx)
                ReDim Preserve mlComponentCacheIdx(lIdx)
                ReDim Preserve mlComponentCacheCompID(lIdx)
                mlComponentCacheUB = lIdx
            End If
            mlComponentCacheIdx(lIdx) = lCacheIdx
            mlComponentCacheID(lIdx) = glComponentCacheIdx(lCacheIdx)
            mlComponentCacheCompID(lIdx) = lComponentID
            Return goComponentCache(lCacheIdx)
        End If
        Return Nothing
    End Function

    Public Function ColonyMineralQuantity(ByVal lMineralID As Int32) As Int32
        Dim lCurUB As Int32 = -1
        If mlMineralCacheIdx Is Nothing = False Then lCurUB = Math.Min(mlMineralCacheUB, mlMineralCacheIdx.GetUpperBound(0))
        If mlMineralCacheID Is Nothing = False Then lCurUB = Math.Min(lCurUB, mlMineralCacheID.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If mlMineralCacheMineralID(X) = lMineralID AndAlso mlMineralCacheIdx(X) > -1 Then
                If glMineralCacheIdx(mlMineralCacheIdx(X)) = mlMineralCacheID(X) Then
                    Return goMineralCache(mlMineralCacheIdx(X)).Quantity
                End If
            End If
        Next X
        Return 0
    End Function

    Public Function ColonyComponentCache(ByVal lComponentID As Int32, ByVal iComponentTypeID As Int16) As ComponentCache
        Dim lCurUB As Int32 = -1
        If mlComponentCacheIdx Is Nothing = False Then lCurUB = Math.Min(mlComponentCacheUB, mlComponentCacheIdx.GetUpperBound(0))
        If mlComponentCacheID Is Nothing = False Then lCurUB = Math.Min(mlComponentCacheUB, mlComponentCacheID.GetUpperBound(0))
        If mlComponentCacheCompID Is Nothing = False Then lCurUB = Math.Min(mlComponentCacheUB, mlComponentCacheCompID.GetUpperBound(0))

        For X As Int32 = 0 To lCurUB
            If mlComponentCacheCompID(X) = lComponentID And mlComponentCacheIdx(X) > -1 Then
                If glComponentCacheIdx(mlComponentCacheIdx(X)) = mlComponentCacheID(X) Then
                    Dim oCC As ComponentCache = goComponentCache(mlComponentCacheIdx(X))
                    If oCC Is Nothing = False Then
                        If oCC.ComponentID = lComponentID AndAlso oCC.ComponentTypeID = iComponentTypeID Then
                            Return oCC
                        End If
                    End If
                End If
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetPerFacilityCargoUsed(ByVal lBaseCargoCap As Int32) As Int32
        Dim lFacilityAvail As Int32 = TotalCargoCapAvailable
        Dim lMineralCacheUsed As Int32 = TotalCargoCapUsed
        If lFacilityAvail = 0 Then Return 0
        Dim fUsage As Single = CSng(lMineralCacheUsed) / CSng(lFacilityAvail)
        If fUsage < 0 Then fUsage = 0
        If fUsage > 1 Then fUsage = 1
        Return CInt(lBaseCargoCap * fUsage)
    End Function

    Public Sub CargoFacilityDestroyed(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal lOriginalCargoCap As Int32)
        Dim lFacilityAvail As Int32 = TotalCargoCapAvailable
        Dim lMineralCacheUsed As Int32 = TotalCargoCapUsed

        If lFacilityAvail >= lMineralCacheUsed OrElse lMineralCacheUsed < 1 Then Return
        Dim lToRemove As Int32 = lMineralCacheUsed - lFacilityAvail
        Dim fRemoval As Single = CSng(lToRemove) / CSng(lMineralCacheUsed)
        If fRemoval < 0 Then fRemoval = 0
        If fRemoval > 1 Then fRemoval = 1

        Dim lCurUB As Int32 = -1
        If mlMineralCacheIdx Is Nothing = False Then lCurUB = Math.Min(mlMineralCacheUB, mlMineralCacheIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If mlMineralCacheIdx(X) > -1 Then
                If glMineralCacheIdx(mlMineralCacheIdx(X)) = mlMineralCacheID(X) Then
                    Dim oCache As MineralCache = goMineralCache(mlMineralCacheIdx(X))
                    If oCache Is Nothing = False Then
                        Dim lQty As Int32 = CInt(oCache.Quantity * fRemoval)
                        oCache.Quantity -= lQty
                    End If
                Else
                    mlMineralCacheIdx(X) = -1
                End If
            End If
        Next X
        lCurUB = -1
        If mlComponentCacheIdx Is Nothing = False Then lCurUB = Math.Min(mlComponentCacheUB, mlComponentCacheIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If mlComponentCacheIdx(X) > -1 Then
                If glComponentCacheIdx(mlComponentCacheIdx(X)) = mlComponentCacheID(X) Then
                    Dim oCache As ComponentCache = goComponentCache(mlComponentCacheIdx(X))
                    If oCache Is Nothing = False Then
                        Dim lQty As Int32 = CInt(oCache.Quantity * fRemoval)
                        oCache.Quantity -= lQty
                    End If
                Else
                    mlComponentCacheIdx(X) = -1
                End If
            End If
        Next X

    End Sub
#End Region

    Public Function GetObjAsString() As Byte()
        'here we will return the object populated as a string
        If mbStringReady = False Then
            'guid-6
            'name-20
            'parentid-4
            'parenttype-2
            'ownerid-4
            'locx-1
            'locy-1
            'morale-1
            'growthrate-1
            'powergeneration-2
            'powerconsumption-2
            'taxrate-1
            'averageincome-2
            'costofliving-2
            'desireforwar-1
            'govscore-1
            'powerlevel-1
            'electronicslevel-1
            'materialslevel-1
            'Population
            ReDim mySendString(57)      '0 to 57 = 58 bytes
            GetGUIDAsString.CopyTo(mySendString, 0)
            ColonyName.CopyTo(mySendString, 6)
            CType(ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, 26)
            System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySendString, 32)
            mySendString(36) = CByte(LocX)
            mySendString(37) = CByte(LocY)
            mySendString(38) = 0        'TODO: No longer used, Morale
            mySendString(39) = 0        'TODO: No longer used, growth rate
            System.BitConverter.GetBytes(PowerGeneration).CopyTo(mySendString, 40)
            System.BitConverter.GetBytes(PowerConsumption).CopyTo(mySendString, 42)
            mySendString(44) = TaxRate
            System.BitConverter.GetBytes(AverageIncome).CopyTo(mySendString, 45)
            System.BitConverter.GetBytes(CostOfLiving).CopyTo(mySendString, 47)
            mySendString(49) = DesireForWar
            mySendString(50) = GovScore
            mySendString(51) = PowerLevel
            mySendString(52) = ElectronicsLevel
            mySendString(53) = MaterialsLevel
            System.BitConverter.GetBytes(Population).CopyTo(mySendString, 54)
            mbStringReady = True
        End If
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
                sSQL = "INSERT INTO tblColony (ColonyName, ParentID, ParentTypeID, OwnerID, LocX, LocY, Morale, " & _
                  "GrowthRate, PowerGeneration, PowerConsumption, TaxRate, AverageIncome, PowerLevel, ElectronicsLevel, " & _
                  "MaterialsLevel, CostOfLiving, DesireForWar, GovScore, Population, CurrentEnlisted, CurrentOfficers, ControlledMorale, " & _
                  "ControlledGrowth, ColonyStarted, MaxPopulation) VALUES ('" & _
                  MakeDBStr(BytesToString(ColonyName)) & "', " & CType(ParentObject, Epica_GUID).ObjectID & ", " & CType(ParentObject, Epica_GUID).ObjTypeID & ", "
                If Owner Is Nothing = False Then
                    sSQL = sSQL & Owner.ObjectID & ", "
                Else
                    sSQL = sSQL & "-1, "
                End If
                sSQL = sSQL & LocX & ", " & LocY & ", 0, 0, " & PowerGeneration & _
                  ", " & PowerConsumption & ", " & TaxRate & ", " & AverageIncome & ", " & PowerLevel & ", " & _
                  ElectronicsLevel & ", " & MaterialsLevel & ", " & CostOfLiving & ", " & DesireForWar & ", " & _
                  GovScore & ", " & Population & ", " & ColonyEnlisted & ", " & ColonyOfficers & ", " & iControlledMorale & ", " & _
                  iControlledGrowth & ", " & ColonyStart & ", " & MaxPopulation & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblColony SET ColonyName = '" & MakeDBStr(BytesToString(ColonyName)) & "', ParentID = " & CType(ParentObject, Epica_GUID).ObjectID & _
                  ", ParentTypeID = " & CType(ParentObject, Epica_GUID).ObjTypeID & ", LocX = " & LocX & ", LocY = " & LocY & ", Morale = 0" & _
                   ", GrowthRate = 0, PowerGeneration = " & PowerGeneration & _
                  ", PowerConsumption = " & PowerConsumption & ", TaxRate = " & TaxRate & ", AverageIncome = " & _
                  AverageIncome & ", PowerLevel = " & PowerLevel & ", ElectronicsLevel = " & ElectronicsLevel & _
                  ", MaterialsLevel = " & MaterialsLevel & ", CostOfLiving = " & CostOfLiving & ", DesireForWar=" & _
                  DesireForWar & ", GovScore = " & GovScore & ", Population = " & Population & ", CurrentEnlisted = " & ColonyEnlisted & _
                  ", CurrentOfficers = " & ColonyOfficers & ", ControlledMorale = " & iControlledMorale & ", ControlledGrowth = " & _
                  iControlledGrowth & ", ColonyStarted = " & ColonyStart & ", MaxPopulation = " & MaxPopulation
                If Owner Is Nothing = False Then
                    sSQL = sSQL & ", OwnerID = " & Owner.ObjectID
                Else
                    sSQL = sSQL & ", OwnerID = -1"
                End If
                sSQL = sSQL & " WHERE ColonyID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(ColonyID) FROM tblColony WHERE ColonyName = '" & MakeDBStr(BytesToString(ColonyName)) & "'"
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

    Public Function GetSaveObjectText() As String
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oSB As New System.Text.StringBuilder

        If ObjectID = -1 Then
            If SaveObject() = False Then Return ""
        End If

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblColony (ColonyName, ParentID, ParentTypeID, OwnerID, LocX, LocY, Morale, " & _
                  "GrowthRate, PowerGeneration, PowerConsumption, TaxRate, AverageIncome, PowerLevel, ElectronicsLevel, " & _
                  "MaterialsLevel, CostOfLiving, DesireForWar, GovScore, Population, CurrentEnlisted, CurrentOfficers, ControlledMorale, " & _
                  "ControlledGrowth, ColonyStarted, MaxPopulation) VALUES ('" & _
                  MakeDBStr(BytesToString(ColonyName)) & "', " & CType(ParentObject, Epica_GUID).ObjectID & ", " & CType(ParentObject, Epica_GUID).ObjTypeID & ", "
                If Owner Is Nothing = False Then
                    sSQL = sSQL & Owner.ObjectID & ", "
                Else
                    sSQL = sSQL & "-1, "
                End If
                sSQL = sSQL & LocX & ", " & LocY & ", 0, 0, " & PowerGeneration & _
                  ", " & PowerConsumption & ", " & TaxRate & ", " & AverageIncome & ", " & PowerLevel & ", " & _
                  ElectronicsLevel & ", " & MaterialsLevel & ", " & CostOfLiving & ", " & DesireForWar & ", " & _
                  GovScore & ", " & Population & ", " & ColonyEnlisted & ", " & ColonyOfficers & ", " & iControlledMorale & ", " & _
                  iControlledGrowth & ", " & ColonyStart & ", " & MaxPopulation & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblColony SET ColonyName = '" & MakeDBStr(BytesToString(ColonyName)) & "', ParentID = " & CType(ParentObject, Epica_GUID).ObjectID & _
                  ", ParentTypeID = " & CType(ParentObject, Epica_GUID).ObjTypeID & ", LocX = " & LocX & ", LocY = " & LocY & ", Morale = 0" & _
                   ", GrowthRate = 0, PowerGeneration = " & PowerGeneration & _
                  ", PowerConsumption = " & PowerConsumption & ", TaxRate = " & TaxRate & ", AverageIncome = " & _
                  AverageIncome & ", PowerLevel = " & PowerLevel & ", ElectronicsLevel = " & ElectronicsLevel & _
                  ", MaterialsLevel = " & MaterialsLevel & ", CostOfLiving = " & CostOfLiving & ", DesireForWar=" & _
                  DesireForWar & ", GovScore = " & GovScore & ", Population = " & Population & ", CurrentEnlisted = " & ColonyEnlisted & _
                  ", CurrentOfficers = " & ColonyOfficers & ", ControlledMorale = " & iControlledMorale & ", ControlledGrowth = " & _
                  iControlledGrowth & ", ColonyStarted = " & ColonyStart & ", MaxPopulation = " & MaxPopulation
                If Owner Is Nothing = False Then
                    sSQL = sSQL & ", OwnerID = " & Owner.ObjectID
                Else
                    sSQL = sSQL & ", OwnerID = -1"
                End If
                sSQL = sSQL & " WHERE ColonyID = " & ObjectID
            End If
            oSB.AppendLine(sSQL)
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return oSB.ToString
    End Function

    Public Sub AddChildFacility(ByRef oFacility As Facility)
        Dim X As Int32
        Dim lIdx As Int32 = -1

        For X = 0 To ChildrenUB
            If lChildrenIdx(X) = -1 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            ChildrenUB += 1
            ReDim Preserve oChildren(ChildrenUB)
            ReDim Preserve lChildrenIdx(ChildrenUB)
            lIdx = ChildrenUB
        End If

        oChildren(lIdx) = oFacility
        lChildrenIdx(lIdx) = oFacility.ObjectID
        oFacility.ColonyArrayIndex = lIdx

        TotalPowerNeeded += oFacility.EntityDef.PowerFactor
    End Sub

    Public Shared Function GetProdCostItemDBCost(ByVal oCost As ProductionCostItem, ByRef oPlayer As Player) As Int64
        'ocost.ItemID
        'ocost.ItemTypeID
        'ocost.QuantityNeeded

        Dim blResult As Int64 = oCost.QuantityNeeded

        Select Case oCost.ItemTypeID
            Case ObjectType.eMineral
                Dim oMineral As Mineral = GetEpicaMineral(oCost.ItemID)
                If oMineral Is Nothing = False Then
                    If oMineral.lAlloyTechID > 0 Then
                        'ok, get that alloy
                        Dim oAlloy As AlloyTech = CType(oPlayer.GetTech(oMineral.lAlloyTechID, ObjectType.eAlloyTech), AlloyTech)
                        If oAlloy Is Nothing = False Then
                            'ok, get those mineral costs...
                            Dim oProdCost As ProductionCost = oAlloy.GetProductionCost
                            If oProdCost Is Nothing = False Then
                                Dim blTemp As Int64 = 0
                                For X As Int32 = 0 To oProdCost.ItemCostUB
                                    blTemp += GetProdCostItemDBCost(oProdCost.ItemCosts(X), oPlayer)
                                Next X
                                blResult = blTemp * CLng(oCost.QuantityNeeded)
                            Else
                                blResult = oMineral.MineralValue * oCost.QuantityNeeded
                            End If
                        Else
                            blResult = oMineral.MineralValue * oCost.QuantityNeeded
                        End If
                    Else
                        blResult = oMineral.MineralValue * oCost.QuantityNeeded
                    End If
                End If
            Case Else
                Dim oTech As Epica_Tech = oPlayer.GetTech(oCost.ItemID, oCost.ItemTypeID)
                If oTech Is Nothing = False Then
                    Dim oProdCost As ProductionCost = oTech.GetProductionCost()
                    If oProdCost Is Nothing = False Then

                        Dim blTemp As Int64 = 0
                        For X As Int32 = 0 To oProdCost.ItemCostUB
                            blTemp += GetProdCostItemDBCost(oProdCost.ItemCosts(X), oPlayer)
                        Next X

                        blTemp += oProdCost.CreditCost
                        blTemp += oProdCost.EnlistedCost
                        blTemp += oProdCost.OfficerCost
                        blTemp += oProdCost.ColonistCost

                        blResult = blTemp * CLng(oCost.QuantityNeeded)
                    End If
                End If
        End Select

        'if all else fails, return the quantity needed
        Return blResult

    End Function

    Public Function HasRequiredResources(ByRef oProdCost As ProductionCost, ByRef oFacility As Facility, ByVal yFlags As eyHasRequiredResourcesFlags) As eResourcesResult
        Dim blCreditCost As Int64 = oProdCost.CreditCost

        If oFacility Is Nothing = False Then
            If oFacility.yProductionType = ProductionType.eResearch Then
                yFlags = yFlags Or eyHasRequiredResourcesFlags.DoNotWait
            End If
        End If

        If (Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase AndAlso Owner.lTutorialStep > 230 AndAlso Owner.lTutorialStep < 285) OrElse gb_IS_TEST_SERVER = True Then Return eResourcesResult.Sufficient

        'If we are BUILDING a component, ignore personnel costs (Colonists, Officers and Enlisted)
        Dim bIgnorePersonnel As Boolean = (oProdCost.ProductionCostType = 0) AndAlso (oProdCost.ObjTypeID = ObjectType.eArmorTech OrElse oProdCost.ObjTypeID = ObjectType.eEngineTech OrElse oProdCost.ObjTypeID = ObjectType.eHullTech OrElse oProdCost.ObjTypeID = ObjectType.eRadarTech OrElse oProdCost.ObjTypeID = ObjectType.eShieldTech OrElse oProdCost.ObjTypeID = ObjectType.eWeaponTech)

        If Owner.DeathBudgetEndTime > glCurrentCycle AndAlso oProdCost.ProductionCostType = 0 Then '  oProdCost.ObjTypeID <> ObjectType.eSpecialTech Then
            'Ok, player will pay with credits for everything
            With oProdCost
                If bIgnorePersonnel = False Then blCreditCost += .EnlistedCost + .OfficerCost + .ColonistCost
                For X As Int32 = 0 To .ItemCostUB
                    blCreditCost += GetProdCostItemDBCost(.ItemCosts(X), Me.Owner)  '.ItemCosts(X).QuantityNeeded
                Next X
            End With

            If Owner.DeathBudgetFundsRemaining >= blCreditCost Then
                If Owner.blCredits >= blCreditCost Then
                    Owner.DeathBudgetFundsRemaining -= CInt(blCreditCost)
                    Owner.blCredits -= CInt(blCreditCost)
                    Return eResourcesResult.Sufficient
                End If
            Else
                blCreditCost -= Owner.DeathBudgetFundsRemaining
                Owner.DeathBudgetFundsRemaining = 0
                Owner.DeathBudgetEndTime = glCurrentCycle
            End If
        End If

        With oProdCost
            If bIgnorePersonnel = False Then
                'enough enlisted?
                If ColonyEnlisted < .EnlistedCost Then
                    Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eEnlisted, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                    Return eResourcesResult.Insufficient_Clear
                End If

                'enough officers?
                If ColonyOfficers < .OfficerCost Then
                    Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eOfficers, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                    Return eResourcesResult.Insufficient_Clear
                End If

                'enough colonists?
                If Population < .ColonistCost Then
                    Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eColonists, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                    Return eResourcesResult.Insufficient_Clear
                End If
            End If

            'Does the player have enough credits?
            'If Owner.blCredits < .CreditCost Then
            If Owner.blCredits < blCreditCost AndAlso blCreditCost > 0 Then
                Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eCredits, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                Return eResourcesResult.Insufficient_Clear
            End If
        End With

        Dim lItemID(oProdCost.ItemCostUB) As Int32
        Dim iTypeID(oProdCost.ItemCostUB) As Int16
        Dim lQty(oProdCost.ItemCostUB) As Int32

        'go ahead and give enough room for all of the prod cost items...
        Dim uResSrc(oProdCost.ItemCostUB) As ResourceSource
        Dim lResSrcUB As Int32 = -1
        Dim iCacheTypeID As Int16

        Dim lSpaceRequired As Int32 = 1
        'If oProdCost.ProductionCostType = 1 Then lSpaceRequired = 0
        If oProdCost.ObjTypeID = ObjectType.eUnitDef AndAlso (oFacility.yProductionType And ProductionType.eNavalProduction) <> ProductionType.eNavalProduction Then
            Dim oUnitDef As Epica_Entity_Def = GetEpicaUnitDef(oProdCost.ObjectID)
            If oUnitDef Is Nothing = False Then lSpaceRequired = oUnitDef.HullSize

            Dim lHangar_Cap As Int32 = 0
            If oFacility.ParentObject Is Nothing = False AndAlso CType(oFacility.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                With CType(oFacility.ParentObject, Facility)
                    lHangar_Cap = .Hangar_Cap
                    If .EntityDef.HasHangarDoorSize(lSpaceRequired) = False Then
                        Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eProduction, -1, ObjectType.eHangarTech, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                        Return eResourcesResult.Insufficient_Clear
                    End If

                    If lSpaceRequired > lHangar_Cap Then
                        Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eProduction, 0, ObjectType.eHangarTech, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                        Return eResourcesResult.Insufficient_Clear
                    End If
                End With
            Else
                lHangar_Cap = oFacility.Hangar_Cap
                If oFacility.EntityDef.HasHangarDoorSize(lSpaceRequired) = False Then
                    Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eProduction, -1, ObjectType.eHangarTech, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                    Return eResourcesResult.Insufficient_Clear
                End If

                If lSpaceRequired > oFacility.Hangar_Cap Then
                    Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eProduction, 0, ObjectType.eHangarTech, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                    Return eResourcesResult.Insufficient_Clear
                End If
            End If

            lSpaceRequired = 0
        ElseIf oProdCost.ObjTypeID = ObjectType.eMineral Then
            If Me.Cargo_Cap < 10 Then
                Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eWareHouse, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                Return eResourcesResult.Insufficient_Clear
            End If
            lSpaceRequired = 0
        End If

        For X As Int32 = 0 To oProdCost.ItemCostUB
            lItemID(X) = oProdCost.ItemCosts(X).ItemID
            iTypeID(X) = oProdCost.ItemCosts(X).ItemTypeID
            lQty(X) = oProdCost.ItemCosts(X).QuantityNeeded

            lSpaceRequired -= lQty(X)
        Next X
        If (lSpaceRequired > 1 OrElse (oProdCost.ObjTypeID = ObjectType.eArmorTech AndAlso oFacility.yProductionType <> ProductionType.eResearch)) AndAlso Me.GetFacilityWithCargoSpace(lSpaceRequired) Is Nothing = True Then
            Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eWareHouse, 0, 0, oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
            Return eResourcesResult.Insufficient_Clear
        End If

        'hit the colony's global array first
        Dim lCurUB As Int32 = -1
        If mlMineralCacheIdx Is Nothing = False Then lCurUB = Math.Min(mlMineralCacheUB, mlMineralCacheIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If mlMineralCacheIdx(X) > -1 Then
                If mlMineralCacheID(X) = glMineralCacheIdx(mlMineralCacheIdx(X)) Then
                    Dim oCache As MineralCache = goMineralCache(mlMineralCacheIdx(X))
                    If oCache Is Nothing = False Then
                        For Y As Int32 = 0 To oProdCost.ItemCostUB
                            If lQty(Y) > 0 AndAlso lItemID(Y) = oCache.oMineral.ObjectID AndAlso iTypeID(Y) = ObjectType.eMineral Then
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
                    End If
                End If
            End If
        Next X
        lCurUB = -1
        If mlComponentCacheIdx Is Nothing = False Then lCurUB = Math.Min(mlComponentCacheUB, mlComponentCacheIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If mlComponentCacheIdx(X) > -1 Then
                If mlComponentCacheID(X) = glComponentCacheIdx(mlComponentCacheIdx(X)) Then
                    Dim oCache As ComponentCache = goComponentCache(mlComponentCacheIdx(X))
                    If oCache Is Nothing = False Then
                        For Y As Int32 = 0 To oProdCost.ItemCostUB
                            If lQty(Y) > 0 AndAlso lItemID(Y) = oCache.ComponentID AndAlso oCache.ComponentTypeID = iTypeID(Y) Then
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
                    End If
                End If
            End If
        Next X

        'Now... loop through... our facilities and find the items...
        For X As Int32 = 0 To ChildrenUB
            If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso (oChildren(X).yProductionType <> ProductionType.eMining) Then 'AndAlso (oChildren(X).yProductionType <> ProductionType.eTradePost) Then
                For lCIdx As Int32 = 0 To oChildren(X).lCargoUB
                    If oChildren(X).lCargoIdx(lCIdx) <> -1 AndAlso oChildren(X).oCargoContents(lCIdx).ObjTypeID <> ObjectType.eAmmunition Then
                        Dim lCargoID As Int32
                        Dim iCargoTypeID As Int16
                        Dim lCargoQty As Int32

                        iCacheTypeID = oChildren(X).oCargoContents(lCIdx).ObjTypeID

                        If iCacheTypeID = ObjectType.eMineralCache Then
                            With CType(oChildren(X).oCargoContents(lCIdx), MineralCache)
                                lCargoID = .oMineral.ObjectID
                                iCargoTypeID = ObjectType.eMineral
                                lCargoQty = .Quantity
                            End With
                        ElseIf iCacheTypeID = ObjectType.eComponentCache Then
                            With CType(oChildren(X).oCargoContents(lCIdx), ComponentCache)
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

                Dim bGood As Boolean = True
                For Y As Int32 = 0 To oProdCost.ItemCostUB
                    If lQty(Y) <> 0 Then
                        bGood = False
                        Exit For
                    End If
                Next Y

                If bGood = True Then Exit For
            End If
        Next X

        If (yFlags And eyHasRequiredResourcesFlags.NoReduction) <> 0 Then
            Dim yRes As eResourcesResult = eResourcesResult.Sufficient
            For X As Int32 = 0 To oProdCost.ItemCostUB
                If lQty(X) <> 0 Then
                    Select Case oProdCost.ItemCosts(X).ItemTypeID
                        Case ObjectType.eMineral, ObjectType.eMineralCache
                            yRes = eResourcesResult.Insufficient_Wait
                        Case Else
                            Return eResourcesResult.Insufficient_Clear
                    End Select
                End If
            Next X
            Return yRes
        End If

        'Now... go through our items...
        Dim bInserted As Boolean = False
        For X As Int32 = 0 To oProdCost.ItemCostUB
            If lQty(X) <> 0 Then
                If iTypeID(X) <> ObjectType.eMineral AndAlso iTypeID(X) <> ObjectType.eMineralCache Then
                    'Ok, insert this production
                    If oFacility Is Nothing OrElse ((oFacility.yProductionType And ProductionType.eProduction) = 0) OrElse oFacility.InsertProduction(lItemID(X), iTypeID(X), lQty(X)) = False Then
                        'Just break out... after telling the facility to begin production if we inserted...
                        If bInserted = True Then oFacility.GetNextProduction(False)
                        Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eMaterials, lItemID(X), iTypeID(X), oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                        Return eResourcesResult.Insufficient_Clear
                    Else : bInserted = True
                    End If
                Else

                    If oFacility.yProductionType = ProductionType.eResearch AndAlso oProdCost.ObjTypeID = ObjectType.eSpecialTech Then Return eResourcesResult.Insufficient_Clear

                    'ok, is the material an alloy
                    Dim oTmpMin As Mineral = GetEpicaMineral(lItemID(X))
                    'Ok, is our facility refinery
                    'If oTmpMin Is Nothing OrElse oTmpMin.lAlloyTechID < 1 OrElse oFacility Is Nothing OrElse ((oFacility.yProductionType And ProductionType.eRefining) = 0) OrElse oFacility.InsertProduction(lItemID(X), ObjectType.eMineral, CInt(Math.Ceiling(lQty(X) / 10.0F))) = False Then
                    If oTmpMin Is Nothing OrElse oTmpMin.lAlloyTechID < 1 OrElse oFacility Is Nothing OrElse oFacility.InsertProduction(lItemID(X), ObjectType.eMineral, CInt(Math.Ceiling(lQty(X) / 10.0F))) = False Then
                        If bInserted = True Then oFacility.GetNextProduction(False)
                        'Need to check if this is from the queue, if it is, do not send the low resources msg
                        If (yFlags And eyHasRequiredResourcesFlags.DoNotAlertLowRes) = 0 Then
                            Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eMaterials, lItemID(X), iTypeID(X), oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                        End If
                        If (yFlags And eyHasRequiredResourcesFlags.DoNotWait) <> 0 Then
                            Return eResourcesResult.Insufficient_Clear
                        Else
                            'Now, ensure that I am not producing
                            oFacility.bProducing = False
                            If oFacility.CurrentProduction Is Nothing = False Then oFacility.CurrentProduction.bPaidFor = False
                            AddToQueue(glCurrentCycle + 1800, QueueItemType.eReprocessFacilityProdQueue, oFacility.ObjectID, 0, 0, 0, 0, 0, 0, 0)
                            Return eResourcesResult.Insufficient_Wait
                        End If
                    Else : bInserted = True
                    End If
                    'Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.eMaterials, lItemID(X), iTypeID(X), oProdCost.ObjectID, oProdCost.ObjTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                    'Return False
                End If
            End If
        Next X

        If bInserted = True Then
            'We've inserted an item before this one... so force the facility to get the next production item (which calls this routine again)
            '   but with the new production item
            If oFacility.GetNextProduction(False) = True Then
                Return eResourcesResult.Sufficient
            Else
                oFacility.ProductionUB = -1
                Return eResourcesResult.Insufficient_Clear
            End If
        Else
            'Ok, deduct our resources... first the basics
            If bIgnorePersonnel = False Then
                ColonyEnlisted -= oProdCost.EnlistedCost
                ColonyOfficers -= oProdCost.OfficerCost
                Population -= oProdCost.ColonistCost
            End If

            'Owner.blCredits -= oProdCost.CreditCost
            Owner.blCredits -= blCreditCost

            'TODO: What sort of race conditions could occur here? It would be potentially unsafe...

            'Now, go back through our uResSrc array and pull out the cargo
            For X As Int32 = 0 To lResSrcUB
                'ok... deduct it...
                If uResSrc(X).iCacheTypeID = ObjectType.eMineralCache OrElse uResSrc(X).iCacheTypeID = ObjectType.eMineral Then

                    If uResSrc(X).lFacilityIdx = Int32.MaxValue Then
                        Me.AdjustColonyMineralCache(uResSrc(X).lItemID, -Math.Abs(uResSrc(X).lQuantity))
                    Else
                        With CType(Me.oChildren(uResSrc(X).lFacilityIdx).oCargoContents(uResSrc(X).lCargoIdx), MineralCache)
                            .Quantity -= uResSrc(X).lQuantity
                        End With
                        'Me.oChildren(uResSrc(X).lFacilityIdx).Cargo_Cap += uResSrc(X).lQuantity
                    End If
                ElseIf uResSrc(X).iCacheTypeID = ObjectType.eComponentCache Then
                    If uResSrc(X).lFacilityIdx = Int32.MaxValue Then
                        Me.AdjustColonyComponentCache(uResSrc(X).lItemID, uResSrc(X).iTypeID, uResSrc(X).lCargoIdx, -Math.Abs(uResSrc(X).lQuantity))
                    Else
                        With CType(Me.oChildren(uResSrc(X).lFacilityIdx).oCargoContents(uResSrc(X).lCargoIdx), ComponentCache)
                            .Quantity -= uResSrc(X).lQuantity
                        End With
                        'Me.oChildren(uResSrc(X).lFacilityIdx).Cargo_Cap += uResSrc(X).lQuantity
                    End If
                Else
                    'TODO: What else could it be? Ammunition?
                End If
            Next X

        End If

        'If we are here, we have everything
        Return eResourcesResult.Sufficient
    End Function

    Public Function GetLowResourcesMsg(ByVal lLowVal As ProductionType, ByVal lLowItemID As Int32, ByVal iLowItemTypeID As Int16, ByVal lProdID As Int32, ByVal iProdTypeID As Int16) As Byte()
        Dim yMsg(46) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eColonyLowResources).CopyTo(yMsg, 0)
        Me.ColonyName.CopyTo(yMsg, 2)
        Me.GetGUIDAsString.CopyTo(yMsg, 22)
        If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
            CType(CType(Me.ParentObject, Facility).ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, 28)
        Else : CType(Me.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, 28)
        End If
        yMsg(34) = lLowVal

        System.BitConverter.GetBytes(lLowItemID).CopyTo(yMsg, 35)
        System.BitConverter.GetBytes(iLowItemTypeID).CopyTo(yMsg, 39)

        System.BitConverter.GetBytes(lProdID).CopyTo(yMsg, 41)
        System.BitConverter.GetBytes(iProdTypeID).CopyTo(yMsg, 45)

        Return yMsg
    End Function

    Public Sub UpdatePowerNeeds(ByVal lCallingFacilityIndex As Int32)
        Dim X As Int32
        Dim lGen As Int32 = 0
        Dim lNeed As Int32 = 0
        Dim lTemp As Int32
        Dim bChanged As Boolean
        Dim bNoticeSent As Boolean = False

        Dim lAttempts As Int32 = 0

        Dim lAgentEffectMult As Int32 = 100
        Dim lAgentEffectAdd As Int32 = 0

        If mbInUpdatePowerNeeds Then Return
        mbInUpdatePowerNeeds = True

        For X = 0 To mlAgentEffectUB
            If EffectValid(X) = True AndAlso moAgentEffects(X).yType = AgentEffectType.eIncreasePowerNeed Then
                If moAgentEffects(X).bAmountAsPerc = True Then
                    lAgentEffectMult += moAgentEffects(X).lAmount
                Else : lAgentEffectAdd += moAgentEffects(X).lAmount
                End If
            End If
        Next X

        Do
            lGen = 0
            lNeed = 0
            TotalPowerNeeded = 0

            lAttempts += 1
            If lAttempts > 5000 Then
                LogEvent(LogEventType.CriticalError, "UpdatePowerNeeds attempts exceeds 5000. ColonyID: " & Me.ObjectID & ", PlayerID: " & Me.Owner.ObjectID)
                Exit Do
            ElseIf lAttempts = 3000 Then
                LogEvent(LogEventType.Warning, "UpdatePowerNeeds attempts exceeds 3000.")
            End If

            bChanged = False
            For X = 0 To ChildrenUB
                If lChildrenIdx(X) <> -1 Then
                    If oChildren(X).Active = True Then
                        lTemp = oChildren(X).PowerConsumption
                        If lTemp < 0 Then
                            lGen += Math.Abs(lTemp)
                        Else
                            lNeed += lTemp
                        End If
                    End If
                    If oChildren(X).EntityDef.PowerFactor > 0 Then
                        TotalPowerNeeded += oChildren(X).EntityDef.PowerFactor
                    End If
                End If
            Next X

            If lAgentEffectMult <> 100 Then
                Dim fMult As Single = (lAgentEffectMult / 100.0F)
                lNeed = CInt(lNeed * fMult)
                TotalPowerNeeded = CInt(TotalPowerNeeded * fMult)
            End If
            lNeed += lAgentEffectAdd
            TotalPowerNeeded += lAgentEffectAdd

            'Now, is our generated greater than need?
            If lGen < lNeed Then
                If bNoticeSent = False Then
                    bNoticeSent = True
                    Me.Owner.SendPlayerMessage(GetLowResourcesMsg(ProductionType.ePowerCenter, 0, 0, -1, -1), False, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget Or AliasingRights.eAddProduction Or AliasingRights.eViewMining)
                End If

                'Nope, gonna need to shut some facilities down... is the calling facility active?
                If lCallingFacilityIndex <> -1 Then
                    If oChildren(lCallingFacilityIndex).Active = True Then
                        'Yes... was our power needs okay before?
                        If PowerGeneration >= PowerConsumption Then
                            'oChildren(lCallingFacilityIndex).CurrentStatus -= elUnitStatus.eFacilityPowered
                            oChildren(lCallingFacilityIndex).SetActive(False)
                            bChanged = True
                        End If
                    End If
                End If

                If bChanged = False Then
                    'Otherwise, we need to go through and determine what facilities won't work now...
                    'First priority is any research facility that is not producing
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).bProducing = False AndAlso oChildren(X).PowerConsumption > 0 AndAlso ((oChildren(X).yProductionType And ProductionType.eResearch) = ProductionType.eResearch) Then
                            If oChildren(X).SetActive(False) = True Then Continue Do
                        End If
                    Next X

                    'Ok, if we are here, then need next priority... production facilitys not producing that are unarmed
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).bProducing = False AndAlso oChildren(X).PowerConsumption > 0 AndAlso ((oChildren(X).yProductionType And ProductionType.eProduction) <> 0) Then
                            'Ok, is it unarmed?
                            If oChildren(X).EntityDef.WeaponDefUB = -1 Then
                                If oChildren(X).SetActive(False) = True Then Continue Do
                            End If
                        End If
                    Next X

                    'Ok, next is armed facilities not producing
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).bProducing = False AndAlso oChildren(X).PowerConsumption > 0 AndAlso ((oChildren(X).yProductionType And ProductionType.eProduction) <> 0) Then
                            If oChildren(X).SetActive(False) = True Then Continue Do
                        End If
                    Next X

                    'Next is unarmed Residential facilities
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).bProducing = False AndAlso oChildren(X).PowerConsumption > 0 AndAlso ((oChildren(X).yProductionType And ProductionType.eColonists) = ProductionType.eColonists) Then
                            'Ok, is it unarmed?
                            If oChildren(X).EntityDef.WeaponDefUB = -1 Then
                                If oChildren(X).SetActive(False) = True Then Continue Do
                            End If
                        End If
                    Next X

                    'Next is armed residential facilities
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).bProducing = False AndAlso oChildren(X).PowerConsumption > 0 AndAlso ((oChildren(X).yProductionType And ProductionType.eColonists) = ProductionType.eColonists) Then
                            If oChildren(X).SetActive(False) = True Then Continue Do
                        End If
                    Next X

                    'Next is mining facilities
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).bProducing = False AndAlso oChildren(X).PowerConsumption > 0 AndAlso ((oChildren(X).yProductionType And ProductionType.eMining) = ProductionType.eMining) Then
                            If oChildren(X).SetActive(False) = True Then Continue Do
                        End If
                    Next X

                    'Refineries and Warehouses
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).PowerConsumption > 0 AndAlso oChildren(X).bProducing = False Then
                            If (oChildren(X).yProductionType And ProductionType.eRefining) = ProductionType.eRefining Then
                                If oChildren(X).SetActive(False) = True Then Continue Do
                            ElseIf (oChildren(X).yProductionType And ProductionType.eWareHouse) = ProductionType.eWareHouse Then
                                If oChildren(X).SetActive(False) = True Then Continue Do
                            End If
                        End If
                    Next X

                    'assisting research facilities
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).PowerConsumption > 0 AndAlso oChildren(X).bProducing = False Then
                            If (oChildren(X).yProductionType And ProductionType.eResearch) = ProductionType.eResearch Then
                                Dim oTech As Epica_Tech = oChildren(X).Owner.GetTech(oChildren(X).CurrentProduction.ProductionID, oChildren(X).CurrentProduction.ProductionTypeID)
                                If oTech Is Nothing OrElse oTech.IsPrimaryResearcher(oChildren(X).ObjectID) = False Then
                                    If oTech Is Nothing = False Then oTech.RemoveResearcher(oChildren(X).ObjectID)
                                    If oChildren(X).SetActive(False) = True Then
                                        oChildren(X).CurrentProduction = Nothing
                                        oChildren(X).bProducing = False
                                        Continue Do
                                    End If
                                End If
                            End If
                        End If
                    Next X

                    'primary research facilities
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).PowerConsumption > 0 AndAlso oChildren(X).bProducing = False Then
                            If (oChildren(X).yProductionType And ProductionType.eResearch) = ProductionType.eResearch Then
                                Dim oTech As Epica_Tech = oChildren(X).Owner.GetTech(oChildren(X).CurrentProduction.ProductionID, oChildren(X).CurrentProduction.ProductionTypeID)
                                If oTech Is Nothing = False Then oTech.RemoveResearcher(oChildren(X).ObjectID)
                                If oChildren(X).SetActive(False) = True Then
                                    oChildren(X).CurrentProduction = Nothing
                                    oChildren(X).bProducing = False
                                    Continue Do
                                End If
                            End If
                        End If
                    Next X

                    'All others...
                    For X = 0 To ChildrenUB
                        If lChildrenIdx(X) <> -1 Then
                            If oChildren(X).Active = True AndAlso oChildren(X).PowerConsumption > 0 Then
                                If oChildren(X).SetActive(False) = True Then
                                    bChanged = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next X
                    If bChanged = False Then Exit Do
                End If
            End If
        Loop Until lNeed <= lGen

        PowerConsumption = lNeed
        PowerGeneration = lGen
        mbInUpdatePowerNeeds = False
    End Sub

    Private mbPreviousPlanetCorrupt As Boolean = False
    Public Sub ProcessGrowth()
        Dim fTempWFE As Single = WorkforceEfficiency
        Dim bRecalcProds As Boolean = False

        If Me.Owner Is Nothing = False AndAlso Me.Owner.bInFullLockDown = True Then Return

        If Me.ParentObject Is Nothing = False Then
            If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                If CType(Me.ParentObject, Facility).yProductionType = ProductionType.eTradePost Then Return
            End If
        End If

        Dim bHasCC As Boolean = False
        For X As Int32 = 0 To ChildrenUB
            If lChildrenIdx(X) <> -1 AndAlso (oChildren(X).yProductionType = ProductionType.eCommandCenterSpecial OrElse oChildren(X).yProductionType = ProductionType.eSpaceStationSpecial) Then
                bHasCC = True
                Exit For
            End If
        Next X

        'Handled in SQL down script
        'If Me.Owner Is Nothing = False Then
        '    If Me.Owner.AccountStatus <> AccountStatusType.eActiveAccount AndAlso Me.Owner.AccountStatus <> AccountStatusType.eTrialAccount AndAlso Me.Owner.AccountStatus <> AccountStatusType.eMondelisActive Then
        '        If Me.Owner.dtAccountWentInactive > Date.MinValue Then
        '            If Me.Population < 200000 Then
        '                If Me.Owner.dtAccountWentInactive.AddDays(14) < Now Then
        '                    Me.Population = 0
        '                End If
        '            Else
        '                If Me.Owner.dtAccountWentInactive.AddDays(30) < Now Then
        '                    Me.Population = 0
        '                End If
        '            End If
        '        End If
        '    End If
        'End If

        'WorkforceEfficiency = IF ( NumberOfJobs = 0, 0, MIN ( ( Population / NumberOfJobs ), 1.05 ) )
        If NumberOfJobs = 0 Then WorkforceEfficiency = 0.0F Else WorkforceEfficiency = Math.Min(CSng(Population / NumberOfJobs), 1.05F)
        bRecalcProds = (fTempWFE <> WorkforceEfficiency)

        'HomelessPopulation = Math.Max(0, Population - TotalHousing)
        Dim HomelessPopulation As Int32 = CInt(Math.Max(0, Population - (PoweredHousing + UnpoweredHousing)))

        Dim lAbsMaxPop As Int32 = CInt(1.15F * (PoweredHousing + UnpoweredHousing))

        'PowerlessPopulation = Math.Max(0, Math.Min( (Population - PoweredHousing), UnpoweredHousing ) )
        Dim PowerlessPopulation As Int32 = CInt(Math.Max(0, Math.Min((Population - PoweredHousing), UnpoweredHousing)))

        'TaxableWorkforce = WorkforceEfficiency * NumberOfJobs
        TaxableWorkforce = CInt(WorkforceEfficiency * NumberOfJobs)

        'Unemployed = Population - TaxableWorkforce
        Dim Unemployed As Int32 = CInt(Population - TaxableWorkforce)
        If Population < 1 Then
            UnemploymentRate = 0
        Else : UnemploymentRate = CByte(CSng(Unemployed / Population) * 100.0F)
        End If

        'fMoraleBonus = MoraleBonus(0) * MoraleBonus(1) * MoraleBonus(2) * MoraleBonus(x)
        'TODO: Calculate Morale Bonus
        Dim fMoraleBonus As Single = 0.0F

        'fAdjustedMoraleBase = IF(fMoraleBonus <> 0, Owner.BaseMorale / fMoraleBonus, Owner.BaseMorale)
        Dim fAdjustedMoraleBase As Single = 0.0F
        'Dim lBaseMorale As Int32 = CInt(Owner.BaseMorale + Owner.MoraleBoost)
        Dim lBaseMorale As Int32 = CInt(Owner.BaseMorale + Owner.oSpecials.yColonyMoraleBoost)
        If fMoraleBonus <> 0 Then
            fAdjustedMoraleBase = lBaseMorale / fMoraleBonus
        Else : fAdjustedMoraleBase = CSng(lBaseMorale)
        End If
        fAdjustedMoraleBase -= Math.Abs(Owner.BadWarDecMoralePenalty)

        'D7 = (HomelessPopulation / Population) * -100
        'D8 = (PowerlessPopulation / Population) * -50
        'D13 = (Unemployed / Population * -200
        'D18 = (((TaxRate * 10) + 1) * (TaxRate * -10))
        'fNegativeMorale = D7 + D8 + D13 + D18
        Dim fAdjustedTaxRate As Single = TaxRate / 100.0F
        Dim fNegativeMorale As Single = 0.0F

        Dim yHighestNegative As LowMoraleReason = 0

        If Population > 0 Then

            Dim lTaxRateMult As Int32 = 10
            Dim lTaxRateMax As Int32 = 80
            'If Population > 2000000 Then
            '	lTaxRateMult = 30
            '	lTaxRateMax = 27
            'ElseIf Population > 1200000 Then
            '	lTaxRateMult = 25
            '	lTaxRateMax = 32
            'ElseIf Population > 400000 Then
            '	lTaxRateMult = 20
            '	lTaxRateMax = 40
            '         End If
            If Population > 2000000 Then
                lTaxRateMult = 40
                lTaxRateMax = 20
            ElseIf Population > 1200000 Then
                lTaxRateMult = 38
                lTaxRateMax = 22
            ElseIf Population > 900000 Then
                lTaxRateMult = 36
                lTaxRateMax = 24
            ElseIf Population > 800000 Then
                lTaxRateMult = 34
                lTaxRateMax = 26
            ElseIf Population > 700000 Then
                lTaxRateMult = 32
                lTaxRateMax = 28
            ElseIf Population > 600000 Then
                lTaxRateMult = 30
                lTaxRateMax = 30
            ElseIf Population > 500000 Then
                lTaxRateMult = 28
                lTaxRateMax = 32
            ElseIf Population > 400000 Then
                lTaxRateMult = 26
                lTaxRateMax = 34
            ElseIf Population > 300000 Then
                lTaxRateMult = 24
                lTaxRateMax = 36
            ElseIf Population > 200000 Then
                lTaxRateMult = 22
                lTaxRateMax = 38
            ElseIf Population > 100000 Then
                lTaxRateMult = 20
                lTaxRateMax = 40
            End If

            Dim fHomelessPerc As Single = CSng(HomelessPopulation / Population)
            'If fHomelessPerc * 100 < Owner.oSpecials.yHomelessAllowedBeforePenalty Then fHomelessPerc = 0
            If Owner.oSpecials.yHomelessAllowedBeforePenalty > 0 Then
                fHomelessPerc = Math.Max(0, fHomelessPerc - (Owner.oSpecials.yHomelessAllowedBeforePenalty * 0.01F))
            End If
            Dim fUnemploymentPerc As Single = CSng(Unemployed / Population)
            'If fUnemploymentPerc * 100 < Owner.oSpecials.yUnemploymentBeforePenalty Then fUnemploymentPerc = 0
            If Owner.oSpecials.yUnemploymentBeforePenalty > 0 Then
                fUnemploymentPerc = Math.Max(0, fUnemploymentPerc - (Owner.oSpecials.yUnemploymentBeforePenalty * 0.01F))
            End If
            'If TaxRate < Owner.oSpecials.yTaxRateBeforePenalty Then fAdjustedTaxRate = 0
            If Owner.oSpecials.yTaxRateBeforePenalty > 0 Then
                fAdjustedTaxRate = Math.Max(0, fAdjustedTaxRate - (Owner.oSpecials.yTaxRateBeforePenalty * 0.01F))
            End If

            Dim fTaxRatePenalty As Single = ((fAdjustedTaxRate * lTaxRateMult) + 1) * (fAdjustedTaxRate * -lTaxRateMult)
            If fTaxRatePenalty < -72.0F Then
                lTaxRateMax = CInt(TaxRate) - lTaxRateMax
                If lTaxRateMax > 0 AndAlso Owner.blCredits > 0 Then
                    Dim fTemp As Single = CSng(Math.Floor(1 + (Owner.blCredits / 1000000000))) * -1.0F
                    fTaxRatePenalty += (fTemp * lTaxRateMax)
                End If
            End If

            fNegativeMorale = 0.0F

            'Process our negative morale effects from agents... 
            Dim fColonyHousingMoraleMult As Single = 1.0F
            Dim fColonyTaxMoraleMult As Single = 1.0F
            Dim fColonyUnemploymentMorale As Single = 1.0F
            For X As Int32 = 0 To mlAgentEffectUB
                If EffectValid(X) = True Then
                    Select Case moAgentEffects(X).yType
                        Case AgentEffectType.eMorale
                            fNegativeMorale += moAgentEffects(X).lAmount
                        Case AgentEffectType.eColonyHousingMorale
                            If moAgentEffects(X).bAmountAsPerc = True Then fColonyHousingMoraleMult += (moAgentEffects(X).lAmount / 100.0F)
                        Case AgentEffectType.eColonyTaxMorale
                            If moAgentEffects(X).bAmountAsPerc = True Then fColonyTaxMoraleMult += (moAgentEffects(X).lAmount / 100.0F)
                        Case AgentEffectType.eColonyUnemploymentMorale
                            If moAgentEffects(X).bAmountAsPerc = True Then fColonyUnemploymentMorale += (moAgentEffects(X).lAmount / 100.0F)
                        Case AgentEffectType.eGovernorMorale
                            fNegativeMorale += moAgentEffects(X).lAmount
                    End Select
                End If
            Next X

            Dim fTempVal As Single = (fHomelessPerc * -100) * fColonyHousingMoraleMult
            Dim fMaxVal As Single = 0
            If fTempVal < fMaxVal Then
                fMaxVal = fTempVal
                yHighestNegative = LowMoraleReason.Homeless
            End If
            fNegativeMorale += fTempVal
            fTempVal = (CSng(PowerlessPopulation / Population) * -50) * fColonyHousingMoraleMult
            If fTempVal < fMaxVal Then
                fMaxVal = fTempVal
                yHighestNegative = LowMoraleReason.UnpoweredHomes
            End If
            fNegativeMorale += fTempVal
            fTempVal = ((fUnemploymentPerc * -200) * fColonyUnemploymentMorale)
            If fTempVal < fMaxVal Then
                fMaxVal = fTempVal
                yHighestNegative = LowMoraleReason.UnemploymentRate
            End If
            fNegativeMorale += fTempVal
            fTempVal = (fTaxRatePenalty * fColonyTaxMoraleMult)
            If fTempVal < fMaxVal Then
                fMaxVal = fTempVal
                yHighestNegative = LowMoraleReason.TaxRate
            End If
            fNegativeMorale += fTempVal

            If Owner.PlayerIsAtWar = True Then
                If Owner.lWarSentiment < 0 Then
                    fTempVal = -(Math.Abs(Owner.lWarSentiment)) '* fColonyHousingMoraleMult
                    If fTempVal < fMaxVal Then
                        fMaxVal = fTempVal
                        yHighestNegative = LowMoraleReason.Sentiment
                    End If
                    fNegativeMorale += fTempVal
                End If
            Else
                If Owner.lWarSentiment > 0 Then
                    fTempVal = -(Math.Abs(Owner.lWarSentiment)) '* fColonyHousingMoraleMult
                    If fTempVal < fMaxVal Then
                        fMaxVal = fTempVal
                        yHighestNegative = LowMoraleReason.Sentiment
                    End If
                    fNegativeMorale += fTempVal
                End If
            End If

            If bHasCC = False Then
                If CType(Me.ParentObject, Epica_GUID).ObjTypeID <> ObjectType.eFacility Then fNegativeMorale += -30
            End If
        Else

            If Me.ParentObject Is Nothing = False Then
                If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                    Me.Population = 10
                    Return
                End If
            End If

            Try
                For X As Int32 = 0 To Me.ChildrenUB
                    If Me.lChildrenIdx(X) > -1 AndAlso Me.oChildren(X) Is Nothing = False Then
                        Dim oEntity As Epica_Entity = oChildren(X)
                        If oEntity Is Nothing = False Then
                            DestroyEntity(oEntity, True, -1, True, "ProcessGrowth:ColonyPosIs0")
                        End If
                    End If
                Next X
            Catch
            End Try

            'TODO: should we delete this colony?
            For X As Int32 = 0 To glColonyUB
                If glColonyIdx(X) = Me.ObjectID Then
                    glColonyIdx(X) = -1
                    Exit For
                End If
            Next X
            DeleteColony(ColonyLostReason.LowMorale)

            ''TODO: Change this in the future, but for now, the facilities belong to the pirates...
            'Dim oPiratePlayer As Player = GetEpicaPlayer(25)
            'Dim iPType As Int16 = CType(Me.ParentObject, Epica_GUID).ObjTypeID
            'Dim oDomain As DomainServer = Nothing
            'If iPType = ObjectType.ePlanet Then
            '	oDomain = CType(Me.ParentObject, Planet).oDomain
            'ElseIf iPType = ObjectType.eSolarSystem Then
            '	oDomain = CType(Me.ParentObject, SolarSystem).oDomain
            'Else
            '	For X As Int32 = 0 To Me.ChildrenUB
            '		If Me.oChildren(X) Is Nothing = False Then
            '			Try
            '				Me.oChildren(X).DeleteEntity(Me.oChildren(X).ServerIndex)
            '			Catch
            '			End Try
            '		End If
            '	Next X
            'End If
            'If oPiratePlayer Is Nothing = False AndAlso oDomain Is Nothing = False Then
            '	For X As Int32 = 0 To Me.ChildrenUB
            '		If Me.oChildren(X) Is Nothing = False Then
            '			Dim oProd As EntityProduction = oChildren(X).CurrentProduction
            '			oChildren(X).bProducing = False
            '			If oProd Is Nothing = False AndAlso oChildren(X).yProductionType = ProductionType.eResearch Then
            '				'ok, get the tech
            '				Dim oTech As Epica_Tech = oChildren(X).Owner.GetTech(oProd.ProductionID, oProd.ProductionTypeID)
            '				If oTech Is Nothing = False Then
            '					oTech.RemoveResearcher(oChildren(X).ObjectID)
            '				End If
            '			End If
            '			oChildren(X).CurrentProduction = Nothing
            '			Me.oChildren(X).Owner = GetEpicaPlayer(25)
            '			Me.oChildren(X).DataChanged()
            '			oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(Me.oChildren(X), GlobalMessageCode.eAddObjectCommand))
            '		End If
            '	Next X
            'End If

            Return
        End If

        'ColonyMorale = b31 + d20
        Dim fColonyMorale As Single = fAdjustedMoraleBase + fNegativeMorale + iControlledMorale

        If Me.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
            fColonyMorale += 50
        End If

        If fColonyMorale < -5.0F Then
            If mbSentGNSNews = False AndAlso Owner.yPlayerPhase = eyPlayerPhase.eFullLivePhase Then
                If glCurrentCycle - mlLastPositiveCycle > 9000 Then     '5 minutes
                    mbSentGNSNews = True
                    Dim yGNS(74) As Byte
                    Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2
                    yGNS(lPos) = NewsItemType.eLowMorale : lPos += 1
                    With CType(Me.ParentObject, Epica_GUID)
                        If .ObjTypeID = ObjectType.eFacility Then
                            With CType(CType(Me.ParentObject, Facility).ParentObject, Epica_GUID)
                                .GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                            End With
                        Else
                            .GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                        End If
                    End With
                    System.BitConverter.GetBytes(Me.Owner.ObjectID).CopyTo(yGNS, lPos) : lPos += 4
                    Me.ColonyName.CopyTo(yGNS, lPos) : lPos += 20
                    yGNS(lPos) = yHighestNegative : lPos += 1

                    Me.Owner.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                    yGNS(lPos) = Me.Owner.yGender : lPos += 1
                    Me.Owner.EmpireName.CopyTo(yGNS, lPos) : lPos += 20

                    goMsgSys.SendToEmailSrvr(yGNS)
                End If
            End If
        Else
            mbSentGNSNews = False
            mlLastPositiveCycle = glCurrentCycle
        End If

        'fMoraleMult = fcolonymorale/100
        MoraleMultiplier = fColonyMorale / 100.0F
        bRecalcProds = bRecalcProds OrElse (MoraleMultiplier <> mfLastMoraleMultiplier)

        'ColonyGrowthRate = CInt(Math.Floor((CInt(Owner.BaseGrowthRate) + Owner.GrowthRateBoost) * MoraleMultiplier))
        ColonyGrowthRate = CInt(Math.Floor((CInt(Owner.BaseGrowthRate) + Owner.oSpecials.yGrowthRateBonus) * MoraleMultiplier))
        ColonyGrowthRate += iControlledGrowth

        'If planet is in corruption, growth rate cannot exceed 0
        If Me.ParentObject Is Nothing = False AndAlso CType(ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
            Dim bCorrupt As Boolean = CType(Me.ParentObject, Planet).PlanetInCorruption(Me.Owner.lStartedEnvirID, Me.Owner.iStartedEnvirTypeID)
            If bCorrupt = True Then ColonyGrowthRate = Math.Min(ColonyGrowthRate, 0)
            bRecalcProds = bRecalcProds OrElse (bCorrupt <> mbPreviousPlanetCorrupt)
            mbPreviousPlanetCorrupt = bCorrupt
        End If

        Population += ColonyGrowthRate

        If Population > lAbsMaxPop AndAlso lAbsMaxPop > 100 Then Population = lAbsMaxPop
        If Population > MaxPopulation Then MaxPopulation = Population

        Dim lTempMax As Int32 = 100000000
        'If Me.ParentObject Is Nothing = False Then
        '	If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
        '		Dim oTmpFac As Facility = CType(Me.ParentObject, Facility)
        '		If oTmpFac Is Nothing = False Then
        '			lTempMax = oTmpFac.EntityDef.Structure_MaxHP \ 50
        '		End If
        '	End If
        'End If

        If Population > lTempMax Then
            'reached max pop... we max it out at 100,000,000
            Population = lTempMax
        ElseIf Population <= 0 Then
            bRecalcProds = False

            ''colony has desolved...
            'Dim bCanDelete As Boolean = True
            'For X As Int32 = 0 To ChildrenUB
            '	If lChildrenIdx(X) <> -1 Then
            '		bCanDelete = False
            '		Exit For
            '	End If
            'Next X

            'If bCanDelete = True Then
            '	DeleteColony(ColonyLostReason.LowMorale)
            'End If

            Population = 0
        End If

        If bRecalcProds = True Then
            For X As Int32 = 0 To ChildrenUB
                If lChildrenIdx(X) > 0 Then
                    Try
                        oChildren(X).RecalcProduction()
                    Catch
                    End Try
                End If
            Next X
        End If

    End Sub

    Public Sub HandleColonyTaxIncome()

        If NumberOfJobs > 0 Then
            Dim lScientist As Int32 = 0
            Dim lFactory As Int32 = 0
            Dim lOther As Int32 = 0
            Dim blTemp As Int64 = 0
            Dim lTotal As Int32 = 0
            Try
                blTemp = CLng(mf_SCIENTIST_ANNUAL_PAY * CSng(TaxableWorkforce * CSng(ScientistJobs / NumberOfJobs)))
                If blTemp > Int32.MaxValue Then lScientist = Int32.MaxValue Else lScientist = CInt(blTemp)
            Catch
            End Try
            Try
                blTemp = CLng(mf_FACTORY_ANNUAL_PAY * CSng(TaxableWorkforce * CSng(FactoryJobs / NumberOfJobs)))
                If blTemp > Int32.MaxValue Then lFactory = Int32.MaxValue Else lFactory = CInt(blTemp)
            Catch
            End Try
            Try
                blTemp = CLng(mf_OTHER_ANNUAL_PAY * CSng(TaxableWorkforce * CSng(OtherJobs / NumberOfJobs)))
                If blTemp > Int32.MaxValue Then lOther = Int32.MaxValue Else lOther = CInt(blTemp)
            Catch
            End Try
            'Dim  As Int32 = 
            'Dim  As Int32 = 
            Try
                blTemp = CLng(lScientist) + CLng(lFactory) + CLng(lOther)
                If blTemp > Int32.MaxValue Then lTotal = Int32.MaxValue Else lTotal = CInt(blTemp)
            Catch
            End Try

            Dim fMult As Single = 1.0F
            If TaxableWorkforce < 20000000 AndAlso TaxableWorkforce > 0 Then
                fMult = CSng(Math.Sqrt(CSng(20000000 / TaxableWorkforce)))
            End If

            Dim lAnnual As Int32 = CInt((fMult * TaxRate * lTotal) / 94608.0F)
            Dim lTaxIncome As Int32 = (lAnnual + 50I) * 5I  'if we should ever change the credit interval, we should change this too

            'If the colony is on a planet and there is corruption...
            If Me.ParentObject Is Nothing = False AndAlso CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                If CType(Me.ParentObject, Planet).PlanetInCorruption(Me.Owner.lStartedEnvirID, Me.Owner.iStartedEnvirTypeID) = True Then lTaxIncome = 0
            End If

            'HW Supply Lines ===================================
            Dim lHWSupply As Int32 = 0

            'Ok, determine my parent system...
            If Population < 2000000 AndAlso Me.ParentObject Is Nothing = False Then
                Dim lSysID As Int32 = -1
                With CType(Me.ParentObject, Epica_GUID)
                    If .ObjTypeID = ObjectType.ePlanet Then
                        If CType(Me.ParentObject, Planet).ParentSystem Is Nothing = False Then lSysID = CType(Me.ParentObject, Planet).ParentSystem.ObjectID
                    ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                        lSysID = .ObjectID
                    ElseIf .ObjTypeID = ObjectType.eFacility Then
                        Dim oP As Epica_GUID = CType(CType(Me.ParentObject, Facility).ParentObject, Epica_GUID)
                        If oP Is Nothing = False Then
                            If oP.ObjTypeID = ObjectType.eSolarSystem Then lSysID = oP.ObjectID
                        End If
                    End If
                End With

                If lSysID <> Me.Owner.lIronCurtainSystem Then
                    lHWSupply = CInt((1 - (Population / 2000000.0F)) * 61000)
                End If
            End If
            'End of HW Supply Lines ============================

            'MAINTENANCE COSTS AND EXPENSES =========

            'facility-specific costs...
            Dim bHasCC As Boolean = False
            Dim lResFacs As Int32 = 0
            Dim lFacs As Int32 = 0
            Dim lSpc As Int32 = 0
            Dim lOth As Int32 = 0
            Dim lTurrets As Int32 = 0
            Dim lExcessCargo As Int32 = 0 ' Math.Abs(Math.Min(0, Me.Cargo_Cap))

            Dim yProdType As Byte

            Dim lTotalFacilities As Int32 = 0

            For X As Int32 = 0 To ChildrenUB
                If lChildrenIdx(X) <> -1 Then
                    yProdType = oChildren(X).yProductionType

                    If yProdType = ProductionType.eSpaceStationSpecial Then
                        lTotalFacilities += 50
                    Else : lTotalFacilities += 1
                    End If

                    If oChildren(X).Active = True OrElse yProdType = ProductionType.eTradePost Then
                        If (oChildren(X).CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                            lExcessCargo += oChildren(X).Cargo_Cap
                        End If
                        If yProdType = ProductionType.eResearch Then
                            lResFacs += 1
                        ElseIf (yProdType And ProductionType.eAerialProduction) = ProductionType.eAerialProduction Then
                            lSpc += 1
                        ElseIf (yProdType And ProductionType.eProduction) <> 0 Then
                            lFacs += 1
                        Else
                            bHasCC = bHasCC OrElse yProdType = ProductionType.eSpaceStationSpecial OrElse yProdType = ProductionType.eCommandCenterSpecial

                            lOth += 1
                            If oChildren(X).EntityDef Is Nothing = False AndAlso oChildren(X).EntityDef.oPrototype Is Nothing = False Then
                                If oChildren(X).EntityDef.oPrototype.oHullTech Is Nothing = False Then
                                    'MSC: In this comparison, we only care about the mesh portion of the model
                                    If (oChildren(X).EntityDef.oPrototype.oHullTech.ModelID And 255) = 16 Then
                                        lOth -= 1
                                        lTurrets += 1
                                    End If
                                End If
                            End If

                            'Now, check if the facility produces tradepost
                            If yProdType = ProductionType.eTradePost Then
                                'ok, get that facility's Slot Usage cnt
                                lOth += oChildren(X).lTradePostBuySlotsUsed + oChildren(X).lTradePostSellSlotsUsed

                                ''also, get its cargo usage...
                                'Dim lCargoUsage As Int32 = oChildren(X).EntityDef.Cargo_Cap - oChildren(X).Cargo_Cap
                                'Dim lHangarUsage As Int32 = oChildren(X).EntityDef.Hangar_Cap - oChildren(X).Hangar_Cap
                                'Dim blTemp As Int64 = CLng(lTradepostContents) + CLng(lCargoUsage) + CLng(lHangarUsage)
                                'If goGTC Is Nothing = False Then blTemp -= goGTC.GetTradepostTradeCargo(oChildren(X).ObjectID)
                                'If blTemp < 0 Then blTemp = 0
                                'If blTemp > Int32.MaxValue Then lTradepostContents = Int32.MaxValue Else lTradepostContents = CInt(blTemp)
                            End If
                        End If
                    End If
                End If
            Next X
            lExcessCargo -= TotalCargoCapUsed
            If lExcessCargo > 0 Then lExcessCargo = 0
            lExcessCargo = Math.Abs(lExcessCargo)

            If Me.ParentObject Is Nothing = False AndAlso CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                Dim oFac As Facility = CType(Me.ParentObject, Facility)
                If oFac Is Nothing = False Then
                    If oFac.ParentObject Is Nothing = False AndAlso CType(oFac.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                        With CType(oFac.ParentObject, Epica_GUID)
                            lTotalFacilities = Owner.oBudget.GetFacilityPointUsage(.ObjectID, .ObjTypeID)
                        End With
                    End If
                End If
            End If

            'Now, calculate our costs...
            Dim lCCMult As Int32 = 1
            If bHasCC = False Then lCCMult = Budget.ml_NoCCMultiplier

            With Owner.oSpecials
                lResFacs = CInt(lResFacs * mfResearchFacMult * .fResearchFacUpkeepMult) * lCCMult
                lFacs = CInt(lFacs * mfFactoryMult * .fFactoryUpkeepMult) * lCCMult
                lSpc = CInt(lSpc * mfSpaceportMult * .fSpaceportUpkeepMult) * lCCMult
                lOth = CInt(lOth * mfOtherMult * .fOtherFacUpkeepMult) * lCCMult
                lTurrets = CInt(lTurrets * mfTurretMult) * lCCMult

                'population upkeep
                Dim lPopUpkeep As Int32 = CInt(Population * mfPopUpkeepMult * .fPopulationUpkeepMult) * lCCMult
                'unemployment
                Dim lUnemployedCost As Int32 = Population - NumberOfJobs
                If lUnemployedCost < 0 Then
                    lUnemployedCost = 0
                Else : lUnemployedCost = CInt(lUnemployedCost * mfUnemploymentMult * .fUnemploymentUpkeepMult) * lCCMult
                End If

                Dim lMaxFacs As Int32 = CInt(.iCPLimit) \ 5

                'If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then lTotalFacilities = 0

                If lTotalFacilities > lMaxFacs Then
                    Dim fFacCPMult As Single = 1.0F + ((lTotalFacilities - lMaxFacs) / 50.0F)
                    lPopUpkeep = CInt(lPopUpkeep * fFacCPMult)
                    lResFacs = CInt(lResFacs * fFacCPMult)
                    lFacs = CInt(lFacs * fFacCPMult)
                    lSpc = CInt(lSpc * fFacCPMult)
                    lOth = CInt(lOth * fFacCPMult)
                    lUnemployedCost = CInt(lUnemployedCost * fFacCPMult)
                    lTurrets = CInt(lTurrets * fFacCPMult)
                    lExcessCargo = CInt(lExcessCargo * fFacCPMult)
                    lHWSupply = CInt(lHWSupply * fFacCPMult)
                End If

                With CType(ParentObject, Epica_GUID)
                    Dim lID As Int32 = .ObjectID
                    If .ObjTypeID = ObjectType.eFacility Then
                        Dim oParentFac As Facility = CType(ParentObject, Facility)
                        Select Case oParentFac.EntityDef.lHullSpecialTraitID
                            Case elModelSpecialTrait.Revenue10
                                lTaxIncome = CInt(lTaxIncome * 1.1F)
                            Case elModelSpecialTrait.Revenue20
                                lTaxIncome = CInt(lTaxIncome * 1.2F)
                        End Select
                        lID = CType(oParentFac.ParentObject, Epica_GUID).ObjectID
                    End If
                    Owner.oBudget.SetColonyDetails(.ObjectID, .ObjTypeID, Me.ObjectID, lTaxIncome, lPopUpkeep, lResFacs, lFacs, lSpc, lOth, lUnemployedCost, bHasCC, TaxRate, lID, lTurrets, lExcessCargo, lTotalFacilities, GovScore, lHWSupply)
                End With

                'TRADE INCOMES (Ally players receive this) =========
                Dim lEnvirID As Int32
                Dim iEnvirTypeID As Int16
                With CType(Me.ParentObject, Epica_GUID)
                    lEnvirID = .ObjectID : iEnvirTypeID = .ObjTypeID
                End With
                If iEnvirTypeID = ObjectType.eFacility Then
                    With CType(CType(Me.ParentObject, Facility).ParentObject, Epica_GUID)
                        lEnvirID = .ObjectID : iEnvirTypeID = .ObjTypeID
                    End With
                End If
                'MSC - 01/02/09 - put this in to scale the trade income down over time
                Dim lTradeIncome As Int32 = 0
                If Me.Owner.yIronCurtainState <> eIronCurtainState.IronCurtainIsUpOnEverything Then
                    If lHWSupply > 0 Then
                        Dim dtForTradeIncomeEffect As Date = CDate("01/02/2009 11:00")
                        Dim fTradeIncomeMult As Single = CSng(CInt(Now.Subtract(dtForTradeIncomeEffect).TotalHours) / 1440.0F) '1440 hours total
                        If fTradeIncomeMult < 0 Then fTradeIncomeMult = 0
                        If fTradeIncomeMult > 1 Then fTradeIncomeMult = 1
                        lTradeIncome = CInt((lTaxIncome - lHWSupply) * 0.05F)
                    Else : lTradeIncome = CInt(lTaxIncome * 0.05F)
                    End If
                    Me.Owner.AddTradeIncomeItemToAllies(lTradeIncome, lEnvirID, iEnvirTypeID)
                End If

                'Now, adjust the credits
                If Owner.bInFullLockDown = False Then
                    Owner.blCredits += (CLng(lTaxIncome) - (CLng(lPopUpkeep) + CLng(lResFacs) + CLng(lFacs) + CLng(lSpc) + CLng(lOth) + CLng(lUnemployedCost) + CLng(lTurrets) + CLng(lExcessCargo) + CLng(lHWSupply)))
                End If
            End With

        End If
    End Sub

    ''' <summary>
    ''' To be called when a facility has changed states (powered/unpowered... added, removed, etc...)
    ''' The values updated are the values that are dependant on the facilities within the colony
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UpdateAllValues(ByVal lCallingFacilityIndex As Int32)
        UpdatePowerNeeds(lCallingFacilityIndex)

        'Reset our values
        ScientistJobs = 0
        FactoryJobs = 0
        OtherJobs = 0
        NumberOfJobs = 0
        PoweredHousing = 0
        UnpoweredHousing = 0

        Dim bColonyDead As Boolean = True

        For X As Int32 = 0 To ChildrenUB
            If lChildrenIdx(X) <> -1 Then
                Dim oFac As Facility = oChildren(X)
                If oFac Is Nothing = False Then
                    bColonyDead = False
                    If oChildren(X).Active = True Then
                        If CLng(NumberOfJobs) + CLng(oChildren(X).MaxWorkers) > Int32.MaxValue Then
                            oChildren(X).CurrentStatus = oChildren(X).CurrentStatus Xor elUnitStatus.eFacilityPowered
                            Continue For
                        End If
                        NumberOfJobs += oChildren(X).MaxWorkers

                        If (oChildren(X).yProductionType And ProductionType.eMiscellaneous) <> 0 Then
                            ScientistJobs += oChildren(X).MaxWorkers
                        ElseIf (oChildren(X).yProductionType And ProductionType.eProduction) <> 0 OrElse (oChildren(X).yProductionType And ProductionType.eFacility) <> 0 Then
                            FactoryJobs += oChildren(X).MaxWorkers
                        Else : OtherJobs += oChildren(X).MaxWorkers
                        End If

                        If oChildren(X).yProductionType = ProductionType.eTradePost Then
                            Dim lTempVal As Int32 = (oChildren(X).EntityDef.WorkerFactor * (oChildren(X).lTradePostBuySlotsUsed + oChildren(X).lTradePostSellSlotsUsed))
                            NumberOfJobs += lTempVal
                            OtherJobs += lTempVal
                        End If

                        If oChildren(X).yProductionType = ProductionType.eColonists OrElse oChildren(X).yProductionType = ProductionType.eCommandCenterSpecial OrElse oChildren(X).yProductionType = ProductionType.eSpaceStationSpecial Then
                            PoweredHousing += oChildren(X).EntityDef.ProdFactor
                        End If
                    ElseIf oChildren(X).yProductionType = ProductionType.eColonists OrElse oChildren(X).yProductionType = ProductionType.eCommandCenterSpecial OrElse oChildren(X).yProductionType = ProductionType.eSpaceStationSpecial Then
                        UnpoweredHousing += oChildren(X).EntityDef.ProdFactor
                    End If
                End If
            End If
        Next X

        If bColonyDead = True Then
            CheckForColonyDeath()
        End If
    End Sub

    Public Enum ColonyLostReason As Byte
        Neglect = 1
        Destruction = 2
        LowMorale = 3
        Abandoned = 4
    End Enum
    Public Sub DeleteColony(ByVal yReason As ColonyLostReason)
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        LogEvent(LogEventType.Informational, "DeleteColony: " & Me.ObjectID & ", Owner: " & Me.Owner.ObjectID & ", Reason: " & yReason & ", TaxRate: " & Me.TaxRate & ", Pop: " & Me.Population & ", Morale: " & Me.MoraleMultiplier & ", ParentID: " & CType(Me.ParentObject, Epica_GUID).ObjectID & ", ParentTypeID: " & CType(Me.ParentObject, Epica_GUID).ObjTypeID)

        Dim oSys As SolarSystem = Nothing
        With CType(Me.ParentObject, Epica_GUID)
            If .ObjTypeID = ObjectType.eFacility Then
                If CType(CType(Me.ParentObject, Facility).ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                    oSys = CType(CType(Me.ParentObject, Facility).ParentObject, SolarSystem)
                End If
            ElseIf .ObjTypeID = ObjectType.ePlanet Then
                Dim oPlanet As Planet = CType(Me.ParentObject, Planet)
                oSys = oPlanet.ParentSystem
                If oPlanet.OwnerID = Me.Owner.ObjectID Then oPlanet.OwnerID = 0
                oPlanet.CheckForColonyLimits(True, False)
            End If
        End With
        If oSys Is Nothing = False AndAlso Owner.yPlayerPhase = eyPlayerPhase.eFullLivePhase Then
            Dim yGNS(102) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2
            yGNS(lPos) = NewsItemType.eLostColony : lPos += 1
            oSys.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
            Me.ColonyName.CopyTo(yGNS, lPos) : lPos += 20
            'CType(ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
            If CType(ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                CType(ParentObject, Planet).PlanetName.CopyTo(yGNS, lPos) : lPos += 20
            Else : oSys.SystemName.CopyTo(yGNS, lPos) : lPos += 20
            End If
            System.BitConverter.GetBytes(Me.Owner.ObjectID).CopyTo(yGNS, lPos) : lPos += 4
            System.BitConverter.GetBytes(ColonyStart).CopyTo(yGNS, lPos) : lPos += 4
            System.BitConverter.GetBytes(MaxPopulation).CopyTo(yGNS, lPos) : lPos += 4
            yGNS(lPos) = yReason : lPos += 1

            Me.Owner.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
            yGNS(lPos) = Me.Owner.yGender : lPos += 1
            Me.Owner.EmpireName.CopyTo(yGNS, lPos) : lPos += 20

            goMsgSys.SendToEmailSrvr(yGNS)
        End If

        If Me.ObjectID < 1 Then Return

        Me.Owner.RemoveFastColonyLookup(Me.ObjectID)

        For X As Int32 = 0 To mlMineralCacheUB
            If mlMineralCacheIdx(X) > -1 Then
                If mlMineralCacheID(X) = glMineralCacheIdx(mlMineralCacheIdx(X)) Then
                    Try
                        Dim oCache As MineralCache = goMineralCache(mlMineralCacheIdx(X))
                        If oCache Is Nothing = False Then
                            oCache.DeleteMe()
                            goMineralCache(mlMineralCacheIdx(X)) = Nothing
                            glMineralCacheIdx(mlMineralCacheIdx(X)) = -1
                        End If
                    Catch
                    End Try
                End If
            End If
        Next X
        For X As Int32 = 0 To mlComponentCacheUB
            If mlComponentCacheIdx(X) > -1 Then
                If mlComponentCacheID(X) = glComponentCacheIdx(mlComponentCacheIdx(X)) Then
                    Try
                        Dim oCache As ComponentCache = goComponentCache(mlComponentCacheIdx(X))
                        If oCache Is Nothing = False Then
                            oCache.DeleteMe()
                            goComponentCache(mlComponentCacheIdx(X)) = Nothing
                            glComponentCacheIdx(mlComponentCacheIdx(X)) = -1
                        End If
                    Catch
                    End Try
                End If
            End If
        Next X

        sSQL = "DELETE FROM tblColony WHERE ColonyID = " & Me.ObjectID
        oComm = New OleDb.OleDbCommand(sSQL, goCN)
        oComm.ExecuteNonQuery()

        'Now, remove me from the globals
        For X As Int32 = 0 To glColonyUB
            If glColonyIdx(X) = ObjectID Then
                If Me.ParentObject Is Nothing = False Then
                    If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                        CType(Me.ParentObject, Planet).RemoveColonyReference(X)
                    End If
                End If

                glColonyIdx(X) = -1
                goColony(X) = Nothing
            End If
        Next X

        With CType(Me.ParentObject, Epica_GUID)
            Dim lX As Int32 = 0
            Dim lZ As Int32 = 0
            If .ObjTypeID = ObjectType.eFacility Then
                lX = CType(Me.ParentObject, Facility).LocX
                lZ = CType(Me.ParentObject, Facility).LocZ
            End If
            Me.Owner.CreateAndSendPlayerAlert(PlayerAlertType.eColonyLost, ObjectID, ObjTypeID, .ObjectID, .ObjTypeID, -1, BytesToString(ColonyName), lX, lZ, "")
        End With

        Me.Owner.HandleCheckForPlayerDeath()
    End Sub

    Public Sub DestroyAllChildrenFacilities()
        For X As Int32 = 0 To Me.ChildrenUB
            If Me.lChildrenIdx(X) <> -1 Then
                Dim oFac As Facility = Me.oChildren(X)
                If oFac Is Nothing = False Then
                    Try
                        If glFacilityIdx(oFac.ServerIndex) = oFac.ObjectID Then
                            goFacility(oFac.ServerIndex) = Nothing
                            glFacilityIdx(oFac.ServerIndex) = -1
                            RemoveLookupFacility(oFac.ObjectID, oFac.ObjTypeID)
                        End If
                        oFac.DeleteEntity(oFac.ServerIndex)
                    Catch
                    End Try
                End If
                '          'Loop through to find the item in our global array
                '          Dim lGArrayIdx As Int32 = -1
                '          For Y As Int32 = 0 To glFacilityUB
                '              If glFacilityIdx(Y) = Me.lChildrenIdx(X) Then
                'glFacilityIdx(Y) = -1
                'goFacility(Y) = Nothing
                '                  lGArrayIdx = Y
                '                  Exit For
                '              End If
                '          Next Y
                '          'And delete the facility
                '          Me.oChildren(X).DeleteEntity(lGArrayIdx)
            End If
        Next X
    End Sub

    Public Function GetFacilityWithCargoSpace(ByVal lQty As Int32) As Facility
        Dim lColonyCargoCap As Int32 = Me.Cargo_Cap
        If lColonyCargoCap < lQty Then Return Nothing

        For X As Int32 = 0 To Me.ChildrenUB
            If Me.lChildrenIdx(X) <> -1 AndAlso oChildren(X).Active = True AndAlso oChildren(X).yProductionType <> ProductionType.eMining AndAlso (oChildren(X).CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 AndAlso oChildren(X).Cargo_Cap >= lQty Then   'AndAlso oChildren(X).yProductionType <> ProductionType.eTradePost 
                If ProductionTypeSharesColonyCargo(oChildren(X).yProductionType) = True Then
                    If Me.Cargo_Cap > lQty Then Return oChildren(X)
                Else : Continue For
                End If
            End If
        Next X
        Return Nothing
    End Function

    ''' <summary>
    ''' Attempts to distribute lQty amongst all children of the colony. Returns the quantity not added. DO NOT PASS NEGATIVE NUMBERS!
    ''' </summary>
    ''' <param name="lObjectID"> Object ID of the specific object (ie: MineralID, or AlloyTechID, etc...) </param>
    ''' <param name="iObjTypeID"> Type ID of the specific object (ie: Mineral, AlloyTech, Ammunition). This determines what sort of caches will be created. </param>
    ''' <param name="lQty"> The quantity to be distributed </param>
    ''' <param name="bNoRemainder"> If true, the remainder will be forcefully placed in facilities of the colony. If false, this method returns the quantity not able to be placed. </param>
    ''' <returns>Returns the quantity not added.</returns>
    ''' <remarks></remarks>
    Public Function AddObjectCaches(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16, ByVal lQty As Int32, ByVal bNoRemainder As Boolean) As Int32
        'ok, objtypeid is the typeid of the object... so if it was Mineral, then we create mineralcaches, if it is
        '  ammunitioncache, we create ammunition caches, if it is a technology, we create componentcaches, etc...
        Dim lTemp As Int32

        Select Case iObjTypeID
            Case ObjectType.eMineral, ObjectType.eMineralCache
                Dim oMineral As Mineral = GetEpicaMineral(lObjectID)
                If oMineral Is Nothing Then Return lQty

                If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                    'Ok, add it only to my parent as I am a space station
                    With CType(Me.ParentObject, Facility)
                        If bNoRemainder = True Then
                            lTemp = lQty
                        Else
                            lTemp = Math.Min(lQty, .Cargo_Cap)
                        End If
                        If lTemp > 0 Then
                            .AddMineralCacheToCargo(oMineral.ObjectID, lTemp)
                            lQty -= lTemp
                        End If
                    End With
                Else
                    If bNoRemainder = True Then
                        lTemp = lQty
                    Else
                        lTemp = Math.Min(Me.Cargo_Cap, lQty)
                    End If
                    If lTemp < 1 Then Return lQty
                    Dim oCache As MineralCache = AdjustColonyMineralCache(lObjectID, lTemp)
                    If oCache Is Nothing = False Then
                        lQty -= lTemp
                    End If
                End If
            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eSpecialTech
                Dim oTech As Epica_Tech = Me.Owner.GetTech(lObjectID, iObjTypeID)
                If oTech Is Nothing Then
                    'Ok, try other players...
                    For X As Int32 = 0 To glPlayerUB
                        If glPlayerIdx(X) <> -1 Then
                            oTech = goPlayer(X).GetTech(lObjectID, iObjTypeID)
                            If oTech Is Nothing Then oTech = goInitialPlayer.GetTech(lObjectID, iObjTypeID)
                            If oTech Is Nothing Then
                                'ok, check our quick lookup
                                oTech = QuickLookupTechnology(lObjectID, iObjTypeID)
                            End If
                            If oTech Is Nothing = False Then Exit For
                        End If
                    Next X

                    If oTech Is Nothing Then Return lQty
                End If

                If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                    'Ok, add it only to my parent as I am a space station
                    With CType(Me.ParentObject, Facility)
                        If bNoRemainder = True Then
                            lTemp = lQty
                        Else
                            lTemp = Math.Min(lQty, .Cargo_Cap)
                        End If
                        If lTemp <> 0 Then
                            .AddComponentCacheToCargo(lObjectID, iObjTypeID, lTemp, oTech.Owner.ObjectID)
                            lQty -= lTemp
                        End If
                    End With
                Else
                    'Ok, add it to my children
                    Dim lPossible As Int32 = 0
                    For X As Int32 = 0 To Me.ChildrenUB
                        If Me.lChildrenIdx(X) <> -1 AndAlso Me.oChildren(X).Active = True AndAlso ProductionTypeSharesColonyCargo(oChildren(X).yProductionType) = True Then 'AndAlso oChildren(X).Cargo_Cap > 0 Then
                            lPossible += 1
                            If oChildren(X).Cargo_Cap > 0 Then
                                With oChildren(X)
                                    lTemp = Math.Min(lQty, .Cargo_Cap)
                                    If lTemp <> 0 Then
                                        .AddComponentCacheToCargo(lObjectID, iObjTypeID, lTemp, oTech.Owner.ObjectID)
                                        lQty -= lTemp
                                    End If

                                    If lQty < 1 Then Exit For
                                End With
                            End If
                        End If
                    Next X
                    If lQty > 0 AndAlso bNoRemainder = True Then
                        'need to force distribute the remaining
                        Dim lDist As Int32 = lQty \ lPossible
                        Dim lRem As Int32 = lQty - (lDist * lPossible)
                        For X As Int32 = 0 To Me.ChildrenUB
                            If Me.lChildrenIdx(X) <> -1 AndAlso Me.oChildren(X).Active = True AndAlso ProductionTypeSharesColonyCargo(oChildren(X).yProductionType) = True Then
                                With oChildren(X)
                                    lTemp = lDist
                                    If lRem <> 0 Then
                                        lTemp += 1
                                        lRem -= 1
                                    End If
                                    If lTemp <> 0 Then
                                        .AddComponentCacheToCargo(lObjectID, iObjTypeID, lTemp, oTech.Owner.ObjectID)
                                        lQty -= lTemp
                                    End If

                                    If lQty < 1 Then Exit For
                                End With
                            End If
                        Next X
                    End If
                End If
            Case ObjectType.eColonists, ObjectType.eEnlisted, ObjectType.eOfficers
                If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                    'Ok, add it only to my parent as I am a space station
                    With CType(Me.ParentObject, Facility)
                        lTemp = .Cargo_Cap
                        lTemp \= CInt(Owner.oSpecials.yPersonnelCargoUsage)
                        lTemp = Math.Min(lQty, lTemp)
                        If lTemp > 0 Then
                            .AddPersonnelCacheToCargo(iObjTypeID, lTemp)
                            lQty -= lTemp
                        End If
                    End With
                Else
                    'Ok, add it to my children
                    For X As Int32 = 0 To Me.ChildrenUB
                        If Me.lChildrenIdx(X) <> -1 AndAlso Me.oChildren(X).Active = True AndAlso ProductionTypeSharesColonyCargo(oChildren(X).yProductionType) = True AndAlso oChildren(X).Cargo_Cap > 0 Then
                            With oChildren(X)
                                lTemp = .Cargo_Cap
                                lTemp \= CInt(Owner.oSpecials.yPersonnelCargoUsage)
                                lTemp = Math.Min(lQty, lTemp)
                                If lTemp > 0 Then
                                    .AddPersonnelCacheToCargo(iObjTypeID, lTemp)
                                    lQty -= lTemp
                                End If

                                If lQty < 1 Then Exit For
                            End With
                        End If
                    Next X
                End If
        End Select

        Return lQty
    End Function
    'Public Function AddObjectCaches(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16, ByVal lQty As Int32) As Int32
    '	'ok, objtypeid is the typeid of the object... so if it was Mineral, then we create mineralcaches, if it is
    '	'  ammunitioncache, we create ammunition caches, if it is a technology, we create componentcaches, etc...
    '	Dim lTemp As Int32

    '	Select Case iObjTypeID
    '		Case ObjectType.eMineral, ObjectType.eMineralCache
    '			Dim oMineral As Mineral = GetEpicaMineral(lObjectID)
    '			If oMineral Is Nothing Then Return lQty

    '			If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
    '				'Ok, add it only to my parent as I am a space station
    '				With CType(Me.ParentObject, Facility)
    '					lTemp = Math.Min(lQty, .Cargo_Cap)
    '					If lTemp <> 0 Then
    '						.AddMineralCacheToCargo(oMineral.ObjectID, lTemp)
    '						lQty -= lTemp
    '					End If
    '				End With
    '			Else
    '				'Ok, add it to my children
    '				For X As Int32 = 0 To Me.ChildrenUB
    '					If Me.lChildrenIdx(X) <> -1 AndAlso Me.oChildren(X).Active = True AndAlso ProductionTypeSharesColonyCargo(oChildren(X).yProductionType) = True AndAlso oChildren(X).Cargo_Cap > 0 Then
    '						With oChildren(X)
    '							lTemp = Math.Min(lQty, .Cargo_Cap)
    '							If lTemp <> 0 Then
    '								.AddMineralCacheToCargo(oMineral.ObjectID, lTemp)
    '								lQty -= lTemp
    '							End If

    '							If lQty < 1 Then Exit For
    '						End With
    '					End If
    '				Next X
    '			End If
    '		Case ObjectType.eAmmunition
    '			'TODO: Finish this...
    '		Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eSpecialTech
    '			Dim oTech As Epica_Tech = Me.Owner.GetTech(lObjectID, iObjTypeID)
    '			If oTech Is Nothing Then
    '				'Ok, try other players...
    '				For X As Int32 = 0 To glPlayerUB
    '					If glPlayerIdx(X) <> -1 Then
    '						oTech = goPlayer(X).GetTech(lObjectID, iObjTypeID)
    '						If oTech Is Nothing Then oTech = goInitialPlayer.GetTech(lObjectID, iObjTypeID)
    '						If oTech Is Nothing Then
    '							'ok, check our quick lookup
    '							oTech = QuickLookupTechnology(lObjectID, iObjTypeID)
    '						End If
    '						If oTech Is Nothing = False Then Exit For
    '					End If
    '				Next X

    '				If oTech Is Nothing Then Return lQty
    '			End If

    '			If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
    '				'Ok, add it only to my parent as I am a space station
    '				With CType(Me.ParentObject, Facility)
    '					lTemp = Math.Min(lQty, .Cargo_Cap)
    '					If lTemp <> 0 Then
    '						.AddComponentCacheToCargo(lObjectID, iObjTypeID, lTemp, oTech.Owner.ObjectID)
    '						lQty -= lTemp
    '					End If
    '				End With
    '			Else
    '				'Ok, add it to my children
    '				For X As Int32 = 0 To Me.ChildrenUB
    '					If Me.lChildrenIdx(X) <> -1 AndAlso Me.oChildren(X).Active = True AndAlso ProductionTypeSharesColonyCargo(oChildren(X).yProductionType) = True AndAlso oChildren(X).Cargo_Cap > 0 Then
    '						With oChildren(X)
    '							lTemp = Math.Min(lQty, .Cargo_Cap)
    '							If lTemp <> 0 Then
    '								.AddComponentCacheToCargo(lObjectID, iObjTypeID, lTemp, oTech.Owner.ObjectID)
    '								lQty -= lTemp
    '							End If

    '							If lQty < 1 Then Exit For
    '						End With
    '					End If
    '				Next X
    '			End If
    '		Case ObjectType.eColonists, ObjectType.eEnlisted, ObjectType.eOfficers
    '			If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
    '				'Ok, add it only to my parent as I am a space station
    '				With CType(Me.ParentObject, Facility)
    '					lTemp = .Cargo_Cap
    '					lTemp \= CInt(Owner.oSpecials.yPersonnelCargoUsage)
    '					lTemp = Math.Min(lQty, lTemp)
    '					If lTemp > 0 Then
    '						.AddPersonnelCacheToCargo(iObjTypeID, lTemp)
    '						lQty -= lTemp
    '					End If
    '				End With
    '			Else
    '				'Ok, add it to my children
    '				For X As Int32 = 0 To Me.ChildrenUB
    '					If Me.lChildrenIdx(X) <> -1 AndAlso Me.oChildren(X).Active = True AndAlso ProductionTypeSharesColonyCargo(oChildren(X).yProductionType) = True AndAlso oChildren(X).Cargo_Cap > 0 Then
    '						With oChildren(X)
    '							lTemp = .Cargo_Cap
    '							lTemp \= CInt(Owner.oSpecials.yPersonnelCargoUsage)
    '							lTemp = Math.Min(lQty, lTemp)
    '							If lTemp > 0 Then
    '								.AddPersonnelCacheToCargo(iObjTypeID, lTemp)
    '								lQty -= lTemp
    '							End If

    '							If lQty < 1 Then Exit For
    '						End With
    '					End If
    '				Next X
    '			End If
    '	End Select

    '	Return lQty
    'End Function

    Public Function GetColonyIntelligence() As Byte
        Dim fIC As Single = 0.0F
        Dim lC As Int32 = Population
        Dim lJR As Int32
        Dim fNJ As Single
        If NumberOfJobs = 0 Then Return 100

        lJR = CInt(TaxableWorkforce * CSng(ScientistJobs / NumberOfJobs))       'Number of Research Jobs
        If lJR = 0 Then Return 100
        fNJ = lJR

        If fNJ > (lC * 0.01F) Then
            fIC += ((lC * 0.01F) * 160.0F)
            fNJ -= (lC * 0.01F)
        Else
            fIC += fNJ * 160.0F
            fNJ = 0.0F
        End If

        If fNJ > (lC * 0.02F) Then
            fIC += ((lC * 0.02F) * 150.0F)
            fNJ -= (lC * 0.02F)
        Else
            fIC += fNJ * 150.0F
            fNJ = 0.0F
        End If

        If fNJ > (lC * 0.02F) Then
            fIC += ((lC * 0.02F) * 140.0F)
            fNJ -= (lC * 0.02F)
        Else
            fIC += fNJ * 140.0F
            fNJ = 0
        End If

        If fNJ > (lC * 0.1F) Then
            fIC += ((lC * 0.1F) * 130.0F)
            fNJ -= (lC * 0.1F)
        Else
            fIC += fNJ * 130.0F
            fNJ = 0.0F
        End If

        If fNJ > (lC * 0.15F) Then
            fIC += ((lC * 0.15F) * 120.0F)
            fNJ -= (lC * 0.15F)
        Else
            fIC += fNJ * 120.0F
            fNJ = 0
        End If

        If fNJ > (lC * 0.2F) Then
            fIC += ((lC * 0.2F) * 110.0F)
            fNJ -= (lC * 0.2F)
        Else
            fIC += fNJ * 110.0F
            fNJ = 0.0F
        End If

        If fNJ > (lC * 0.25F) Then
            fIC += ((lC * 0.25F) * 100.0F)
            fNJ -= (lC * 0.25F)
        Else
            fIC += fNJ * 100.0F
            fNJ = 0.0F
        End If

        If fNJ > (lC * 0.25F) Then
            fIC += ((lC * 0.25F) * 90.0F)
            fNJ -= (lC * 0.25F)
        Else
            fIC += fNJ * 90.0F
            fNJ = 0
        End If

        lC = CInt(fIC / lJR)
        If lC < 0 Then lC = 0
        If lC > 255 Then lC = 255
        Return CByte(lC)
    End Function
    Public Function GetResearchJobCount() As Int32
        If NumberOfJobs = 0 Then Return 0
        Return CInt(TaxableWorkforce * CSng(ScientistJobs / NumberOfJobs))       'Number of Research Jobs
    End Function

    Public Sub ConsumeColony(ByRef oColony As Colony)
        With Me
            .ColonyEnlisted += oColony.ColonyEnlisted
            .ColonyOfficers += oColony.ColonyOfficers
            .Population += oColony.Population
        End With

        'Now, go through all of the facilities and change their parent to me
        For X As Int32 = 0 To oColony.ChildrenUB
            If oColony.lChildrenIdx(X) <> -1 Then
                oColony.oChildren(X).ParentColony = Me
                oColony.oChildren(X).Owner = Me.Owner
                Me.AddChildFacility(oColony.oChildren(X))

                oColony.oChildren(X).SetHangarContentsOwner(Me.Owner)
            End If
        Next X

        Me.UpdateAllValues(-1)
    End Sub

    Public Sub CheckForColonyDeath()
        For X As Int32 = 0 To Me.ChildrenUB
            If Me.lChildrenIdx(X) <> -1 Then Return
        Next X
        Owner.RemoveFastColonyLookup(Me.ObjectID)
        Me.Population = 0
        Owner.HandleCheckForPlayerDeath()
    End Sub

    Public Shared Function ProductionTypeSharesColonyCargo(ByVal yProdType As Byte) As Boolean
        If yProdType = ProductionType.eMining Then Return False
        'If yProdType = ProductionType.eTradePost Then Return False
        Return True
    End Function

#Region "  Agent Effects versus this Colony  "
    Protected moAgentEffects() As AgentEffect
    Private myAgentEffectUsed() As Byte
    Protected mlAgentEffectUB As Int32 = -1

    Public Function AddAgentEffect(ByVal lStartCycle As Int32, ByVal lDuration As Int32, ByVal yType As AgentEffectType, ByVal lAmount As Int32, ByVal bAsPercentage As Boolean, ByVal lCausedByID As Int32) As AgentEffect

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlAgentEffectUB
            If myAgentEffectUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            mlAgentEffectUB += 1
            ReDim Preserve myAgentEffectUsed(mlAgentEffectUB)
            ReDim Preserve moAgentEffects(mlAgentEffectUB)
            lIdx = mlAgentEffectUB
        End If

        moAgentEffects(lIdx) = New AgentEffect
        With moAgentEffects(lIdx)
            .bAmountAsPerc = bAsPercentage
            .lAmount = lAmount
            .lDuration = lDuration
            .lStartCycle = lStartCycle
            .yType = yType
            .lCausedByID = lCausedByID
        End With
        myAgentEffectUsed(lIdx) = 255

        Return moAgentEffects(lIdx)
    End Function

    Protected Sub RemoveAgentEffect(ByVal lIdx As Int32)
        myAgentEffectUsed(lIdx) = 0
    End Sub

    Protected Function EffectValid(ByVal lIdx As Int32) As Boolean
        Try
            If myAgentEffectUsed(lIdx) = 0 Then Return False
            If moAgentEffects(lIdx).lLastVerification = glCurrentCycle Then Return True
            If moAgentEffects(lIdx).lStartCycle <= glCurrentCycle Then
                If moAgentEffects(lIdx).lStartCycle + moAgentEffects(lIdx).lDuration > glCurrentCycle Then
                    moAgentEffects(lIdx).lLastVerification = glCurrentCycle
                    Return True
                Else : RemoveAgentEffect(lIdx)
                End If
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Colony.EffectValid: " & ex.Message)
        End Try
        Return False
    End Function

    Protected Function SaveAgentEffects() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            sSQL = "DELETE FROM tblAgentEffects WHERE EffectedItemID = " & Me.ObjectID & " AND EffectedItemTypeID = " & Me.ObjTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            'Now, all effects are inserts
            For X As Int32 = 0 To mlAgentEffectUB
                If myAgentEffectUsed(X) <> 0 Then
                    With moAgentEffects(X)
                        Try
                            sSQL = "INSERT INTO tblAgentEffects (EffectedItemID, EffectedItemTypeID, RemainingCycles, EffectType, EffectAmount, " & _
                              "yAmountAsPerc, CausedByID) VALUES (" & Me.ObjectID & ", " & Me.ObjTypeID & ", " & .lDuration - (glCurrentCycle - .lStartCycle) & _
                              ", " & CByte(.yType) & ", " & .lAmount & ", "
                            If .bAmountAsPerc = True Then sSQL &= "1," Else sSQL &= "0,"
                            sSQL &= .lCausedByID & ")"
                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                            oComm.ExecuteNonQuery()
                            oComm.Dispose()
                            oComm = Nothing
                        Catch ex As Exception
                            LogEvent(LogEventType.CriticalError, "Unable to Save Agent Effect: " & ex.Message)
                        End Try
                    End With
                End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save agent effect. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Function HasEffectActive(ByVal yType As AgentEffectType) As Boolean
        For X As Int32 = 0 To mlAgentEffectUB
            If myAgentEffectUsed(X) > 0 Then
                Dim oEffect As AgentEffect = moAgentEffects(X)
                If oEffect Is Nothing = False AndAlso oEffect.yType = yType Then
                    If EffectValid(X) = True Then Return True
                End If
            End If
        Next X
        Return False
    End Function

    Public Function AuditRemoveEffects(ByVal lTargetNum As Int32) As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To mlAgentEffectUB
            If EffectValid(X) = True Then
                If moAgentEffects(X).yType <> AgentEffectType.eGovernorMorale Then
                    If Rnd() * 100 < lTargetNum Then
                        RemoveAgentEffect(X)
                        lCnt += 1
                    End If
                End If
            End If
        Next X
        Return lCnt
    End Function
    Public Function AuditRemoveFactoryProdEffects(ByVal lTargetNum As Int32) As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To ChildrenUB
            If lChildrenIdx(X) > 0 Then
                Dim oFac As Facility = oChildren(X)
                If oFac Is Nothing = False Then lCnt += oFac.AuditRemoveEffects(lTargetNum)
            End If
        Next X
        Return lCnt
    End Function
#End Region

#Region "  Rebuild Commander  "
    'The UNIT doing the rebuilding
    Public lRebuilderUnitID As Int32 = -1
    Public bWaitingForRebuilder As Boolean = False
    Public lRebuilderQueueIdx As Int32 = -1
    Public lCancelRebuilderQueueIdx As Int32 = -1

    Private muRebuildItems() As RebuildAIItem
    Private mlRebuildItemUB As Int32 = -1

    Public Sub AddRebuildItem(ByVal lDefID As Int32, ByVal iDefTypeID As Int16, ByVal lParentID As Int32, ByVal iParentTypeID As Int16, ByVal lX As Int32, ByVal lZ As Int32, ByVal iA As Int16)
        Dim oDef As FacilityDef = Nothing
        If iDefTypeID = ObjectType.eFacilityDef Then
            oDef = GetEpicaFacilityDef(lDefID)
        End If
        If oDef Is Nothing Then Return

        Dim uNewItem As RebuildAIItem
        With uNewItem
            .lDefID = lDefID
            .iDefTypeID = iDefTypeID
            .lParentID = lParentID
            .iParentTypeID = iParentTypeID
            .LocX = lX
            .LocZ = lZ
            .LocA = iA
            .bRebuilt = False
            .bSkipped = False
            .PowerFactor = oDef.PowerFactor
            .sDefName = BytesToString(oDef.DefName)
        End With

        If muRebuildItems Is Nothing Then ReDim muRebuildItems(-1)

        SyncLock muRebuildItems
            mlRebuildItemUB += 1
            ReDim Preserve muRebuildItems(mlRebuildItemUB)
            muRebuildItems(mlRebuildItemUB) = uNewItem
        End SyncLock

        If lRebuilderQueueIdx = -1 AndAlso bWaitingForRebuilder = False AndAlso lRebuilderUnitID = -1 Then
            'ok, first, send an external email alert for cancellation
            If (Me.Owner.iInternalEmailSettings And eEmailSettings.eRebuildCancel) <> 0 Then
                Dim oPC As PlayerComm = Me.Owner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, BytesToString(Me.ColonyName) & " has lost facilities and will begin the rebuild process in 15 minutes from the time this email was sent.", "Rebuilding Colony", Me.Owner.ObjectID, GetDateAsNumber(Now), False, Me.Owner.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False AndAlso (Me.Owner.iEmailSettings And eEmailSettings.eRebuildCancel) <> 0 Then
                    goMsgSys.SendOutboundEmail(oPC, Me.Owner, GlobalMessageCode.eRebuildAISetting, Me.ObjectID, 0, 0, 0, 0, 0, "Respond with Cancel if you wish to cancel the rebuild AI.")
                End If
            End If

            'Then, we add a 15 minute timer on rebuilding to begin.
            lRebuilderQueueIdx = AddToQueueResult(glCurrentCycle + 27000, QueueItemType.eBeginRebuilderAI, Me.ObjectID, -1, -1, -1, 0, 0, 0, 0)
        End If

        If lCancelRebuilderQueueIdx <> -1 Then
            AlterQueueItem(lCancelRebuilderQueueIdx, Int32.MinValue, QueueItemType.eCancelRebuilderAI, -1, -1, -1, -1)
            lCancelRebuilderQueueIdx = -1
            Dim oUnit As Unit = GetEpicaUnit(lRebuilderUnitID)
            If oUnit Is Nothing = False Then
                DoNextRebuildOrder(oUnit)
            End If
        End If
    End Sub

    Public Sub ClearRebuildCommander()
        'ok, we do this when a player logs in or 15 minutes after the last facility is rebuilt
        'let's compose our email

        Dim lSkipped As Int32 = 0
        Dim lRebuilt As Int32 = 0

        For X As Int32 = 0 To mlRebuildItemUB
            If muRebuildItems(X).bSkipped = False AndAlso muRebuildItems(X).bRebuilt = True Then
                lRebuilt += 1
            End If
            If muRebuildItems(X).bRebuilt = False AndAlso muRebuildItems(X).bSkipped = True Then
                lSkipped += 1
            End If
        Next X
        If lSkipped <> 0 OrElse lRebuilt <> 0 Then
            If (Me.Owner.iInternalEmailSettings And eEmailSettings.eRebuildCancel) <> 0 Then
                Dim oSB As New System.Text.StringBuilder()
                oSB.AppendLine("While you were offline, a commander began rebuilding facilities that were lost in the " & BytesToString(Me.ColonyName) & " colony." & vbCrLf)

                If lRebuilt <> 0 Then
                    oSB.AppendLine("Of the facilities known to be lost, the following rebuilt:")
                    For X As Int32 = 0 To mlRebuildItemUB
                        If muRebuildItems(X).bSkipped = False AndAlso muRebuildItems(X).bRebuilt = True Then
                            oSB.AppendLine("  " & muRebuildItems(X).sDefName)
                        End If
                    Next X
                    oSB.AppendLine()
                End If
                If lSkipped <> 0 Then
                    oSB.AppendLine("The following facilities could not be rebuilt either because the area to rebuild them was blocked, there were insufficient resources or the area to rebuild them was still in conflict:")
                    For X As Int32 = 0 To mlRebuildItemUB
                        If muRebuildItems(X).bRebuilt = False AndAlso muRebuildItems(X).bSkipped = True Then
                            oSB.AppendLine("  " & muRebuildItems(X).sDefName)
                        End If
                    Next X
                End If

                Dim oPC As PlayerComm = Me.Owner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "AI Rebuild Report", Me.Owner.ObjectID, GetDateAsNumber(Now), False, Me.Owner.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then
                    Me.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If
            End If
        End If

        'TODO: Order the "Rebuilder" to self-destruct

        If Me.lRebuilderQueueIdx <> -1 Then
            AlterQueueItem(lRebuilderQueueIdx, Int32.MinValue, QueueItemType.eBeginRebuilderAI, -1, -1, -1, -1)
        End If

        'Now, clear out the rebuilder...
        If Me.lRebuilderUnitID <> -1 Then
            Dim oUnit As Unit = GetEpicaUnit(Me.lRebuilderUnitID)
            If oUnit Is Nothing = False Then oUnit.oRebuilderFor = Nothing
        End If
        Me.lRebuilderUnitID = -1
        'and clear out our rebuild array
        mlRebuildItemUB = -1
        ReDim muRebuildItems(-1)
    End Sub

    Public Sub BeginRebuildProcess()
        Dim lLastValidIdx As Int32 = -1

        Me.bWaitingForRebuilder = True

        'ok, find a command center, spaceport or factory
        For X As Int32 = 0 To ChildrenUB
            If lChildrenIdx(X) <> -1 Then
                Dim oFac As Facility = oChildren(X)
                If oFac Is Nothing = False Then
                    If oFac.Active = False Then Continue For
                    If oFac.yProductionType = ProductionType.eCommandCenterSpecial Then
                        'build an engineer
                        If oFac.bProducing = False Then
                            '30
                            If oFac.AddProduction(30, ObjectType.eUnitDef, 254, 1, 0) = True Then
                                AddEntityProducing(oFac.ServerIndex, ObjectType.eFacility, oFac.ObjectID)
                                Return
                            End If
                        Else : lLastValidIdx = X
                        End If
                    ElseIf oFac.yProductionType = ProductionType.eLandProduction Then
                        'build a cargo truck
                        If oFac.bProducing = False Then
                            '28
                            If oFac.AddProduction(28, ObjectType.eUnitDef, 254, 1, 0) = True Then
                                AddEntityProducing(oFac.ServerIndex, ObjectType.eFacility, oFac.ObjectID)
                                Return
                            End If
                        Else : lLastValidIdx = X
                        End If
                    ElseIf oFac.yProductionType = ProductionType.eAerialProduction Then
                        'build a space engineer
                        If oFac.bProducing = False Then
                            '29
                            If oFac.AddProduction(29, ObjectType.eUnitDef, 254, 1, 0) = True Then
                                AddEntityProducing(oFac.ServerIndex, ObjectType.eFacility, oFac.ObjectID)
                                Return
                            End If
                        Else : lLastValidIdx = X
                        End If
                    End If
                End If
            End If
        Next X

        If lLastValidIdx <> -1 Then
            Dim oFac As Facility = oChildren(lLastValidIdx)
            If oFac Is Nothing = False Then
                If oFac.yProductionType = ProductionType.eCommandCenterSpecial Then
                    'build an engineer
                    '30
                    If oFac.AddProduction(30, ObjectType.eUnitDef, 254, 1, 0) = True Then
                        'AddEntityProducing(oFac.ServerIndex, ObjectType.eFacility, oFac.ObjectID)
                        'Return
                    End If
                ElseIf oFac.yProductionType = ProductionType.eLandProduction Then
                    'build a cargo truck
                    '28
                    If oFac.AddProduction(28, ObjectType.eUnitDef, 254, 1, 0) = True Then
                        'AddEntityProducing(oFac.ServerIndex, ObjectType.eFacility, oFac.ObjectID)
                        'Return
                    End If
                ElseIf oFac.yProductionType = ProductionType.eAerialProduction Then
                    'build a space engineer
                    '29
                    If oFac.AddProduction(29, ObjectType.eUnitDef, 254, 1, 0) = True Then
                        'AddEntityProducing(oFac.ServerIndex, ObjectType.eFacility, oFac.ObjectID)
                        'Return
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub DoNextRebuildOrder(ByRef oRebuilder As Unit)
        If oRebuilder Is Nothing = False AndAlso oRebuilder.ParentObject Is Nothing = False Then
            'ok... let's get our list... sorted by power factor ascending
            Dim lSorted(mlRebuildItemUB) As Int32
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To mlRebuildItemUB
                Dim lIdx As Int32 = -1
                For Y As Int32 = 0 To lSortedUB
                    If muRebuildItems(lSorted(Y)).PowerFactor > muRebuildItems(X).PowerFactor Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                End If
            Next X

            'Now, do our deal...
            For X As Int32 = 0 To mlRebuildItemUB
                If muRebuildItems(lSorted(X)).bAttempted = False Then
                    'ok, here we go... send the order to build this item
                    muRebuildItems(lSorted(X)).bAttempted = True

                    'Now, send off our order...
                    Dim yMsg(23) As Byte
                    Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yMsg, lPos) : lPos += 2
                    oRebuilder.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    With muRebuildItems(lSorted(X))
                        System.BitConverter.GetBytes(.lDefID).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.iDefTypeID).CopyTo(yMsg, lPos) : lPos += 2
                        System.BitConverter.GetBytes(.LocX).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.LocZ).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.LocA).CopyTo(yMsg, lPos) : lPos += 2
                    End With

                    Dim iTemp As Int16 = CType(oRebuilder.ParentObject, Epica_GUID).ObjTypeID
                    If iTemp = ObjectType.ePlanet Then
                        CType(oRebuilder.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                    ElseIf iTemp = ObjectType.eSolarSystem Then
                        CType(oRebuilder.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                    Else
                        LogEvent(LogEventType.CriticalError, "Parent Object of rebuilder is not planet or system. Type: " & iTemp)
                    End If

                    Return
                End If
            Next X

            'Implement a queue event for clearing the rebuild order... cancel it if the rebuild ai is restarted
            If lCancelRebuilderQueueIdx = -1 Then
                lCancelRebuilderQueueIdx = AddToQueueResult(glCurrentCycle + 27000, QueueItemType.eCancelRebuilderAI, Me.ObjectID, -1, -1, -1, 0, 0, 0, 0)
            End If
        End If
    End Sub
#End Region

#Region "  Research Queue  "
    Private mlRQ_ID() As Int32 = Nothing
    Private miRQ_TypeID() As Int16 = Nothing
    Private myRQ_Cnt() As Byte = Nothing
    Private myRQ_Qty() As Byte = Nothing
    Private mlRQUB As Int32 = -1

    Public Sub AddColonyQueueItem(ByVal lTechID As Int32, ByVal iTechTypeID As Int16, ByVal yCnt As Byte, ByVal yQty As Byte)
        Try
            For X As Int32 = 0 To mlRQUB
                If mlRQ_ID(X) = lTechID AndAlso miRQ_TypeID(X) = iTechTypeID Then
                    myRQ_Cnt(X) = yCnt
                    myRQ_Qty(X) = yQty
                    Return
                End If
            Next X
            Dim lUB As Int32 = mlRQUB + 1

            ReDim Preserve mlRQ_ID(lUB)
            ReDim Preserve miRQ_TypeID(lUB)
            ReDim Preserve myRQ_Cnt(lUB)
            ReDim Preserve myRQ_Qty(lUB)
            mlRQ_ID(lUB) = lTechID
            miRQ_TypeID(lUB) = iTechTypeID
            myRQ_Cnt(lUB) = yCnt
            myRQ_Qty(lUB) = yQty
            mlRQUB = lUB
        Catch
        End Try
    End Sub

    Public Sub RemoveColonyQueueItem(ByVal lTechID As Int32, ByVal iTechTypeID As Int16)
        Try
            For X As Int32 = 0 To mlRQUB
                If mlRQ_ID(X) = lTechID AndAlso miRQ_TypeID(X) = iTechTypeID Then
                    'ok, found it
                    For Y As Int32 = X To mlRQUB - 1
                        mlRQ_ID(Y) = mlRQ_ID(Y + 1)
                        miRQ_TypeID(Y) = miRQ_TypeID(Y + 1)
                        myRQ_Cnt(Y) = myRQ_Cnt(Y + 1)
                    Next Y
                    mlRQUB -= 1
                    Exit For
                End If
            Next X
        Catch
        End Try
    End Sub

    Public Function HandleGetColonyResearchList() As Byte()
        Dim yMsg() As Byte = Nothing
        Dim lPos As Int32 = 0

        Try

            Dim lResCnt As Int32 = 0
            For X As Int32 = 0 To Me.ChildrenUB
                If Me.oChildren(X) Is Nothing = False AndAlso (Me.oChildren(X).yProductionType = ProductionType.eResearch OrElse (Me.oChildren(X).yProductionType And ProductionType.eProduction) <> 0) Then lResCnt += 1
            Next X

            ReDim yMsg(14 + ((mlRQUB + 1) * 8) + (lResCnt * 26))

            System.BitConverter.GetBytes(GlobalMessageCode.eGetColonyResearchList).CopyTo(yMsg, lPos) : lPos += 2
            CType(Me.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4

            'Queue Count
            yMsg(lPos) = CByte(mlRQUB + 1) : lPos += 1

            For X As Int32 = 0 To mlRQUB
                System.BitConverter.GetBytes(mlRQ_ID(X)).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(miRQ_TypeID(X)).CopyTo(yMsg, lPos) : lPos += 2
                yMsg(lPos) = myRQ_Cnt(X) : lPos += 1
                yMsg(lPos) = myRQ_Qty(X) : lPos += 1
            Next X

            'Researcher Counter
            System.BitConverter.GetBytes(CShort(lResCnt)).CopyTo(yMsg, lPos) : lPos += 2
            For X As Int32 = 0 To Me.ChildrenUB
                Dim oFac As Facility = Me.oChildren(X)
                If oFac Is Nothing = False AndAlso (oFac.yProductionType = ProductionType.eResearch OrElse (oFac.yProductionType And ProductionType.eProduction) <> 0) Then
                    oFac.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    System.BitConverter.GetBytes(oFac.EntityDef.ProdFactor).CopyTo(yMsg, lPos) : lPos += 4
                    yMsg(lPos) = oFac.yProductionType : lPos += 1

                    If oFac.bProducing = True AndAlso oFac.CurrentProduction Is Nothing = False Then
                        oFac.RecalcProduction()

                        With oFac.CurrentProduction
                            Dim blPointsProd As Int64 = .PointsProduced
                            Dim lLastUpdate As Int32 = .lLastUpdateCycle
                            Dim lProd As Int32 = oFac.mlProdPoints
                            Dim lFinish As Int32 = .lFinishCycle

                            Dim bIsPrimary As Boolean = False
                            If oFac.yProductionType = ProductionType.eResearch Then
                                Dim oTech As Epica_Tech = Owner.GetTech(.ProductionID, .ProductionTypeID)
                                If oTech Is Nothing = False Then
                                    oTech.FillPrimarysProductionStatus(blPointsProd, lLastUpdate, lProd, lFinish)
                                    bIsPrimary = oTech.IsPrimaryResearcher(oFac.ObjectID)
                                ElseIf Math.Abs(.ProductionTypeID) = ObjectType.eMineralTech Then
                                    blPointsProd = .PointsProduced
                                    lLastUpdate = .lLastUpdateCycle
                                End If
                            Else
                                bIsPrimary = oFac.bExcludeFromColonyQueue
                            End If


                            If blPointsProd = 0 AndAlso lLastUpdate = 0 Then
                                oFac.CurrentProduction = Nothing
                                oFac.bProducing = False

                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                                System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
                                yMsg(lPos) = 0 : lPos += 1
                                System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4

                                Continue For
                            End If

                            System.BitConverter.GetBytes(oFac.CurrentProduction.ProductionID).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(oFac.CurrentProduction.ProductionTypeID).CopyTo(yMsg, lPos) : lPos += 2

                            Dim blTemp As Int64 = blPointsProd + ((glCurrentCycle - lLastUpdate) * lProd)
                            Dim blRequired As Int64 = .PointsRequired
                            If oFac.yProductionType = ProductionType.eResearch Then blRequired = CLng(.PointsRequired * Owner.fFactionResearchTimeMultiplier)

                            Dim yTemp As Byte = CByte((blTemp / blRequired) * 100.0F)
                            Dim lCyclesRemaining As Int32 = lFinish - glCurrentCycle
                            If bIsPrimary = True Then yTemp = CByte(yTemp Or 128)

                            yMsg(lPos) = yTemp : lPos += 1
                            System.BitConverter.GetBytes(lCyclesRemaining).CopyTo(yMsg, lPos) : lPos += 4

                            System.BitConverter.GetBytes(.lProdCount).CopyTo(yMsg, lPos) : lPos += 4
                        End With

                    Else
                        System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
                        If oFac.yProductionType = ProductionType.eResearch OrElse oFac.bExcludeFromColonyQueue = False Then
                            yMsg(lPos) = 0 : lPos += 1
                        Else
                            yMsg(lPos) = 128 : lPos += 1
                        End If
                        System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4
                    End If
                End If
            Next X

        Catch
            yMsg = Nothing
        End Try
        Return yMsg
    End Function

    'returns 0 if the queue is done or a positive number representing number of cycles to wait before attempting again
    Public Function ProcessColonyProdQueues() As Int32 ' Boolean
        'ok, return true if we have a researcher that is unavailable or if the queue is empty
        If mlRQUB = -1 Then Return 0

        If Me.Owner.bInFullLockDown = True Then Return 9000

        Try
            Dim oMaxResFac As Facility = Nothing
            Dim lMaxResFacProd As Int32 = 0
            Dim oMaxProdFac As Facility = Nothing
            Dim lMaxProdFacProd As Int32 = 0

            Dim oMaxNvlProdFac As Facility = Nothing
            Dim oMaxAirProdFac As Facility = Nothing
            Dim oMaxGrnProdFac As Facility = Nothing
            Dim lMaxNvlProdFac As Int32 = 0
            Dim lMaxAirProdFac As Int32 = 0
            Dim lMaxGrnProdFac As Int32 = 0

            Dim lRequiredHangar As Int32 = 0
            For X As Int32 = 0 To mlRQUB
                If myRQ_Cnt(X) > 0 Then
                    If miRQ_TypeID(X) = ObjectType.eUnitDef Then
                        Dim oDef As Epica_Entity_Def = GetEpicaUnitDef(mlRQ_ID(X))
                        If oDef Is Nothing = False Then lRequiredHangar = oDef.HullSize
                        oDef = Nothing
                    ElseIf miRQ_TypeID(X) = ObjectType.eFacilityDef Then
                        Dim oDef As Epica_Entity_Def = GetEpicaFacilityDef(mlRQ_ID(X))
                        If oDef Is Nothing = False Then lRequiredHangar = oDef.HullSize
                        oDef = Nothing
                    End If
                    Exit For
                End If
            Next X

            For X As Int32 = 0 To Me.ChildrenUB
                Dim oFac As Facility = Me.oChildren(X)
                If oFac Is Nothing = False AndAlso oFac.Active = True AndAlso oFac.bExcludeFromColonyQueue = False Then
                    If oFac.bProducing = False AndAlso oFac.CurrentProduction Is Nothing Then
                        If oFac.yProductionType = ProductionType.eResearch Then
                            'ok, found one...
                            If oFac.EntityDef.ProdFactor > lMaxResFacProd Then
                                oMaxResFac = oFac
                                lMaxResFacProd = oFac.EntityDef.ProdFactor
                            End If
                        ElseIf (oFac.yProductionType And ProductionType.eProduction) <> 0 Then
                            'ok, found one...
                            If oFac.EntityDef.ProdFactor > lMaxProdFacProd Then
                                oMaxProdFac = oFac
                                lMaxProdFacProd = oFac.EntityDef.ProdFactor
                            End If

                            If lRequiredHangar > 0 Then
                                If oFac.EntityDef.HasHangarDoorSize(lRequiredHangar) = False Then Continue For
                                If oFac.Hangar_Cap < lRequiredHangar Then Continue For
                            End If

                            If oFac.yProductionType = ProductionType.eLandProduction Then
                                If oFac.EntityDef.ProdFactor > lMaxGrnProdFac Then
                                    oMaxGrnProdFac = oFac
                                    lMaxGrnProdFac = oFac.EntityDef.ProdFactor
                                End If
                            ElseIf oFac.yProductionType = ProductionType.eAerialProduction Then
                                If oFac.EntityDef.ProdFactor > lMaxAirProdFac Then
                                    oMaxAirProdFac = oFac
                                    lMaxAirProdFac = oFac.EntityDef.ProdFactor
                                End If
                            ElseIf oFac.yProductionType = ProductionType.eNavalProduction Then
                            End If
                        End If

                    End If
                End If
            Next X

            For X As Int32 = 0 To mlRQUB
                If myRQ_Cnt(X) > 0 Then
                    'Ok, determine the type...
                    Select Case miRQ_TypeID(X)
                        Case ObjectType.eUnitDef, ObjectType.eFacilityDef
                            Dim oDef As Epica_Entity_Def = Nothing
                            If miRQ_TypeID(X) = ObjectType.eUnitDef Then oDef = GetEpicaUnitDef(mlRQ_ID(X)) Else oDef = GetEpicaFacilityDef(mlRQ_ID(X))
                            If oDef Is Nothing Then
                                RemoveColonyQueueItem(mlRQ_ID(X), miRQ_TypeID(X))
                                Return 30
                            End If
                            If oDef.OwnerID <> Me.Owner.ObjectID AndAlso oDef.OwnerID > 0 Then
                                LogEvent(LogEventType.PossibleCheat, "Invalid DefGUID because owners mismatch in ProcessColonyProdQueues: " & Me.Owner.ObjectID)
                                RemoveColonyQueueItem(mlRQ_ID(X), miRQ_TypeID(X))
                                Return 30
                            End If

                            oMaxProdFac = Nothing
                            Select Case oDef.RequiredProductionTypeID
                                Case ProductionType.eLandProduction
                                    oMaxProdFac = oMaxGrnProdFac
                                Case ProductionType.eAerialProduction
                                    oMaxProdFac = oMaxAirProdFac
                                Case ProductionType.eNavalProduction
                                    oMaxProdFac = oMaxNvlProdFac
                                Case Else
                                    RemoveColonyQueueItem(mlRQ_ID(X), miRQ_TypeID(X))
                                    Return 30
                            End Select
                            If oMaxProdFac Is Nothing Then Return 0


                            myRQ_Cnt(X) -= CByte(1)
                            Dim lID As Int32 = mlRQ_ID(X)
                            Dim iTypeID As Int16 = miRQ_TypeID(X)
                            Dim lQty As Int32 = myRQ_Qty(X)

                            If myRQ_Cnt(X) < 1 Then
                                RemoveColonyQueueItem(lID, iTypeID)
                            End If

                            If oMaxProdFac.AddProduction(lID, iTypeID, 254, lQty, 0) = True Then
                                AddEntityProducing(oMaxProdFac.ServerIndex, ObjectType.eFacility, oMaxProdFac.ObjectID)
                                Return 30
                            Else
                                Return 9000
                            End If
                        Case Else
                            'Ok, get the tech... but, if the tech is already researched, then, try to build it
                            Dim lTechID As Int32 = mlRQ_ID(X)
                            Dim iTechTypeID As Int16 = miRQ_TypeID(X)

                            Dim oTech As Epica_Tech = Owner.GetTech(lTechID, iTechTypeID)
                            If oTech Is Nothing OrElse (oTech.ObjTypeID = ObjectType.eSpecialTech AndAlso CType(oTech, PlayerSpecialTech).bInTheTank = True) Then
                                myRQ_Cnt(X) = 0
                                RemoveColonyQueueItem(mlRQ_ID(X), miRQ_TypeID(X))
                                Return 30
                            End If

                            'Ok, now is the tech researched?
                            If oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                'yes, ok try to produce
                                If oMaxProdFac Is Nothing Then Return 0

                                'If oTech.ObjTypeID = ObjectType.eAlloyTech Then
                                '    With CType(oTech, AlloyTech)
                                '        If .AlloyResult Is Nothing Then Return 0
                                '        lTechID = .AlloyResult.ObjectID
                                '        iTechTypeID = ObjectType.eMineral
                                '    End With
                                'End If

                                myRQ_Cnt(X) -= CByte(1)

                                Dim lQty As Int32 = myRQ_Qty(X)

                                If myRQ_Cnt(X) < 1 Then
                                    RemoveColonyQueueItem(lTechID, iTechTypeID)
                                End If

                                If iTechTypeID = ObjectType.ePrototype Then
                                    'Ok, this is a special case
                                    If CType(oTech, Prototype).ResultingDef Is Nothing Then Return 30
                                    With CType(oTech, Prototype).ResultingDef
                                        If .ProductionRequirements Is Nothing Then Return 30
                                        For Y As Int32 = 0 To .ProductionRequirements.ItemCostUB
                                            Select Case .ProductionRequirements.ItemCosts(Y).ItemTypeID
                                                Case ObjectType.eMineral, ObjectType.eHullTech
                                                    Continue For
                                            End Select
                                            oMaxProdFac.AddProduction(.ProductionRequirements.ItemCosts(Y).ItemID, .ProductionRequirements.ItemCosts(Y).ItemTypeID, 254, .ProductionRequirements.ItemCosts(Y).QuantityNeeded, 0)
                                        Next Y
                                        AddEntityProducing(oMaxProdFac.ServerIndex, ObjectType.eFacility, oMaxProdFac.ObjectID)
                                        Return 30
                                    End With
                                End If

                                If oMaxProdFac.AddProduction(lTechID, iTechTypeID, 254, lQty, 0) = True Then
                                    AddEntityProducing(oMaxProdFac.ServerIndex, ObjectType.eFacility, oMaxProdFac.ObjectID)
                                    Return 30
                                Else
                                    Return 9000
                                End If
                            Else
                                If oMaxResFac Is Nothing Then Return 0

                                'ok, try to research
                                myRQ_Cnt(X) -= CByte(1)

                                If myRQ_Cnt(X) < 1 Then
                                    RemoveColonyQueueItem(mlRQ_ID(X), miRQ_TypeID(X))
                                End If

                                If oMaxResFac.AddProduction(lTechID, iTechTypeID, 254, 1, 0) = True Then
                                    AddEntityProducing(oMaxResFac.ServerIndex, ObjectType.eFacility, oMaxResFac.ObjectID)
                                    Return 30
                                Else
                                    Return 9000
                                End If
                            End If

                    End Select

                End If
            Next X
        Catch
            Return 30 'return true to try again later
        End Try
        Return 0
    End Function
#End Region

    Public Sub LoadFromDataReader(ByRef oData As OleDb.OleDbDataReader)
        With Me
            .AverageIncome = CShort(oData("AverageIncome"))
            .ColonyName = StringToBytes(CStr(oData("ColonyName")))
            .CostOfLiving = CShort(oData("CostOfLiving"))
            .DesireForWar = CByte(oData("DesireForWar"))
            .ElectronicsLevel = CByte(oData("ElectronicsLevel"))
            .GovScore = CByte(oData("GovScore"))
            '.GrowthRate = CByte(oData("GrowthRate"))
            .LocX = CInt(oData("LocX"))
            .LocY = CInt(oData("LocY"))
            .MaterialsLevel = CByte(oData("MaterialsLevel"))
            '.Morale = CByte(oData("Morale"))
            .ObjectID = CInt(oData("ColonyID"))
            .ObjTypeID = ObjectType.eColony
            .Owner = GetEpicaPlayer(CInt(oData("OwnerID")))
            .ParentObject = GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
            .Population = CInt(oData("Population"))
            .ColonyEnlisted = CInt(oData("CurrentEnlisted"))
            .ColonyOfficers = CInt(oData("CurrentOfficers"))
            .iControlledGrowth = CShort(oData("ControlledGrowth"))
            .iControlledMorale = CShort(oData("ControlledMorale"))
            .PowerConsumption = CInt(oData("PowerConsumption"))
            .PowerGeneration = CInt(oData("PowerGeneration"))
            .PowerLevel = CByte(oData("PowerLevel"))
            .TaxRate = CByte(oData("TaxRate"))
            .MaxPopulation = CInt(oData("MaxPopulation"))
            .ColonyStart = CInt(oData("ColonyStarted"))

            'set the sent GNS News to true so that we don't pound GNS with a bunch of bad stories
            ' it'll get reset if necessary
            .mbSentGNSNews = True
        End With
    End Sub
End Class

Public Enum LowMoraleReason As Byte
	UnemploymentRate = 1
	TaxRate = 2
	Homeless = 4
	UnpoweredHomes = 8
	Sentiment = 16
End Enum

Public Structure RebuildAIItem
	'GUID of the Entity def to rebuild
	Public lDefID As Int32
	Public iDefTypeID As Int16

	Public PowerFactor As Int32	'power production/consumption of this item for sorting purposes
	Public sDefName As String

	'Parent GUID
	Public lParentID As Int32
	Public iParentTypeID As Int16

	'Location to attempt to rebuild
	Public LocX As Int32
	Public LocZ As Int32
	Public LocA As Int16

	Public bRebuilt As Boolean
	Public bSkipped As Boolean
	Public bAttempted As Boolean
End Structure

Public Structure ResearchQueueItem
	Public lTechID As Int32
	Public iTechTypeID As Int16
	Public yResearcherCount As Byte
End Structure