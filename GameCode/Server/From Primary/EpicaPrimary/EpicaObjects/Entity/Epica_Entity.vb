
Public Class Epica_Entity
    Inherits Epica_GUID

    Public EntityName(19) As Byte       'inherits the EntityDef.oPrototype.PrototypeName originally

    Private moOwner As Player
    Public Property Owner() As Player
        Get
            Return moOwner
        End Get
        Set(ByVal value As Player)
            moOwner = value
            bNeedsSaved = True
        End Set
    End Property

#Region "  Parent Object  "
    '#If EXTENSIVELOGGING = 1 Then
    Private moParentObject As Object
    Public Property ParentObject() As Object
        Get
            Return moParentObject
        End Get
        Set(ByVal value As Object)
            Dim lFromID As Int32 = -1
            Dim iFromTypeID As Int16 = -1
            If moParentObject Is Nothing = False Then
                With CType(moParentObject, Epica_GUID)
                    lFromID = .ObjectID
                    iFromTypeID = .ObjTypeID
                End With
            End If
            Dim lToID As Int32 = -1
            Dim iToTypeID As Int16 = -1
            If value Is Nothing = False Then
                With CType(value, Epica_GUID)
                    lToID = .ObjectID
                    iToTypeID = .ObjTypeID
                End With
            End If
            If gbServerInitializing = False AndAlso (lFromID <> lToID OrElse iFromTypeID <> iToTypeID) Then LogEvent(LogEventType.ExtensiveLogging, "PrntChg:" & Me.ObjectID & ", " & Me.ObjTypeID & " from " & lFromID & ", " & iFromTypeID & " to " & lToID & ", " & iToTypeID)
            moParentObject = value

            If Me.ObjTypeID = ObjectType.eUnit Then
                With CType(Me, Unit)
                    If .lFleetID > 0 Then
                        Dim oFleet As UnitGroup = GetEpicaUnitGroup(.lFleetID)
                        If oFleet Is Nothing = False Then
                            oFleet.UpdateUnitGroupLocation()
                        End If
                    End If
                End With
            End If
        End Set
    End Property
    '#Else
    '	Public ParentObject As Object
    '#End If
#End Region

    Private mlLocX As Int32
    Public Property LocX() As Int32
        Get
            Return mlLocX
        End Get
        Set(ByVal value As Int32)
            mlLocX = value
            bNeedsSaved = True
        End Set
    End Property
    Private mlLocZ As Int32
    Public Property LocZ() As Int32
        Get
            Return mlLocZ
        End Get
        Set(ByVal value As Int32)
            mlLocZ = value
            bNeedsSaved = True
        End Set
    End Property
    Private miLocAngle As Int16
    Public Property LocAngle() As Int16
        Get
            Return miLocAngle
        End Get
        Set(ByVal value As Int16)
            miLocAngle = value
            bNeedsSaved = True
        End Set
    End Property

    Public bLaunchedForCounterAttack As Boolean = False

	'Current Armor Hitpoints
    Private mlQ1_HP As Int32
    Private mlQ2_HP As Int32
    Private mlQ3_HP As Int32
    Private mlQ4_HP As Int32
    Private mlStructure_HP As Int32    'remaining structural hitpoints
    Private mlCurrentStatus As Int32           'bit-wise representation of the ship's status and activity
    Private mlShield_HP As Int32       'current shield hp
    Private myExpLevel As Byte = 0

    Public Property Q1_HP() As Int32
        Get
            Return mlQ1_HP
        End Get
        Set(ByVal value As Int32)
            mlQ1_HP = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property Q2_HP() As Int32
        Get
            Return mlQ2_HP
        End Get
        Set(ByVal value As Int32)
            mlQ2_HP = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property Q3_HP() As Int32
        Get
            Return mlQ3_HP
        End Get
        Set(ByVal value As Int32)
            mlQ3_HP = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property Q4_HP() As Int32
        Get
            Return mlQ4_HP
        End Get
        Set(ByVal value As Int32)
            mlQ4_HP = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property Structure_HP() As Int32
        Get
            Return mlStructure_HP
        End Get
        Set(ByVal value As Int32)
            mlStructure_HP = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property Shield_HP() As Int32
        Get
            Return mlShield_HP
        End Get
        Set(ByVal value As Int32)
            mlShield_HP = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property CurrentStatus() As Int32
        Get
            Return mlCurrentStatus
        End Get
        Set(ByVal value As Int32)
            mlCurrentStatus = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property ExpLevel() As Byte
        Get
            Return myExpLevel
        End Get
        Set(ByVal value As Byte)
            myExpLevel = value
            bNeedsSaved = True
        End Set
    End Property

    'If the servers are shut down while the entity is in movement, we need to store the entity's current speed
    Public CurrentSpeed As Byte

    Public bProducing As Boolean = False
    Public bMining As Boolean = False
    Public lDockeeCnt As Int32 = 0
    Public CurrentProduction As EntityProduction

    'Because Facilities and Units can both mine from a mineral cache...
    Public lCacheIndex As Int32 = -1         'SERVER INDEX NOT ID
    Public lCacheID As Int32 = -1
    Public iCacheTypeID As Int16 = -1       'for differentiating between minerals and components

    Public Fuel_Cap As Int32        'fuel_cap used

    Public iTargetingTactics As Int16
    Public iCombatTactics As Int32
    'TODO: need to include the extended values in the add object message and stuff...
    Public lExtendedCT_1 As Int32
    Public lExtendedCT_2 As Int32

    Public lCurrentAmmo() As Int32          'NOT REAL TIME, they are used to store the data over long term

    Public mlProdPoints As Int32

    'Hangar and Cargo Holds... 
    Public oHangarContents() As Epica_GUID
    Public lHangarIdx() As Int32
    Public lHangarUB As Int32 = -1
    Public oCargoContents() As Epica_GUID
    Public lCargoIdx() As Int32
    Public lCargoUB As Int32 = -1

    Public yProductionType As Byte
    Public bNeedsSaved As Boolean = False
    Public lDeathStatus As Int32        'status the entity had when it died (only used when in RemoveObject)

    Friend uUndockQueueItem As QueueItem

    Public ReadOnly Property Cargo_Cap() As Int32
        Get
            If oCargoContents Is Nothing Then lCargoUB = -1
            If yProductionType <> ProductionType.eSpaceStationSpecial Then
                If (CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then Return 0
            End If

            Dim lCapUsed As Int32 = 0
            Dim lTmpUB As Int32 = -1
            If lCargoIdx Is Nothing = False Then lTmpUB = Math.Min(lCargoUB, lCargoIdx.GetUpperBound(0))
            For X As Int32 = 0 To lTmpUB
                If lCargoIdx(X) <> -1 Then
                    Dim oTmp As Epica_GUID = oCargoContents(X)
                    If oTmp Is Nothing = False Then
                        Select Case oTmp.ObjTypeID
                            Case ObjectType.eMineralCache
                                lCapUsed += CType(oTmp, MineralCache).Quantity
                            Case ObjectType.eComponentCache
                                lCapUsed += CType(oTmp, ComponentCache).Quantity
                            Case ObjectType.eEnlisted, ObjectType.eOfficers, ObjectType.eColonists
                                lCapUsed += (oTmp.ObjectID * CInt(Owner.oSpecials.yPersonnelCargoUsage))
                        End Select
                    Else
                        lCargoIdx(X) = -1
                    End If
                End If
            Next X

            Dim lCargoCap As Int32 = 0
            If (CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                If Me.ObjTypeID = ObjectType.eUnit Then
                    lCargoCap = CType(Me, Unit).EntityDef.Cargo_Cap
                ElseIf Me.ObjTypeID = ObjectType.eFacility Then
                    lCargoCap = CType(Me, Facility).EntityDef.Cargo_Cap
                End If
            End If

            'Need to take into account the per facility cargo usage of the colony's cargo
            'If Me.ObjTypeID = ObjectType.eFacility AndAlso Colony.ProductionTypeSharesColonyCargo(Me.yProductionType) = True Then
            '	With CType(Me, Facility)
            '		If .ParentColony Is Nothing = False Then lCapUsed += .ParentColony.GetPerFacilityCargoUsed(lCargoCap)
            '	End With
            'End If

            If yProductionType = ProductionType.eSpaceStationSpecial Then
                If CType(Me, Facility).ParentColony Is Nothing = False Then
                    Dim oColony As Colony = CType(Me, Facility).ParentColony
                    For X As Int32 = 0 To oColony.ChildrenUB
                        If oColony.lChildrenIdx(X) <> -1 AndAlso oColony.lChildrenIdx(X) <> ObjectID Then
                            If (oColony.oChildren(X).CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                                lCargoCap += oColony.oChildren(X).Cargo_Cap
                            End If
                        End If
                    Next X
                End If
            End If

            'Now, apply any agent effects
            If mlAgentEffectUB <> -1 Then
                Dim fCargoCapMult As Single = 1.0F
                For X As Int32 = 0 To mlAgentEffectUB
                    If EffectValid(X) = True Then
                        If moAgentEffects(X).yType = AgentEffectType.eCargoBay Then
                            If moAgentEffects(X).bAmountAsPerc = True Then
                                fCargoCapMult *= (moAgentEffects(X).lAmount / 100.0F)
                            Else : lCargoCap += moAgentEffects(X).lAmount
                            End If
                        End If
                    End If
                Next X
                lCargoCap = CInt(lCargoCap * fCargoCapMult)
            End If
            If lCargoCap < 0 Then lCargoCap = 0

            Return lCargoCap - lCapUsed
        End Get
    End Property
    Public ReadOnly Property Hangar_Cap() As Int32
        Get
            If (CurrentStatus And elUnitStatus.eHangarOperational) = 0 Then Return 0
            If oHangarContents Is Nothing Then lHangarUB = -1
            Dim lCapUsed As Int32 = 0
            Dim lTmpUB As Int32 = -1
            If lHangarIdx Is Nothing = False Then lTmpUB = Math.Min(lHangarUB, lHangarIdx.GetUpperBound(0))
            For X As Int32 = 0 To lTmpUB
                If lHangarIdx(X) <> -1 Then
                    Dim oTmp As Epica_GUID = oHangarContents(X)
                    If oTmp Is Nothing = False Then
                        Select Case oTmp.ObjTypeID
                            Case ObjectType.eUnit
                                lCapUsed += CType(oTmp, Unit).EntityDef.HullSize
                            Case ObjectType.eFacility
                                lCapUsed += CType(oTmp, Facility).EntityDef.HullSize
                        End Select
                    Else
                        lHangarIdx(X) = -1
                    End If
                End If
            Next X

            Dim lHangarCap As Int32 = 0
            If Me.ObjTypeID = ObjectType.eUnit Then
                lHangarCap = CType(Me, Unit).EntityDef.Hangar_Cap
            ElseIf Me.ObjTypeID = ObjectType.eFacility Then
                lHangarCap = CType(Me, Facility).EntityDef.Hangar_Cap
            End If

            If yProductionType = ProductionType.eSpaceStationSpecial Then
                If CType(Me, Facility).ParentColony Is Nothing = False Then
                    Dim oColony As Colony = CType(Me, Facility).ParentColony
                    For X As Int32 = 0 To oColony.ChildrenUB
                        If oColony.lChildrenIdx(X) <> -1 AndAlso oColony.lChildrenIdx(X) <> ObjectID Then
                            If (oColony.oChildren(X).CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then
                                lHangarCap += oColony.oChildren(X).Hangar_Cap
                            End If
                        End If
                    Next X
                End If
            End If

            'Now, apply any agent effects
            Dim fHangarCapMult As Single = 1.0F
            For X As Int32 = 0 To mlAgentEffectUB
                If EffectValid(X) = True Then
                    If moAgentEffects(X).yType = AgentEffectType.eHangarBay Then
                        If moAgentEffects(X).bAmountAsPerc = True Then
                            fHangarCapMult *= (moAgentEffects(X).lAmount / 100.0F)
                        Else : lHangarCap += moAgentEffects(X).lAmount
                        End If
                    End If
                End If
            Next X
            lHangarCap = CInt(lHangarCap * fHangarCapMult)
            If lHangarCap < 0 Then lHangarCap = 0

            Return lHangarCap - lCapUsed
        End Get
    End Property

#Region "Rally Location Data"
    Public mlRallyPointX As Int32 = Int32.MinValue
    Public mlRallyPointZ As Int32 = Int32.MinValue
    Public miRallyPointA As Int16
    Public mlRallyPointEnvirID As Int32
    Public miRallyPointEnvirTypeID As Int16
    Public Property RallyPointX() As Int32
        Get
            RallyPointX = mlRallyPointX
        End Get
        Set(ByVal value As Int32)
            mlRallyPointX = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property RallyPointZ() As Int32
        Get
            RallyPointZ = mlRallyPointZ
        End Get
        Set(ByVal value As Int32)
            mlRallyPointZ = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property RallyPointA() As Int16
        Get
            RallyPointA = miRallyPointA
        End Get
        Set(ByVal value As Int16)
            miRallyPointA = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property RallyPointEnvirID() As Int32
        Get
            RallyPointEnvirID = mlRallyPointEnvirID
        End Get
        Set(ByVal value As Int32)
            mlRallyPointEnvirID = value
            bNeedsSaved = True
        End Set
    End Property
    Public Property RallyPointEnvirTypeID() As Int16
        Get
            RallyPointEnvirTypeID = miRallyPointEnvirTypeID
        End Get
        Set(ByVal value As Int16)
            miRallyPointEnvirTypeID = value
            bNeedsSaved = True
        End Set
    End Property
#End Region

    Protected mlNextLaunch() As Int32

    'Public Sub TransferCargo(ByRef oEntity As Epica_Entity, ByVal bLoadFromColony As Boolean)
    '	'ok, determine how much cargo we are transferring...
    '	Dim lQty As Int32

    '	Dim oReceiver As Epica_Entity = oEntity

    '	'If colCargo Is Nothing Then Exit Sub 'nothing to transfer
    '	If lCargoUB = -1 Then Return

    '	If (Me.CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then Return
    '	If (oReceiver.CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then Return

    '	For X As Int32 = 0 To lCargoUB
    '		If lCargoIdx(X) <> -1 Then
    '			If oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then

    '				'We are transferring a mineral cache... is the receiver a refinery or warehouse?
    '				If oReceiver.ObjTypeID = ObjectType.eFacility Then
    '					If oReceiver.yProductionType = ProductionType.eRefining OrElse oReceiver.yProductionType = ProductionType.eWareHouse Then
    '						'Yes, it is... is the receiver full?
    '						If oReceiver.Cargo_Cap = 0 Then
    '							'yes, it is... so find the next warehouse...
    '							Dim bFound As Boolean = False
    '							If CType(oReceiver, Facility).ParentColony Is Nothing = False Then
    '								With CType(oReceiver, Facility).ParentColony
    '									For lTmpIdx As Int32 = 0 To .ChildrenUB
    '										If .lChildrenIdx(lTmpIdx) <> -1 AndAlso (.oChildren(lTmpIdx).yProductionType And ProductionType.eProduction) = 0 AndAlso .oChildren(lTmpIdx).yProductionType <> ProductionType.eMining AndAlso .oChildren(lTmpIdx).yProductionType <> ProductionType.eTradePost Then
    '											'Ok, possible, is it active and is the cargo bay operational?
    '											If .oChildren(lTmpIdx).Active = True AndAlso (.oChildren(lTmpIdx).CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
    '												'Yes... is it full?
    '												If .oChildren(lTmpIdx).Cargo_Cap > 0 Then
    '													'No, it is not... so use this as the receiver now...
    '													oReceiver = .oChildren(lTmpIdx)
    '													bFound = True
    '													Exit For
    '												End If
    '											End If
    '										End If
    '									Next lTmpIdx
    '								End With
    '							End If
    '							'Nobody has any more room... so we're done...
    '							If bFound = False Then
    '								'TODO: Send an Insufficient capacity notice...
    '								Return
    '							End If
    '						End If

    '					End If
    '				ElseIf oReceiver.ObjTypeID = ObjectType.eUnit Then
    '					Dim oMin As Mineral = CType(oCargoContents(X), MineralCache).oMineral
    '					If oMin Is Nothing = False Then
    '						If CType(oReceiver, Unit).AcceptingCargo(oMin.ObjectID, oMin.ObjTypeID) = False Then Continue For
    '					Else : Continue For
    '					End If
    '				End If

    '				With CType(oCargoContents(X), MineralCache)
    '					lQty = .Quantity

    '					'If oReceiver.ObjTypeID = ObjectType.eFacility AndAlso (oReceiver.yProductionType = ProductionType.eRefining OrElse oReceiver.yProductionType = ProductionType.eWareHouse) Then
    '					If oReceiver.ObjTypeID = ObjectType.eFacility AndAlso oReceiver.yProductionType <> ProductionType.eMining Then
    '                           If oReceiver.Owner.CheckFirstContactWithMineral(.oMineral.ObjectID) = True Then
    '                               If .oMineral.ObjectID = 157 Then
    '								lQty += 3000000
    '                               End If
    '                           End If
    '					End If

    '					Dim cExistingCargo As MineralCache = Nothing
    '					For lTmpX As Int32 = 0 To oReceiver.lCargoUB
    '						If oReceiver.lCargoIdx(lTmpX) <> -1 AndAlso oReceiver.oCargoContents(lTmpX).ObjTypeID = .ObjTypeID Then
    '							If CType(oReceiver.oCargoContents(lTmpX), MineralCache).oMineral.ObjectID = .oMineral.ObjectID Then
    '								cExistingCargo = CType(oReceiver.oCargoContents(lTmpX), MineralCache)
    '								Exit For
    '							End If
    '						End If
    '					Next lTmpX

    '					'Now, determine the quantity to transfer...
    '					Dim lReceiverCargoCap As Int32 = oReceiver.Cargo_Cap
    '					If lReceiverCargoCap < lQty Then lQty = lReceiverCargoCap

    '					'Now, check our costs
    '					If Me.ObjTypeID = ObjectType.eFacility AndAlso Me.yProductionType = ProductionType.eMining Then
    '						Dim blTotalCost As Int64 = CLng(lQty) * 5L
    '						If blTotalCost > oReceiver.Owner.blCredits Then
    '							If oReceiver.Owner.blCredits > 4 Then
    '								lQty = CInt(oReceiver.Owner.blCredits \ 5)
    '								blTotalCost = CLng(lQty) * 5L
    '							Else
    '								Return
    '							End If
    '						End If
    '						oReceiver.Owner.blCredits -= blTotalCost

    '						'Now, give the owner of this facility 1 credit per mineral
    '						Me.Owner.blCredits += lQty
    '					End If

    '					If lQty = 0 Then Return

    '					'Now... do the transfer
    '					'if the target already has cargo of this mineral...
    '					If cExistingCargo Is Nothing = False Then
    '						'Then, increase the existing cargo bay's quantity of mineral by lQty...
    '						cExistingCargo.Quantity += lQty

    '						'Decrease our quantity... if it decreases to 0 it will delete itself
    '						.Quantity -= lQty

    '						If oCargoContents(X) Is Nothing = False Then
    '							If .Quantity = 0 Then
    '								CType(oCargoContents(X), MineralCache).DeleteMe()
    '								lCargoIdx(X) = -1
    '								oCargoContents(X) = Nothing
    '							End If
    '						End If
    '					Else
    '						'Ok, the target does not already have a cache there... are we transferring ALL of the cache?
    '						If lQty = .Quantity Then
    '							'Yes, so... we'll add it to the entity
    '							.ParentObject = oReceiver
    '							oReceiver.AddCargoRef(oCargoContents(X))
    '							'And remove it from us...
    '							lCargoIdx(X) = -1
    '							oCargoContents(X) = Nothing
    '						Else
    '							'No, we are doing a partial transfer... so create a new mineral cache
    '							oReceiver.AddMineralCacheToCargo(.oMineral.ObjectID, lQty)

    '							'Decrease our quantity... if it decrease to 0 it will delete itself
    '							.Quantity -= lQty
    '						End If
    '					End If

    '					'Regardless of what happened to the mineral, update our cargo-cap based on what was transferred
    '					'Cargo_Cap += lQty
    '					'And the cargo on the receiver
    '					'oReceiver.Cargo_Cap -= lQty
    '				End With
    '			End If
    '		End If
    '	Next X

    '	'Ok, now, is the receiver nothing? and does it have cargo cap remaining?
    '	If bLoadFromColony = True AndAlso oReceiver Is Nothing = False AndAlso oReceiver.Cargo_Cap > 0 Then
    '		'ok, am I a facility who shares cargo with the colony?
    '		If Me.ObjTypeID = ObjectType.eFacility AndAlso Colony.ProductionTypeSharesColonyCargo(Me.yProductionType) = True Then
    '			Dim oColony As Colony = CType(Me, Facility).ParentColony
    '			If oColony Is Nothing = False AndAlso oColony.lChildrenIdx Is Nothing = False Then
    '				'I need to check the rest of the colony for minerals (only) and if a specific mineral is
    '				'chosen, only that mineral... if I find them I can transfer from them too.

    '				Try
    '					'Ok, loop through the children
    '					Dim lMaxChild As Int32 = Math.Min(oColony.ChildrenUB, oColony.lChildrenIdx.GetUpperBound(0))
    '					For X As Int32 = 0 To lMaxChild
    '						If oColony.lChildrenIdx(X) > -1 Then
    '							Dim oChild As Facility = oColony.oChildren(X)
    '							If oChild Is Nothing = False AndAlso Colony.ProductionTypeSharesColonyCargo(oChild.yProductionType) = True Then
    '								If (oChild.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
    '									If oChild.EntityDef Is Nothing = False AndAlso oChild.EntityDef.Cargo_Cap > 0 Then
    '										oChild.TransferCargo(oReceiver, False)

    '										If oReceiver.Cargo_Cap = 0 Then Exit For
    '									End If
    '								End If
    '							End If
    '						End If
    '					Next X
    '				Catch
    '					'do nothing, we'll be okay
    '				End Try

    '			End If
    '		End If
    '	End If


    'End Sub

    Public Sub TransferCargo(ByRef oEntity As Epica_Entity, ByVal bLoadFromColony As Boolean)
        'ok, determine how much cargo we are transferring...
        Dim lQty As Int32

        Dim oReceiver As Epica_Entity = oEntity

        'If colCargo Is Nothing Then Exit Sub 'nothing to transfer
        If lCargoUB = -1 AndAlso bLoadFromColony = False Then Return

        If (Me.CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then Return
        If (oReceiver.CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then Return

        If oReceiver.ObjTypeID = ObjectType.eUnit Then
            With CType(oReceiver, Unit)
                If .AcceptingColonists = True Then
                    If Me.ObjTypeID = ObjectType.eFacility AndAlso bLoadFromColony = True Then
                        Dim oColony As Colony = CType(Me, Facility).ParentColony
                        If oColony Is Nothing = False Then
                            lQty = Math.Min(oColony.Population, oReceiver.Cargo_Cap)

                            LogEvent(LogEventType.ExtensiveLogging, "TransferColonists From Colony " & oColony.ObjectID & ", Rec: " & oReceiver.ObjectID & ", Qty: " & lQty & ", ColonyPop: " & oColony.Population & ", Owner: " & oColony.Owner.ObjectID)

                            oColony.Population -= lQty
                            oReceiver.AddPersonnelCacheToCargo(ObjectType.eColonists, lQty)
                            Return
                        End If
                    End If
                ElseIf .AcceptingEnlisted = True Then
                    If Me.ObjTypeID = ObjectType.eFacility AndAlso bLoadFromColony = True Then
                        Dim oColony As Colony = CType(Me, Facility).ParentColony
                        If oColony Is Nothing = False Then
                            lQty = Math.Min(oColony.ColonyEnlisted, oReceiver.Cargo_Cap)
                            oColony.ColonyEnlisted -= lQty
                            oReceiver.AddPersonnelCacheToCargo(ObjectType.eEnlisted, lQty)
                            Return
                        End If
                    End If
                ElseIf .AcceptingOfficers = True Then
                    If Me.ObjTypeID = ObjectType.eFacility AndAlso bLoadFromColony = True Then
                        Dim oColony As Colony = CType(Me, Facility).ParentColony
                        If oColony Is Nothing = False Then
                            lQty = Math.Min(oColony.ColonyOfficers, oReceiver.Cargo_Cap)
                            oColony.ColonyOfficers -= lQty
                            oReceiver.AddPersonnelCacheToCargo(ObjectType.eOfficers, lQty)
                            Return
                        End If
                    End If
                End If
            End With
        End If

        For X As Int32 = 0 To lCargoUB
            If lCargoIdx(X) <> -1 Then
                'If oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then

                Dim bAddToColony As Boolean = False
                Dim oColony As Colony = Nothing
                Dim lColonyCargoCap As Int32 = 0

                'We are transferring a mineral cache... is the receiver a refinery or warehouse?
                If oReceiver.ObjTypeID = ObjectType.eFacility Then
                    If Colony.ProductionTypeSharesColonyCargo(oReceiver.yProductionType) = True Then   'oReceiver.yProductionType = ProductionType.eRefining OrElse oReceiver.yProductionType = ProductionType.eWareHouse OrElse oReceiver.yProductionType = ProductionType.eSpaceStationSpecial OrElse oReceiver.yProductionType = ProductionType.eTradePost Then
                        'Yes, it is... get the colony
                        With CType(oReceiver, Facility)
                            If .ParentColony Is Nothing = False Then
                                lColonyCargoCap = .ParentColony.Cargo_Cap
                                If lColonyCargoCap <= 0 AndAlso oCargoContents(X).ObjTypeID <> ObjectType.eColonists AndAlso oCargoContents(X).ObjTypeID <> ObjectType.eEnlisted AndAlso oCargoContents(X).ObjTypeID <> ObjectType.eOfficers Then
                                    'TODO: Send an insufficient capacity notice
                                    Return
                                End If
                                bAddToColony = True
                                oColony = .ParentColony
                            Else
                                'ok, try to put it in the facility's cargo bay directly
                                If .Cargo_Cap = 0 Then
                                    'TODO: send an insufficient capacity notice
                                    Return
                                End If
                            End If
                        End With
                    End If
                ElseIf oReceiver.ObjTypeID = ObjectType.eUnit Then
                    If oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then
                        Dim oMin As Mineral = CType(oCargoContents(X), MineralCache).oMineral
                        If oMin Is Nothing = False Then
                            If CType(oReceiver, Unit).AcceptingCargo(oMin.ObjectID, oMin.ObjTypeID, ObjectType.eMineralCache) = False Then Continue For
                        Else : Continue For
                        End If
                    ElseIf oCargoContents(X).ObjTypeID = ObjectType.eColonists Then
                        If CType(oReceiver, Unit).AcceptingColonists = False Then Continue For
                    ElseIf oCargoContents(X).ObjTypeID = ObjectType.eEnlisted Then
                        If CType(oReceiver, Unit).AcceptingEnlisted = False Then Continue For
                    ElseIf oCargoContents(X).ObjTypeID = ObjectType.eOfficers Then
                        If CType(oReceiver, Unit).AcceptingOfficers = False Then Continue For
                    Else
                        Dim oTech As Epica_Tech = CType(oCargoContents(X), ComponentCache).GetComponent
                        If oTech Is Nothing = False Then
                            If CType(oReceiver, Unit).AcceptingCargo(oTech.ObjectID, oTech.ObjTypeID, ObjectType.eComponentCache) = False Then Continue For
                        Else : Continue For
                        End If
                    End If
                End If

                If oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then

                    With CType(oCargoContents(X), MineralCache)
                        lQty = .Quantity
                        If lQty < 1 Then Continue For

                        'If oReceiver.ObjTypeID = ObjectType.eFacility AndAlso (oReceiver.yProductionType = ProductionType.eRefining OrElse oReceiver.yProductionType = ProductionType.eWareHouse) Then
                        If oReceiver.ObjTypeID = ObjectType.eFacility AndAlso oReceiver.yProductionType <> ProductionType.eMining Then
                            If oReceiver.Owner.CheckFirstContactWithMineral(.oMineral.ObjectID) = True Then
                                If .oMineral.ObjectID = 157 AndAlso oReceiver.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                                    lQty += 3000000
                                End If
                            End If
                        End If

                        If bAddToColony = True AndAlso oColony Is Nothing = False Then
                            If lQty > lColonyCargoCap Then lQty = lColonyCargoCap
                            If lQty = 0 Then Return
                            oColony.AdjustColonyMineralCache(.oMineral.ObjectID, lQty)
                            .Quantity -= lQty
                            Continue For
                        End If

                        Dim cExistingCargo As MineralCache = Nothing
                        For lTmpX As Int32 = 0 To oReceiver.lCargoUB
                            If oReceiver.lCargoIdx(lTmpX) <> -1 AndAlso oReceiver.oCargoContents(lTmpX).ObjTypeID = .ObjTypeID Then
                                If CType(oReceiver.oCargoContents(lTmpX), MineralCache).oMineral.ObjectID = .oMineral.ObjectID Then
                                    cExistingCargo = CType(oReceiver.oCargoContents(lTmpX), MineralCache)
                                    Exit For
                                End If
                            End If
                        Next lTmpX

                        'Now, determine the quantity to transfer...
                        Dim lReceiverCargoCap As Int32 = oReceiver.Cargo_Cap
                        If lReceiverCargoCap < lQty Then lQty = lReceiverCargoCap

                        'Now, check our costs
                        If Me.ObjTypeID = ObjectType.eFacility AndAlso Me.yProductionType = ProductionType.eMining Then
                            Dim blTotalCost As Int64 = CLng(lQty) * 5L
                            If blTotalCost > oReceiver.Owner.blCredits Then
                                If oReceiver.Owner.blCredits > 4 Then
                                    lQty = CInt(oReceiver.Owner.blCredits \ 5)
                                    blTotalCost = CLng(lQty) * 5L
                                Else
                                    Return
                                End If
                            End If
                            oReceiver.Owner.blCredits -= blTotalCost

                            'Now, give the owner of this facility 1 credit per mineral
                            Me.Owner.blCredits += lQty
                        End If

                        If lQty = 0 Then Return

                        'Now... do the transfer
                        'if the target already has cargo of this mineral...
                        If cExistingCargo Is Nothing = False Then
                            'Then, increase the existing cargo bay's quantity of mineral by lQty...
                            cExistingCargo.Quantity += lQty

                            'Decrease our quantity... if it decreases to 0 it will delete itself
                            .Quantity -= lQty

                            'If oCargoContents(X) Is Nothing = False Then
                            '	If .Quantity = 0 Then
                            '		CType(oCargoContents(X), MineralCache).DeleteMe()
                            '		lCargoIdx(X) = -1
                            '		oCargoContents(X) = Nothing
                            '	End If
                            'End If
                        Else
                            'Ok, the target does not already have a cache there... are we transferring ALL of the cache?
                            If lQty = .Quantity Then
                                'Yes, so... we'll add it to the entity
                                .ParentObject = oReceiver
                                oReceiver.AddCargoRef(oCargoContents(X))
                                'And remove it from us...
                                lCargoIdx(X) = -1
                                oCargoContents(X) = Nothing
                            Else
                                'No, we are doing a partial transfer... so create a new mineral cache
                                oReceiver.AddMineralCacheToCargo(.oMineral.ObjectID, lQty)

                                'Decrease our quantity... if it decrease to 0 it will delete itself
                                .Quantity -= lQty
                            End If
                        End If

                    End With
                ElseIf oCargoContents(X).ObjTypeID = ObjectType.eColonists Then
                    lQty = oCargoContents(X).ObjectID
                    If lQty < 1 Then Continue For
                    If bAddToColony = True AndAlso oColony Is Nothing = False Then
                        oColony.Population += lQty
                        lCargoIdx(X) = -1
                        oCargoContents(X).ObjectID = 0
                        Continue For
                    Else
                        lQty = Math.Min(lQty, oReceiver.Cargo_Cap)
                        oReceiver.AddPersonnelCacheToCargo(ObjectType.eColonists, lQty)
                        oCargoContents(X).ObjectID -= lQty
                        If oCargoContents(X).ObjectID <= 1 Then lCargoIdx(X) = -1
                        Continue For
                    End If
                ElseIf oCargoContents(X).ObjTypeID = ObjectType.eEnlisted Then
                    lQty = oCargoContents(X).ObjectID
                    If lQty < 1 Then Continue For
                    If bAddToColony = True AndAlso oColony Is Nothing = False Then
                        oColony.ColonyEnlisted += lQty
                        lCargoIdx(X) = -1
                        oCargoContents(X).ObjectID = 0
                        Continue For
                    Else
                        lQty = Math.Min(lQty, oReceiver.Cargo_Cap)
                        oReceiver.AddPersonnelCacheToCargo(ObjectType.eEnlisted, lQty)
                        oCargoContents(X).ObjectID -= lQty
                        If oCargoContents(X).ObjectID <= 1 Then lCargoIdx(X) = -1
                        Continue For
                    End If
                ElseIf oCargoContents(X).ObjTypeID = ObjectType.eOfficers Then
                    lQty = oCargoContents(X).ObjectID
                    If lQty < 1 Then Continue For
                    If bAddToColony = True AndAlso oColony Is Nothing = False Then
                        oColony.ColonyOfficers += lQty
                        lCargoIdx(X) = -1
                        oCargoContents(X).ObjectID = 0
                        Continue For
                    Else
                        lQty = Math.Min(lQty, oReceiver.Cargo_Cap)
                        oReceiver.AddPersonnelCacheToCargo(ObjectType.eOfficers, lQty)
                        oCargoContents(X).ObjectID -= lQty
                        If oCargoContents(X).ObjectID <= 1 Then lCargoIdx(X) = -1
                        Continue For
                    End If
                Else
                    With CType(oCargoContents(X), ComponentCache)
                        lQty = .Quantity
                        If lQty < 1 Then Continue For

                        If bAddToColony = True AndAlso oColony Is Nothing = False Then
                            If lQty > lColonyCargoCap Then lQty = lColonyCargoCap
                            If lQty = 0 Then Return
                            oColony.AdjustColonyComponentCache(.ComponentID, .ComponentTypeID, .ComponentOwnerID, lQty)
                            .Quantity -= lQty
                            Continue For
                        End If

                        Dim cExistingCargo As ComponentCache = Nothing
                        For lTmpX As Int32 = 0 To oReceiver.lCargoUB
                            If oReceiver.lCargoIdx(lTmpX) <> -1 AndAlso oReceiver.oCargoContents(lTmpX).ObjTypeID = .ObjTypeID Then
                                If CType(oReceiver.oCargoContents(lTmpX), ComponentCache).ComponentID = .ComponentID AndAlso CType(oReceiver.oCargoContents(lTmpX), ComponentCache).ComponentTypeID = .ComponentTypeID Then
                                    cExistingCargo = CType(oReceiver.oCargoContents(lTmpX), ComponentCache)
                                    Exit For
                                End If
                            End If
                        Next lTmpX

                        'Now, determine the quantity to transfer...
                        Dim lReceiverCargoCap As Int32 = oReceiver.Cargo_Cap
                        If lReceiverCargoCap < lQty Then lQty = lReceiverCargoCap

                        If lQty = 0 Then Return

                        'Now... do the transfer
                        'if the target already has cargo of this mineral...
                        If cExistingCargo Is Nothing = False Then
                            'Then, increase the existing cargo bay's quantity of mineral by lQty...
                            cExistingCargo.Quantity += lQty

                            'Decrease our quantity... if it decreases to 0 it will delete itself
                            .Quantity -= lQty
                        Else
                            'Ok, the target does not already have a cache there... are we transferring ALL of the cache?
                            If lQty = .Quantity Then
                                'Yes, so... we'll add it to the entity
                                .ParentObject = oReceiver
                                oReceiver.AddCargoRef(oCargoContents(X))
                                'And remove it from us...
                                lCargoIdx(X) = -1
                                oCargoContents(X) = Nothing
                            Else
                                'No, we are doing a partial transfer... so create a new mineral cache
                                oReceiver.AddComponentCacheToCargo(.ComponentID, .ComponentTypeID, lQty, .ComponentOwnerID)

                                'Decrease our quantity... if it decrease to 0 it will delete itself
                                .Quantity -= lQty
                            End If
                        End If

                    End With

                End If


            End If
        Next X

        'Ok, now, is the receiver nothing? and does it have cargo cap remaining?
        If bLoadFromColony = True AndAlso oReceiver Is Nothing = False AndAlso oReceiver.Cargo_Cap > 0 Then
            'ok, am I a facility who shares cargo with the colony?
            If Me.ObjTypeID = ObjectType.eFacility AndAlso Colony.ProductionTypeSharesColonyCargo(Me.yProductionType) = True Then
                Dim oColony As Colony = CType(Me, Facility).ParentColony
                If oColony Is Nothing = False AndAlso oColony.lChildrenIdx Is Nothing = False Then
                    'I need to check the rest of the colony for minerals (only) and if a specific mineral is
                    'chosen, only that mineral... if I find them I can transfer from them too.

                    Try

                        Dim lCurUB As Int32 = -1
                        With oColony
                            If CType(oReceiver, Unit).NeedToTryMinerals() = True Then
                                If .mlMineralCacheIdx Is Nothing = False Then lCurUB = Math.Min(.mlMineralCacheUB, .mlMineralCacheIdx.GetUpperBound(0))
                                For X As Int32 = 0 To lCurUB
                                    If .mlMineralCacheIdx(X) > -1 Then
                                        If glMineralCacheIdx(.mlMineralCacheIdx(X)) = .mlMineralCacheID(X) Then
                                            If CType(oReceiver, Unit).AcceptingCargo(.mlMineralCacheMineralID(X), ObjectType.eMineral, ObjectType.eMineralCache) = True Then
                                                Dim oCache As MineralCache = goMineralCache(.mlMineralCacheIdx(X))
                                                If oCache Is Nothing = False Then
                                                    Dim lNewQty As Int32 = Math.Max(0, Math.Min(oReceiver.Cargo_Cap, oCache.Quantity))
                                                    If lNewQty < 1 Then Continue For
                                                    oCache.Quantity -= lNewQty
                                                    oReceiver.AddMineralCacheToCargo(.mlMineralCacheMineralID(X), lNewQty)
                                                    If oReceiver.Cargo_Cap = 0 Then Return
                                                End If
                                            End If
                                        End If
                                    End If
                                Next X
                            End If

                            If CType(oReceiver, Unit).NeedToTryComponents() = True Then
                                lCurUB = -1
                                If .mlComponentCacheIdx Is Nothing = False Then lCurUB = Math.Min(.mlComponentCacheUB, .mlComponentCacheIdx.GetUpperBound(0))
                                For X As Int32 = 0 To lCurUB
                                    If .mlComponentCacheIdx(X) > -1 Then
                                        If glComponentCacheIdx(.mlComponentCacheIdx(X)) = .mlComponentCacheID(X) Then
                                            If CType(oReceiver, Unit).AcceptingCargo(.mlComponentCacheID(X), goComponentCache(.mlComponentCacheIdx(X)).ObjTypeID, ObjectType.eComponentCache) = True Then
                                                Dim oCache As ComponentCache = goComponentCache(.mlComponentCacheIdx(X))
                                                If oCache Is Nothing = False Then
                                                    Dim lNewQty As Int32 = Math.Max(0, Math.Min(oReceiver.Cargo_Cap, oCache.Quantity))
                                                    If lNewQty < 1 Then Continue For
                                                    oCache.Quantity -= lNewQty
                                                    oReceiver.AddComponentCacheToCargo(oCache.ComponentID, oCache.ComponentTypeID, lNewQty, oCache.ComponentOwnerID)
                                                    If oReceiver.Cargo_Cap = 0 Then Return
                                                End If
                                            End If
                                        End If
                                    End If
                                Next X
                            End If
                        End With


                        ''Ok, loop through the children
                        'Dim lMaxChild As Int32 = Math.Min(oColony.ChildrenUB, oColony.lChildrenIdx.GetUpperBound(0))
                        'For X As Int32 = 0 To lMaxChild
                        '    If oColony.lChildrenIdx(X) > -1 Then
                        '        Dim oChild As Facility = oColony.oChildren(X)
                        '        If oChild Is Nothing = False AndAlso Colony.ProductionTypeSharesColonyCargo(oChild.yProductionType) = True Then
                        '            If (oChild.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                        '                If oChild.EntityDef Is Nothing = False AndAlso oChild.EntityDef.Cargo_Cap > 0 Then
                        '                    oChild.TransferCargo(oReceiver, False)
                        '                    If oReceiver.Cargo_Cap = 0 Then Exit For
                        '                End If
                        '            End If
                        '        End If
                        '    End If
                        'Next X
                    Catch
                        'do nothing, we'll be okay
                    End Try

                End If
            End If
        End If


    End Sub

    Public Sub DeleteEntity(ByVal lGlobalArrayIndex As Int32)
        'call this to remove a unit or facility entity from the database
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand = Nothing

        'Remove the object from the pathfinding server
        goMsgSys.SendPathfindingRemoveObject(CType(Me, Epica_GUID))

        'If we are producing a command center... ensure that any colony in the environment has bCCInProduction reset
        If Me.bProducing = True AndAlso Me.CurrentProduction Is Nothing = False Then
            'Ok, check my production type
            If Me.CurrentProduction.ProductionTypeID = ObjectType.eFacilityDef Then
                Dim oTmpFacDef As FacilityDef = GetEpicaFacilityDef(Me.CurrentProduction.ProductionID)
                If oTmpFacDef Is Nothing = False Then
                    If oTmpFacDef.ProductionTypeID = ProductionType.eCommandCenterSpecial Then
                        With CType(Me.ParentObject, Epica_GUID)
                            Dim lTempIdx As Int32 = Owner.GetColonyFromParent(.ObjectID, .ObjTypeID)
                            If lTempIdx <> -1 AndAlso goColony(lTempIdx) Is Nothing = False Then goColony(lTempIdx).bCCInProduction = False
                        End With
                    ElseIf oTmpFacDef.ProductionTypeID = ProductionType.eTradePost Then
                        With CType(Me.ParentObject, Epica_GUID)
                            Dim lTempIdx As Int32 = Owner.GetColonyFromParent(.ObjectID, .ObjTypeID)
                            If lTempIdx <> -1 AndAlso goColony(lTempIdx) Is Nothing = False Then goColony(lTempIdx).bTradepostInProduction = False
                        End With
                    End If
                End If
            ElseIf Me.yProductionType = ProductionType.eResearch AndAlso Me.ObjTypeID = ObjectType.eFacility Then
                Dim oTech As Epica_Tech = Me.Owner.GetTech(Me.CurrentProduction.ProductionID, Me.CurrentProduction.ProductionTypeID)
                If oTech Is Nothing = False Then oTech.RemoveResearcher(Me.ObjectID)
            End If
        End If

        'Remove me from the unit group I am in (If I am a unit that is in a unit group)
        If Me.ObjTypeID = ObjectType.eUnit Then

            If Me.yProductionType = ProductionType.eFacility Then
                Me.Owner.HandleCheckForPlayerDeath()
            End If

            Dim lFleetID As Int32 = CType(Me, Unit).lFleetID
            If lFleetID <> -1 Then
                Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(lFleetID)
                If oUnitGroup Is Nothing = False Then oUnitGroup.RemoveUnit(Me.ObjectID, True, True)
            End If

            If Me.Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID Then Me.Owner.lMilitaryScore -= CType(Me, Unit).EntityDef.CombatRating
            'ElseIf Me.ObjTypeID = ObjectType.eFacility Then
            '    If Me.ParentObject Is Nothing = False Then
            '        With CType(Me.ParentObject, Epica_GUID)
            '            If .ObjTypeID = ObjectType.eSolarSystem Then
            '                If Me.Owner Is Nothing = False Then
            '                    If CType(Me.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
            '                        Me.Owner.RemoveSpaceFacIdx(CType(Me, Facility).ServerIndex)
            '                    End If
            '                End If
            '            End If
            '        End With
            '    End If
        End If

        'Clear any "Being Mined By" locks I may have
        If lCacheIndex <> -1 AndAlso glMineralCacheIdx(lCacheIndex) = lCacheID Then
            Dim oCache As MineralCache = goMineralCache(lCacheIndex)
            If oCache Is Nothing = False AndAlso oCache.BeingMinedBy Is Nothing = False Then
                With oCache
                    If .BeingMinedBy.ObjectID = Me.ObjectID AndAlso .BeingMinedBy.ObjTypeID = Me.ObjTypeID Then
                        .BeingMinedBy = Nothing
                    End If
                End With
            End If
        End If

        Try
            If Me.ParentObject Is Nothing = False Then
                Dim iParentType As Int16 = CType(Me.ParentObject, Epica_GUID).ObjTypeID
                If iParentType = ObjectType.ePlanet OrElse iParentType = ObjectType.eSolarSystem Then
                    Dim lParent As Int32 = CType(Me.ParentObject, Epica_GUID).ObjectID
                    If Rnd() * 200.0F < 1.0F Then
                        'Ok, determine what drops... find a component that is still active
                        Dim oDef As Epica_Entity_Def = Nothing
                        If Me.ObjTypeID = ObjectType.eUnit Then
                            oDef = CType(Me, Unit).EntityDef
                        ElseIf Me.ObjTypeID = ObjectType.eFacility Then
                            oDef = CType(Me, Facility).EntityDef
                        End If

                        If oDef Is Nothing = False AndAlso oDef.oPrototype Is Nothing = False Then
                            Dim lCnt As Int32 = 0
                            Dim lVals(10) As Int32
                            If (lDeathStatus And elUnitStatus.eAftWeapon1) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eAftWeapon1
                            End If
                            If (lDeathStatus And elUnitStatus.eAftWeapon2) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eAftWeapon2
                            End If
                            If (lDeathStatus And elUnitStatus.eEngineOperational) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eEngineOperational
                            End If
                            If (lDeathStatus And elUnitStatus.eForwardWeapon1) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eForwardWeapon1
                            End If
                            If (lDeathStatus And elUnitStatus.eForwardWeapon2) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eForwardWeapon2
                            End If
                            If (lDeathStatus And elUnitStatus.eLeftWeapon1) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eLeftWeapon1
                            End If
                            If (lDeathStatus And elUnitStatus.eLeftWeapon2) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eLeftWeapon2
                            End If
                            If (lDeathStatus And elUnitStatus.eRadarOperational) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eRadarOperational
                            End If
                            If (lDeathStatus And elUnitStatus.eRightWeapon1) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eRightWeapon1
                            End If
                            If (lDeathStatus And elUnitStatus.eRightWeapon2) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eRightWeapon2
                            End If
                            If (lDeathStatus And elUnitStatus.eShieldOperational) <> 0 Then
                                lCnt += 1 : lVals(lCnt - 1) = elUnitStatus.eShieldOperational
                            End If

                            Dim yCacheType As Byte = 0
                            If (oDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
                                yCacheType = yCacheType Or MineralCacheType.eGround
                            ElseIf (oDef.yChassisType And ChassisType.eNavalBased) <> 0 Then
                                yCacheType = yCacheType Or MineralCacheType.eNaval
                            ElseIf (oDef.yChassisType And ChassisType.eAtmospheric) <> 0 OrElse (oDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                                yCacheType = yCacheType Or MineralCacheType.eFlying
                            End If

                            If lCnt <> 0 Then
                                Dim lDropVal As Int32 = CInt(Rnd() * lCnt)
                                If lDropVal > 10 Then lDropVal = 10
                                If lVals(lDropVal) <> 0 Then
                                    'Ok, what dropped?
                                    Dim yCompMsg() As Byte = Nothing
                                    Select Case lVals(lDropVal)
                                        Case elUnitStatus.eAftWeapon1, elUnitStatus.eAftWeapon2
                                            Dim oWpn As BaseWeaponTech = oDef.oPrototype.GetRandomWeaponComponentOnSide(UnitArcs.eBackArc)
                                            If oWpn Is Nothing = False Then
                                                'ok, put this in space....
                                                Dim lCacheIdx As Int32 = AddComponentCache(lParent, iParentType, 1, Me.LocX, Me.LocZ, oWpn.ObjectID, oWpn.ObjTypeID, oWpn.Owner.ObjectID, yCacheType)
                                                If lCacheIdx <> -1 Then
                                                    Dim oTmpCache As ComponentCache = goComponentCache(lCacheIdx)
                                                    If oTmpCache Is Nothing = False Then yCompMsg = goMsgSys.GetAddObjectMessage(oTmpCache, GlobalMessageCode.eAddObjectCommand)
                                                End If
                                            End If
                                            oWpn = Nothing
                                        Case elUnitStatus.eEngineOperational
                                            Dim oEngine As EngineTech = oDef.oPrototype.oEngineTech
                                            If oEngine Is Nothing = False Then
                                                'ok, put this in space....
                                                Dim lCacheIdx As Int32 = AddComponentCache(lParent, iParentType, 1, Me.LocX, Me.LocZ, oEngine.ObjectID, oEngine.ObjTypeID, oEngine.Owner.ObjectID, yCacheType)
                                                If lCacheIdx <> -1 Then
                                                    Dim oTmpCache As ComponentCache = goComponentCache(lCacheIdx)
                                                    If oTmpCache Is Nothing = False Then yCompMsg = goMsgSys.GetAddObjectMessage(oTmpCache, GlobalMessageCode.eAddObjectCommand)
                                                End If
                                            End If
                                            oEngine = Nothing
                                        Case elUnitStatus.eForwardWeapon1, elUnitStatus.eForwardWeapon2
                                            Dim oWpn As BaseWeaponTech = oDef.oPrototype.GetRandomWeaponComponentOnSide(UnitArcs.eForwardArc)
                                            If oWpn Is Nothing = False Then
                                                'ok, put this in space....
                                                Dim lCacheIdx As Int32 = AddComponentCache(lParent, iParentType, 1, Me.LocX, Me.LocZ, oWpn.ObjectID, oWpn.ObjTypeID, oWpn.Owner.ObjectID, yCacheType)
                                                If lCacheIdx <> -1 Then
                                                    Dim oTmpCache As ComponentCache = goComponentCache(lCacheIdx)
                                                    If oTmpCache Is Nothing = False Then yCompMsg = goMsgSys.GetAddObjectMessage(oTmpCache, GlobalMessageCode.eAddObjectCommand)
                                                End If
                                            End If
                                            oWpn = Nothing
                                        Case elUnitStatus.eLeftWeapon1, elUnitStatus.eLeftWeapon2
                                            Dim oWpn As BaseWeaponTech = oDef.oPrototype.GetRandomWeaponComponentOnSide(UnitArcs.eLeftArc)
                                            If oWpn Is Nothing = False Then
                                                'ok, put this in space....
                                                Dim lCacheIdx As Int32 = AddComponentCache(lParent, iParentType, 1, Me.LocX, Me.LocZ, oWpn.ObjectID, oWpn.ObjTypeID, oWpn.Owner.ObjectID, yCacheType)
                                                If lCacheIdx <> -1 Then
                                                    Dim oTmpCache As ComponentCache = goComponentCache(lCacheIdx)
                                                    If oTmpCache Is Nothing = False Then yCompMsg = goMsgSys.GetAddObjectMessage(oTmpCache, GlobalMessageCode.eAddObjectCommand)
                                                End If
                                            End If
                                            oWpn = Nothing
                                        Case elUnitStatus.eRadarOperational
                                            Dim oRadar As RadarTech = oDef.oPrototype.oRadarTech
                                            If oRadar Is Nothing = False Then
                                                'ok, put this in space....
                                                Dim lCacheIdx As Int32 = AddComponentCache(lParent, iParentType, 1, Me.LocX, Me.LocZ, oRadar.ObjectID, oRadar.ObjTypeID, oRadar.Owner.ObjectID, yCacheType)
                                                If lCacheIdx <> -1 Then
                                                    Dim oTmpCache As ComponentCache = goComponentCache(lCacheIdx)
                                                    If oTmpCache Is Nothing = False Then yCompMsg = goMsgSys.GetAddObjectMessage(oTmpCache, GlobalMessageCode.eAddObjectCommand)
                                                End If
                                            End If
                                            oRadar = Nothing
                                        Case elUnitStatus.eRightWeapon1, elUnitStatus.eRightWeapon2
                                            Dim oWpn As BaseWeaponTech = oDef.oPrototype.GetRandomWeaponComponentOnSide(UnitArcs.eRightArc)
                                            If oWpn Is Nothing = False Then
                                                'ok, put this in space....
                                                Dim lCacheIdx As Int32 = AddComponentCache(lParent, iParentType, 1, Me.LocX, Me.LocZ, oWpn.ObjectID, oWpn.ObjTypeID, oWpn.Owner.ObjectID, yCacheType)
                                                If lCacheIdx <> -1 Then
                                                    Dim oTmpCache As ComponentCache = goComponentCache(lCacheIdx)
                                                    If oTmpCache Is Nothing = False Then yCompMsg = goMsgSys.GetAddObjectMessage(oTmpCache, GlobalMessageCode.eAddObjectCommand)
                                                End If
                                            End If
                                            oWpn = Nothing
                                        Case elUnitStatus.eShieldOperational
                                            Dim oShield As ShieldTech = oDef.oPrototype.oShieldTech
                                            If oShield Is Nothing = False Then
                                                'ok, put this in space....
                                                Dim lCacheIdx As Int32 = AddComponentCache(lParent, iParentType, 1, Me.LocX, Me.LocZ, oShield.ObjectID, oShield.ObjTypeID, oShield.Owner.ObjectID, yCacheType)
                                                If lCacheIdx <> -1 Then
                                                    Dim oTmpCache As ComponentCache = goComponentCache(lCacheIdx)
                                                    If oTmpCache Is Nothing = False Then yCompMsg = goMsgSys.GetAddObjectMessage(oTmpCache, GlobalMessageCode.eAddObjectCommand)
                                                End If
                                            End If
                                            oShield = Nothing
                                    End Select

                                    If yCompMsg Is Nothing = False Then
                                        If iParentType = ObjectType.ePlanet Then
                                            CType(Me.ParentObject, Planet).oDomain.DomainSocket.SendData(yCompMsg)
                                        ElseIf iParentType = ObjectType.eSolarSystem Then
                                            CType(Me.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yCompMsg)
                                        End If
                                        gfrmDisplayForm.AddEventLine("Component Dropped")
                                    End If

                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "DeleteEntity.CreateRandomComponentCache: " & ex.Message)
        End Try


        'Make calls to these methods to ensure any cargo or hangar contents are destroyed properly...
        HangarDestroyed(Me.LocX, Me.LocZ)
        CargoDestroyed(Me.LocX, Me.LocZ)

        'Now, check if we are facility...
        If Me.ObjTypeID = ObjectType.eFacility Then
            With CType(Me, Facility)

                If .yProductionType = ProductionType.eTradePost Then
                    If goGTC Is Nothing = False Then goGTC.ClearAllTradeDeliveries(.ObjectID)
                End If

                If .ParentColony Is Nothing = False Then
                    .ParentColony.lChildrenIdx(.ColonyArrayIndex) = -1
                    .ParentColony.oChildren(.ColonyArrayIndex) = Nothing

                    If .yProductionType = ProductionType.eSpaceStationSpecial Then
                        If CType(.ParentColony.ParentObject, Epica_GUID).ObjectID = .ObjectID AndAlso CType(.ParentColony.ParentObject, Epica_GUID).ObjTypeID = .ObjTypeID Then
                            'Remove all facilities within the colony
                            Try
                                .ParentColony.DestroyAllChildrenFacilities()
                            Catch
                            End Try

                            'Remove the colony
                            .ParentColony.DeleteColony(Colony.ColonyLostReason.Destruction)
                        End If
                    ElseIf .ParentColony Is Nothing = False Then
                        .ParentColony.UpdateAllValues(-1)
                        .ParentColony.CheckForColonyDeath()
                    End If
                End If

                'Now, remove any units that have me as a mining facility or refinery
                If lGlobalArrayIndex <> -1 Then
                    For X As Int32 = 0 To glUnitUB
                        If glUnitIdx(X) <> -1 Then
                            Dim oUnit As Unit = goUnit(X)
                            If oUnit Is Nothing = False Then
                                If oUnit.RouteContainsFacility(.ObjectID) = True Then
                                    'clear the route
                                    oUnit.lRouteUB = -1
                                End If
                            End If
                        End If
                    Next X
                End If
            End With
        End If

        Try
            sSQL = "DELETE FROM "
            If Me.ObjTypeID = ObjectType.eUnit Then
                sSQL &= " tblUnit WHERE UnitID = "
            Else : sSQL &= " tblStructure WHERE StructureID = "
            End If
            sSQL &= Me.ObjectID

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                LogEvent(LogEventType.CriticalError, "Unable to delete entity. ID: " & Me.ObjectID & ", Type: " & Me.ObjTypeID)
            End If
        Catch
            LogEvent(LogEventType.CriticalError, "Unable to delete entity. ID: " & Me.ObjectID & ", Type: " & Me.ObjTypeID)
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' This method returns a component cache in the cargo hold of this entity... if it doesn't exist
    '''    it is created on the server and that new creation is returned
    ''' </summary>
    ''' <param name="lComponentID"></param>
    ''' <param name="iComponentTypeID"></param>
    ''' <param name="lQty"></param>
    ''' <param name="lComponentOwnerID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddComponentCacheToCargo(ByVal lComponentID As Int32, ByVal iComponentTypeID As Int16, ByVal lQty As Int32, ByVal lComponentOwnerID As Int32) As ComponentCache
        If Me.ObjTypeID = ObjectType.eFacility Then
            If CType(Me, Facility).ParentColony Is Nothing = False AndAlso Colony.ProductionTypeSharesColonyCargo(Me.yProductionType) = True Then
                Return CType(Me, Facility).ParentColony.AdjustColonyComponentCache(lComponentID, iComponentTypeID, lComponentOwnerID, lQty)
            End If
        End If

        For X As Int32 = 0 To lCargoUB
            If lCargoIdx(X) <> -1 Then
                If oCargoContents(X).ObjTypeID = ObjectType.eComponentCache Then
                    With CType(oCargoContents(X), ComponentCache)
                        If .ComponentID = lComponentID AndAlso .ComponentTypeID = iComponentTypeID Then
                            If .Quantity < 0 Then .Quantity = lQty Else .Quantity += lQty
                            .ComponentOwnerID = lComponentOwnerID
                            Return CType(oCargoContents(X), ComponentCache)
                        End If
                    End With
                End If
            End If
        Next X

        'Ok, if we are here, then the component cache does not already exist... so we'll create it
        Dim lCacheIdx As Int32 = AddComponentCache(Me.ObjectID, Me.ObjTypeID, lQty, 0, 0, lComponentID, iComponentTypeID, lComponentOwnerID, 0)
        If lCacheIdx = -1 Then
            Return Nothing
        Else
            Me.AddCargoRef(CType(goComponentCache(lCacheIdx), Epica_GUID))
            Return goComponentCache(lCacheIdx)
        End If
    End Function

    ''' <summary>
    ''' This method returns a mineral cache in the cargo hold of this entity... if it doesn't exist
    '''   it is created on the server and that new creation is returned
    ''' </summary>
    ''' <param name="lMineralID"></param>
    ''' <param name="lQty"></param>
    ''' <returns> the resulting mineral cache </returns>
    ''' <remarks></remarks>
    Public Function AddMineralCacheToCargo(ByVal lMineralID As Int32, ByVal lQty As Int32) As MineralCache

        If Me.ObjTypeID = ObjectType.eFacility Then
            If CType(Me, Facility).ParentColony Is Nothing = False AndAlso Colony.ProductionTypeSharesColonyCargo(Me.yProductionType) = True Then
                Return CType(Me, Facility).ParentColony.AdjustColonyMineralCache(lMineralID, lQty)
            End If
        End If

        For X As Int32 = 0 To lCargoUB
            If lCargoIdx(X) <> -1 Then
                If oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then
                    With CType(oCargoContents(X), MineralCache)
                        If .oMineral Is Nothing = False AndAlso .oMineral.ObjectID = lMineralID Then
                            If .Quantity < 0 Then .Quantity = lQty Else .Quantity += lQty
                            Return CType(oCargoContents(X), MineralCache)
                        End If
                    End With
                End If
            End If
        Next X

        'Ok, if we are here, then the mineral cache does not already exist... so we'll create it
        Dim lCacheIdx As Int32 = AddMineralCache(Me.ObjectID, Me.ObjTypeID, MineralCacheType.eGround, lQty, lQty, 0, 0, GetEpicaMineral(lMineralID))
        If lCacheIdx = -1 Then
            Return Nothing
        Else
            Dim oCache As MineralCache = goMineralCache(lCacheIdx)
            If oCache Is Nothing = False Then Me.AddCargoRef(CType(oCache, Epica_GUID))
            Return oCache
        End If

    End Function

    ''' <summary>
    ''' This method returns a ammunition cache in the cargo hold of this entity... if it doesn't exist
    '''    it is created on the server and that new creation is returned
    ''' </summary>
    ''' <param name="oWeaponTech"></param>
    ''' <param name="lQty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddAmmunitionCacheToCargo(ByRef oWeaponTech As BaseWeaponTech, ByVal lQty As Int32) As AmmunitionCache
        Dim lWeaponID As Int32 = oWeaponTech.ObjectID
        For X As Int32 = 0 To lCargoUB
            If lCargoIdx(X) <> -1 Then
                If oCargoContents(X).ObjTypeID = ObjectType.eAmmunition Then
                    With CType(oCargoContents(X), AmmunitionCache)
                        If .oWeaponTech.ObjectID = lWeaponID Then
                            .Quantity += lQty
                            Return CType(oCargoContents(X), AmmunitionCache)
                        End If
                    End With
                End If
            End If
        Next X

        'Ok, if we are here, then the cache does not already exist... so we'll create it
        Dim oCache As AmmunitionCache = AddAmmunitionCache(Me.ObjectID, Me.ObjTypeID, lQty, 0, 0, oWeaponTech)
        If oCache Is Nothing = False Then
            Me.AddCargoRef(CType(oCache, Epica_GUID))
        End If
        Return oCache

    End Function

    ''' <summary>
    ''' This method adds the personnel to the cargo hold of this entity. If it doesn't exist, it is added.
    ''' If I am a facility, it will add the personnel to the parent colony's totals
    ''' </summary>
    ''' <param name="iObjTypeID"></param>
    ''' <param name="lQty"></param>
    ''' <remarks></remarks>
    Public Sub AddPersonnelCacheToCargo(ByVal iObjTypeID As Int16, ByVal lQty As Int32)
        If Me.ObjTypeID = ObjectType.eFacility Then
            With CType(Me, Facility)
                If .yProductionType <> ProductionType.eMining AndAlso .ParentColony Is Nothing = False Then '.yProductionType <> ProductionType.eTradePost AndAlso 
                    Select Case iObjTypeID
                        Case ObjectType.eColonists
                            .ParentColony.Population += lQty
                        Case ObjectType.eEnlisted
                            .ParentColony.ColonyEnlisted += lQty
                        Case ObjectType.eOfficers
                            .ParentColony.ColonyOfficers += lQty
                    End Select
                End If
            End With
        End If

        For lCIdx As Int32 = 0 To lCargoUB
            If lCargoIdx(lCIdx) <> -1 AndAlso oCargoContents(lCIdx).ObjTypeID = iObjTypeID Then
                'Because it is personnel, the ObjectID is the quantity - 1
                oCargoContents(lCIdx).ObjectID += lQty
                lCargoIdx(lCIdx) = oCargoContents(lCIdx).ObjectID
                Return
            End If
        Next lCIdx

        Dim oNew As New Epica_GUID
        oNew.ObjectID = lQty - 1
        oNew.ObjTypeID = iObjTypeID
        AddCargoRef(oNew)
    End Sub

    ''' <summary>
    ''' Does the actual process of adding cargo to the cargo bay and ensures that its parent is me
    ''' </summary>
    ''' <param name="oObj"></param>
    ''' <remarks></remarks>
    Public Sub AddCargoRef(ByRef oObj As Epica_GUID)
        If oCargoContents Is Nothing Then lCargoUB = -1

        If oObj.ObjTypeID = ObjectType.eComponentCache Then
            With CType(oObj, ComponentCache)
                If .ComponentOwnerID <> Me.Owner.ObjectID AndAlso .ComponentOwnerID <> 0 AndAlso .GetComponent Is Nothing = False Then
                    'Check for player tech knowledge insert...
                    If Me.Owner.HasTechKnowledge(.ComponentID, .ComponentTypeID, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2) = False Then
                        'add it to the player
                        Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(Me.Owner, .GetComponent, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2, False)
                        'if we added a player tech knowledge, then send it to the player
                        If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
                    End If
                End If

                .ParentObject = Me
            End With
        ElseIf oObj.ObjTypeID = ObjectType.eMineralCache Then
            With CType(oObj, MineralCache)
                If Me.ObjTypeID = ObjectType.eFacility AndAlso yProductionType <> ProductionType.eMining AndAlso .oMineral Is Nothing = False Then
                    If Owner.CheckFirstContactWithMineral(.oMineral.ObjectID) = True Then
                        If .oMineral.ObjectID = 157 Then .Quantity += 30000
                    End If
                End If
                .ParentObject = Me
            End With
        ElseIf oObj.ObjTypeID = ObjectType.eAmmunition Then
            CType(oObj, AmmunitionCache).ParentObject = Me
        End If

        For X As Int32 = 0 To lCargoUB
            If lCargoIdx(X) = -1 Then
                oCargoContents(X) = oObj
                lCargoIdx(X) = oObj.ObjectID
                Return
            End If
        Next X

        lCargoUB += 1
        ReDim Preserve oCargoContents(lCargoUB)
        ReDim Preserve lCargoIdx(lCargoUB)

        oCargoContents(lCargoUB) = oObj
        lCargoIdx(lCargoUB) = oObj.ObjectID
    End Sub

    ''' <summary>
    ''' Does the actual process of adding entity to the hangar bay and ensures that its parent is me
    ''' </summary>
    ''' <param name="oObj"></param>
    ''' <remarks></remarks>
    Public Sub AddHangarRef(ByRef oObj As Epica_GUID)
        If oHangarContents Is Nothing Then lHangarUB = -1
        For X As Int32 = 0 To lHangarUB
            If lHangarIdx(X) = -1 Then
                oHangarContents(X) = oObj
                lHangarIdx(X) = oObj.ObjectID
                Return
            End If
        Next X

        lHangarUB += 1
        ReDim Preserve oHangarContents(lHangarUB)
        ReDim Preserve lHangarIdx(lHangarUB)

        oHangarContents(lHangarUB) = oObj
        lHangarIdx(lHangarUB) = oObj.ObjectID
    End Sub

    Public Function GetHangarEntityCount() As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lHangarUB
            If lHangarIdx(X) <> -1 Then lCnt += 1
        Next X
        Return lCnt
    End Function

    Public Sub HangarDestroyed(ByVal lLocX As Int32, ByVal lLocZ As Int32)
        Dim X As Int32
        Dim Y As Int32
        Dim lIdx As Int32
        Dim iTemp As Int16

        For lHIdx As Int32 = 0 To lHangarUB
            Try
                If lHangarIdx(lHIdx) <> -1 Then
                    Dim oChild As Epica_Entity = CType(oHangarContents(lHIdx), Epica_Entity)

                    'Units are the only things contained in hangars
                    For X = 0 To glUnitUB
                        If glUnitIdx(X) = oChild.ObjectID Then
                            Dim oUnit As Unit = goUnit(X)
                            If oUnit Is Nothing = False Then
                                For Y = 0 To oUnit.EntityDef.lEntityDefMineralUB

                                    Dim yCacheType As Byte = 0
                                    If (oUnit.EntityDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
                                        yCacheType = yCacheType Or MineralCacheType.eGround
                                    ElseIf (oUnit.EntityDef.yChassisType And ChassisType.eNavalBased) <> 0 Then
                                        yCacheType = yCacheType Or MineralCacheType.eNaval
                                    ElseIf (oUnit.EntityDef.yChassisType And ChassisType.eAtmospheric) <> 0 OrElse (oUnit.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                                        yCacheType = yCacheType Or MineralCacheType.eFlying
                                    End If

                                    lIdx = AddMineralCache(CType(ParentObject, Epica_GUID).ObjectID, CType(ParentObject, Epica_GUID).ObjTypeID, yCacheType, oUnit.EntityDef.EntityDefMinerals(Y).lQuantity, oUnit.EntityDef.EntityDefMinerals(Y).lQuantity, CInt((Rnd() * 20) - 10) + lLocX, CInt((Rnd() * 20) - 10) + lLocZ, oUnit.EntityDef.EntityDefMinerals(Y).oMineral)
                                    If lIdx = -1 Then Continue For
                                    AddToQueue(glCurrentCycle + 9000, QueueItemType.eRemoveMineralCache, goMineralCache(lIdx).ObjectID, goMineralCache(lIdx).ObjTypeID, -1, -1, 0, 0, 0, 0)

                                    iTemp = CType(ParentObject, Epica_GUID).ObjTypeID
                                    If iTemp = ObjectType.ePlanet Then
                                        CType(ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(goMineralCache(lIdx), GlobalMessageCode.eAddObjectCommand))
                                    ElseIf iTemp = ObjectType.eSolarSystem Then
                                        CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(goMineralCache(lIdx), GlobalMessageCode.eAddObjectCommand))
                                    End If

                                Next Y
                                oUnit.DeleteEntity(X)
                                glUnitIdx(X) = -1
                                goUnit(X) = Nothing
                            End If
                            'Exit For
                        End If
                    Next X

                    oChild = Nothing
                End If
            Catch
            End Try
        Next lHIdx

        Erase oHangarContents
        Erase lHangarIdx
        lHangarUB = -1
    End Sub

    Public Sub CargoDestroyed(ByVal lLocX As Int32, ByVal lLocZ As Int32)
        Dim X As Int32
        Dim lIdx As Int32
        Dim Y As Int32
        Dim iTemp As Int16

        'Personnel do not leave anything behind... except LOOK AT THE BONES!!!! >:D
        'Ammunition will not leave anything behind either

        If Me.ObjTypeID = ObjectType.eFacility Then
            With CType(Me, Facility)
                If .ParentColony Is Nothing = False AndAlso Colony.ProductionTypeSharesColonyCargo(Me.yProductionType) = True Then
                    .ParentColony.CargoFacilityDestroyed(lLocX, lLocZ, .EntityDef.Cargo_Cap)
                End If
            End With
        End If

        'Go through cargo bay...
        Try
            If lCargoIdx Is Nothing Then Return
            Dim lCurUB As Int32 = Math.Min(lCargoUB, lCargoIdx.GetUpperBound(0))

            For lCIdx As Int32 = 0 To lCargoUB
                'TODO: Units and Cache's and Personnel and Enlisted and Officers... etc... 
                'For now, just units and caches
                If lCargoIdx(lCIdx) <> -1 Then
                    Dim oChild As Epica_GUID = oCargoContents(lCIdx)
                    If oChild Is Nothing Then Continue For

                    If oChild.ObjTypeID = ObjectType.eUnit Then
                        Dim lTmpUnitUB As Int32 = Math.Min(glUnitUB, glUnitIdx.GetUpperBound(0))
                        For X = 0 To lTmpUnitUB
                            If glUnitIdx(X) = oChild.ObjectID Then
                                Dim oUnit As Unit = goUnit(X)
                                glUnitIdx(X) = -1
                                goUnit(X) = Nothing
                                If oUnit Is Nothing = False Then
                                    For Y = 0 To oUnit.EntityDef.lEntityDefMineralUB
                                        Dim yCacheType As Byte = 0
                                        If (oUnit.EntityDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
                                            yCacheType = yCacheType Or MineralCacheType.eGround
                                        ElseIf (oUnit.EntityDef.yChassisType And ChassisType.eNavalBased) <> 0 Then
                                            yCacheType = yCacheType Or MineralCacheType.eNaval
                                        ElseIf (oUnit.EntityDef.yChassisType And ChassisType.eAtmospheric) <> 0 OrElse (oUnit.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                                            yCacheType = yCacheType Or MineralCacheType.eFlying
                                        End If

                                        lIdx = AddMineralCache(CType(ParentObject, Epica_GUID).ObjectID, CType(ParentObject, Epica_GUID).ObjTypeID, yCacheType, oUnit.EntityDef.EntityDefMinerals(Y).lQuantity, oUnit.EntityDef.EntityDefMinerals(Y).lQuantity, CInt((Rnd() * 20) - 10) + lLocX, CInt((Rnd() * 20) - 10) + lLocZ, oUnit.EntityDef.EntityDefMinerals(Y).oMineral)
                                        If lIdx = -1 Then Continue For

                                        AddToQueue(glCurrentCycle + 9000, QueueItemType.eRemoveMineralCache, goMineralCache(lIdx).ObjectID, goMineralCache(lIdx).ObjTypeID, -1, -1, 0, 0, 0, 0)

                                        iTemp = CType(ParentObject, Epica_GUID).ObjTypeID
                                        If iTemp = ObjectType.ePlanet Then
                                            CType(ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(goMineralCache(lIdx), GlobalMessageCode.eAddObjectCommand))
                                        ElseIf iTemp = ObjectType.eSolarSystem Then
                                            CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(goMineralCache(lIdx), GlobalMessageCode.eAddObjectCommand))
                                        End If

                                    Next Y
                                End If
                                'Exit For
                            End If
                        Next X
                    ElseIf oChild.ObjTypeID = ObjectType.eMineralCache Then
                        With CType(oChild, MineralCache)
                            .ParentObject = ParentObject

                            If .ParentObject Is Nothing = False Then
                                .Quantity \= 2      'if it is 0 it will be deleted
                                .LocX = Me.LocX
                                .LocZ = Me.LocZ
                                If .Quantity <> 0 Then
                                    AddToQueue(glCurrentCycle + 9000, QueueItemType.eRemoveMineralCache, .ObjectID, .ObjTypeID, -1, -1, 0, 0, 0, 0)
                                    .DataChanged()
                                    iTemp = CType(ParentObject, Epica_GUID).ObjTypeID
                                    If iTemp = ObjectType.ePlanet Then
                                        CType(ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oChild, GlobalMessageCode.eAddObjectCommand))
                                    ElseIf iTemp = ObjectType.eSolarSystem Then
                                        CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oChild, GlobalMessageCode.eAddObjectCommand))
                                    End If
                                End If
                            Else : .Quantity = 0
                            End If
                        End With
                    ElseIf oChild.ObjTypeID = ObjectType.eComponentCache Then
                        '3 percent chance to drop the component unharmed
                        If Rnd() * 100.0F < 3.0F Then
                            With CType(oChild, ComponentCache)
                                AddToQueue(glCurrentCycle + 9000, QueueItemType.eRemoveComponentCache, .ObjectID, .ObjTypeID, -1, -1, 0, 0, 0, 0)
                                .ParentObject = Me.ParentObject
                                .LocX = Me.LocX
                                .LocZ = Me.LocZ
                                .DataChanged()

                                iTemp = CType(ParentObject, Epica_GUID).ObjTypeID
                                If iTemp = ObjectType.ePlanet Then
                                    CType(ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oChild, GlobalMessageCode.eAddObjectCommand))
                                ElseIf iTemp = ObjectType.eSolarSystem Then
                                    CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oChild, GlobalMessageCode.eAddObjectCommand))
                                End If
                            End With
                        Else
                            'TODO: If the component does not survive, we need to produce the mineral cache's for the component debris
                            '  this *might* be based on the EntityDefMineral
                            CType(oChild, ComponentCache).Quantity = 0       'to delete the cache
                        End If
                    ElseIf oChild.ObjTypeID = ObjectType.eAmmunition Then
                        'Ammunition is always destroyed
                        CType(oChild, AmmunitionCache).Quantity = 0
                    End If
                    Erase oCargoContents
                    Erase lCargoIdx
                    lCargoUB = -1
                End If
            Next lCIdx
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Checks if there is a Hangar Door available for docking that will fit this item based on passed in hull size.
    ''' Returns -2 for failed, -1 for Success, or > 0 for the soonest a door with that size will be available
    ''' </summary>
    ''' <param name="lHullSize"> The Hull Size of the Docker </param>
    ''' <param name="yDockeeManeuver"> Docker's Maneuver value </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AttemptDockCommand(ByVal lHullSize As Int32, ByVal yDockeeManeuver As Byte) As Int32
        Dim oMyEntityDef As Epica_Entity_Def = Nothing

        If gb_IS_TEST_SERVER = True Then Return -1

        If Me.ObjTypeID = ObjectType.eUnit Then
            oMyEntityDef = CType(Me, Unit).EntityDef
        ElseIf Me.ObjTypeID = ObjectType.eFacility Then
            oMyEntityDef = CType(Me, Facility).EntityDef
        Else : Return -2
        End If

        If Me.Owner.bInFullLockDown = True AndAlso Me.yProductionType <> ProductionType.eMining Then
            Return 60
        End If

        If Me.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
            Return -1
        End If

        If oMyEntityDef.HasHangarDoorSize(lHullSize) = False AndAlso (oMyEntityDef.ProductionTypeID And ProductionType.eNavalProduction) <> ProductionType.eNavalProduction Then Return -2
        If oMyEntityDef.lHangarDoorUB < 0 Then Return -2

        If mlNextLaunch Is Nothing OrElse mlNextLaunch.GetUpperBound(0) <> oMyEntityDef.lHangarDoorUB Then
            ReDim Preserve mlNextLaunch(oMyEntityDef.lHangarDoorUB)
        End If

        Dim lSoonest As Int32 = Int32.MaxValue

        'Ok, now... go through my timers
        For X As Int32 = 0 To mlNextLaunch.GetUpperBound(0)
            If oMyEntityDef.GetDoorSize(X) >= lHullSize Then
                If mlNextLaunch(X) < glCurrentCycle Then
                    'Ok, we can launch this one...
                    Dim lVal As Int32 = Math.Max(1500I - ((CInt(yDockeeManeuver) - 1) * Me.Owner.lHangarManMult * 30), 30)
                    Dim fMult As Single = 1.0F
                    With CType(Me.ParentObject, Epica_GUID)
                        fMult = Owner.oBudget.GetDockingDelayMult(.ObjectID, .ObjTypeID)
                    End With

                    Dim fVal As Single = lVal
                    If fMult > 1.0F Then fVal *= fMult

                    Dim lSpecTraitID As elModelSpecialTrait = elModelSpecialTrait.NoSpecialTrait
                    If Me.ObjTypeID = ObjectType.eUnit Then
                        lSpecTraitID = CType(Me, Unit).EntityDef.lHullSpecialTraitID
                    Else
                        lSpecTraitID = CType(Me, Facility).EntityDef.lHullSpecialTraitID
                    End If
                    If lSpecTraitID = elModelSpecialTrait.Launch10 Then
                        fVal *= 0.9F
                    ElseIf lSpecTraitID = elModelSpecialTrait.Launch20 Then
                        fVal *= 0.8F
                    ElseIf lSpecTraitID = elModelSpecialTrait.Launch6 Then
                        fVal *= 0.94F
                    End If

                    mlNextLaunch(X) = CInt(glCurrentCycle + fVal)

                    Return -1
                ElseIf mlNextLaunch(X) < lSoonest Then
                    lSoonest = mlNextLaunch(X)
                End If
            End If
        Next X

        'Nothing is available yet, so return lSoonest
        Return lSoonest
    End Function

    'ONLY UNITS CAN UNDOCK!!!
    Public Sub SendUndockFirstWaypoint(ByRef oUndocker As Epica_Entity)
        Dim yMsg(23) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMsg, 0)

        If RallyPointX = Int32.MinValue OrElse RallyPointZ = Int32.MinValue Then

            With oUndocker
                'TODO: Dummy +3000 Z and MY angle isn't taken into account

                System.BitConverter.GetBytes(.LocX).CopyTo(yMsg, 2)
                If .LocZ > 0 Then
                    System.BitConverter.GetBytes(CInt(.LocZ - 1500)).CopyTo(yMsg, 6)
                Else : System.BitConverter.GetBytes(CInt(.LocZ + 1500)).CopyTo(yMsg, 6)
                End If
                System.BitConverter.GetBytes(.LocAngle).CopyTo(yMsg, 10)

                'The undocker's ParentObject should be set already now

                CType(.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, 12)
                .GetGUIDAsString.CopyTo(yMsg, 18)

                Dim iTemp As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID
                If iTemp = ObjectType.ePlanet Then
                    CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                ElseIf iTemp = ObjectType.eSolarSystem Then
                    CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                End If
            End With
        Else
            System.BitConverter.GetBytes(RallyPointX).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(RallyPointZ).CopyTo(yMsg, 6)
            System.BitConverter.GetBytes(RallyPointA).CopyTo(yMsg, 10)
            System.BitConverter.GetBytes(RallyPointEnvirID).CopyTo(yMsg, 12)
            System.BitConverter.GetBytes(RallyPointEnvirTypeID).CopyTo(yMsg, 16)

            With oUndocker
                .GetGUIDAsString.CopyTo(yMsg, 18)

                Dim iTemp As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID
                If iTemp = ObjectType.ePlanet Then
                    CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                ElseIf iTemp = ObjectType.eSolarSystem Then
                    CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                End If
            End With
        End If

    End Sub

    ''' <summary>
    ''' To be called anytime a unit's Parent changes... this makes sure the speed setting of the unit group is updated
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CheckUpdateUnitGroup()
        'Determine if the entity is a unit
        If Me.ObjTypeID = ObjectType.eUnit Then
            'does it belong to a unit group?
            Dim lUnitGroupID As Int32 = CType(Me, Unit).lFleetID
            If lUnitGroupID <> -1 Then
                'Yes, it does... find the unitgroup
                Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(lUnitGroupID)
                If oUnitGroup Is Nothing = False Then
                    oUnitGroup.CalculateActualSystemSpeed()
                    'Ok, verified, so send an update of the unit's parent to the client
                    If Me.Owner.lConnectedPrimaryID > -1 OrElse Me.Owner.HasOnlineAliases(AliasingRights.eViewBattleGroups) = True Then
                        Dim yMsg(13) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eUpdateEntityParent).CopyTo(yMsg, 0)
                        Me.GetGUIDAsString.CopyTo(yMsg, 2)
                        CType(Me.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, 8)
                        'Me.Owner.oSocket.SendData(yMsg)
                        Me.Owner.SendPlayerMessage(yMsg, False, AliasingRights.eViewBattleGroups)
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SetHangarContentsOwner(ByRef oOwner As Player)
        With Me
            If (.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then
                For lHIdx As Int32 = 0 To .lHangarUB
                    If .lHangarIdx(lHIdx) <> -1 Then
                        'Set its owner to no one
                        CType(.oHangarContents(lHIdx), Epica_Entity).Owner = oOwner
                        'And alert the owning group/fleet
                        If .oHangarContents(lHIdx).ObjTypeID = ObjectType.eUnit Then
                            With CType(.oHangarContents(lHIdx), Unit)
                                'oOwner.blWarpoints -= .EntityDef.WarpointUpkeep
                                If .lFleetID > 0 Then
                                    Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(.lFleetID)
                                    If oUnitGroup Is Nothing = False Then oUnitGroup.RemoveUnit(.ObjectID, True, False)
                                End If
                                .DataChanged()
                            End With
                            CType(.oHangarContents(lHIdx), Epica_Entity).QueueMeToSave()
                        End If
                    End If
                Next lHIdx
            End If
        End With
    End Sub

    'Public Sub ProduceResupply(ByRef oTarget As Epica_Entity)

    '    If ((yProductionType And ProductionType.eProduction) <> 0) OrElse (yProductionType = ProductionType.eCommandCenterSpecial) Then
    '        'Ok, reload ammo
    '        Dim oDef As Epica_Entity_Def
    '        If oTarget.ObjTypeID = ObjectType.eUnit Then
    '            oDef = CType(oTarget, Unit).EntityDef
    '            CType(oTarget, Unit).DataChanged()
    '        Else
    '            oDef = CType(oTarget, Facility).EntityDef
    '            CType(oTarget, Facility).DataChanged()
    '        End If

    '        For X As Int32 = 0 To oDef.WeaponDefUB
    '            If oDef.WeaponDefs(X).mlAmmoCap <> -1 Then
    '                If oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon Is Nothing = False Then
    '                    'Ok, only projectile weapons automatically reload
    '                    If oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon.WeaponClassTypeID = WeaponClassType.eProjectile Then

    '                        Dim oAmmoProdCost As ProductionCost = CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, ProjectileWeaponTech).GetAmmoCost
    '                        If oAmmoProdCost Is Nothing = False Then
    '                            If Me.ObjTypeID = ObjectType.eFacility Then
    '                                If CType(Me, Facility).ParentColony.HasRequiredResources(oAmmoProdCost, CType(Me, Facility), False) = True Then
    '                                    oTarget.lCurrentAmmo(X) = oDef.WeaponDefs(X).mlAmmoCap
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next X
    '    End If
    'End Sub

    'Public Function GetRequestEntityAmmoResponse() As Byte()
    '    Dim lWpnCnt As Int32 = 0
    '    Dim lCargoCnt As Int32 = 0

    '    Dim oDef As Epica_Entity_Def = Nothing
    '    If Me.ObjTypeID = ObjectType.eUnit Then
    '        oDef = CType(Me, Unit).EntityDef
    '    Else : oDef = CType(Me, Facility).EntityDef
    '    End If
    '    If oDef Is Nothing Then Return Nothing

    '    If lCurrentAmmo Is Nothing = False Then lWpnCnt = lCurrentAmmo.GetUpperBound(0) + 1

    '    'Now, for our cargo bay
    '    If (Me.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
    '        For X As Int32 = 0 To lCargoUB
    '            If lCargoIdx(X) <> -1 AndAlso oCargoContents(X).ObjTypeID = ObjectType.eAmmunition Then
    '                lCargoCnt += 1
    '            End If
    '        Next X
    '    End If

    '    'Ok, our Ammo values ARE real time... Now, dim our result
    '    Dim yResp(19 + (lWpnCnt * 17) + (lCargoCnt * 12)) As Byte
    '    Dim lPos As Int32 = 0

    '    'Msgcode (2)
    '    System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityAmmo).CopyTo(yResp, lPos) : lPos += 2
    '    'GUID (6)
    '    Me.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6

    '    'TotalCargoCapacity (4)
    '    System.BitConverter.GetBytes(oDef.Cargo_Cap).CopyTo(yResp, lPos) : lPos += 4

    '    'AvailableCargoCapacity (4)
    '    System.BitConverter.GetBytes(Me.Cargo_Cap).CopyTo(yResp, lPos) : lPos += 4

    '    'iWpnCnt (2)
    '    System.BitConverter.GetBytes(CShort(lWpnCnt)).CopyTo(yResp, lPos) : lPos += 2

    '    'Per Weapon on Unit...
    '    For X As Int32 = 0 To lWpnCnt - 1
    '        'EDW_ID(4)
    '        System.BitConverter.GetBytes(oDef.WeaponDefs(X).ObjectID).CopyTo(yResp, lPos) : lPos += 4
    '        'WeaponTechID(4)
    '        System.BitConverter.GetBytes(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon.ObjectID).CopyTo(yResp, lPos) : lPos += 4
    '        'Arc(1)
    '        yResp(lPos) = oDef.WeaponDefs(X).yArcID : lPos += 1
    '        'MaxAmmo(4)
    '        System.BitConverter.GetBytes(oDef.WeaponDefs(X).mlAmmoCap).CopyTo(yResp, lPos) : lPos += 4
    '        'CurrentAmmo(4)
    '        System.BitConverter.GetBytes(lCurrentAmmo(X)).CopyTo(yResp, lPos) : lPos += 4
    '    Next X

    '    'iCargoCnt(2)
    '    System.BitConverter.GetBytes(CShort(lCargoCnt)).CopyTo(yResp, lPos) : lPos += 2
    '    'Per Ammo Cache in Cargo
    '    For X As Int32 = 0 To lCargoUB
    '        If lCargoIdx(X) <> -1 AndAlso oCargoContents(X).ObjTypeID = ObjectType.eAmmunition Then
    '            With CType(oCargoContents(X), AmmunitionCache)
    '                System.BitConverter.GetBytes(.oWeaponTech.ObjectID).CopyTo(yResp, lPos) : lPos += 4
    '                System.BitConverter.GetBytes(.Quantity).CopyTo(yResp, lPos) : lPos += 4
    '                System.BitConverter.GetBytes(.oWeaponTech.GetAmmoSize).CopyTo(yResp, lPos) : lPos += 4
    '            End With
    '        End If
    '    Next X

    '    Return yResp
    'End Function

    'Public Function HandleLoadAmmo(ByRef oTarget As Epica_Entity, ByVal lWpnID As Int32, ByVal lQty As Int32, ByVal yType As Byte) As Boolean
    '    'Ok, get the target's def
    '    If oTarget Is Nothing Then Return False

    '    Dim oDef As Epica_Entity_Def = Nothing
    '    If oTarget.ObjTypeID = ObjectType.eUnit Then
    '        oDef = CType(oTarget, Unit).EntityDef
    '    Else : oDef = CType(oTarget, Facility).EntityDef
    '    End If
    '    If oDef Is Nothing Then Return False

    '    Dim bResult As Boolean = False

    '    Try
    '        'Ok, now, what are we changing?
    '        Select Case yType
    '            Case 0, 1
    '                '0 = lWpnID is an lEDW_ID meaning we are loading a specific weapon
    '                '1 = lWpnID is a WeaponTechID and we are loading the cargo bay
    '                Dim lMaxQty As Int32 = 0
    '                Dim lWpnQty As Int32 = 0
    '                Dim lCargoQty As Int32 = 0
    '                Dim lWpnTechID As Int32 = -1
    '                Dim lWpnDefIdx As Int32 = -1

    '                If yType = 0 Then
    '                    For X As Int32 = 0 To oDef.WeaponDefUB
    '                        If oDef.WeaponDefs(X).ObjectID = lWpnID Then
    '                            lWpnTechID = oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon.ObjectID
    '                            lWpnDefIdx = X
    '                            lMaxQty = oDef.WeaponDefs(X).mlAmmoCap - oTarget.lCurrentAmmo(X)
    '                            lWpnQty = Math.Min(lMaxQty, lQty)
    '                            lCargoQty = lQty - lWpnQty

    '                            If (oTarget.CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then
    '                                lQty -= lCargoQty
    '                                lCargoQty = 0
    '                            End If
    '                            Exit For
    '                        End If
    '                    Next X
    '                Else
    '                    lWpnTechID = lWpnID
    '                    lMaxQty = 0
    '                    lWpnQty = 0
    '                    lCargoQty = lQty
    '                    If (oTarget.CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then Return False
    '                End If

    '                'Set our max qty to the total qty here in the event that I am a unit and do not have any cargo
    '                lMaxQty = lQty

    '                'Pull from my cargo hold first...
    '                'TODO: need to implement this cargo bay check!
    '                'If (Me.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
    '                For X As Int32 = 0 To lCargoUB
    '                    If lCargoIdx(X) <> -1 AndAlso oCargoContents(X).ObjTypeID = ObjectType.eAmmunition Then
    '                        If CType(oCargoContents(X), AmmunitionCache).oWeaponTech.ObjectID = lWpnTechID Then
    '                            lMaxQty = PullAmmoFromCache(Me, oTarget, CType(oCargoContents(X), AmmunitionCache), lWpnQty, lCargoQty, lWpnDefIdx)
    '                            bResult = True
    '                            Exit For
    '                        End If
    '                    End If
    '                Next X
    '                'End If

    '                'Now, check lMaxQty... if it is not 0
    '                If lMaxQty <> 0 Then
    '                    'Ok, we didn't get everything... are we a unit?
    '                    If Me.ObjTypeID = ObjectType.eUnit Then
    '                        'Now, use lMaxQty as the number of items to produce (if we can produce it)
    '                        If (Me.yProductionType And ProductionType.eProduction) <> 0 OrElse Me.yProductionType = ProductionType.eCommandCenterSpecial Then
    '                            'TODO: Add this as production to the unit's production queue (somehow)
    '                        Else
    '                            If Me.Owner.oSocket Is Nothing = False OrElse Me.Owner.HasOnlineAliases(AliasingRights.eAddProduction Or AliasingRights.eTransferCargo) = True Then
    '                                'TODO: Mining!?!?
    '                                Me.Owner.SendPlayerMessage(CType(Me, Unit).GetEnvirLowResourcesMsg(ProductionType.eMining, 0, 0, 0, -1), False, AliasingRights.eAddProduction Or AliasingRights.eTransferCargo)
    '                            End If
    '                        End If
    '                    Else
    '                        'Okie dokie... we are a facility, pull from the rest of the colony
    '                        Dim oColony As Colony = CType(Me, Facility).ParentColony
    '                        If oColony Is Nothing = False Then
    '                            'Ok, now, go through the children
    '                            For X As Int32 = 0 To oColony.ChildrenUB
    '                                If oColony.lChildrenIdx(X) <> -1 AndAlso (oColony.oChildren(X).yProductionType <> ProductionType.eMining) Then
    '                                    'Ok, this can be a source
    '                                    If (oColony.oChildren(X).CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
    '                                        'Ok, check the cargo bay for ammo
    '                                        With oColony.oChildren(X)
    '                                            For Y As Int32 = 0 To .lCargoUB
    '                                                If .lCargoIdx(Y) <> -1 AndAlso .oCargoContents(Y).ObjTypeID = ObjectType.eAmmunition Then
    '                                                    If CType(oCargoContents(X), AmmunitionCache).oWeaponTech.ObjectID = lWpnTechID Then
    '                                                        lMaxQty = PullAmmoFromCache(CType(oColony.oChildren(X), Epica_Entity), oTarget, CType(oCargoContents(X), AmmunitionCache), lWpnQty, lCargoQty, lWpnDefIdx)
    '                                                        bResult = True
    '                                                        Exit For
    '                                                    End If
    '                                                End If
    '                                            Next Y

    '                                            'Ok, is lMaxQty 0? if so, we are done
    '                                            If lMaxQty = 0 Then Exit For
    '                                        End With
    '                                    End If
    '                                End If
    '                            Next X
    '                        End If

    '                        If lMaxQty <> 0 Then
    '                            'Now, use lMaxQty as the number of items to produce (if we can produce it)
    '                            If (Me.yProductionType And ProductionType.eProduction) <> 0 OrElse Me.yProductionType = ProductionType.eCommandCenterSpecial Then
    '                                'Add this as production to the facilities's production queue (somehow)
    '                                Dim oTech As BaseWeaponTech = CType(Me.Owner.GetTech(lWpnTechID, ObjectType.eWeaponTech), BaseWeaponTech)
    '                                If oTech Is Nothing = False Then
    '                                    Dim fAmmoSize As Single = 0 'oTech.GetAmmoSize()
    '                                    If fAmmoSize <= 0.0F Then fAmmoSize = 0.000001F
    '                                    Select Case oTech.WeaponClassTypeID
    '                                        'TODO: Add more here as needed
    '                                        Case WeaponClassType.eMissile
    '                                            'CType(Me, Facility).AddAmmoProdItem(lWpnID, CType(oTech, MissileWeaponTech).GetAmmoCost, lMaxQty, oTarget.ObjectID, oTarget.ObjTypeID, yType, fAmmoSize, lWpnTechID)
    '                                        Case WeaponClassType.eProjectile
    '                                            'CType(Me, Facility).AddAmmoProdItem(lWpnID, CType(oTech, ProjectileWeaponTech).GetAmmoCost, lMaxQty, oTarget.ObjectID, oTarget.ObjTypeID, yType, fAmmoSize, lWpnTechID)
    '                                    End Select
    '                                End If
    '                            Else
    '                                If Me.Owner.oSocket Is Nothing = False OrElse Me.Owner.HasOnlineAliases(AliasingRights.eAddProduction Or AliasingRights.eTransferCargo) = True Then
    '                                    'TODO: Mining?!?
    '                                    Me.Owner.SendPlayerMessage(CType(Me, Facility).ParentColony.GetLowResourcesMsg(ProductionType.eMining, 0, 0, -1, -1), False, AliasingRights.eAddProduction Or AliasingRights.eTransferCargo)
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            Case Else       'lWpnID and Qty are not used and instead, we are loading all weapons to their max capacity
    '                'Easy enough, we'll just recurse...
    '                For X As Int32 = 0 To oDef.WeaponDefUB
    '                    If oDef.WeaponDefs(X).mlAmmoCap > -1 Then
    '                        If oTarget.lCurrentAmmo(X) <> oDef.WeaponDefs(X).mlAmmoCap Then
    '                            'DO NOT OrElse this!!!
    '                            bResult = bResult Or HandleLoadAmmo(oTarget, oDef.WeaponDefs(X).ObjectID, oDef.WeaponDefs(X).mlAmmoCap - oTarget.lCurrentAmmo(X), 0)
    '                        End If
    '                    End If
    '                Next X
    '        End Select
    '    Catch ex As Exception
    '        'Do Nothing
    '        LogEvent(LogEventType.CriticalError, "HandleLoadAmmo: " & ex.Message)
    '    End Try

    '    Return bResult
    'End Function

    'Public Shared Function PullAmmoFromCache(ByRef oFrom As Epica_Entity, ByVal oTarget As Epica_Entity, ByRef oCache As AmmunitionCache, ByRef lWpnQty As Int32, ByRef lCargoQty As Int32, ByVal lWpnDefIdx As Int32) As Int32
    '    Dim lQty As Int32 = lWpnQty + lCargoQty
    '    Dim lMaxQty As Int32 = lQty

    '    With oCache
    '        Dim fAmmoSize As Single = 0 ' .oWeaponTech.GetAmmoSize()
    '        If fAmmoSize <= 0.0F Then fAmmoSize = 0.000001F

    '        If .Quantity >= lQty Then
    '            'Ok, we have enough
    '            If lWpnQty <> 0 Then oTarget.lCurrentAmmo(lWpnDefIdx) += lWpnQty

    '            'Ensure the target has enough cargo space and that the cargo hold is available
    '            If lCargoQty <> 0 Then
    '                Dim fCargoSpace As Single = fAmmoSize * lCargoQty
    '                If (oTarget.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
    '                    If oTarget.Cargo_Cap - fCargoSpace < 0 Then
    '                        'not enough room in cargo bay, so fill as much as we can
    '                        lCargoQty = CInt(Math.Floor(oTarget.Cargo_Cap / fAmmoSize))
    '                    End If

    '                    If lCargoQty <> 0 Then
    '                        'oTarget.Cargo_Cap -= CInt(lCargoQty * fAmmoSize)
    '                        oTarget.AddAmmunitionCacheToCargo(.oWeaponTech, lCargoQty)
    '                    End If
    '                Else : lCargoQty = 0
    '                End If
    '            End If
    '            .Quantity -= (lWpnQty + lCargoQty)
    '            'oFrom.Cargo_Cap += CInt((lWpnQty + lCargoQty) * fAmmoSize)

    '            lMaxQty = 0         'set this to 0 for the next phase of our check
    '        Else
    '            'ok, we do not have enough, load what we have now...
    '            If .Quantity >= lWpnQty Then
    '                If lWpnQty <> 0 Then
    '                    oTarget.lCurrentAmmo(lWpnDefIdx) += lWpnQty
    '                    .Quantity -= lWpnQty
    '                    'oFrom.Cargo_Cap += CInt(lWpnQty * fAmmoSize)
    '                    lWpnQty = 0
    '                End If

    '                'Now, load the cargo with as much as possible
    '                lMaxQty = Math.Min(.Quantity, lCargoQty)

    '                ' Ensure the target has enough cargo space and that the cargo hold is available
    '                If lMaxQty <> 0 Then
    '                    Dim fCargoSpace As Single = fAmmoSize * lMaxQty
    '                    If (oTarget.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
    '                        If oTarget.Cargo_Cap - fCargoSpace < 0 Then
    '                            'Not enough room in cargo bay, so fill as much as we can
    '                            lMaxQty = CInt(Math.Floor(oTarget.Cargo_Cap / fAmmoSize))
    '                            lCargoQty = lMaxQty
    '                        End If

    '                        oTarget.AddAmmunitionCacheToCargo(.oWeaponTech, lMaxQty)
    '                        lCargoQty -= lMaxQty
    '                        .Quantity -= lMaxQty

    '                        'oFrom.Cargo_Cap += CInt(lMaxQty * fAmmoSize)
    '                        'oTarget.Cargo_Cap -= CInt(lMaxQty * fAmmoSize)
    '                    Else
    '                        lMaxQty = 0
    '                        lCargoQty = 0
    '                    End If
    '                End If
    '            Else
    '                oTarget.lCurrentAmmo(lWpnDefIdx) += .Quantity
    '                lWpnQty -= .Quantity
    '                'oFrom.Cargo_Cap += CInt(.Quantity * fAmmoSize)
    '                .Quantity = 0
    '            End If

    '            'set lMaxQty to what we are missing...
    '            lMaxQty = lWpnQty + lCargoQty
    '        End If
    '    End With

    '    Return lMaxQty
    'End Function

    'Public Sub HandleAmmoProduced(ByRef oProd As EntityProduction)

    '    'TODO: We need to handle cargo_cap better here
    '    Dim oTech As BaseWeaponTech = CType(Me.Owner.GetTech(oProd.lAmmoWpnTechID, ObjectType.eWeaponTech), BaseWeaponTech)
    '    If oTech Is Nothing Then
    '        oProd.lProdCount = 0
    '        Return
    '    End If
    '    Me.AddAmmunitionCacheToCargo(oTech, 1)

    '    If oProd.lProdCount <= 1 Then
    '        'Me.Cargo_Cap -= CInt(Math.Ceiling(oTech.GetAmmoSize() * oProd.lAmmoQty))

    '        If oProd.lExtendedID > 0 AndAlso oProd.iExtendedTypeID > 0 Then
    '            Dim oEntity As Epica_Entity = Nothing
    '            If oProd.iExtendedTypeID = ObjectType.eUnit Then
    '                oEntity = GetEpicaUnit(oProd.lExtendedID)
    '            ElseIf oProd.iExtendedTypeID = ObjectType.eFacility Then
    '                oEntity = GetEpicaFacility(oProd.lExtendedID)
    '            End If
    '            If oEntity Is Nothing Then Return

    '            'check my hangar contents
    '            With CType(oEntity.ParentObject, Epica_GUID)
    '                If .ObjectID = Me.ObjectID AndAlso .ObjTypeID = Me.ObjTypeID Then
    '                    HandleLoadAmmo(oEntity, oProd.ProductionID, oProd.lAmmoQty, oProd.yExtendedType)
    '                    'If HandleLoadAmmo(oEntity, oProd.ProductionID, oProd.lAmmoQty, oProd.yExtendedType) = True Then
    '                    '    If oEntity.Owner.oSocket Is Nothing = False Then
    '                    '        oEntity.Owner.oSocket.SendData(oEntity.GetRequestEntityAmmoResponse())
    '                    '    End If
    '                    'End If
    '                End If
    '            End With
    '        End If
    '    End If

    'End Sub

    Public Function GetRootParentEnvir() As Epica_GUID
        'entities can be contained in: Units, Facilities, Systems, Planets, Galaxies (and maybe trade?)
        '  this will return the ENVIRONMENT (Planet or System) they reside in
        Try
            Dim oParent As Epica_GUID = CType(Me.ParentObject, Epica_GUID)
            Dim lIterations As Int32 = 0
            While oParent Is Nothing = False
                lIterations += 1
                If lIterations > 20 Then Return Nothing
                If oParent.ObjTypeID = ObjectType.ePlanet Then
                    Return oParent
                ElseIf oParent.ObjTypeID = ObjectType.eSolarSystem Then
                    Return oParent
                ElseIf oParent.ObjTypeID = ObjectType.eUnit Then
                    oParent = CType(CType(oParent, Epica_Entity).ParentObject, Epica_GUID)
                ElseIf oParent.ObjTypeID = ObjectType.eFacility Then
                    Dim oTmpParent As Epica_GUID = CType(CType(oParent, Epica_Entity).ParentObject, Epica_GUID)
                    If oTmpParent Is Nothing = False AndAlso oTmpParent.ObjTypeID = ObjectType.eSolarSystem AndAlso CType(oParent, Facility).yProductionType = ProductionType.eSpaceStationSpecial Then Return oParent
                    oParent = oTmpParent
                Else : Exit While
                End If
            End While
        Catch
            Return Nothing
        End Try
        Return Nothing
    End Function

#Region "  Agent Effects versus this Entity  "
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
        bNeedsSaved = True
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
            LogEvent(LogEventType.CriticalError, "Epica_Entity.EffectValid: " & ex.Message)
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
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save agent effect " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Protected Function GetSaveAgentEffectsText() As String
        Dim sSQL As String

        Try
            Dim oSB As New System.Text.StringBuilder

            'sSQL = "DELETE FROM tblAgentEffects WHERE EffectedItemID = " & Me.ObjectID & " AND EffectedItemTypeID = " & Me.ObjTypeID
            'oSB.AppendLine(sSQL)

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
                            oSB.AppendLine(sSQL)
                        Catch ex As Exception
                            LogEvent(LogEventType.CriticalError, "Unable to Save Agent Effect: " & ex.Message)
                        End Try
                    End With
                End If
            Next X

            Return oSB.ToString
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save agent effect " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Public Function AuditRemoveEffects(ByVal lTargetNum As Int32) As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To mlAgentEffectUB
            If EffectValid(X) = True Then
                If Rnd() * 100 < lTargetNum Then
                    RemoveAgentEffect(X)
                    lCnt += 1
                End If
            End If
        Next X
        Return lCnt
    End Function
#End Region

#Region "  Repair Orders  "
    'entities can now repair, not just facilities

    Public Sub HandleRepairItem(ByRef oEntity As Epica_Entity, ByRef oDef As Epica_Entity_Def, ByVal lRepairItem As Int32)
        'Now, what are we repairing?
        Select Case lRepairItem
            Case -2, -3, -4, -5
                'Armor side
                Dim lCurrent As Int32
                Dim lMax As Int32

                If lRepairItem = -2 Then
                    lCurrent = oEntity.Q1_HP : lMax = oDef.Q1_MaxHP
                ElseIf lRepairItem = -3 Then
                    lCurrent = oEntity.Q2_HP : lMax = oDef.Q2_MaxHP
                ElseIf lRepairItem = -4 Then
                    lCurrent = oEntity.Q3_HP : lMax = oDef.Q3_MaxHP
                Else : lCurrent = oEntity.Q4_HP : lMax = oDef.Q4_MaxHP
                End If

                If lCurrent <> lMax AndAlso lMax <> 0 Then
                    Dim lRepNeeded As Int32 = 0
                    If oDef.oPrototype.oArmorTech Is Nothing Then lRepNeeded = lMax - lCurrent Else lRepNeeded = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
                    AddRepairItem(lRepairItem, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                End If

            Case -6
                'structure
                Dim lCurrent As Int32 = oEntity.Structure_HP
                Dim lMax As Int32 = oDef.Structure_MaxHP

                If lCurrent <> lMax AndAlso lMax <> 0 Then
                    Dim lRepNeeded As Int32 = lMax - lCurrent
                    AddRepairItem(lRepairItem, lRepNeeded, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                End If
            Case -7
                'all defensive values
                Dim lCurrent As Int32 = oEntity.Structure_HP
                Dim lMax As Int32 = oDef.Structure_MaxHP
                If lCurrent <> lMax AndAlso lMax <> 0 Then
                    Dim lRepNeeded As Int32 = lMax - lCurrent
                    AddRepairItem(-6, lRepNeeded, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                End If
                lCurrent = oEntity.Q1_HP : lMax = oDef.Q1_MaxHP
                If lCurrent <> lMax AndAlso lMax <> 0 Then
                    Dim lRepNeeded As Int32 = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
                    AddRepairItem(-2, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                End If
                lCurrent = oEntity.Q2_HP : lMax = oDef.Q2_MaxHP
                If lCurrent <> lMax AndAlso lMax <> 0 Then
                    Dim lRepNeeded As Int32 = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
                    AddRepairItem(-3, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                End If
                lCurrent = oEntity.Q3_HP : lMax = oDef.Q3_MaxHP
                If lCurrent <> lMax AndAlso lMax <> 0 Then
                    Dim lRepNeeded As Int32 = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
                    AddRepairItem(-4, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                End If
                lCurrent = oEntity.Q4_HP : lMax = oDef.Q4_MaxHP
                If lCurrent <> lMax AndAlso lMax <> 0 Then
                    Dim lRepNeeded As Int32 = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
                    AddRepairItem(-5, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                End If
            Case Int32.MaxValue
                'repair all components
                Dim lStatus As Int32 = oEntity.CurrentStatus
                Dim lDefStatus As Int32 = 0
                For X As Int32 = 0 To oDef.lSideCrits.GetUpperBound(0)
                    lDefStatus = lDefStatus Or oDef.lSideCrits(X)
                Next X

                If (lDefStatus And elUnitStatus.eEngineOperational) <> 0 Then
                    If (lStatus And elUnitStatus.eEngineOperational) = 0 Then
                        AddRepairItem(elUnitStatus.eEngineOperational, 1, oDef.oPrototype.oEngineTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                    End If
                End If
                If (lDefStatus And elUnitStatus.eFuelBayOperational) <> 0 Then
                    If (lStatus And elUnitStatus.eFuelBayOperational) = 0 Then
                        AddRepairItem(elUnitStatus.eFuelBayOperational, 1, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                    End If
                End If
                If (lDefStatus And elUnitStatus.eRadarOperational) <> 0 Then
                    If (lStatus And elUnitStatus.eRadarOperational) = 0 Then
                        AddRepairItem(elUnitStatus.eRadarOperational, 1, oDef.oPrototype.oRadarTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                    End If
                End If
                If (lDefStatus And elUnitStatus.eShieldOperational) <> 0 Then
                    If (lStatus And elUnitStatus.eShieldOperational) = 0 Then
                        AddRepairItem(elUnitStatus.eShieldOperational, 1, oDef.oPrototype.oShieldTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                    End If
                End If
                If (lDefStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                    If (lStatus And elUnitStatus.eCargoBayOperational) = 0 Then
                        AddRepairItem(elUnitStatus.eCargoBayOperational, 1, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                    End If
                End If
                If (lDefStatus And elUnitStatus.eHangarOperational) <> 0 Then
                    If (lStatus And elUnitStatus.eHangarOperational) = 0 Then
                        AddRepairItem(elUnitStatus.eHangarOperational, 1, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                    End If
                End If

                'Now for weapons
                If (lDefStatus And elUnitStatus.eAftWeapon1) <> 0 Then
                    If (lStatus And elUnitStatus.eAftWeapon1) = 0 Then
                        For X As Int32 = 0 To oDef.WeaponDefUB
                            If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eAftWeapon1 Then
                                AddRepairItem(elUnitStatus.eAftWeapon1, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                            End If
                        Next X
                    End If
                End If
                If (lDefStatus And elUnitStatus.eAftWeapon2) <> 0 Then
                    If (lStatus And elUnitStatus.eAftWeapon2) = 0 Then
                        For X As Int32 = 0 To oDef.WeaponDefUB
                            If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eAftWeapon2 Then
                                AddRepairItem(elUnitStatus.eAftWeapon2, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                            End If
                        Next X
                    End If
                End If
                If (lDefStatus And elUnitStatus.eLeftWeapon1) <> 0 Then
                    If (lStatus And elUnitStatus.eLeftWeapon1) = 0 Then
                        For X As Int32 = 0 To oDef.WeaponDefUB
                            If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eLeftWeapon1 Then
                                AddRepairItem(elUnitStatus.eLeftWeapon1, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                            End If
                        Next X
                    End If
                End If
                If (lDefStatus And elUnitStatus.eLeftWeapon2) <> 0 Then
                    If (lStatus And elUnitStatus.eLeftWeapon2) = 0 Then
                        For X As Int32 = 0 To oDef.WeaponDefUB
                            If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eLeftWeapon2 Then
                                AddRepairItem(elUnitStatus.eLeftWeapon2, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                            End If
                        Next X
                    End If
                End If
                If (lDefStatus And elUnitStatus.eForwardWeapon1) <> 0 Then
                    If (lStatus And elUnitStatus.eForwardWeapon1) = 0 Then
                        For X As Int32 = 0 To oDef.WeaponDefUB
                            If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eForwardWeapon1 Then
                                AddRepairItem(elUnitStatus.eForwardWeapon1, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                            End If
                        Next X
                    End If
                End If
                If (lDefStatus And elUnitStatus.eForwardWeapon2) <> 0 Then
                    If (lStatus And elUnitStatus.eForwardWeapon2) = 0 Then
                        For X As Int32 = 0 To oDef.WeaponDefUB
                            If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eForwardWeapon2 Then
                                AddRepairItem(elUnitStatus.eForwardWeapon2, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                            End If
                        Next X
                    End If
                End If
                If (lDefStatus And elUnitStatus.eRightWeapon1) <> 0 Then
                    If (lStatus And elUnitStatus.eRightWeapon1) = 0 Then
                        For X As Int32 = 0 To oDef.WeaponDefUB
                            If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eRightWeapon1 Then
                                AddRepairItem(elUnitStatus.eRightWeapon1, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                            End If
                        Next X
                    End If
                End If
                If (lDefStatus And elUnitStatus.eRightWeapon2) <> 0 Then
                    If (lStatus And elUnitStatus.eRightWeapon2) = 0 Then
                        For X As Int32 = 0 To oDef.WeaponDefUB
                            If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eRightWeapon2 Then
                                AddRepairItem(elUnitStatus.eRightWeapon2, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                            End If
                        Next X
                    End If
                End If

            Case Is > 0
                'repair a component
                Dim lDefStatus As Int32 = 0
                For X As Int32 = 0 To oDef.lSideCrits.GetUpperBound(0)
                    lDefStatus = lDefStatus Or oDef.lSideCrits(X)
                Next X
                If (lDefStatus And lRepairItem) <> 0 Then
                    If (oEntity.CurrentStatus And lRepairItem) = 0 Then
                        Select Case lRepairItem
                            Case elUnitStatus.eEngineOperational
                                AddRepairItem(lRepairItem, 1, oDef.oPrototype.oEngineTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                            Case elUnitStatus.eRadarOperational
                                AddRepairItem(lRepairItem, 1, oDef.oPrototype.oRadarTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                            Case elUnitStatus.eShieldOperational
                                AddRepairItem(lRepairItem, 1, oDef.oPrototype.oShieldTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                            Case elUnitStatus.eCargoBayOperational, elUnitStatus.eHangarOperational, elUnitStatus.eFuelBayOperational
                                AddRepairItem(lRepairItem, 1, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
                            Case Else
                                Dim lTemp As Int32

                                lTemp = elUnitStatus.eAftWeapon1
                                For X As Int32 = 0 To oDef.WeaponDefUB
                                    If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
                                        AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                                    End If
                                Next X
                                lTemp = elUnitStatus.eAftWeapon2
                                For X As Int32 = 0 To oDef.WeaponDefUB
                                    If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
                                        AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                                    End If
                                Next X
                                lTemp = elUnitStatus.eForwardWeapon1
                                For X As Int32 = 0 To oDef.WeaponDefUB
                                    If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
                                        AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                                    End If
                                Next X
                                lTemp = elUnitStatus.eForwardWeapon2
                                For X As Int32 = 0 To oDef.WeaponDefUB
                                    If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
                                        AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                                    End If
                                Next X
                                lTemp = elUnitStatus.eRightWeapon1
                                For X As Int32 = 0 To oDef.WeaponDefUB
                                    If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
                                        AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                                    End If
                                Next X
                                lTemp = elUnitStatus.eRightWeapon2
                                For X As Int32 = 0 To oDef.WeaponDefUB
                                    If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
                                        AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                                    End If
                                Next X
                                lTemp = elUnitStatus.eLeftWeapon1
                                For X As Int32 = 0 To oDef.WeaponDefUB
                                    If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
                                        AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                                    End If
                                Next X
                                lTemp = elUnitStatus.eLeftWeapon2
                                For X As Int32 = 0 To oDef.WeaponDefUB
                                    If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
                                        AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
                                    End If
                                Next X
                        End Select
                    End If
                End If

            Case Else
                'tell the parent to cancel repair for this entity 
                If Me.ObjTypeID = ObjectType.eFacility Then
                    With CType(Me, Facility)
                        For X As Int32 = 0 To .ProductionUB
                            If .ProductionIdx(X) <> -1 AndAlso .Production(X).ProductionTypeID = ObjectType.eRepairItem Then
                                If .Production(X).lExtendedID = oEntity.ObjectID AndAlso .Production(X).iExtendedTypeID = oEntity.ObjTypeID Then
                                    .ProductionIdx(X) = -1
                                End If
                            End If
                        Next X
                    End With
                End If
        End Select
    End Sub

    Public Sub RepairCompleted(ByRef oProd As EntityProduction)
        Dim oEntity As Epica_Entity = CType(GetEpicaObject(oProd.lExtendedID, oProd.iExtendedTypeID), Epica_Entity)
        'Verify that it is still in the hangar (for facility repairs)
        Dim bFound As Boolean = False
        Dim bRepairMyself As Boolean = False
        If oEntity Is Nothing Then Return

        If oEntity.ObjectID = Me.ObjectID AndAlso oEntity.ObjTypeID = Me.ObjTypeID Then
            bFound = True
            bRepairMyself = True
        Else
            For X As Int32 = 0 To Me.lHangarUB
                If Me.lHangarIdx(X) = oProd.lExtendedID AndAlso Me.oHangarContents(X).ObjTypeID = oProd.iExtendedTypeID Then
                    bFound = True
                    Exit For
                End If
            Next X
        End If

        If oEntity Is Nothing OrElse bFound = False Then
            'Ok, that entity no longer exists... or no longer is in my hangar
            ClearProductionByExtended(oProd.lExtendedID, oProd.iExtendedTypeID)
            Return
        End If

        Dim oDef As Epica_Entity_Def
        If oEntity.ObjTypeID = ObjectType.eUnit Then
            oDef = CType(oEntity, Unit).EntityDef
        Else : oDef = CType(oEntity, Facility).EntityDef
        End If
        If oDef Is Nothing Then Return
        If oDef.oPrototype Is Nothing Then Return

        Dim lQty As Int32 = 1

        Select Case oProd.ProductionID
            Case -2 'q1
                lQty = oEntity.Q1_HP
                If oDef.oPrototype.oArmorTech Is Nothing Then oEntity.Q1_HP += 10 Else oEntity.Q1_HP += oDef.oPrototype.oArmorTech.lHPPerPlate * 10
                If oEntity.Q1_HP > oDef.Q1_MaxHP Then oEntity.Q1_HP = oDef.Q1_MaxHP
                lQty = oEntity.Q1_HP - lQty
            Case -3 'q2
                lQty = oEntity.Q2_HP
                If oDef.oPrototype.oArmorTech Is Nothing Then oEntity.Q2_HP += 10 Else oEntity.Q2_HP += oDef.oPrototype.oArmorTech.lHPPerPlate * 10
                If oEntity.Q2_HP > oDef.Q2_MaxHP Then oEntity.Q2_HP = oDef.Q2_MaxHP
                lQty = oEntity.Q2_HP - lQty
            Case -4 'q3
                lQty = oEntity.Q3_HP
                If oDef.oPrototype.oArmorTech Is Nothing Then oEntity.Q3_HP += 10 Else oEntity.Q3_HP += oDef.oPrototype.oArmorTech.lHPPerPlate * 10
                If oEntity.Q3_HP > oDef.Q3_MaxHP Then oEntity.Q3_HP = oDef.Q3_MaxHP
                lQty = oEntity.Q3_HP - lQty
            Case -5 'q4
                lQty = oEntity.Q4_HP
                If oDef.oPrototype.oArmorTech Is Nothing Then oEntity.Q4_HP += 10 Else oEntity.Q4_HP += oDef.oPrototype.oArmorTech.lHPPerPlate * 10
                If oEntity.Q4_HP > oDef.Q4_MaxHP Then oEntity.Q4_HP = oDef.Q4_MaxHP
                lQty = oEntity.Q4_HP - lQty
            Case -6
                'structure
                oEntity.Structure_HP = oDef.Structure_MaxHP
            Case Is > 0
                'repair a component
                oEntity.CurrentStatus = oEntity.CurrentStatus Or oProd.ProductionID
        End Select

        If bRepairMyself = True Then
            'Ok, how do we do this?
            Dim iTemp As Int16 = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
            If iTemp = ObjectType.ePlanet OrElse iTemp = ObjectType.eSolarSystem Then
                Dim yMsg(15) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eRepairCompleted).CopyTo(yMsg, 0)
                oEntity.GetGUIDAsString.CopyTo(yMsg, 2)

                System.BitConverter.GetBytes(oProd.ProductionID).CopyTo(yMsg, 8)
                System.BitConverter.GetBytes(lQty).CopyTo(yMsg, 12)

                If iTemp = ObjectType.ePlanet Then
                    CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                Else : CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                End If
            End If
        End If
 

    End Sub

    Private Sub AddRepairItem(ByVal lRepairItemID As Int32, ByVal lQuantity As Int32, ByRef oComponent As Epica_Tech, ByVal blEntityDefPointsRequired As Int64, ByRef oEntity As Epica_Entity)
        If Me.ObjTypeID = ObjectType.eFacility Then
            CType(Me, Facility).fac_DONOTCALL_AddRepairItem(lRepairItemID, lQuantity, oComponent, blEntityDefPointsRequired, oEntity)
        Else
            With Me
                .bProducing = False
                .CurrentProduction = Nothing
                .CurrentProduction = CreateRepairProduction(lRepairItemID, lQuantity, oComponent, blEntityDefPointsRequired, oEntity)
            End With
            With CType(Me, Unit)

                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To glUnitUB
                    If glUnitIdx(X) = .ObjectID Then
                        lIdx = X
                        Exit For
                    End If
                Next X

                If .PopulateRequirements() = True Then
                    .bProducing = True
                    AddEntityProducing(lIdx, ObjectType.eUnit, .ObjectID)
                End If
            End With
        End If
    End Sub

    Protected Function CreateRepairProduction(ByVal lRepairItemID As Int32, ByVal lQuantity As Int32, ByRef oComponent As Epica_Tech, ByVal blEntityDefPointsRequired As Int64, ByRef oEntity As Epica_Entity) As EntityProduction
        Dim oProd As EntityProduction = New EntityProduction()
        With oProd
            .oParent = Me
            .ProductionID = lRepairItemID
            .ProductionTypeID = ObjectType.eRepairItem
            .OrderNumber = 254
            .PointsProduced = 0
            .lProdCount = lQuantity

            .ProdCost = New ProductionCost()
            .ProdCost.ColonistCost = 0

            .lExtendedID = oEntity.ObjectID
            .iExtendedTypeID = oEntity.ObjTypeID

            Dim lProdCostQty As Int32 = 1

            If oComponent Is Nothing Then
                .ProdCost.CreditCost = 1000
                .ProdCost.PointsRequired = 5000
                lProdCostQty = 10
                .lProdCount = CInt(Math.Ceiling(lQuantity / 10.0F))
            ElseIf lRepairItemID = -6 Then
                .lProdCount = 1
                If oComponent.GetProductionCost.ItemCostUB <> -1 AndAlso oComponent.Owner.ObjectID > 0 Then
                    Dim lBaseMins As Int32 = oComponent.GetProductionCost.ItemCosts(0).QuantityNeeded
                    Dim lMinID As Int32 = oComponent.GetProductionCost.ItemCosts(0).ItemID
                    Dim lVal As Int32 = 0
                    If oEntity.ObjTypeID = ObjectType.eUnit Then
                        lVal = CType(oEntity, Unit).EntityDef.Structure_MaxHP
                    ElseIf oEntity.ObjTypeID = ObjectType.eFacility Then
                        lVal = CType(oEntity, Facility).EntityDef.Structure_MaxHP
                    End If
                    If lVal = 0 Then
                        .ProdCost.PointsRequired = 500 * lQuantity
                        .ProdCost.CreditCost = 1000 * lQuantity
                    Else
                        Dim fMult As Single = CSng(oEntity.Structure_HP / lVal)
                        If fMult > 1 Then fMult = 1.0F
                        fMult = 1.0F - fMult
                        lVal = CInt(Math.Ceiling(lBaseMins * fMult))

                        .ProdCost.AddProductionCostItem(lMinID, ObjectType.eMineral, lVal)
                        .ProdCost.PointsRequired = CLng(oComponent.GetProductionCost.PointsRequired * fMult)
                        .ProdCost.CreditCost = CLng(oComponent.GetProductionCost.CreditCost * fMult)
                    End If
                Else
                    .ProdCost.PointsRequired = 500 * lQuantity
                    .ProdCost.CreditCost = 1000 * lQuantity
                End If

            ElseIf oComponent.ObjTypeID = ObjectType.eArmorTech Then
                .ProdCost.CreditCost = 1000
                .ProdCost.PointsRequired = 5000
                If lQuantity < 10 Then
                    lProdCostQty = lQuantity
                Else
                    lProdCostQty = 10
                End If
                .lProdCount = CInt(Math.Ceiling(lQuantity / 10.0F))
            ElseIf oComponent.ObjTypeID = ObjectType.eEngineTech Then
                .ProdCost.CreditCost = 100000
                .ProdCost.PointsRequired = blEntityDefPointsRequired \ 5
            ElseIf oComponent.ObjTypeID = ObjectType.eShieldTech Then
                .ProdCost.CreditCost = 160000
                .ProdCost.PointsRequired = blEntityDefPointsRequired \ 6
            ElseIf oComponent.ObjTypeID = ObjectType.eWeaponTech Then
                .ProdCost.CreditCost = 500000
                .ProdCost.PointsRequired = oComponent.GetProductionCost.PointsRequired \ 5
            Else
                .ProdCost.CreditCost = 100000
                .ProdCost.PointsRequired = 10000
            End If

            If (Me.yProductionType And ProductionType.eProduction) = 0 Then
                'Ok, need to scale for the differences between facility production abilities
                Dim fMultiplier As Single = 1.0F
                Select Case Me.yProductionType
                    Case ProductionType.ePowerCenter
                        fMultiplier = Me.mlProdPoints / 100.0F
                    Case ProductionType.eCommandCenterSpecial, ProductionType.eSpaceStationSpecial
                        fMultiplier = Me.mlProdPoints / 10.0F
                    Case ProductionType.eColonists
                        fMultiplier = Me.mlProdPoints / 150.0F
                    Case ProductionType.eWareHouse
                        fMultiplier = Me.mlProdPoints / 200.0F
                    Case ProductionType.eMining
                        fMultiplier = Me.mlProdPoints / 200.0F
                End Select
                .ProdCost.PointsRequired = CLng(.ProdCost.PointsRequired * fMultiplier)
                '.ProdCost.CreditCost = CLng(.ProdCost.CreditCost * fMultiplier)
            End If

            Dim yReqProd As Byte
            If oEntity.ObjTypeID = ObjectType.eUnit Then
                yReqProd = CType(oEntity, Unit).EntityDef.RequiredProductionTypeID
            Else : yReqProd = CType(oEntity, Facility).EntityDef.RequiredProductionTypeID
            End If
            'If this facility can produce this unit, it gets a half bonus
            If (yReqProd And Me.yProductionType) = yReqProd Then
                .ProdCost.CreditCost \= 2
                .ProdCost.PointsRequired \= 2
            End If

            Dim lSpeedMult As Int32 = 100 - Owner.oSpecials.yRepairSpeedBonus
            If lSpeedMult <> 100 Then
                .ProdCost.PointsRequired = CLng(.ProdCost.PointsRequired * (lSpeedMult / 100.0F))
            End If

            If lRepairItemID <> -6 AndAlso oComponent Is Nothing = False AndAlso oComponent.ObjTypeID <> ObjectType.eHullTech AndAlso oComponent.Owner.ObjectID > 0 Then .ProdCost.AddProductionCostItem(oComponent.ObjectID, oComponent.ObjTypeID, lProdCostQty)
        End With

        Return oProd
    End Function

    Private Sub ClearProductionByExtended(ByVal lExtendedID As Int32, ByVal iExtendedTypeID As Int16)
        If CurrentProduction Is Nothing = False Then
            If CurrentProduction.lExtendedID = lExtendedID AndAlso CurrentProduction.iExtendedTypeID = iExtendedTypeID Then
                CurrentProduction.lProdCount = 0
            End If
        End If

        If Me.ObjTypeID = ObjectType.eFacility Then
            With CType(Me, Facility)
                For X As Int32 = 0 To .ProductionUB
                    If .ProductionIdx(X) <> -1 AndAlso .Production(X).lExtendedID = lExtendedID AndAlso .Production(X).iExtendedTypeID = iExtendedTypeID Then
                        .ProductionIdx(X) = -1
                    End If
                Next X
            End With
        End If
        
    End Sub


#End Region

    Public Sub ProcessChangePrimaryServers(ByVal oDest As Epica_GUID)
        'first, we need to force save me
        If Me.ObjTypeID = ObjectType.eFacility Then
            CType(Me, Facility).SaveObject()
        ElseIf Me.ObjTypeID = ObjectType.eUnit Then
            CType(Me, Unit).SaveObject()
        Else : Return
        End If

        For X As Int32 = 0 To lCargoUB
            If lCargoIdx(X) > -1 Then
                Select Case oCargoContents(X).ObjTypeID
                    Case ObjectType.eUnit
                        CType(oCargoContents(X), Unit).SaveObject()
                        Dim lCurUB As Int32 = glUnitUB
                        For Y As Int32 = 0 To lCurUB
                            If glUnitIdx(Y) = oCargoContents(X).ObjectID Then
                                Dim oUnit As Unit = goUnit(Y)
                                If oUnit Is Nothing = False Then
                                    oUnit.SaveObject()
                                End If
                                glUnitIdx(Y) = -1
                                Exit For
                            End If
                        Next Y
                    Case ObjectType.eMineralCache
                        With CType(oCargoContents(X), MineralCache)
                            .SaveObject()
                            glMineralIdx(.lServerIndex) = -1
                        End With
                    Case ObjectType.eComponentCache
                        With CType(oCargoContents(X), ComponentCache)
                            .SaveObject()
                            glComponentCacheIdx(X) = -1
                        End With
                End Select
            End If
        Next X

        For X As Int32 = 0 To lHangarUB
            If lHangarIdx(X) > -1 Then
                Select Case oHangarContents(X).ObjTypeID
                    Case ObjectType.eUnit
                        With CType(oHangarContents(X), Unit)
                            .SaveObject()
                            Dim lCurUB As Int32 = glUnitUB
                            For Y As Int32 = 0 To lCurUB
                                If glUnitIdx(Y) = .ObjectID Then
                                    glUnitIdx(Y) = -1
                                    Exit For
                                End If
                            Next Y
                        End With
                End Select
            End If
        Next X

        If Me.ObjTypeID = ObjectType.eFacility Then
            If glFacilityIdx(CType(Me, Facility).ServerIndex) = Me.ObjectID Then
                glFacilityIdx(CType(Me, Facility).ServerIndex) = -1
                RemoveLookupFacility(Me.ObjectID, Me.ObjTypeID)
            End If
        ElseIf Me.ObjTypeID = ObjectType.eUnit Then
            Dim lCurUB As Int32 = glUnitUB
            For X As Int32 = 0 To lCurUB
                If glUnitIdx(X) = Me.ObjectID Then
                    glUnitIdx(X) = -1
                    Exit For
                End If
            Next X
        End If

        Dim yMsg(23) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eEntityChangingPrimary).CopyTo(yMsg, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(Me.Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        'Dest is going to be a unit, facility, planet or system
        oDest.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        'if the dest is a unit or facility, we need to get that entity's parent guid
        If oDest.ObjTypeID = ObjectType.eUnit OrElse oDest.ObjTypeID = ObjectType.eFacility Then
            CType(CType(oDest, Epica_Entity).ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, lPos)
        End If
        lPos += 6

        goMsgSys.SendMsgToOperator(yMsg)

    End Sub

    Public Function RecycleMe(ByVal EntityDef As Epica_Entity_Def) As Collection
        Dim colResult As New Collection

        If EntityDef Is Nothing = False AndAlso EntityDef.oPrototype Is Nothing = False Then
            'Refund:
            Dim oRand As New Random()
            With EntityDef.oPrototype
                '1) 95% chance of recovery of undamaged componets includeing engines, radar weapons, sheilds. Each compnenet would get a individual "roll" to see if it recoverd.  
                If .oEngineTech Is Nothing = False Then
                    If (CurrentStatus And elUnitStatus.eEngineOperational) <> 0 Then
                        If oRand.Next(0, 100) < 95 Then
                            Dim oNew As New RecycledPartsList
                            oNew.lObjID = .oEngineTech.ObjectID
                            oNew.iObjTypeID = .oEngineTech.ObjTypeID
                            oNew.lQuantity = 1
                            If .oEngineTech.Owner Is Nothing = False Then oNew.lExtended = .oEngineTech.Owner.ObjectID
                            colResult.Add(oNew)
                        End If
                    End If
                End If
                If .oRadarTech Is Nothing = False Then
                    If (CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then
                        If oRand.Next(0, 100) < 95 Then
                            Dim oNew As New RecycledPartsList
                            oNew.lObjID = .oRadarTech.ObjectID
                            oNew.iObjTypeID = .oRadarTech.ObjTypeID
                            oNew.lQuantity = 1
                            If .oRadarTech.Owner Is Nothing = False Then oNew.lExtended = .oRadarTech.Owner.ObjectID
                            colResult.Add(oNew)
                        End If
                    End If
                End If
                If .oShieldTech Is Nothing = False Then
                    If (CurrentStatus And elUnitStatus.eShieldOperational) <> 0 Then
                        If oRand.Next(0, 100) < 95 Then
                            Dim oNew As New RecycledPartsList
                            oNew.lObjID = .oShieldTech.ObjectID
                            oNew.iObjTypeID = .oShieldTech.ObjTypeID
                            oNew.lQuantity = 1
                            If .oShieldTech.Owner Is Nothing = False Then oNew.lExtended = .oShieldTech.Owner.ObjectID
                            colResult.Add(oNew)
                        End If
                    End If
                End If

                '2) 15-25% of intact (no damage) armor plates  
                If .oArmorTech Is Nothing = False Then
                    Dim lQ1Remain As Int32 = Q1_HP
                    Dim lQ2Remain As Int32 = Q2_HP
                    Dim lQ3Remain As Int32 = Q3_HP
                    Dim lQ4Remain As Int32 = Q4_HP

                    'Now, get our plate count
                    lQ1Remain = lQ1Remain \ .oArmorTech.lHPPerPlate
                    lQ2Remain = lQ2Remain \ .oArmorTech.lHPPerPlate
                    lQ3Remain = lQ3Remain \ .oArmorTech.lHPPerPlate
                    lQ4Remain = lQ4Remain \ .oArmorTech.lHPPerPlate

                    Dim lSpecTraitID As elModelSpecialTrait = elModelSpecialTrait.NoSpecialTrait
                    Dim oModelDef As ModelDef = goModelDefs.GetModelDef(.oHullTech.ModelID)
                    If oModelDef Is Nothing = False Then
                        lSpecTraitID = CType(oModelDef.lSpecialTraitID, elModelSpecialTrait)
                    End If

                    If lSpecTraitID = elModelSpecialTrait.Armor1000 Then
                        lQ1Remain \= 10
                        lQ2Remain \= 10
                        lQ3Remain \= 10
                        lQ4Remain \= 10
                    End If
                    If .oHullTech Is Nothing = False AndAlso .oHullTech.yTypeID = 2 AndAlso .oHullTech.ySubTypeID <> 9 Then
                        lQ1Remain \= 10
                        lQ2Remain \= 10
                        lQ3Remain \= 10
                        lQ4Remain \= 10
                    End If

                    Dim lTotalArmorRefund As Int32 = lQ1Remain + lQ2Remain + lQ3Remain + lQ4Remain
                    lTotalArmorRefund \= 4

                    If lTotalArmorRefund > 0 Then
                        Dim oNew As New RecycledPartsList
                        oNew.lObjID = .oArmorTech.ObjectID
                        oNew.iObjTypeID = .oArmorTech.ObjTypeID
                        oNew.lQuantity = lTotalArmorRefund
                        If .oArmorTech.Owner Is Nothing = False Then oNew.lExtended = .oArmorTech.Owner.ObjectID
                        colResult.Add(oNew)
                    End If
                End If
            End With

            With EntityDef
                For X As Int32 = 0 To .WeaponDefUB
                    If .WeaponDefs(X) Is Nothing = False AndAlso .WeaponDefs(X).oWeaponDef Is Nothing = False AndAlso .WeaponDefs(X).oWeaponDef.RelatedWeapon Is Nothing = False Then
                        If (CurrentStatus And .WeaponDefs(X).lEntityStatusGroup) <> 0 Then
                            If oRand.Next(0, 100) < 95 Then
                                Dim oNew As New RecycledPartsList
                                oNew.lObjID = .WeaponDefs(X).oWeaponDef.RelatedWeapon.ObjectID
                                oNew.iObjTypeID = .WeaponDefs(X).oWeaponDef.RelatedWeapon.ObjTypeID
                                oNew.lQuantity = 1
                                If .WeaponDefs(X).oWeaponDef.RelatedWeapon.Owner Is Nothing = False Then oNew.lExtended = .WeaponDefs(X).oWeaponDef.RelatedWeapon.Owner.ObjectID
                                colResult.Add(oNew)
                            End If
                        End If
                    End If
                Next X
            End With
        End If

        Return colResult
    End Function
End Class

Public Class RecycledPartsList
    Public lObjID As Int32
    Public iObjTypeID As Int16
    Public lQuantity As Int32
    Public lExtended As Int32
End Class
