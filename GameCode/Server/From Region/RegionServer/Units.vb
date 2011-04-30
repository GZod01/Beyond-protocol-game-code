Option Strict On

Public Class WeaponDef
    Inherits Epica_GUID

    Public WeaponName(19) As Byte
    Public yWeaponType As Byte       'indicates what type of weapon this is, how it fires, etc.
    Public WpnGroup As Byte         'indicates whether the weapon is Primary, Secondary, or Point Defense, etc...
    Public ROF_Delay As Int16       'Added to NextFireCycle to determine when the unit fires again
    Public iDefRange As Int16       'Value from the WeaponDef
    Public lRange As Int32          'Value adjusted for the Entity Def it belngs to
    Public Accuracy As Byte

    'For Performance
    Public lPD_AccVal As Int32
    Public lNrm_AccVal As Int32

    Public ParentObject As Object

    Public ArcID As Byte            'what facing the weapon shoots from

    Public lEntityStatusGroup As Int32      'this is the Bit-Wise identifier of what group this weapon def belongs to in CurrentStatus (elUnitStatus)
    Public lFirePowerRating As Int32

    Public PiercingMaxMinRange As Int32
    Public PiercingMinDmg As Int32
    Public PiercingMaxDmg As Int32
    Public ImpactMaxMinRange As Int32
    Public ImpactMinDmg As Int32
    Public ImpactMaxDmg As Int32
    Public BeamMaxMinRange As Int32
    Public BeamMinDmg As Int32
    Public BeamMaxDmg As Int32
    Public ECMMaxMinRange As Int32
    Public ECMMinDmg As Int32
    Public ECMMaxDmg As Int32
    Public FlameMaxMinRange As Int32
    Public FlameMinDmg As Int32
    Public FlameMaxDmg As Int32
    Public ChemicalMaxMinRange As Int32
    Public ChemicalMinDmg As Int32
    Public ChemicalMaxDmg As Int32

    'Public MaxRangeAccuracy As Single       'stored value to aid in computations
    Public blOverallMaxDmg As Int64

    Public AOERange As Byte

    Public lAmmoCap As Int32

    Public WeaponSpeed As Byte = 100
    Public Maneuver As Byte = 0

    Public fPulseDegradation As Single = 0.0F

    Public lEntityWpnDefID As Int32         'For example, UnitDefWeapon.UDW_ID

    'These are optimized values set when the entitydef is optimized
    Public fOneOverRange As Single
    Public fOneOverOptMinusRange As Single
    Public fOneOverWeaponSpeed As Single

    Public yAOEDmgType As Byte = 0

    Public lStructHP As Int32       'used for missiles... if the weapon is a missile, this is the beam min dmg... the beam min dmg is then set to 0

    Public Sub SetAOEDmgType()
        If AOERange = 0 Then Return

        If yWeaponType < WeaponType.eMetallicProjectile_Gold Then
            'this is any energy beam weapon - pulse, solid, flicker
            yAOEDmgType = eyAOEDmgType.eBeamDmg
        ElseIf yWeaponType < WeaponType.eBomb_Green Then
            'projectile and missile weapons - can be explosive, chemical or magnetic unless all are 0's at which point it is impact
            Dim lMax As Int32 = Math.Max(Math.Max(FlameMaxDmg, ECMMaxDmg), ChemicalMaxDmg)
            If lMax = 0 Then
                If ImpactMaxDmg > 0 Then
                    yAOEDmgType = eyAOEDmgType.eImpactDmg
                Else
                    yAOEDmgType = 0
                End If
            Else
                If lMax = FlameMaxDmg Then
                    yAOEDmgType = eyAOEDmgType.eFlameDmg
                ElseIf lMax = ECMMaxDmg Then
                    yAOEDmgType = eyAOEDmgType.eECMDmg
                Else
                    yAOEDmgType = eyAOEDmgType.eChemicalDmg
                End If
            End If
        Else
            'bombs for now
            yAOEDmgType = 255       'set to 255, all dmg types apply
        End If

    End Sub
End Class

Public Class ObjectCache
    Inherits Epica_GUID

    Public Const l_CACHE_MSG_LEN As Int32 = 19

    Public LocX As Int32
    Public LocZ As Int32

    Public lDetail1 As Int32        'Mineral Cache is MineralID, ComponentCache is ComponentID

    Public CacheTypeID As Byte      'Mineral Cache is the mineral cache type, component cache is ComponentTypeID (as a byte)

    Public lQuantity As Int32

    Public lDetail2 As Int32        'Mineral Cache is Concentration, ComponentCache is ComponentOwnerID

    Private mySmallSendString() As Byte

    Public Function GetObjectAsSmallString() As Byte()
        If mbStringReady = False Then
            ReDim mySmallSendString(l_CACHE_MSG_LEN - 1)  '0 to 18 = 19 bytes
            GetGUIDAsString.CopyTo(mySmallSendString, 0)
            System.BitConverter.GetBytes(LocX).CopyTo(mySmallSendString, 6)
            System.BitConverter.GetBytes(LocZ).CopyTo(mySmallSendString, 10)
            System.BitConverter.GetBytes(lDetail1).CopyTo(mySmallSendString, 14)
            mySmallSendString(18) = CacheTypeID
            mbStringReady = True
        End If
        Return mySmallSendString
    End Function

    Public Function GetObjectAsString() As Byte()
        Dim yResp(26) As Byte

        GetGUIDAsString.CopyTo(yResp, 0)
        System.BitConverter.GetBytes(LocX).CopyTo(yResp, 6)
        System.BitConverter.GetBytes(LocZ).CopyTo(yResp, 10)
        System.BitConverter.GetBytes(lDetail1).CopyTo(yResp, 14)
        yResp(18) = CacheTypeID
        System.BitConverter.GetBytes(lQuantity).CopyTo(yResp, 19)
        System.BitConverter.GetBytes(lDetail2).CopyTo(yResp, 23)

        Return yResp
    End Function

End Class

'Public Class Mineral
'    Inherits Epica_GUID

'End Class

Public Class Epica_Entity
    Inherits Epica_GUID

    'Both UNITS and FACILITIES fall under this class
    Public Enum eyHullTypeDmgMod As Byte
        eAllOthers = 0
        eEscorts = 1
        eCorvettes = 2
        eFrigates = 3
        eDestroyers = 4
        eCruisers = 5
        eBattlecruisers = 6
        eBattleships = 7
        eStructures = 8
    End Enum
    Private Shared mfHullTypeDmgModMult() As Single = {1.0F, 0.5F, 0.25F, 0.125F, 0.0625F, 0.03125F, 0.015625F, 0.007813F, 0.003906F}

    Public UnitName(19) As Byte

    Public ServerIndex As Int32     ' the index of the server array, used internally on the server for targeting

    Public LocX As Int32
    Public LocY As Int32 = 0      'Used for LOS and such, when an object moves, its LocY is set to 0, when LOS needs it, it calculates this value
    Public LocZ As Int32
    Public LocAngle As Short    'Degree angle of the unit's forward facing

    Public MoveLocX As Single   'used for movement... we use a single for precision
    Public MoveLocZ As Single   'used for movement... we use a single for precision

    Public VelX As Single
    Public VelZ As Single

    Public DestX As Int32
    Public DestZ As Int32
    Public DestAngle As Short       'used for temporary usage

    Public TrueDestAngle As Short   'when done moving, this is the final angle the unit will have

    Public TotalVelocity As Single

    Public Fuel_Cap As Int32

    Public Exp_Level As Byte

    'Now, these will likely directly mimic the values in the entity def, but we want to keep them here so we have the ability to change them
    Public ToHitMod As Int16 = 0
    Public TurnAmount As Byte
    Public TurnAmountTimes100 As Int16
    Public MaxSpeed As Byte
    Public Acceleration As Single
    Public DamageMod As Int32 = 0
    Public Maneuver As Byte

    'Formation-Specifics...
    Public yFormationMaxSpeed As Byte = 0
    Public fFormationAcceleration As Single = 0.0F
    Public yFormationManeuver As Byte = 0
    Public yFormationTurnAmount As Byte = 0
    Public iFormationTurnAmount100 As Int16 = 0
    'end of formation specifics

    Public lGridIndex As Int32         'the grid index of the environment object
    Public lGridEntityIdx As Int32     'a result from the AddEntity call to the EnvirGrid object (used for removing later)
    Public lSmallSectorID As Int32
    Public lTinyX As Int32
    Public lTinyZ As Int32

    Public lEntityDefServerIndex As Int32 = -1      'this is the server index of the Global Entity Def Array
    Public lEntityDefServerID As Int32 = -1
    Public lModelRangeOffset As Int32 = 0

    'Ok, we track the next fire cycles here for weapons... because it is on a per instance basis
    Public lWpnNextFireCycle() As Int32
    Public lWpnAmmoCnt() As Int32

    Public bForceAggressionTest As Boolean = False
    Public bNewAddedEntity As Boolean = False

    Public iTargetingTactics As Int16   'uses the eiTacticalAttrs enum
    Public iCombatTactics As Int32 = 513      'uses the eiBehaviorPatterns enum

    'When using EVADE as a combat tactic, these values are the LocX and LocZ of the Evade to...
    '  when using Dock With target, these values are the Entity ID and TypeID of the dock target
    Public lExtendedCT_1 As Int32
    Public lExtendedCT_2 As Int32
    'Used for Maneuver pattern
    Public MinSpeed As Single = Single.MinValue
    Public myManPointNum As Byte
    Public iManApproach As Int16
    Public lLastManCalcCycle As Int32 = Int32.MinValue
    Public yResetMoveByPlayer As Byte = 0

    Public lPrimaryTargetServerIdx As Int32 = -1
    Public lTargetsServerIdx(3) As Int32       'index for each facing's target..
    Public lPreviousTargetServerIdx(3) As Int32     'MSC 05/13/08 - added for new fire weapon optimization
    Public lPrimaryTargetRange As Int32
    Public Ranges(3) As Int32

    Public lFighterTargetServerIdx(3) As Int32
    Public lFighterTargetRange(3) As Int32

#If EXTENSIVELOGGING = 1 Then
    Public Sub DisplayTargetMsg()
        Try
            gfrmDisplayForm.AddEventLine("TargetUpdate: " & Me.ObjectID & ", " & Me.ObjTypeID & " (" & Me.ServerIndex & ") has targets: " & lPrimaryTargetServerIdx & ", " & lTargetsServerIdx(0) & ", " & lTargetsServerIdx(1) & ", " & lTargetsServerIdx(2) & ", " & lTargetsServerIdx(3))
        Catch
        End Try
    End Sub
#End If

    Public lNextFireWpnEvent As Int32 = Int32.MaxValue      'my shot going out
    Public lNextWpnEvent As Int32 = Int32.MaxValue          'shot against me

    Public ParentEnvir As Envir     'the environment this unit currently resides
    Public Owner As Player
    Public lOwnerID As Int32        'for not having to check Owner

    'This array contains server indices of the goEntity array that are currently targeting this object
    Public lTargetedByIdx() As Int32
    Public lTargetedByID() As Int32     'for verification purposes
    Public lTargetedByUB As Int32 = -1

    Public CurrentStatus As Int32   'a bit-wise long indicating the status of the unit currently (see CurrentStatus.xls)
    Public LastCycleMoved As Int32  'the last cycle number that the movement engine calculated this unit

    Private mlShieldHP As Int32        'amount of shield hps currently remaining
    Public lLastShieldRechargeCycle As Int32
    Public Property ShieldHP() As Int32
        Get
            If (Me.CurrentStatus And elUnitStatus.eShieldOperational) <> 0 Then
                If lLastShieldRechargeCycle <> glCurrentCycle Then
                    If lEntityDefServerIndex <> -1 Then
                        Dim oDef As Epica_Entity_Def = goEntityDefs(lEntityDefServerIndex)
                        If oDef Is Nothing = False Then
                            If oDef.Shield_MaxHP <> mlShieldHP AndAlso oDef.ShieldRechargeFreq <> 0 Then
                                Dim lRechargeCycles As Int32 = (glCurrentCycle - lLastShieldRechargeCycle) \ oDef.ShieldRechargeFreq
                                If lRechargeCycles > 0 Then
                                    Try
                                        Dim lTemp As Int32 = (oDef.ShieldRecharge * lRechargeCycles)
                                        mlShieldHP += lTemp
                                    Catch
                                        mlShieldHP = oDef.Shield_MaxHP
                                    End Try
                                    lRechargeCycles *= oDef.ShieldRechargeFreq
                                    lLastShieldRechargeCycle += lRechargeCycles
                                End If
                            Else : lLastShieldRechargeCycle = glCurrentCycle
                            End If
                            If mlShieldHP > oDef.Shield_MaxHP Then mlShieldHP = oDef.Shield_MaxHP
                        End If
                    End If
                End If
            Else
                mlShieldHP = 0
                lLastShieldRechargeCycle = glCurrentCycle
            End If
            If mlShieldHP < 0 Then mlShieldHP = 0
            Return mlShieldHP
        End Get
        Set(ByVal value As Int32)
            mlShieldHP = value
        End Set
    End Property

    'Public ArmorHP(3) As Int32      'amount of armor hps (per quadrant) remaining
    Private mlArmorHP(3) As Int32
    Private mlSideMaxHP() As Int32 = Nothing
    Private Shared mbDoSideHPCheck As Boolean = True
    Public Property ArmorHP(ByVal lSide As Int32) As Int32
        Get
            Return mlArmorHP(lSide)
        End Get
        Set(ByVal value As Int32)
            If mbDoSideHPCheck = True Then
                If mlSideMaxHP Is Nothing Then
                    'get our definition
                    If lEntityDefServerIndex <> -1 Then
                        If glEntityDefIdx(lEntityDefServerIndex) <> -1 Then
                            Dim oCurDef As Epica_Entity_Def = goEntityDefs(lEntityDefServerIndex)
                            If oCurDef Is Nothing = False Then
                                ReDim mlSideMaxHP(3)
                                For Y As Int32 = 0 To 3
                                    mlSideMaxHP(Y) = oCurDef.Armor_MaxHP(Y)
                                Next Y
                            End If
                        End If
                    End If
                Else
                    If mlSideMaxHP(lSide) < value Then
                        value = mlSideMaxHP(lSide)
                    End If
                End If
            End If
            mlArmorHP(lSide) = value
        End Set
    End Property
    Public StructuralHP As Int32

    Private mySideHasCrits(3) As Byte   'indicates whether a side has valid critical hit locations remaining (performance)

    Public oWpnEvents() As WeaponEvent
    Public lWpnEventUB As Int32 = -1
    Public yWpnEventUsed() As Byte

    'Public bPlayerMoveRequestPending As Boolean = False
    Private mbPlayerMoveRequestPending As Boolean = False
    Public Property bPlayerMoveRequestPending() As Boolean
        Get
            Return mbPlayerMoveRequestPending
        End Get
        Set(ByVal value As Boolean)
            mbPlayerMoveRequestPending = value
            If value = True Then
                lTetherPointX = Int32.MinValue
                lTetherPointZ = Int32.MinValue
            End If
        End Set
    End Property

    Public bAIMoveRequestPending As Boolean = False

    Public bDecelerating As Boolean = False

    Public yChangeEnvironments As Byte = 0

    Public yProductionType As Byte

    Public bFinalizeStopEvent As Boolean = False        'indicates that this unit has been requested to finalize its stop event in the event that it disengages

    Public CPUsage As Int32 = 0         'command point usage

    Public lUpdatePrimaryWithHPRequest As Int32 = Int32.MinValue
    'For UnitHPMessageUpdates
    Public PreviousStatus As Int32
    Public PreviousArmorPerc(3) As Byte
    Public PreviousStructPerc As Byte
    Public PreviousShieldPerc As Byte

    Public bAIControlled As Boolean = False
    Public Pre_AI_LocX As Int32
    Public Pre_AI_LocZ As Int32

    Public bInCombatRegister As Boolean = False
    Public bInMovementRegister As Boolean = False

    Public lLastFighterCritCycle As Int32 = 0

    Public lTetherPointX As Int32 = Int32.MinValue
    Public lTetherPointZ As Int32 = Int32.MinValue
    Public lLaunchedFromID As Int32 = Int32.MinValue
    Public iLaunchedFromTypeID As Int16 = Int16.MinValue
    Public lLastEngagement As Int32 = 0
    Public lLastTetherCheck As Int32 = 0

    Public bInAILaunchAll As Boolean = False
    Public lLastForceAggressionTest As Int32 = 0        'Used for managing performance - no unit can force an aggression test within a certain threshold from the previous aggression test

#Region "  Incoming Missile Management  "
    Public lIncMissileIdx() As Int32        'index in the missile manager's missile array
    Public lIncMissileID() As Int32         'ID of the missile for comparison
    Public lIncMissileUB As Int32 = -1

    Public lIncMissileCnt As Int32 = 0

    Private mlMissileTargetIdx(3) As Int32
    Private mlMissileTargetRange(3) As Int32
    Private mlMissileTargetID(3) As Int32
    Private mlLastFindMissileTargetCycle As Int32 = 0
    Public Function GetMissileTarget(ByVal yArc As Byte, ByVal lRange As Int32, ByRef lMissileTargetRange As Int32, ByVal oCurDef As Epica_Entity_Def) As MissileMgr.Missile
        If yArc = UnitArcs.eAllArcs Then
            'ok, different... does any of our sides have a target?
            For X As Int32 = 0 To 3
                If mlMissileTargetIdx(X) > -1 Then
                    If goMissileMgr.mlMissileIdx(mlMissileTargetIdx(X)) = mlMissileTargetID(X) Then
                        'If mlMissileTargetRange(X) <= lRange Then
                        lMissileTargetRange = mlMissileTargetRange(X)
                        Return goMissileMgr.moMissile(mlMissileTargetIdx(X))
                        'End If
                    Else
                        mlMissileTargetIdx(X) = -1
                    End If
                End If
            Next X
        Else
            'does the arc already have a target?
            If mlMissileTargetIdx(yArc) > -1 Then
                If goMissileMgr.mlMissileIdx(mlMissileTargetIdx(yArc)) = mlMissileTargetID(yArc) Then

                    If mlLastFindMissileTargetCycle < glCurrentCycle Then
                        'ok, verify the missile's side...
                        Dim oMissile As MissileMgr.Missile = goMissileMgr.moMissile(mlMissileTargetIdx(yArc))
                        If oMissile Is Nothing = False Then
                            Dim lAdjust() As Int32 = ParentEnvir.lGridIdxAdjust
                            If ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                Dim lModVal As Int32 = lGridIndex Mod ParentEnvir.lGridsPerRow
                                If lModVal = 0 Then
                                    'I'm on the left edge, use the left edge adjust
                                    lAdjust = ParentEnvir.lLeftEdgeGridIdxAdjust
                                ElseIf lModVal = ParentEnvir.lGridsPerRow - 1 Then
                                    'I'm on the right edge, use the right edge adjust
                                    lAdjust = ParentEnvir.lRightEdgeGridIdxAdjust
                                End If
                            End If

                            Dim lGrid As Int32 = -1
                            If lGridIndex = oMissile.lGridIndex Then
                                lGrid = 0
                            Else
                                For Y As Int32 = 0 To lAdjust.GetUpperBound(0)
                                    If lGridIndex + lAdjust(Y) = oMissile.lGridIndex Then
                                        lGrid = Y
                                        Exit For
                                    End If
                                Next Y
                            End If
                            If lGrid <> -1 Then
                                With oMissile

                                    Dim lTemp As Int32 = giRelativeSmall(lGrid, lSmallSectorID, oMissile.lSmallSectorID)
                                    Dim lRelTinyX As Int32
                                    Dim lRelTinyZ As Int32
                                    If lTemp <> -1 Then
                                        lRelTinyX = glBaseRelTinyX(lTemp) + .lTinyX + (gl_HALF_FINAL_PER_ROW - lTinyX)
                                        If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                            lRelTinyZ = glBaseRelTinyZ(lTemp) + .lTinyZ + (gl_HALF_FINAL_PER_ROW - lTinyZ)
                                            If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                                Dim ySideCanFire As Byte = WhatSideCanFire(Me, lRelTinyX, lRelTinyZ)

                                                If ySideCanFire = yArc Then
                                                    lMissileTargetRange = mlMissileTargetRange(yArc)
                                                    Return goMissileMgr.moMissile(mlMissileTargetIdx(yArc))
                                                End If
                                            End If
                                        End If
                                    End If
                                End With
                            End If

                        End If
                        mlMissileTargetIdx(yArc) = -1
                    Else
                        lMissileTargetRange = mlMissileTargetRange(yArc)
                        Return goMissileMgr.moMissile(mlMissileTargetIdx(yArc))
                    End If
                End If
            End If
        End If

        'If we are here, find a new target...
        If mlLastFindMissileTargetCycle < glCurrentCycle Then
            mlLastFindMissileTargetCycle = glCurrentCycle
            'Determine what adjust array to use (for map wrapping of planets)
            Dim lAdjust() As Int32 = ParentEnvir.lGridIdxAdjust
            If ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                Dim lModVal As Int32 = lGridIndex Mod ParentEnvir.lGridsPerRow
                If lModVal = 0 Then
                    'I'm on the left edge, use the left edge adjust
                    lAdjust = ParentEnvir.lLeftEdgeGridIdxAdjust
                ElseIf lModVal = ParentEnvir.lGridsPerRow - 1 Then
                    'I'm on the right edge, use the right edge adjust
                    lAdjust = ParentEnvir.lRightEdgeGridIdxAdjust
                End If
            End If

            Dim lMaxRadarRange As Int32 = oCurDef.lOptRadarRange + lJamRangeMod + 1     'we add 1 so we do not need to do a <=

            'Let's verify our targets here too
            For X As Int32 = 0 To 3
                If mlMissileTargetIdx(X) > -1 Then
                    If mlMissileTargetID(X) <> goMissileMgr.mlMissileIdx(mlMissileTargetIdx(X)) Then
                        mlMissileTargetIdx(X) = -1
                        mlMissileTargetID(X) = -1
                        mlMissileTargetRange(X) = 999999
                    Else
                        Dim oMissile As MissileMgr.Missile = goMissileMgr.moMissile(mlMissileTargetIdx(X))
                        If oMissile Is Nothing = True Then
                            mlMissileTargetIdx(X) = -1
                            mlMissileTargetID(X) = -1
                            mlMissileTargetRange(X) = 999999
                            Continue For
                        End If

                        Dim lEnvirDist As Int32 = Int32.MaxValue
                        Dim lGrid As Int32
                        With oMissile
                            If .lGridIndex <> lGridIndex Then
                                'Determine the grid
                                lGrid = -1
                                For Y As Int32 = 0 To lAdjust.GetUpperBound(0)
                                    If lGridIndex + lAdjust(Y) = .lGridIndex Then
                                        lGrid = Y
                                        Exit For
                                    End If
                                Next Y
                                If lGrid = -1 Then Continue For
                            Else
                                'if the grid index is the same as my index, then the grid to use is 0... always. This saves us a search thru an array
                                lGrid = 0
                            End If
                            If lGrid <> -1 Then
                                Dim lTemp As Int32 = giRelativeSmall(lGrid, lSmallSectorID, oMissile.lSmallSectorID)
                                Dim lRelTinyX As Int32
                                Dim lRelTinyZ As Int32
                                If lTemp <> -1 Then
                                    lRelTinyX = glBaseRelTinyX(lTemp) + .lTinyX + (gl_HALF_FINAL_PER_ROW - lTinyX)
                                    If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                        lRelTinyZ = glBaseRelTinyZ(lTemp) + .lTinyZ + (gl_HALF_FINAL_PER_ROW - lTinyZ)
                                        If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                            lEnvirDist = gyDistances(lRelTinyX, lRelTinyZ)
                                        End If
                                    End If
                                End If
                            End If

                            If lEnvirDist < lMaxRadarRange Then
                                mlMissileTargetRange(X) = lEnvirDist
                            Else
                                mlMissileTargetIdx(X) = -1
                                mlMissileTargetID(X) = -1
                                mlMissileTargetRange(X) = 999999
                            End If
                        End With
                    End If
                End If
            Next X

            'to do so, we basically run a mini-aggression engine... to do this, we create a temporary array of targets and ranges
            Dim lTempMissileTargetIdx(3) As Int32
            Dim lTempMissileTargetRange(3) As Int32
            Dim lTempMissileTargetID(3) As Int32
            For X As Int32 = 0 To 3
                lTempMissileTargetIdx(X) = -1
                lTempMissileTargetRange(X) = 999999
                lTempMissileTargetID(X) = -1
            Next X

            'First search priority is the Incoming Missile list... self-preservation first...
            For X As Int32 = 0 To lIncMissileUB
                If lIncMissileIdx(X) > -1 AndAlso goMissileMgr.mlMissileIdx(lIncMissileIdx(X)) = lIncMissileID(X) Then
                    Dim oMissile As MissileMgr.Missile = goMissileMgr.moMissile(lIncMissileIdx(X))
                    If oMissile Is Nothing Then Continue For

                    With oMissile
                        'ok, determine what side can fire...
                        Dim lEnvirDist As Int32 = Int32.MaxValue
                        Dim lGrid As Int32
                        If .lGridIndex <> lGridIndex Then
                            'Determine the grid
                            lGrid = -1
                            For Y As Int32 = 0 To lAdjust.GetUpperBound(0)
                                If lGridIndex + lAdjust(Y) = .lGridIndex Then
                                    lGrid = Y
                                    Exit For
                                End If
                            Next Y
                            If lGrid = -1 Then Continue For
                        Else
                            'if the grid index is the same as my index, then the grid to use is 0... always. This saves us a search thru an array
                            lGrid = 0
                        End If

                        Dim lTemp As Int32 = giRelativeSmall(lGrid, lSmallSectorID, oMissile.lSmallSectorID)
                        Dim lRelTinyX As Int32
                        Dim lRelTinyZ As Int32
                        If lTemp <> -1 Then
                            lRelTinyX = glBaseRelTinyX(lTemp) + .lTinyX + (gl_HALF_FINAL_PER_ROW - lTinyX)
                            If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                lRelTinyZ = glBaseRelTinyZ(lTemp) + .lTinyZ + (gl_HALF_FINAL_PER_ROW - lTinyZ)
                                If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                    lEnvirDist = gyDistances(lRelTinyX, lRelTinyZ)
                                End If
                            End If
                        End If

                        If lEnvirDist < lMaxRadarRange Then
                            'it is in range... does the side that can fire at the missile have a target?
                            Dim ySideCanFire As Byte = WhatSideCanFire(Me, lRelTinyX, lRelTinyZ)
                            If ySideCanFire < 4 Then
                                'is that side without a missile target?
                                If mlMissileTargetIdx(ySideCanFire) < 0 Then
                                    'ok, we are without a target on this side... do we have a temporary target?
                                    If lTempMissileTargetRange(ySideCanFire) > lEnvirDist Then  'was a <
                                        lTempMissileTargetRange(ySideCanFire) = lEnvirDist
                                        lTempMissileTargetIdx(ySideCanFire) = lIncMissileIdx(X)
                                        lTempMissileTargetID(ySideCanFire) = .lMissileID
                                    End If
                                End If
                            End If
                        End If

                    End With
                End If
            Next X

            'Ok, if we are here, store any non-negative 1 Idx's into our missile target array
            Dim bNeedSecondSearch As Boolean = False
            For X As Int32 = 0 To 3
                If lTempMissileTargetIdx(X) > -1 Then
                    mlMissileTargetIdx(X) = lTempMissileTargetIdx(X)
                    mlMissileTargetID(X) = lTempMissileTargetID(X)
                    mlMissileTargetRange(X) = lTempMissileTargetRange(X)

                    lTempMissileTargetID(X) = -1
                    lTempMissileTargetIdx(X) = -1
                    lTempMissileTargetRange(X) = 999999
                End If
                If mlMissileTargetIdx(X) < 0 Then bNeedSecondSearch = True
            Next X

            'second search priority is the Envir Grid Missile List...
            If bNeedSecondSearch = True Then
                For lGrid As Int32 = 0 To lAdjust.GetUpperBound(0)
                    Dim lLargeSector As Int32 = lGridIndex + lAdjust(lGrid)
                    If lLargeSector > -1 AndAlso lLargeSector < ParentEnvir.lGridUB + 1 Then
                        Dim oTmpGrid As EnvirGrid = ParentEnvir.oGrid(lLargeSector)
                        If oTmpGrid Is Nothing Then Continue For
                        For Y As Int32 = 0 To oTmpGrid.lMissileUB
                            Dim lMissileIdx As Int32 = oTmpGrid.lMissiles(Y)
                            If lMissileIdx <> -1 Then
                                Dim oMissile As MissileMgr.Missile = goMissileMgr.moMissile(lMissileIdx)
                                If oMissile Is Nothing Then
                                    oTmpGrid.lMissiles(Y) = -1
                                    Continue For
                                End If

                                With oMissile
                                    'ignore missiles targeting me as I have already checked them
                                    If .lTargetID = ObjectID AndAlso .lTargetIdx = ServerIndex Then Continue For

                                    If .oParentEnvir.ObjectID <> ParentEnvir.ObjectID OrElse .oParentEnvir.ObjTypeID <> ParentEnvir.ObjTypeID Then
                                        oTmpGrid.RemoveMissile(Y)
                                        Continue For
                                    End If

                                    'ok, now, is the missile within my radar range?
                                    Dim lEnvirDist As Int32 = Int32.MaxValue
                                    Dim lTemp As Int32 = giRelativeSmall(lGrid, lSmallSectorID, oMissile.lSmallSectorID)
                                    Dim lRelTinyX As Int32
                                    Dim lRelTinyZ As Int32
                                    If lTemp <> -1 Then
                                        lRelTinyX = glBaseRelTinyX(lTemp) + .lTinyX + (gl_HALF_FINAL_PER_ROW - lTinyX)
                                        If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                            lRelTinyZ = glBaseRelTinyZ(lTemp) + .lTinyZ + (gl_HALF_FINAL_PER_ROW - lTinyZ)
                                            If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                                lEnvirDist = gyDistances(lRelTinyX, lRelTinyZ)
                                            End If
                                        End If
                                    End If
                                    If lEnvirDist < lMaxRadarRange Then
                                        'Does the side that can fire have a target?
                                        Dim ySideCanFire As Byte = WhatSideCanFire(Me, lRelTinyX, lRelTinyZ)
                                        If ySideCanFire < 4 Then
                                            'is that side without a missile target?
                                            If mlMissileTargetIdx(ySideCanFire) < 0 Then
                                                'ok, we are without a target on this side... do we have a temporary target?
                                                If lTempMissileTargetRange(ySideCanFire) > lEnvirDist Then
                                                    'is the owner of the missile an enemy?
                                                    If Me.Owner.GetPlayerRelScore(.lAttackerOwnerID, False, -1) <= elRelTypes.eWar Then
                                                        'ok, their missile cannot be good... shoot it down
                                                        lTempMissileTargetRange(ySideCanFire) = lEnvirDist
                                                        lTempMissileTargetIdx(ySideCanFire) = lMissileIdx
                                                        lTempMissileTargetID(ySideCanFire) = .lMissileID
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If

                                End With

                            End If
                        Next Y
                    End If
                Next lGrid

                'Ok, we're done with the final search save our values
                For X As Int32 = 0 To 3
                    If lTempMissileTargetIdx(X) > -1 Then
                        mlMissileTargetIdx(X) = lTempMissileTargetIdx(X)
                        mlMissileTargetID(X) = lTempMissileTargetID(X)
                        mlMissileTargetRange(X) = lTempMissileTargetRange(X)
                    End If
                Next X

            End If

            'Now, we are finally done, so we reask the original question
            If yArc = UnitArcs.eAllArcs Then
                'ok, different... does any of our sides have a target?
                For X As Int32 = 0 To 3
                    If mlMissileTargetIdx(X) > -1 Then
                        If goMissileMgr.mlMissileIdx(mlMissileTargetIdx(X)) = mlMissileTargetID(X) Then
                            'If mlMissileTargetRange(X) <= lRange Then
                            lMissileTargetRange = mlMissileTargetRange(X)
                            Return goMissileMgr.moMissile(mlMissileTargetIdx(X))
                            'End If
                        Else
                            mlMissileTargetIdx(X) = -1
                        End If
                    End If
                Next X
            Else
                'does the arc already have a target?
                If mlMissileTargetIdx(yArc) > -1 Then
                    If goMissileMgr.mlMissileIdx(mlMissileTargetIdx(yArc)) = mlMissileTargetID(yArc) Then
                        'If mlMissileTargetRange(yArc) <= lRange Then
                        lMissileTargetRange = mlMissileTargetRange(yArc)
                        Return goMissileMgr.moMissile(mlMissileTargetIdx(yArc))
                        'Else : Return Nothing
                        'End If
                    End If
                End If
            End If

        End If

        Return Nothing
    End Function

    Public Function AddIncomingMissile(ByVal lMissileIndex As Int32, ByVal lMissileID As Int32) As Int32
        Dim lIdx As Int32 = -1

        For X As Int32 = 0 To lIncMissileUB
            If lIncMissileIdx(X) = lMissileIndex Then
                Return X
            ElseIf lIdx = -1 AndAlso lIncMissileIdx(X) = -1 Then
                lIdx = X
            End If
        Next X

        lIncMissileCnt += 1

        If lIdx <> -1 Then
            lIncMissileIdx(lIdx) = lMissileIndex
            lIncMissileID(lIdx) = lMissileID
            Return lIdx
        End If

        lIncMissileUB += 1
        ReDim Preserve lIncMissileIdx(lIncMissileUB)
        ReDim Preserve lIncMissileID(lIncMissileUB)
        lIncMissileIdx(lIncMissileUB) = lMissileIndex
        lIncMissileID(lIncMissileUB) = lMissileID
        Return lIncMissileUB

    End Function

    Public Sub RemoveIncomingMissile(ByVal lArrayIdx As Int32, ByVal lMissileID As Int32)
        If lIncMissileID Is Nothing Then Return
        If lIncMissileID.GetUpperBound(0) < lArrayIdx Then Return
        If lIncMissileID(lArrayIdx) = lMissileID Then
            lIncMissileIdx(lArrayIdx) = -1
            lIncMissileCnt -= 1
            If lIncMissileCnt < 0 Then lIncMissileCnt = 0
        End If
    End Sub
#End Region

    '#Region "  Player Damage  "
    '    Private mlPlayerID() As Int32
    '    Private mfDmgDone() As Single
    '    Private mfLastDmgDone() As Single
    '    Private mlPlayerDmgUB As Int32 = -1
    '    Public bHasWarpointsToDistribute As Boolean = False
    '    Public Sub AddPlayerDmg(ByVal lID As Int32, ByVal lDmg As Int32, ByVal yAttackerType As Byte, ByVal yTargetType As Byte)
    '        If (iCombatTactics And eiBehaviorPatterns.eEngagement_Hold_Fire) <> 0 Then Return

    '        Dim lVal As Int32 = CInt(yAttackerType) - CInt(yTargetType)
    '        If lVal < 0 Then lVal = 0
    '        Dim fActualDmg As Single = lDmg * mfHullTypeDmgModMult(lVal)

    '        bHasWarpointsToDistribute = True

    '        If mlPlayerID Is Nothing = False Then
    '            Dim lUB As Int32 = Math.Min(mlPlayerDmgUB, mlPlayerID.GetUpperBound(0))
    '            For X As Int32 = 0 To lUB
    '                If mlPlayerID(X) = lID Then
    '                    mfDmgDone(X) += fActualDmg
    '                    Return
    '                End If
    '            Next X
    '        Else
    '            ReDim mlPlayerID(-1)
    '        End If
    '        SyncLock mlPlayerID
    '            mlPlayerDmgUB += 1
    '            ReDim Preserve mlPlayerID(mlPlayerDmgUB)
    '            ReDim Preserve mfDmgDone(mlPlayerDmgUB)
    '            ReDim Preserve mfLastDmgDone(mlPlayerDmgUB)
    '            mlPlayerID(mlPlayerDmgUB) = lID
    '            mfDmgDone(mlPlayerDmgUB) = fActualDmg
    '            mfLastDmgDone(mlPlayerDmgUB) = 0
    '        End SyncLock
    '    End Sub
    '    Public Sub DistributeWarpoints(ByVal bRunningTotal As Boolean)
    '        If Me.ObjTypeID = ObjectType.eFacility OrElse mlPlayerDmgUB < 0 Then
    '            bHasWarpointsToDistribute = False
    '            Return
    '        End If

    '        Dim oCurDef As Epica_Entity_Def = Nothing
    '        If lEntityDefServerIndex > -1 Then
    '            oCurDef = goEntityDefs(lEntityDefServerIndex)
    '            If oCurDef Is Nothing Then
    '                bHasWarpointsToDistribute = False
    '                Return
    '            End If
    '        End If 

    '        If mlPlayerDmgUB < 0 Then Return
    '        'Dim fTotalDmg As Single = 0
    '        'For X As Int32 = 0 To mlPlayerDmgUB
    '        '    fTotalDmg += mfDmgDone(X)
    '        'Next X
    '        'If fTotalDmg < 1 Then Return
    '        'Dim fOneOver As Single = 1.0F / fTotalDmg
    '        Dim fOneOver As Single = 1.0F / oCurDef.lAbsoluteTotalHP

    '        Dim yMsg(39) As Byte

    '        If bRunningTotal = False Then
    '            'nope, we're actually rewarding the warpoints
    '            System.BitConverter.GetBytes(GlobalMessageCode.eRewardWarpoints).CopyTo(yMsg, 0)
    '        Else
    '            'We are only reporting warpoints
    '            System.BitConverter.GetBytes(GlobalMessageCode.eReportWarpoints).CopyTo(yMsg, 0)
    '        End If

    '        Me.GetGUIDAsString.CopyTo(yMsg, 2)
    '        If Me.ParentEnvir Is Nothing = False Then
    '            Me.ParentEnvir.GetGUIDAsString.CopyTo(yMsg, 8)
    '        End If
    '        System.BitConverter.GetBytes(Me.LocX).CopyTo(yMsg, 14)
    '        System.BitConverter.GetBytes(Me.LocZ).CopyTo(yMsg, 18)

    '        System.BitConverter.GetBytes(Me.lOwnerID).CopyTo(yMsg, 30)
    '        System.BitConverter.GetBytes(oCurDef.ModelID).CopyTo(yMsg, 34)

    '        System.BitConverter.GetBytes(oCurDef.CombatRating).CopyTo(yMsg, 36)

    '        Dim lMaxInf As Int32 = 1
    '        If lEntityDefServerIndex <> -1 Then
    '            If glEntityDefIdx(lEntityDefServerIndex) <> -1 Then
    '                lMaxInf = goEntityDefs(lEntityDefServerIndex).lWarpointValue
    '            End If
    '        End If

    '        Dim oPlayerInEnvironment As Player = Nothing

    '        For X As Int32 = 0 To mlPlayerDmgUB
    '            If mlPlayerID(X) = Me.lOwnerID Then Continue For

    '            If bRunningTotal = True Then
    '                'Ok, we can reduce packets by finding if the player is in the environment
    '                oPlayerInEnvironment = ParentEnvir.GetPlayerInEnvironment(mlPlayerID(X))
    '                If oPlayerInEnvironment Is Nothing Then Continue For
    '            End If

    '            Dim fThisDmgDone As Single = mfDmgDone(X)
    '            If bRunningTotal = True Then fThisDmgDone = mfDmgDone(X) - mfLastDmgDone(X)
    '            Dim fPerc As Single = fThisDmgDone * fOneOver
    '            If bRunningTotal = False Then mfDmgDone(X) = 0.0F Else mfLastDmgDone(X) = mfDmgDone(X)
    '            Dim lWarpoints As Int32 = CInt(Math.Floor(fPerc * lMaxInf))
    '            If lWarpoints = 0 Then Continue For

    '            'Now, award player mlPlayerID(X), lWarpoints warpoints for killing me
    '            Dim lPlayerID As Int32 = mlPlayerID(X)
    '            System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, 22)
    '            System.BitConverter.GetBytes(lWarpoints).CopyTo(yMsg, 26)

    '            If bRunningTotal = True Then
    '                If oPlayerInEnvironment Is Nothing = False AndAlso oPlayerInEnvironment.oSocket Is Nothing = False Then
    '                    'Still send it thru the primary so it can calculate the points based on the rules
    '                    goMsgSys.SendToPrimary(yMsg)
    '                End If
    '            Else
    '                goMsgSys.SendToPrimary(yMsg)
    '            End If

    '        Next X

    '        bHasWarpointsToDistribute = False
    '    End Sub

    '    Public Sub FacAddPlayerDmg(ByVal lID As Int32, ByVal lDmg As Int32)
    '        If True = True Then Return
    '        If (iCombatTactics And eiBehaviorPatterns.eEngagement_Hold_Fire) <> 0 Then Return

    '        If mlPlayerID Is Nothing = False Then
    '            Dim lUB As Int32 = Math.Min(mlPlayerDmgUB, mlPlayerID.GetUpperBound(0))
    '            For X As Int32 = 0 To lUB
    '                If mlPlayerID(X) = lID Then
    '                    mfDmgDone(X) += lDmg
    '                    Return
    '                End If
    '            Next X
    '        Else
    '            ReDim mlPlayerID(-1)
    '        End If
    '        SyncLock mlPlayerID
    '            mlPlayerDmgUB += 1
    '            ReDim Preserve mlPlayerID(mlPlayerDmgUB)
    '            ReDim Preserve mfDmgDone(mlPlayerDmgUB)
    '            ReDim Preserve mfLastDmgDone(mlPlayerDmgUB)
    '            mlPlayerID(mlPlayerDmgUB) = lID
    '            mfDmgDone(mlPlayerDmgUB) = lDmg
    '            mfLastDmgDone(mlPlayerDmgUB) = 0
    '        End SyncLock
    '    End Sub
    '    Public Sub FacDistributeWarpoints(ByVal iModelID As Int16, ByVal bRunningTotal As Boolean)
    '        If True = True Then Return
    '        If mlPlayerDmgUB < 0 Then Return
    '        Dim fTotalDmg As Single = 0
    '        For X As Int32 = 0 To mlPlayerDmgUB
    '            fTotalDmg += mfDmgDone(X)
    '        Next X
    '        If fTotalDmg < 1 Then Return
    '        Dim fOneOver As Single = 1.0F / fTotalDmg

    '        Dim yMsg(39) As Byte

    '        If bRunningTotal = False Then
    '            'nope, we're actually rewarding the warpoints
    '            System.BitConverter.GetBytes(GlobalMessageCode.eRewardWarpoints).CopyTo(yMsg, 0)
    '        Else
    '            'We are only reporting warpoints
    '            System.BitConverter.GetBytes(GlobalMessageCode.eReportWarpoints).CopyTo(yMsg, 0)
    '        End If

    '        Me.GetGUIDAsString.CopyTo(yMsg, 2)
    '        If Me.ParentEnvir Is Nothing = False Then
    '            Me.ParentEnvir.GetGUIDAsString.CopyTo(yMsg, 8)
    '        End If
    '        System.BitConverter.GetBytes(Me.LocX).CopyTo(yMsg, 14)
    '        System.BitConverter.GetBytes(Me.LocZ).CopyTo(yMsg, 18)

    '        System.BitConverter.GetBytes(Me.lOwnerID).CopyTo(yMsg, 30)
    '        System.BitConverter.GetBytes(iModelID).CopyTo(yMsg, 34)

    '        System.BitConverter.GetBytes(1I).CopyTo(yMsg, 36)

    '        Dim lMaxInf As Int32 = 1000

    '        Dim oPlayerInEnvironment As Player = Nothing

    '        For X As Int32 = 0 To mlPlayerDmgUB
    '            If mlPlayerID(X) = Me.lOwnerID Then Continue For

    '            If bRunningTotal = True Then
    '                'Ok, we can reduce packets by finding if the player is in the environment
    '                oPlayerInEnvironment = ParentEnvir.GetPlayerInEnvironment(mlPlayerID(X))
    '                If oPlayerInEnvironment Is Nothing Then Continue For
    '            End If

    '            Dim fThisDmgDone As Single = mfDmgDone(X)
    '            If bRunningTotal = True Then fThisDmgDone = mfDmgDone(X) - mfLastDmgDone(X)
    '            Dim fPerc As Single = fThisDmgDone * fOneOver
    '            If bRunningTotal = False Then mfDmgDone(X) = 0.0F Else mfLastDmgDone(X) = mfDmgDone(X)
    '            Dim lWarpoints As Int32 = CInt(Math.Floor(fPerc * lMaxInf))
    '            If lWarpoints = 0 Then Continue For

    '            'Now, award player mlPlayerID(X), lWarpoints warpoints for killing me
    '            Dim lPlayerID As Int32 = mlPlayerID(X)
    '            System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, 22)
    '            System.BitConverter.GetBytes(lWarpoints).CopyTo(yMsg, 26)

    '            If bRunningTotal = True Then
    '                If oPlayerInEnvironment Is Nothing = False AndAlso oPlayerInEnvironment.oSocket Is Nothing = False Then
    '                    'Still send it thru the primary so it can calculate the points based on the rules
    '                    goMsgSys.SendToPrimary(yMsg)
    '                End If
    '            Else
    '                goMsgSys.SendToPrimary(yMsg)
    '            End If
    '        Next X
    '    End Sub

    '#End Region

#Region "  Jamming Target Management  "
    'Essentially, we store all aggressions here
    Private mlJamTargetIdx(-1) As Int32           'server index (goEntity)
    Private mlJamTargetScore(-1) As Int32         'score for this target (higher scores get preferred targetting)
    Private mlJamTargetID(-1) As Int32            'the ID to ensure no cross-array indexing occurs

    Public lJamRangeMod As Int32 = 0            'adjusts to Max and Opt Radar Range
    Public lJamResMod As Int32 = 0              'adjusts to scan resolution and disruption resistance
    Public lJamWpnAccMod As Int32 = 0           'adjusts to weapon accuracy
    Public lJamStrMod As Int32 = 0              'adjusts to jamming strength
    Public lJamTargetsMod As Int32 = 0          'adjusts to jam targets

    Public Sub ClearJamModifiers()
        lJamRangeMod = 0
        lJamResMod = 0
        lJamWpnAccMod = 0
        lJamStrMod = 0
        lJamTargetsMod = 0
    End Sub

    Public Sub ClearJamTargets()
        If mlJamTargetIdx Is Nothing Then
            Dim lUB As Int32 = goEntityDefs(lEntityDefServerIndex).JamTargets
            If lUB = 0 AndAlso goEntityDefs(lEntityDefServerIndex).JamStrength = 0 Then lUB = -1 Else lUB = 255
            ReDim mlJamTargetIdx(lUB)
            ReDim mlJamTargetScore(lUB)
            ReDim mlJamTargetID(lUB)
        End If

        For X As Int32 = 0 To mlJamTargetIdx.GetUpperBound(0)
            mlJamTargetIdx(X) = -1
            mlJamTargetScore(X) = 0
            mlJamTargetID(X) = -1
        Next X
    End Sub

    Public Sub AddJamTarget(ByVal lServerIndex As Int32, ByVal lTargetScore As Int32)
        'we'll sort highest to lowest
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlJamTargetIdx.GetUpperBound(0) - lJamTargetsMod
            If mlJamTargetIdx(X) = lServerIndex Then Return
            If lIdx = -1 AndAlso mlJamTargetScore(X) < lTargetScore Then
                lIdx = X
            End If
        Next X

        If lIdx = -1 Then
            'I am the lowest and I do not fit within the target list so return
            Return
        Else
            For X As Int32 = (mlJamTargetIdx.GetUpperBound(0) - lJamTargetsMod) To lIdx + 1 Step -1
                mlJamTargetIdx(X) = mlJamTargetIdx(X - 1)
                mlJamTargetScore(X) = mlJamTargetScore(X - 1)
                mlJamTargetID(X) = mlJamTargetID(X - 1)
            Next X
            mlJamTargetIdx(lIdx) = lServerIndex
            mlJamTargetScore(lIdx) = lTargetScore
            mlJamTargetID(lIdx) = glEntityIdx(mlJamTargetIdx(lIdx))
        End If
    End Sub

    Public Sub ProcessJamming()
        Dim oDef As Epica_Entity_Def = goEntityDefs(Me.lEntityDefServerIndex)
        Dim lStr As Int32 = oDef.JamStrength
        Dim lEff As Int32 = oDef.JamEffect

        If lEff < 1 Then Return
        If lStr < 1 Then Return

        For X As Int32 = 0 To mlJamTargetIdx.GetUpperBound(0) - lJamTargetsMod
            If mlJamTargetIdx(X) <> -1 AndAlso glEntityIdx(mlJamTargetIdx(X)) = mlJamTargetID(X) Then
                'Ok, attempt to jam it
                Dim oEntity As Epica_Entity = goEntity(mlJamTargetIdx(X))
                If oEntity Is Nothing = False Then
                    If oEntity.lEntityDefServerIndex <> -1 Then
                        If glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
                            oDef = goEntityDefs(oEntity.lEntityDefServerIndex)

                            'now, do our test... if the target is not immune
                            If lEff <> oDef.JamImmunity Then
                                'If oEntity.bInCombatRegister = False Then AddEntityCombat(oEntity.ServerIndex, oEntity.ObjectID)
                                'the test is the following:
                                'Random value from 0 to Modified Jamming Strength vs. Random Value from 0 to Modified Disrupt Resist
                                Dim lAttack As Int32 = CInt(Rnd() * (lStr + lJamStrMod))
                                Dim lDefend As Int32 = CInt(Rnd() * (oDef.DisruptionResist + oEntity.lJamResMod))

                                'Higher roll wins
                                If lAttack > lDefend Then
                                    'Ok, jam the target
                                    Select Case lEff
                                        Case 1  'system degradation: reduces opt and max radar ranges
                                            Dim lVal As Int32 = Math.Max(oDef.lOptRadarRange, oDef.lMaxRadarRange)
                                            lVal \= 2
                                            lJamRangeMod = -lVal
                                        Case 2  'system interference: reduces scan resolution and disruption resistance
                                            Dim lVal As Int32 = Math.Max(oDef.ScanResolution, oDef.DisruptionResist)
                                            lVal \= 2
                                            lJamResMod = -lVal
                                        Case 3  'system clutter: reduces weapon accuracy
                                            lJamWpnAccMod = oDef.Weapon_Acc \ -2
                                        Case 4  'anti-jamming: reduces jamming strength
                                            lJamStrMod = oDef.JamStrength \ -2
                                        Case 5  'decoy clutter: reduces jamming max targets (ineffective against AOE Jamming)
                                            If oDef.JamTargets <> 0 Then lJamTargetsMod = oDef.JamTargets \ 2
                                    End Select
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next X
    End Sub
#End Region

    Public Sub SendAsyncUpdateToPrimary()
        Dim yMsg(56) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eUpdateEntityAttrs).CopyTo(yMsg, 0)
        GetGUIDAsString.CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(CurrentStatus).CopyTo(yMsg, 8)
        System.BitConverter.GetBytes(ShieldHP).CopyTo(yMsg, 12)
        System.BitConverter.GetBytes(ArmorHP(0)).CopyTo(yMsg, 16)
        System.BitConverter.GetBytes(ArmorHP(1)).CopyTo(yMsg, 20)
        System.BitConverter.GetBytes(ArmorHP(2)).CopyTo(yMsg, 24)
        System.BitConverter.GetBytes(ArmorHP(3)).CopyTo(yMsg, 28)
        System.BitConverter.GetBytes(StructuralHP).CopyTo(yMsg, 32)
        System.BitConverter.GetBytes(Fuel_Cap).CopyTo(yMsg, 36)
        yMsg(40) = Exp_Level
        System.BitConverter.GetBytes(CInt(MoveLocX)).CopyTo(yMsg, 41)
        System.BitConverter.GetBytes(CInt(MoveLocZ)).CopyTo(yMsg, 45)
        System.BitConverter.GetBytes(DestX).CopyTo(yMsg, 49)
        System.BitConverter.GetBytes(DestZ).CopyTo(yMsg, 53)

        goMsgSys.SendToPrimary(yMsg)
        lUpdatePrimaryWithHPRequest = Int32.MinValue
    End Sub

    Public Sub DeductFuel()
        Dim lOwnerID As Int32 = Me.Owner.ObjectID
        Dim lDeduct As Int32 = 1

        With Me.ParentEnvir
            For X As Int32 = 0 To .lPlayersWhoHaveUnitsHereUB
                If .lPlayersWhoHaveUnitsHereIdx(X) = lOwnerID Then
                    If .lPlayersWhoHaveUnitsHereCP(X) > Owner.lCPLimit Then lDeduct = 2
                    Exit For
                End If
            Next X
        End With

        If Fuel_Cap > 0 Then Fuel_Cap -= lDeduct
    End Sub

    Public Sub SetDest(ByVal lX As Int32, ByVal lZ As Int32, ByVal iAngle As Int16)

        If Me.ParentEnvir Is Nothing = False Then
            If LocX < ParentEnvir.lMinPosX Then LocX = ParentEnvir.lMinPosX
            If LocZ < ParentEnvir.lMinPosZ Then LocZ = ParentEnvir.lMinPosZ
            If LocX > ParentEnvir.lMaxPosX Then LocX = ParentEnvir.lMaxPosX
            If LocZ > ParentEnvir.lMaxPosZ Then LocZ = ParentEnvir.lMaxPosZ
            If lX < ParentEnvir.lMinPosX Then lX = ParentEnvir.lMinPosX
            If lX > ParentEnvir.lMaxPosX Then lX = ParentEnvir.lMaxPosX
            If lZ < ParentEnvir.lMinPosZ Then lZ = ParentEnvir.lMinPosZ
            If lZ > ParentEnvir.lMaxPosZ Then lZ = ParentEnvir.lMaxPosZ
        End If

        'ok, here we go... first, we need to set our move locs up...
        MoveLocX = LocX
        MoveLocZ = LocZ

        'then, we can set our dest up
        DestX = lX
        DestZ = lZ
        TrueDestAngle = iAngle
        DestAngle = CShort(LineAngleDegrees(LocX, LocZ, DestX, DestZ) * 10)

        LastCycleMoved = glCurrentCycle
        'finally, set our movement status
        CurrentStatus = CurrentStatus Or elUnitStatus.eUnitMoving

        yResetMoveByPlayer = 0

        'If ParentEnvir Is Nothing = False AndAlso ParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
        '	'me.ParentEnvir.oGeoObject
        'End If

        bDecelerating = False

        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, Me.ServerIndex, Me.ObjectID) 'AddEntityMoving(Me.ServerIndex, Me.ObjectID)
    End Sub

    Public Sub InitializeWeaponCycles()
        Dim X As Int32
        Dim oTmpDef As Epica_Entity_Def

        If lEntityDefServerIndex <> -1 Then
            If glEntityDefIdx(lEntityDefServerIndex) <> -1 Then
                oTmpDef = goEntityDefs(lEntityDefServerIndex)

                ReDim lWpnNextFireCycle(oTmpDef.WeaponDefUB)
                For X = 0 To oTmpDef.WeaponDefUB
                    lWpnNextFireCycle(X) = 0
                Next X

                oTmpDef = Nothing
            End If
        End If

    End Sub

    Public Function GetObjectAsSmallString() As Byte()
        Dim oTmpDef As Epica_Entity_Def = Nothing
        Dim mySmallSendString() As Byte

        If lEntityDefServerIndex <> -1 Then
            If glEntityDefIdx(lEntityDefServerIndex) <> -1 Then
                oTmpDef = goEntityDefs(lEntityDefServerIndex)
            End If
        Else
            For X As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(X) = lEntityDefServerID Then
                    If (ObjTypeID = ObjectType.eUnit AndAlso goEntityDefs(X).ObjTypeID = ObjectType.eUnitDef) OrElse _
                       (ObjTypeID = ObjectType.eFacility AndAlso goEntityDefs(X).ObjTypeID = ObjectType.eFacilityDef) Then
                        With Me
                            .lEntityDefServerIndex = X
                            .lModelRangeOffset = goEntityDefs(X).lModelRangeOffset
                            .Acceleration = goEntityDefs(X).BaseAcceleration
                            .Maneuver = goEntityDefs(X).BaseManeuver
                            .MaxSpeed = goEntityDefs(X).BaseMaxSpeed
                            .TurnAmount = goEntityDefs(X).BaseTurnAmount
                            .TurnAmountTimes100 = goEntityDefs(X).BaseTurnAmountTimes100

                            ReDim .lWpnAmmoCnt(goEntityDefs(X).WeaponDefUB)
                            For Y As Int32 = 0 To .lWpnAmmoCnt.Length - 1
                                .lWpnAmmoCnt(Y) = -1
                            Next Y

                            oTmpDef = goEntityDefs(X)

                            Exit For
                        End With
                    End If
                End If
            Next X
        End If

        If oTmpDef Is Nothing Then Return Nothing

        'Create our burn FX value
        Dim yBurnFX As Byte

        'First, check structure
        If StructuralHP <= 0 OrElse (CSng(StructuralHP) / oTmpDef.Structure_MaxHP) < 0.5F Then
            yBurnFX = 255
        Else
            'Ok, now, check our armor sides
            Dim lIdx As Int32 = 0
            If ArmorHP(lIdx) < 0 Then ArmorHP(lIdx) = 0
            If oTmpDef.Armor_MaxHP(lIdx) <> 0 AndAlso ArmorHP(lIdx) > 0 Then
                If ArmorHP(lIdx) < oTmpDef.Armor_MaxHP(lIdx) * 0.9F Then
                    yBurnFX += CByte(1)
                    If (CSng(ArmorHP(lIdx)) / oTmpDef.Armor_MaxHP(lIdx)) < 0.5F Then
                        yBurnFX += CByte(16)
                    End If
                End If
            End If

            lIdx = 1
            If ArmorHP(lIdx) < 0 Then ArmorHP(lIdx) = 0
            If oTmpDef.Armor_MaxHP(lIdx) <> 0 AndAlso ArmorHP(lIdx) > 0 Then
                If ArmorHP(lIdx) < oTmpDef.Armor_MaxHP(lIdx) * 0.9F Then
                    yBurnFX += CByte(2)
                    If (CSng(ArmorHP(lIdx)) / oTmpDef.Armor_MaxHP(lIdx)) < 0.5F Then
                        yBurnFX += CByte(32)
                    End If
                End If
            End If

            lIdx = 2
            If ArmorHP(lIdx) < 0 Then ArmorHP(lIdx) = 0
            If oTmpDef.Armor_MaxHP(lIdx) <> 0 AndAlso ArmorHP(lIdx) > 0 Then
                If ArmorHP(lIdx) < oTmpDef.Armor_MaxHP(lIdx) * 0.9F Then
                    yBurnFX += CByte(4)
                    If (CSng(ArmorHP(lIdx)) / oTmpDef.Armor_MaxHP(lIdx)) < 0.5F Then
                        yBurnFX += CByte(64)
                    End If
                End If
            End If

            lIdx = 3
            If ArmorHP(lIdx) < 0 Then ArmorHP(lIdx) = 0
            If oTmpDef.Armor_MaxHP(lIdx) <> 0 AndAlso ArmorHP(lIdx) > 0 Then
                If ArmorHP(lIdx) < oTmpDef.Armor_MaxHP(lIdx) * 0.9F Then
                    yBurnFX += CByte(8)
                    If (CSng(ArmorHP(lIdx)) / oTmpDef.Armor_MaxHP(lIdx)) < 0.5F Then
                        yBurnFX += CByte(128)
                    End If
                End If
            End If
        End If

        'If mbStringReady = False Then
        If (CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
            'unit moving byte array...
            ReDim mySmallSendString(45)
            GetGUIDAsString.CopyTo(mySmallSendString, 0)

            Dim lPos As Int32 = 6

            'Begin Encryption Code
            Select Case ObjectID Mod 4
                Case 0
                    System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(DestX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(DestZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(TrueDestAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oTmpDef.ModelID).CopyTo(mySmallSendString, lPos) : lPos += 2
                    mySmallSendString(lPos) = CByte(TotalVelocity) : lPos += 1
                    mySmallSendString(lPos) = MaxSpeed : lPos += 1
                    mySmallSendString(lPos) = Maneuver : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yDefOptRadarRange : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yDefMaxRadarRange : lPos += 1
                    mySmallSendString(lPos) = yProductionType : lPos += 1
                    mySmallSendString(lPos) = yBurnFX : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yFXColor : lPos += 1
                    mySmallSendString(lPos) = yFormationManeuver : lPos += 1
                    mySmallSendString(lPos) = yFormationMaxSpeed : lPos += 1
                Case 1
                    mySmallSendString(lPos) = MaxSpeed : lPos += 1
                    mySmallSendString(lPos) = Maneuver : lPos += 1
                    mySmallSendString(lPos) = yBurnFX : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yFXColor : lPos += 1
                    System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySmallSendString, lPos) : lPos += 4
                    mySmallSendString(lPos) = CByte(TotalVelocity) : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yDefOptRadarRange : lPos += 1
                    System.BitConverter.GetBytes(DestX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(DestZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(TrueDestAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    mySmallSendString(lPos) = yFormationManeuver : lPos += 1
                    mySmallSendString(lPos) = yFormationMaxSpeed : lPos += 1
                    System.BitConverter.GetBytes(oTmpDef.ModelID).CopyTo(mySmallSendString, lPos) : lPos += 2
                    mySmallSendString(lPos) = oTmpDef.yDefMaxRadarRange : lPos += 1
                    mySmallSendString(lPos) = yProductionType : lPos += 1
                Case 2
                    System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(TrueDestAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(LocX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(DestZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oTmpDef.ModelID).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(DestX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    mySmallSendString(lPos) = Maneuver : lPos += 1
                    mySmallSendString(lPos) = CByte(TotalVelocity) : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yDefOptRadarRange : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yDefMaxRadarRange : lPos += 1
                    mySmallSendString(lPos) = MaxSpeed : lPos += 1
                    mySmallSendString(lPos) = yFormationMaxSpeed : lPos += 1
                    mySmallSendString(lPos) = yFormationManeuver : lPos += 1
                    mySmallSendString(lPos) = yProductionType : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yFXColor : lPos += 1
                    mySmallSendString(lPos) = yBurnFX : lPos += 1
                Case Else
                    mySmallSendString(lPos) = yBurnFX : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yFXColor : lPos += 1
                    mySmallSendString(lPos) = Maneuver : lPos += 1
                    mySmallSendString(lPos) = MaxSpeed : lPos += 1
                    System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySmallSendString, lPos) : lPos += 4
                    mySmallSendString(lPos) = oTmpDef.yDefOptRadarRange : lPos += 1
                    mySmallSendString(lPos) = CByte(TotalVelocity) : lPos += 1
                    System.BitConverter.GetBytes(DestZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(DestX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(LocZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(TrueDestAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(oTmpDef.ModelID).CopyTo(mySmallSendString, lPos) : lPos += 2
                    mySmallSendString(lPos) = oTmpDef.yDefMaxRadarRange : lPos += 1
                    mySmallSendString(lPos) = yProductionType : lPos += 1
                    mySmallSendString(lPos) = yFormationManeuver : lPos += 1
                    mySmallSendString(lPos) = yFormationMaxSpeed : lPos += 1
            End Select
        Else
            'unit not moving byte array...
            ReDim mySmallSendString(32)
            GetGUIDAsString.CopyTo(mySmallSendString, 0)

            Dim lPos As Int32 = 6

            'Begin Encryption Code
            Select Case ObjectID Mod 4
                Case 0
                    System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oTmpDef.ModelID).CopyTo(mySmallSendString, lPos) : lPos += 2
                    mySmallSendString(lPos) = MaxSpeed : lPos += 1
                    mySmallSendString(lPos) = Maneuver : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yDefOptRadarRange : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yDefMaxRadarRange : lPos += 1
                    mySmallSendString(lPos) = yProductionType : lPos += 1
                    mySmallSendString(lPos) = yBurnFX : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yFXColor : lPos += 1
                Case 1
                    mySmallSendString(lPos) = MaxSpeed : lPos += 1
                    mySmallSendString(lPos) = Maneuver : lPos += 1
                    mySmallSendString(lPos) = yBurnFX : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yFXColor : lPos += 1
                    System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySmallSendString, lPos) : lPos += 4
                    mySmallSendString(lPos) = oTmpDef.yDefOptRadarRange : lPos += 1
                    System.BitConverter.GetBytes(LocZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(oTmpDef.ModelID).CopyTo(mySmallSendString, lPos) : lPos += 2
                    mySmallSendString(lPos) = oTmpDef.yDefMaxRadarRange : lPos += 1
                    mySmallSendString(lPos) = yProductionType : lPos += 1
                Case 2
                    System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oTmpDef.ModelID).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(LocAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    mySmallSendString(lPos) = Maneuver : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yDefOptRadarRange : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yDefMaxRadarRange : lPos += 1
                    mySmallSendString(lPos) = MaxSpeed : lPos += 1
                    mySmallSendString(lPos) = yProductionType : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yFXColor : lPos += 1
                    mySmallSendString(lPos) = yBurnFX : lPos += 1
                Case Else
                    mySmallSendString(lPos) = yBurnFX : lPos += 1
                    mySmallSendString(lPos) = oTmpDef.yFXColor : lPos += 1
                    mySmallSendString(lPos) = Maneuver : lPos += 1
                    mySmallSendString(lPos) = MaxSpeed : lPos += 1
                    System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySmallSendString, lPos) : lPos += 4
                    mySmallSendString(lPos) = oTmpDef.yDefOptRadarRange : lPos += 1
                    System.BitConverter.GetBytes(LocX).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(LocAngle).CopyTo(mySmallSendString, lPos) : lPos += 2
                    System.BitConverter.GetBytes(LocZ).CopyTo(mySmallSendString, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oTmpDef.ModelID).CopyTo(mySmallSendString, lPos) : lPos += 2
                    mySmallSendString(lPos) = oTmpDef.yDefMaxRadarRange : lPos += 1
                    mySmallSendString(lPos) = yProductionType : lPos += 1
            End Select
        End If
        oTmpDef = Nothing

        Return mySmallSendString
    End Function

    Public Sub New()
        Dim X As Int32
        mbExpRewarded = False
        For X = 0 To 3
            lTargetsServerIdx(X) = -1
            Ranges(X) = 999999999
            lFighterTargetServerIdx(X) = -1
            lFighterTargetRange(X) = 999999999
            mlMissileTargetIdx(X) = -1
            mlMissileTargetRange(X) = 999999999
            mlMissileTargetID(X) = -1
            mySideHasCrits(X) = 1
        Next X
    End Sub

    Public Function FighterSubsystemHit(ByVal yArcID As Byte, ByVal yCritHitType As eyCriticalHitType) As Boolean
        'ok, the fighter has a target it is trying to hit indicated by iTargetTactic
        'the arcid is the side hit on ME
        'if my crit loc is not on the arc hit, it is no good

        Dim lStatus As Int32 = 0

        Select Case yCritHitType
            Case eyCriticalHitType.FighterSubsystemEngine
                lStatus = elUnitStatus.eEngineOperational
            Case eyCriticalHitType.FighterSubsystemHangar
                lStatus = elUnitStatus.eHangarOperational
            Case eyCriticalHitType.FighterSubsystemRadar
                lStatus = elUnitStatus.eRadarOperational
            Case eyCriticalHitType.FighterSubsystemShields
                lStatus = elUnitStatus.eShieldOperational
            Case eyCriticalHitType.FighterSubsystemWeapons
                'ok, weapons are special, now, we need to determine arc
                Select Case yArcID
                    Case UnitArcs.eForwardArc
                        If Rnd() * 100.0F > 50 Then lStatus = elUnitStatus.eForwardWeapon1 Else lStatus = elUnitStatus.eForwardWeapon2
                    Case UnitArcs.eLeftArc
                        If Rnd() * 100.0F > 50 Then lStatus = elUnitStatus.eLeftWeapon1 Else lStatus = elUnitStatus.eLeftWeapon2
                    Case UnitArcs.eBackArc
                        If Rnd() * 100.0F > 50 Then lStatus = elUnitStatus.eAftWeapon1 Else lStatus = elUnitStatus.eAftWeapon2
                    Case UnitArcs.eRightArc
                        If Rnd() * 100.0F > 50 Then lStatus = elUnitStatus.eRightWeapon1 Else lStatus = elUnitStatus.eRightWeapon2
                    Case Else
                        Return False
                End Select
            Case Else
                Return False
        End Select

        'Ok, is our status something?
        If lStatus <> 0 Then
            'ok, now, determine if the arc being shot at has this status
            'Get this entity's def
            Dim oDef As Epica_Entity_Def = Nothing
            If lEntityDefServerIndex <> -1 Then
                If glEntityDefIdx(lEntityDefServerIndex) <> -1 Then
                    oDef = goEntityDefs(lEntityDefServerIndex)
                Else
                    Return False
                End If
            Else
                Return False
            End If
            If oDef Is Nothing Then Return False

            'all-arc hits never do this type of dmg
            If yArcID > 3 Then Return False

            'does this side have crits left?
            If mySideHasCrits(yArcID) = 0 Then Return False

            'does this status still exist?
            If (CurrentStatus And lStatus) = 0 Then Return False

            If lStatus = elUnitStatus.eEngineOperational Then
                HandleEngineCritical(oDef)
            End If

            Select Case yArcID
                Case UnitArcs.eForwardArc
                    If oDef.lSide1Crits Is Nothing = False Then
                        For X As Int32 = 0 To oDef.lSide1Crits.GetUpperBound(0)
                            If (oDef.lSide1Crits(X) And lStatus) <> 0 Then
                                'ok, this side has a crit, so we hit it...
                                CurrentStatus = CurrentStatus Xor lStatus
                                Return True
                            End If
                        Next X
                    End If
                Case UnitArcs.eLeftArc
                    If oDef.lSide2Crits Is Nothing = False Then
                        For X As Int32 = 0 To oDef.lSide2Crits.GetUpperBound(0)
                            If (oDef.lSide2Crits(X) And lStatus) <> 0 Then
                                'ok, this side has a crit, so we hit it...
                                CurrentStatus = CurrentStatus Xor lStatus
                                Return True
                            End If
                        Next X
                    End If
                Case UnitArcs.eBackArc
                    If oDef.lSide3Crits Is Nothing = False Then
                        For X As Int32 = 0 To oDef.lSide3Crits.GetUpperBound(0)
                            If (oDef.lSide3Crits(X) And lStatus) <> 0 Then
                                'ok, this side has a crit, so we hit it...
                                CurrentStatus = CurrentStatus Xor lStatus
                                Return True
                            End If
                        Next X
                    End If
                Case UnitArcs.eRightArc
                    If oDef.lSide4Crits Is Nothing = False Then
                        For X As Int32 = 0 To oDef.lSide4Crits.GetUpperBound(0)
                            If (oDef.lSide4Crits(X) And lStatus) <> 0 Then
                                'ok, this side has a crit, so we hit it...
                                CurrentStatus = CurrentStatus Xor lStatus
                                Return True
                            End If
                        Next X
                    End If
                Case Else
                    Return False
            End Select
        End If

        Return False
    End Function

    Public Function CriticalHit(ByVal yArcID As Byte) As Boolean    'returns true if the ship is destroyed
        Dim lCritLoc As Int32
        Dim X As Int32
        Dim lLastValidIdx As Int32

        Dim lCritAryLen As Int32

        Dim lStatusChanged As Int32             'indicates which critical was hit...

        Dim bDmgPassThru As Boolean = False     'indicates that damage passed thru one side to the other

        Dim oDef As Epica_Entity_Def

        'Does this side have any more crits left???
        If mySideHasCrits(yArcID) = 0 Then
            'No, so set the new arc ID (damage pass thru goes to opposite end)
            Select Case yArcID
                Case UnitArcs.eForwardArc : yArcID = UnitArcs.eBackArc
                Case UnitArcs.eLeftArc : yArcID = UnitArcs.eRightArc
                Case UnitArcs.eRightArc : yArcID = UnitArcs.eLeftArc
                Case UnitArcs.eBackArc : yArcID = UnitArcs.eForwardArc
            End Select
            bDmgPassThru = True
        End If

        'Get this entity's def
        If lEntityDefServerIndex <> -1 Then
            If glEntityDefIdx(lEntityDefServerIndex) <> -1 Then
                oDef = goEntityDefs(lEntityDefServerIndex)
            Else
                Return False
            End If
        Else
            Return False
        End If

        With oDef
            If .mlCriticalHitChanceUB <> -1 Then
                'Ok, we can use the new method for critical hits... get our 1D100
                Dim lRoll As Int32 = CInt(Rnd() * 100)
                Dim lCumulative As Int32 = 0                'indicates the current to hit number, roll must be less than this value
                Dim lCritical As Int32 = 0                  'indicates the resulting critical

                'Loop through our critical hit chances
                For X = 0 To .mlCriticalHitChanceUB
                    'is this on the right side?
                    If .muCriticalHitChances(X).ySide = yArcID Then
                        'add it to cumulative
                        lCumulative += .muCriticalHitChances(X).yChance
                        'is roll less than cumulative?
                        If lRoll < lCumulative Then
                            'yes, we found our critical
                            lCritical = .muCriticalHitChances(X).lCritical
                            Exit For
                        End If
                    End If
                Next X
                If lCritical = 0 Then
                    'Ok, need to check if this side has any criticals left...
                    If mySideHasCrits(yArcID) <> 0 Then
                        Dim bHasCriticals As Boolean = False
                        For X = 0 To .mlCriticalHitChanceUB
                            If .muCriticalHitChances(X).ySide = yArcID Then
                                If (CurrentStatus And .muCriticalHitChances(X).lCritical) <> 0 Then
                                    bHasCriticals = True
                                    Exit For
                                End If
                            End If
                        Next X
                        If bHasCriticals = False Then mySideHasCrits(yArcID) = 0
                    End If
                Else
                    If (CurrentStatus And lCritical) <> 0 Then
                        CurrentStatus = CurrentStatus Xor lCritical 'CurrentStatus -= lCritical
                        lStatusChanged = lCritical
                    End If
                End If
            Else
                Select Case yArcID
                    Case UnitArcs.eForwardArc
                        'How many critical locs are there?
                        lCritAryLen = .lSide1Crits.Length

                        'Get our critical loc (assuming all critical locs are valid)
                        lCritLoc = CInt(Int(Rnd() * lCritAryLen))

                        'set our last valid index so that if no valid indices are found, we can set this side as unuseable
                        lLastValidIdx = -1
                        'Now, loop through our critical array
                        For X = 0 To lCritAryLen - 1
                            'does this critical still exist?
                            If (CurrentStatus And .lSide1Crits(X)) <> 0 Then
                                'yes, so set our last valid idx
                                lLastValidIdx = X

                                'is this the critical we were looking for?
                                If X = lCritLoc Then
                                    'ok, found our crit location
                                    CurrentStatus = CurrentStatus Xor .lSide1Crits(X) 'CurrentStatus -= .lSide1Crits(X)
                                    lStatusChanged = .lSide1Crits(X)
                                    Exit Select
                                End If
                            End If
                        Next X

                        'Ok, if we are here, we did not find our crit loc... do we have a valid idx?
                        If lLastValidIdx <> -1 Then
                            'yes, so use that
                            CurrentStatus = CurrentStatus Xor .lSide1Crits(lLastValidIdx) 'CurrentStatus -= .lSide1Crits(lLastValidIdx)
                            lStatusChanged = .lSide1Crits(lLastValidIdx)
                            Exit Select
                        Else
                            'Ok, no, we didn't... set our critical hit loc to the opposite side
                            If bDmgPassThru = True Then
                                Return False        'no criticals remain along this axis, so do nothing
                            Else
                                mySideHasCrits(yArcID) = 0
                            End If
                            Return False    ' Exit Function       'the pass thru crit is a gimmie
                        End If
                    Case UnitArcs.eLeftArc
                        lCritAryLen = .lSide2Crits.Length

                        lCritLoc = CInt(Int(Rnd() * lCritAryLen))

                        'ok, hit a regular crit loc...
                        lLastValidIdx = -1
                        For X = 0 To lCritAryLen - 1
                            If (CurrentStatus And .lSide2Crits(X)) <> 0 Then
                                lLastValidIdx = X

                                If X = lCritLoc Then
                                    'ok, found our crit location
                                    CurrentStatus = CurrentStatus Xor .lSide2Crits(X) ' CurrentStatus -= .lSide2Crits(X)
                                    lStatusChanged = .lSide2Crits(X)
                                    Exit Select
                                End If
                            End If
                        Next X

                        'Ok, if we are here, we did not find our crit loc... do we have a valid idx?
                        If lLastValidIdx <> -1 Then
                            'yes, so use that
                            CurrentStatus = CurrentStatus Xor .lSide2Crits(lLastValidIdx) 'CurrentStatus -= .lSide2Crits(lLastValidIdx)
                            lStatusChanged = .lSide2Crits(lLastValidIdx)
                            Exit Select
                        Else
                            'Ok, no, we didn't... set our critical hit loc to the opposite side
                            If bDmgPassThru = True Then
                                Return False 'ship not destroyed by default
                            Else
                                mySideHasCrits(yArcID) = 0
                            End If
                            Return False    ' Exit Function       'the pass thru crit is a gimmie
                        End If
                    Case UnitArcs.eBackArc
                        lCritAryLen = .lSide3Crits.Length

                        lCritLoc = CInt(Int(Rnd() * lCritAryLen))

                        'ok, hit a regular crit loc...
                        lLastValidIdx = -1
                        For X = 0 To lCritAryLen - 1
                            If (CurrentStatus And .lSide3Crits(X)) <> 0 Then
                                lLastValidIdx = X

                                If X = lCritLoc Then
                                    'ok, found our crit location
                                    CurrentStatus = CurrentStatus Xor .lSide3Crits(X) ' CurrentStatus -= .lSide3Crits(X)
                                    lStatusChanged = .lSide3Crits(X)
                                    Exit Select
                                End If
                            End If
                        Next X

                        'Ok, if we are here, we did not find our crit loc... do we have a valid idx?
                        If lLastValidIdx <> -1 Then
                            'yes, so use that
                            CurrentStatus = CurrentStatus Xor .lSide3Crits(lLastValidIdx) ' CurrentStatus -= .lSide3Crits(lLastValidIdx)
                            lStatusChanged = .lSide3Crits(lLastValidIdx)
                            Exit Select
                        Else
                            'Ok, no, we didn't... set our critical hit loc to the opposite side
                            If bDmgPassThru = True Then
                                Return False 'ship not destroyed by default
                            Else
                                mySideHasCrits(yArcID) = 0
                            End If
                            Return False 'Exit Function       'the pass thru crit is a gimmie
                        End If
                    Case UnitArcs.eRightArc
                        lCritAryLen = .lSide4Crits.Length

                        lCritLoc = CInt(Int(Rnd() * lCritAryLen))

                        'ok, hit a regular crit loc...
                        lLastValidIdx = -1
                        For X = 0 To lCritAryLen - 1
                            If (CurrentStatus And .lSide4Crits(X)) <> 0 Then
                                lLastValidIdx = X

                                If X = lCritLoc Then
                                    'ok, found our crit location
                                    CurrentStatus = CurrentStatus Xor .lSide4Crits(X) ' CurrentStatus -= .lSide4Crits(X)
                                    lStatusChanged = .lSide4Crits(X)
                                    Exit Select
                                End If
                            End If
                        Next X

                        'Ok, if we are here, we did not find our crit loc... do we have a valid idx?
                        If lLastValidIdx <> -1 Then
                            'yes, so use that
                            CurrentStatus = CurrentStatus Xor .lSide4Crits(lLastValidIdx) 'CurrentStatus -= .lSide4Crits(lLastValidIdx)
                            lStatusChanged = .lSide4Crits(lLastValidIdx)
                            Exit Select
                        Else
                            'Ok, no, we didn't... set our critical hit loc to the opposite side
                            If bDmgPassThru = True Then
                                Return False ' True 'ship destroyed by default
                            Else
                                mySideHasCrits(yArcID) = 0
                            End If
                            'Exit Function       'the pass thru crit is a gimmie
                            Return False
                        End If
                End Select
            End If

            'Ok, now, based on the status changed...
            Select Case lStatusChanged
                Case elUnitStatus.eEngineOperational ', elUnitStatus.eFuelBayOperational

                    'Is this a flying unit that is in a planet atmosphere?
                    'If (oDef.yChassisType And ChassisType.eAtmospheric) <> 0 AndAlso Me.ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                    If (oDef.yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) = 0 AndAlso Me.ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                        'yes, so, this kills the unit...
                        Return True
                    End If

                    HandleEngineCritical(oDef)

                    Return False    'no longer destroys the unit
                Case elUnitStatus.eHangarOperational, elUnitStatus.eCargoBayOperational
                    goMsgSys.SendHangarCargoDestroyed(Me, lStatusChanged)
                    Return False
                Case Else
                    Return False
            End Select


        End With
    End Function

    Public Sub HandleEngineCritical(ByRef oDef As Epica_Entity_Def)
        'are we set to maneuver?
        If (Me.iCombatTactics And eiBehaviorPatterns.eTactics_Maneuver) <> 0 Then
            'remove that setting...
            Me.iCombatTactics -= eiBehaviorPatterns.eTactics_Maneuver
        End If

        'Tell the Pathfinding Server to not continue any dest lists for this object
        goMsgSys.SendClearDestList(Me)

        'Need to determine the final resting place for this unit
        Dim fTemp As Single = Math.Max(Math.Abs(VelX), Math.Abs(VelZ))
        If fTemp <> 0 AndAlso (Me.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
            fTemp = fTemp / Acceleration

            'So, at this point we have a VelX and VelZ
            Dim fVelX As Single = Math.Abs(VelX)
            Dim fVelZ As Single = Math.Abs(VelZ)
            Dim fVecAngleRads As Single
            If fVelX = 0 Then
                fVecAngleRads = gdPieAndAHalf
            ElseIf fVelZ = 0 Then
                fVecAngleRads = 0
            Else
                fVecAngleRads = CSng(Math.Atan(Math.Abs(fVelZ / fVelX)))
                fVecAngleRads = gdTwoPie - fVecAngleRads
            End If
            Const fPiOver180 As Single = gdPi / 180.0F
            fVecAngleRads = (LocAngle * 0.1F) * fPiOver180 - fVecAngleRads
            Dim fTmpCos As Single = CSng(Math.Cos(fVecAngleRads))
            Dim fTmpSin As Single = CSng(Math.Sin(fVecAngleRads))
            Dim fTempVelX As Single = (fVelX * fTmpCos) + (fVelZ * fTmpSin)
            Dim fTempVelZ As Single = -((fVelX * fTmpSin) - (fVelZ * fTmpCos))

            'Now... velx and velz are never negative...
            Dim fDX As Single = MoveLocX + (fTempVelX * fTemp)
            Dim fDZ As Single = MoveLocZ + (fTempVelZ * fTemp)

            If fDX < Int32.MinValue OrElse fDX > Int32.MaxValue Then fDX = MoveLocX
            If fDZ < Int32.MinValue OrElse fDZ > Int32.MaxValue Then fDZ = MoveLocZ

            'Finally... tell everyone the move loc
            SetDest(CInt(fDX), CInt(fDZ), LocAngle)
            yChangeEnvironments = 0
            Dim yMsg(17) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yMsg, 0)
            oDef.GetGUIDAsString.CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(DestX).CopyTo(yMsg, 8)
            System.BitConverter.GetBytes(DestZ).CopyTo(yMsg, 12)
            System.BitConverter.GetBytes(LocAngle).CopyTo(yMsg, 16)
            If ParentEnvir.lPlayersInEnvirCnt > 0 Then goMsgSys.BroadcastToEnvironmentClients(yMsg, ParentEnvir)
        End If
    End Sub

    Public Sub AddTargetedBy(ByVal lServerIdx As Int32, ByVal lEntityID As Int32)
        Dim X As Int32
        Dim lIdx As Int32

        lIdx = -1
        For X = 0 To lTargetedByUB
            If lTargetedByIdx(X) = -1 OrElse lTargetedByIdx(X) = lServerIdx Then        'msc 12/13/08 - was just an Or... save some time asking OrElse
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            lTargetedByUB += 1
            ReDim Preserve lTargetedByIdx(lTargetedByUB)
            ReDim Preserve lTargetedByID(lTargetedByUB)
            lIdx = lTargetedByUB
        End If

        lTargetedByIdx(lIdx) = lServerIdx
        lTargetedByID(lIdx) = lEntityID
    End Sub
    Public Sub RemoveTargetedBy(ByVal lServerIdx As Int32)
        Dim X As Int32

        For X = 0 To lTargetedByUB
            If lTargetedByIdx(X) = lServerIdx Then
                lTargetedByIdx(X) = -1
                lTargetedByID(X) = -1
                Exit For
            End If
        Next X
    End Sub

    Private Shared mbDoNewCodeA As Boolean = True
    Private Shared mbDoNewCodeB As Boolean = True
    Public Sub VerifyTargets()
        Dim X As Int32
        Dim Y As Int32
        Dim lFaceTargets(3) As Int32
        Dim lFacing As Int32
        Dim lDist As Int32

        Dim lTemp As Int32

        Dim oTmpUnit As Epica_Entity
        Dim yGridResult As Byte

        Dim lRelTinyX As Int32
        Dim lRelTinyZ As Int32

        Dim lFighterTargets(3) As Int32

        'Dim bHadTargets As Boolean

        'Ok, first, let's go through who I have targeted...
        For Y = 0 To 3
            lFaceTargets(Y) = lTargetsServerIdx(Y)
            lTargetsServerIdx(Y) = -1
            lFighterTargets(Y) = lFighterTargetServerIdx(Y)
            lFighterTargetServerIdx(Y) = -1
        Next Y
        'bHadTargets = ((CurrentStatus And elUnitStatus.eSide1HasTarget) <> 0) Or _
        '    ((CurrentStatus And elUnitStatus.eSide2HasTarget) <> 0) Or _
        '    ((CurrentStatus And elUnitStatus.eSide3HasTarget) <> 0) Or _
        '    ((CurrentStatus And elUnitStatus.eSide4HasTarget) <> 0)

        'clear our engagement flags
        If (CurrentStatus And elUnitStatus.eSide1HasTarget) <> 0 Then CurrentStatus = CurrentStatus Xor elUnitStatus.eSide1HasTarget ' CurrentStatus -= elUnitStatus.eSide1HasTarget
        If (CurrentStatus And elUnitStatus.eSide2HasTarget) <> 0 Then CurrentStatus = CurrentStatus Xor elUnitStatus.eSide2HasTarget 'CurrentStatus -= elUnitStatus.eSide2HasTarget
        If (CurrentStatus And elUnitStatus.eSide3HasTarget) <> 0 Then CurrentStatus = CurrentStatus Xor elUnitStatus.eSide3HasTarget 'CurrentStatus -= elUnitStatus.eSide3HasTarget
        If (CurrentStatus And elUnitStatus.eSide4HasTarget) <> 0 Then CurrentStatus = CurrentStatus Xor elUnitStatus.eSide4HasTarget 'CurrentStatus -= elUnitStatus.eSide4HasTarget

        'And update them
        For Y = 0 To 3
            If lFaceTargets(Y) <> -1 Then
                lFacing = Int32.MaxValue
                lDist = Int32.MaxValue
                If goEntity(lFaceTargets(Y)) Is Nothing = False Then
                    With goEntity(lFaceTargets(Y))
                        If (.CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 Then Continue For
                        yGridResult = gyLargeGridArray(.lGridIndex, lGridIndex, .ParentEnvir.lGridsPerRow)
                        If yGridResult <> 255 Then
                            lTemp = giRelativeSmall(yGridResult, lSmallSectorID, .lSmallSectorID)
                            If lTemp <> -1 Then
                                lRelTinyX = glBaseRelTinyX(lTemp) + .lTinyX + (gl_HALF_FINAL_PER_ROW - lTinyX)
                                If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                    lRelTinyZ = glBaseRelTinyZ(lTemp) + .lTinyZ + (gl_HALF_FINAL_PER_ROW - lTinyZ)
                                    If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                        lDist = gyDistances(lRelTinyX, lRelTinyZ) - .lModelRangeOffset
                                        lFacing = glFacing(lRelTinyX, lRelTinyZ)
                                    End If
                                End If
                            End If

                            lFacing -= LocAngle
                            If lFacing < 0 Then
                                lFacing = 3600 + (lFacing Mod 3600)
                            Else
                                lFacing = lFacing Mod 3600
                            End If
                            lFacing = lFacing \ 10 ' CInt(lFacing / 10)
                            lFacing = AngleToQuadrant(lFacing)
                        End If
                    End With
                End If

                If lDist <= goEntityDefs(lEntityDefServerIndex).lOptRadarRange + lJamRangeMod Then
                    lTargetsServerIdx(lFacing) = lFaceTargets(Y)
                    Ranges(lFacing) = lDist
                End If
            End If
            If lFighterTargets(Y) <> -1 Then
                lFacing = Int32.MaxValue
                lDist = Int32.MaxValue
                If goEntity(lFighterTargets(Y)) Is Nothing = False Then
                    With goEntity(lFighterTargets(Y))
                        If (.CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 Then Continue For
                        yGridResult = gyLargeGridArray(.lGridIndex, lGridIndex, .ParentEnvir.lGridsPerRow)
                        If yGridResult <> 255 Then
                            lTemp = giRelativeSmall(yGridResult, lSmallSectorID, .lSmallSectorID)
                            If lTemp <> -1 Then
                                lRelTinyX = glBaseRelTinyX(lTemp) + .lTinyX + (gl_HALF_FINAL_PER_ROW - lTinyX)
                                If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                    lRelTinyZ = glBaseRelTinyZ(lTemp) + .lTinyZ + (gl_HALF_FINAL_PER_ROW - lTinyZ)
                                    If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                        lDist = gyDistances(lRelTinyX, lRelTinyZ) - .lModelRangeOffset
                                        lFacing = glFacing(lRelTinyX, lRelTinyZ)
                                    End If
                                End If
                            End If

                            lFacing -= LocAngle
                            If lFacing < 0 Then
                                lFacing = 3600 + (lFacing Mod 3600)
                            Else
                                lFacing = lFacing Mod 3600
                            End If
                            lFacing = lFacing \ 10 ' CInt(lFacing / 10)
                            lFacing = AngleToQuadrant(lFacing)
                        End If
                    End With
                End If

                If lDist <= goEntityDefs(lEntityDefServerIndex).lOptRadarRange + lJamRangeMod Then
                    lFighterTargetServerIdx(lFacing) = lFighterTargets(Y)
                    lFighterTargetRange(lFacing) = lDist
                End If
            End If
        Next Y

        'Check Primary Target
        If lPrimaryTargetServerIdx <> -1 Then
            If glEntityIdx(lPrimaryTargetServerIdx) > 0 Then
                Dim oPrimEntity As Epica_Entity = goEntity(lPrimaryTargetServerIdx)
                If oPrimEntity Is Nothing = False Then
                    With oPrimEntity
                        yGridResult = gyLargeGridArray(.lGridIndex, lGridIndex, .ParentEnvir.lGridsPerRow)
                        lDist = Int32.MaxValue
                        If yGridResult <> 255 Then
                            lTemp = giRelativeSmall(yGridResult, lSmallSectorID, .lSmallSectorID)
                            If lTemp <> -1 Then
                                lRelTinyX = glBaseRelTinyX(lTemp) + .lTinyX + (gl_HALF_FINAL_PER_ROW - lTinyX)
                                If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                    lRelTinyZ = glBaseRelTinyZ(lTemp) + .lTinyZ + (gl_HALF_FINAL_PER_ROW - lTinyZ)
                                    If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                        lDist = gyDistances(lRelTinyX, lRelTinyZ) - .lModelRangeOffset
                                    End If
                                End If
                            End If
                        End If


                        If lDist > goEntityDefs(lEntityDefServerIndex).lOptRadarRange + .lJamRangeMod OrElse (.CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 Then
                            'ok, check our behavior
                            If ((CurrentStatus And elUnitStatus.eUnitMoving) <> 0 AndAlso (CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0) OrElse (.CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 Then
                                'yes, we were, then kill our target
                                lPrimaryTargetServerIdx = -1
                                bForceAggressionTest = True
                                SyncLockMovementRegisters(MovementCommand.AddEntityMoving, ServerIndex, ObjectID) 'AddEntityMoving(ServerIndex, ObjectID)
                                .RemoveTargetedBy(ServerIndex)
                                If (CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                                    CurrentStatus = CurrentStatus Xor elUnitStatus.eUnitEngaged 'CurrentStatus -= elUnitStatus.eUnitEngaged
                                    If mbDoNewCodeA = True Then
                                        RemoveEntityInCombat(ObjectID, ObjTypeID)
                                    Else
                                        RemoveEntityInCombat(.ObjectID, .ObjTypeID)
                                    End If
                                End If
                            End If

                            'did we kill our target?
                            If lPrimaryTargetServerIdx <> -1 Then
                                'no, ok, check our tactics, do we pursue or engage?
                                If (((iCombatTactics And eiBehaviorPatterns.eEngagement_Pursue) <> 0) OrElse _
                                   ((iCombatTactics And eiBehaviorPatterns.eEngagement_Engage) <> 0)) AndAlso bAIMoveRequestPending = False AndAlso MaxSpeed <> 0 Then
                                    'Ok, the unit needs to move to keep the target in range...
                                    Dim lDX As Int32 = .LocX '- Me.LocX
                                    Dim lDZ As Int32 = .LocZ '- Me.LocZ
                                    'Dim fTime As Single = Math.Abs((lDX + lDZ) / CSng(MaxSpeed)) ' goEntityDefs(lEntityDefServerIndex).MaxSpeed)

                                    'lDX = .LocX + CInt(.VelX * fTime)
                                    'lDZ = .LocZ + CInt(.VelZ * fTime)
                                    'fTime += ((Math.Abs(.VelX) + Math.Abs(.VelZ)) * fTime) / CSng(MaxSpeed) ' goEntityDefs(lEntityDefServerIndex).MaxSpeed

                                    'Dim fTmpDX As Single = .LocX + (.VelX * fTime)
                                    'Dim fTmpDZ As Single = .LocZ + (.VelZ * fTime)
                                    'If fTmpDX > Int32.MinValue AndAlso fTmpDX < Int32.MaxValue Then
                                    '    If fTmpDZ > Int32.MinValue AndAlso fTmpDZ < Int32.MaxValue Then
                                    '        lDX = CInt(fTmpDX)
                                    '        lDZ = CInt(fTmpDZ)
                                    '        'lDX and lDZ supposedly now have our destination, we need to send this to the pathfinding server
                                    goMsgSys.SendAIMoveRequestToPathfinding(Me, lDX, lDZ, 0S)
                                    '    End If
                                    'End If
                                Else
                                    'no, clear the target
                                    lPrimaryTargetServerIdx = -1
                                    bForceAggressionTest = True
                                    SyncLockMovementRegisters(MovementCommand.AddEntityMoving, ServerIndex, ObjectID) 'AddEntityMoving(ServerIndex, ObjectID)
                                    .RemoveTargetedBy(ServerIndex)
                                    If (CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                                        'RemoveEntityCombat(.ServerIndex, .ObjectID)
                                        If mbDoNewCodeB = True Then
                                            CurrentStatus = CurrentStatus Xor elUnitStatus.eUnitEngaged ' .CurrentStatus -= elUnitStatus.eUnitEngaged
                                            RemoveEntityInCombat(ObjectID, ObjTypeID)
                                        Else
                                            .CurrentStatus = .CurrentStatus Xor elUnitStatus.eUnitEngaged ' .CurrentStatus -= elUnitStatus.eUnitEngaged
                                            RemoveEntityInCombat(.ObjectID, .ObjTypeID)
                                        End If

                                    End If
                                End If
                            End If
                        End If
                    End With
                Else : lPrimaryTargetServerIdx = -1
                End If
            Else : lPrimaryTargetServerIdx = -1
            End If
        End If

        'finally, update our statuses
        If lTargetsServerIdx(0) <> -1 Then CurrentStatus = CurrentStatus Or elUnitStatus.eSide1HasTarget
        If lTargetsServerIdx(1) <> -1 Then CurrentStatus = CurrentStatus Or elUnitStatus.eSide2HasTarget
        If lTargetsServerIdx(2) <> -1 Then CurrentStatus = CurrentStatus Or elUnitStatus.eSide3HasTarget
        If lTargetsServerIdx(3) <> -1 Then CurrentStatus = CurrentStatus Or elUnitStatus.eSide4HasTarget

        'If we lose a target, then we need to remove us from the Targeted By of that target!!!
        For X = 0 To 3
            If lFaceTargets(X) <> -1 Then
                For Y = 0 To 3
                    If lTargetsServerIdx(Y) = lFaceTargets(X) Then
                        lFaceTargets(X) = -1
                        Exit For
                    End If
                Next Y

                If lFaceTargets(X) <> -1 Then
                    If glEntityIdx(lFaceTargets(X)) > 0 Then
                        goEntity(lFaceTargets(X)).RemoveTargetedBy(Me.ServerIndex)
                    End If
                End If
            End If
        Next X

        'Now, go through and check those that target us...
        For X = 0 To lTargetedByUB
            If lTargetedByIdx(X) <> -1 Then
                If glEntityIdx(lTargetedByIdx(X)) > 0 Then
                    'Ok, valid object
                    oTmpUnit = goEntity(lTargetedByIdx(X))
                    If oTmpUnit Is Nothing OrElse oTmpUnit.ObjectID <> lTargetedByID(X) Then
                        lTargetedByIdx(X) = -1
                        lTargetedByID(X) = -1
                        Continue For
                    End If
                    With oTmpUnit
                        'check this unit's sides for me...
                        For Y = 0 To 3
                            If oTmpUnit.lTargetsServerIdx(Y) = ServerIndex Then
                                'Ok, being targeted by this side...
                                lFacing = Int32.MaxValue
                                lDist = Int32.MaxValue

                                'yGridResult = gyLargeGridArray(.lGridIndex, lGridIndex, .ParentEnvir.lGridsPerRow)
                                yGridResult = gyLargeGridArray(lGridIndex, .lGridIndex, .ParentEnvir.lGridsPerRow)
                                If yGridResult <> 255 Then
                                    lTemp = giRelativeSmall(yGridResult, .lSmallSectorID, lSmallSectorID)
                                    If lTemp <> -1 Then
                                        lRelTinyX = glBaseRelTinyX(lTemp) + lTinyX + (gl_HALF_FINAL_PER_ROW - .lTinyX)
                                        If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                            lRelTinyZ = glBaseRelTinyZ(lTemp) + lTinyZ + (gl_HALF_FINAL_PER_ROW - .lTinyZ)
                                            If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                                lDist = gyDistances(lRelTinyX, lRelTinyZ) - lModelRangeOffset
                                                lFacing = glFacing(lRelTinyX, lRelTinyZ)
                                            End If
                                        End If
                                    End If

                                    lFacing -= .LocAngle
                                    If lFacing < 0 Then
                                        lFacing = 3600 + (lFacing Mod 3600)
                                    Else
                                        lFacing = lFacing Mod 3600
                                    End If
                                    lFacing = lFacing \ 10 'CInt(lFacing / 10)
                                    lFacing = AngleToQuadrant(lFacing)
                                End If      'If yGridResult <> 255 Then

                                If lFacing <> Y OrElse lDist > goEntityDefs(.lEntityDefServerIndex).lOptRadarRange + .lJamRangeMod OrElse (CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 Then
                                    'need to remove this from the target list
                                    lTemp = 0
                                    Select Case Y
                                        Case 0
                                            lTemp = elUnitStatus.eSide1HasTarget
                                        Case 1
                                            lTemp = elUnitStatus.eSide2HasTarget
                                        Case 2
                                            lTemp = elUnitStatus.eSide3HasTarget
                                        Case 3
                                            lTemp = elUnitStatus.eSide4HasTarget
                                    End Select

                                    If lTemp <> 0 AndAlso (.CurrentStatus And lTemp) <> 0 Then .CurrentStatus = .CurrentStatus Xor lTemp ' .CurrentStatus -= lTemp
                                    If (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then
                                        .bForceAggressionTest = True
                                        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                    End If
                                Else
                                    .Ranges(lFacing) = lDist
                                End If

                            End If
                        Next Y

                        'Need check its primary target...
                        If .lPrimaryTargetServerIdx = ServerIndex Then
                            'It does, okay, is the distance outside this unit's max radar range?
                            'If lDist > goEntityDefs(.lEntityDefServerIndex).lMaxRadarRange Then
                            '    'ok, the unit is forced to lose the target, for it no longer is trackable on radar
                            '    'TODO: One day, it might be a good idea to determine if the unit is visible to any unit
                            '    '  of this unit's owning player to determine if the unit can still be seen *somewhere*
                            '    .lPrimaryTargetServerIdx = -1
                            '    .bForceAggressionTest = True
                            '    RemoveTargetedBy(.ServerIndex)
                            '    If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then .CurrentStatus -= elUnitStatus.eUnitEngaged
                            'ElseIf lDist > goEntityDefs(.lEntityDefServerIndex).lOptRadarRange Then
                            If lDist > goEntityDefs(.lEntityDefServerIndex).lOptRadarRange + .lJamRangeMod OrElse (CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 Then
                                'ok, check our behavior
                                If ((.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 AndAlso (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0) OrElse (CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 Then
                                    'yes, we were, then kill our target
                                    .lPrimaryTargetServerIdx = -1
                                    .bForceAggressionTest = True
                                    SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                    RemoveTargetedBy(.ServerIndex)
                                    If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                                        .CurrentStatus = .CurrentStatus Xor elUnitStatus.eUnitEngaged ' .CurrentStatus -= elUnitStatus.eUnitEngaged
                                        RemoveEntityInCombat(.ObjectID, .ObjTypeID)
                                    End If
                                End If

                                'did we kill our target?
                                If .lPrimaryTargetServerIdx <> -1 Then
                                    'no, ok, check our tactics, do we pursue or engage?
                                    If .MaxSpeed <> 0 AndAlso ((((.iCombatTactics And eiBehaviorPatterns.eEngagement_Pursue) <> 0) OrElse _
                                       ((.iCombatTactics And eiBehaviorPatterns.eEngagement_Engage) <> 0)) AndAlso _
                                       .bAIMoveRequestPending = False AndAlso ((.CurrentStatus And elUnitStatus.eRadarOperational) <> 0)) Then
                                        'Ok, the unit needs to move to keep the target in range...
                                        Dim lDX As Int32 = .LocX
                                        Dim lDZ As Int32 = .LocZ
                                        'Dim fTime As Single = Math.Abs((lDX + lDZ) / CSng(.MaxSpeed)) ' goEntityDefs(.lEntityDefServerIndex).MaxSpeed)

                                        'lDX = LocX + CInt(VelX * fTime)
                                        'lDZ = LocZ + CInt(VelZ * fTime)
                                        'fTime += ((Math.Abs(VelX) + Math.Abs(VelZ)) * fTime) / CSng(.MaxSpeed) ' goEntityDefs(.lEntityDefServerIndex).MaxSpeed

                                        'Dim fTmpDX As Single = LocX + (VelX * fTime)
                                        'Dim fTmpDZ As Single = LocZ + (VelZ * fTime)
                                        'If fTmpDX > Int32.MinValue AndAlso fTmpDX < Int32.MaxValue Then
                                        '    If fTmpDZ > Int32.MinValue AndAlso fTmpDZ < Int32.MaxValue Then
                                        '        lDX = CInt(fTmpDX)
                                        '        lDZ = CInt(fTmpDZ)
                                        '        'lDX and lDZ supposedly now have our destination, we need to send this to the pathfinding server
                                        goMsgSys.SendAIMoveRequestToPathfinding(oTmpUnit, lDX, lDZ, 0S)
                                        '    End If
                                        'End If
                                    Else
                                        'no, clear the target
                                        .lPrimaryTargetServerIdx = -1
                                        If (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then
                                            .bForceAggressionTest = True
                                            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                        End If
                                        RemoveTargetedBy(.ServerIndex)
                                        If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                                            .CurrentStatus = .CurrentStatus Xor elUnitStatus.eUnitEngaged '.CurrentStatus -= elUnitStatus.eUnitEngaged
                                            RemoveEntityInCombat(.ObjectID, .ObjTypeID)
                                        End If
                                    End If      'If ((.iCombatTactics And eiBehaviorPatterns.eEngagement_Pursue) <> 0) OrElse _
                                End If      'If .lPrimaryTargetServerIdx <> -1 Then
                            End If      'If lDist > goEntityDefs(.lEntityDefServerIndex).MaxRadarRange Then
                        End If      'If .lPrimaryTargetServerIdx = ServerIndex Then

                    End With

                End If      'If glEntityIdx(lTargetedByIdx(X)) <> -1 Then
            End If      'If lTargetedByIdx(X) <> -1 Then
        Next X

    End Sub

    Private mbExpRewarded As Boolean = False
    Public Sub ClearTargetLists(ByVal bApplyExpGain As Boolean, ByVal lKillShotID As Int32)
        Dim lTmpTargetByUB As Int32 = -1
        If lTargetedByIdx Is Nothing = False AndAlso lTargetedByID Is Nothing = False Then
            lTmpTargetByUB = Math.Min(lTargetedByID.GetUpperBound(0), Math.Min(lTargetedByUB, lTargetedByIdx.GetUpperBound(0)))
        End If

        Dim oCurDef As Epica_Entity_Def = Nothing
        If Me.lEntityDefServerIndex > -1 AndAlso glEntityDefIdx(Me.lEntityDefServerIndex) <> -1 Then
            oCurDef = goEntityDefs(Me.lEntityDefServerIndex)
        End If
        If oCurDef Is Nothing Then bApplyExpGain = False

        If mbExpRewarded = True Then
            bApplyExpGain = False
        End If
        If bApplyExpGain = True Then mbExpRewarded = True
        'If Me.lOwnerID <> gl_HARDCODE_PIRATE_PLAYER_ID OrElse lKillShotID <> -1 Then
        '    If (oCurDef.yChassisType And ChassisType.eGroundBased) = 0 OrElse oCurDef.ProductionTypeID <> ProductionType.eFacility Then
        '        If Me.ObjTypeID = ObjectType.eFacility OrElse Me.lOwnerID = gl_HARDCODE_PIRATE_PLAYER_ID Then
        '            FacDistributeWarpoints(oCurDef.ModelID, False)
        '        Else
        '            DistributeWarpoints(False)
        '        End If
        '    End If
        'End If

        For lTmp As Int32 = 0 To lTmpTargetByUB
            If lTargetedByIdx(lTmp) <> -1 Then
                If glEntityIdx(lTargetedByIdx(lTmp)) > 0 AndAlso lTargetedByID(lTmp) = glEntityIdx(lTargetedByIdx(lTmp)) Then
                    Dim oTmpUnit As Epica_Entity = goEntity(lTargetedByIdx(lTmp))
                    If oTmpUnit Is Nothing = False Then
                        If bApplyExpGain = True Then oTmpUnit.ApplyExpGain(oCurDef.CombatRating, lKillShotID = lTargetedByID(lTmp))

                        If oTmpUnit.lPrimaryTargetServerIdx = ServerIndex Then
                            If (oTmpUnit.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                                oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eUnitEngaged
                                RemoveEntityInCombat(oTmpUnit.ObjectID, oTmpUnit.ObjTypeID)
                            End If
                            oTmpUnit.lPrimaryTargetServerIdx = -1
                            oTmpUnit.bForceAggressionTest = True
                            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                        End If
                        If oTmpUnit.lTargetsServerIdx(0) = ServerIndex Then
                            oTmpUnit.lTargetsServerIdx(0) = -1
                            If (oTmpUnit.CurrentStatus And elUnitStatus.eSide1HasTarget) <> 0 Then
                                oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eSide1HasTarget ' oTmpUnit.CurrentStatus -= elUnitStatus.eSide1HasTarget
                            End If
                            oTmpUnit.bForceAggressionTest = True
                            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                        End If
                        If oTmpUnit.lTargetsServerIdx(1) = ServerIndex Then
                            oTmpUnit.lTargetsServerIdx(1) = -1
                            If (oTmpUnit.CurrentStatus And elUnitStatus.eSide2HasTarget) <> 0 Then
                                oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eSide2HasTarget ' oTmpUnit.CurrentStatus -= elUnitStatus.eSide2HasTarget
                            End If
                            oTmpUnit.bForceAggressionTest = True
                            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                        End If
                        If oTmpUnit.lTargetsServerIdx(2) = ServerIndex Then
                            oTmpUnit.lTargetsServerIdx(2) = -1
                            If (oTmpUnit.CurrentStatus And elUnitStatus.eSide3HasTarget) <> 0 Then
                                oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eSide3HasTarget ' oTmpUnit.CurrentStatus -= elUnitStatus.eSide3HasTarget
                            End If
                            oTmpUnit.bForceAggressionTest = True
                            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                        End If
                        If oTmpUnit.lTargetsServerIdx(3) = ServerIndex Then
                            oTmpUnit.lTargetsServerIdx(3) = -1
                            If (oTmpUnit.CurrentStatus And elUnitStatus.eSide4HasTarget) <> 0 Then
                                oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eSide4HasTarget 'oTmpUnit.CurrentStatus -= elUnitStatus.eSide4HasTarget
                            End If
                            oTmpUnit.bForceAggressionTest = True
                            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                        End If
                    End If
                End If
            End If
        Next lTmp

        If Me.lTargetsServerIdx Is Nothing = False AndAlso Me.lTargetedByID Is Nothing = False Then
            Dim lMaxUB As Int32 = Math.Min(Me.lTargetsServerIdx.GetUpperBound(0), Me.lTargetedByID.GetUpperBound(0))
            For lTmp As Int32 = 0 To lMaxUB
                Me.lTargetsServerIdx(lTmp) = -1
                Me.lTargetedByID(lTmp) = -1
            Next lTmp
            Me.lPrimaryTargetServerIdx = -1
            Dim lTemp As Int32 = (Me.CurrentStatus And (elUnitStatus.eUnitEngaged Or elUnitStatus.eSide1HasTarget Or elUnitStatus.eSide2HasTarget Or elUnitStatus.eSide3HasTarget Or elUnitStatus.eSide4HasTarget))
            If lTemp <> 0 Then
                If (Me.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                    RemoveEntityInCombat(ObjectID, ObjTypeID)
                End If
                Me.CurrentStatus = Me.CurrentStatus Xor lTemp
            End If
        End If
    End Sub

    Public Function GetObjectUpdateMsg(ByVal iMsgCode As Int16, ByVal yRemoveType As Byte) As Byte()
        Dim yResp() As Byte
        Dim lPos As Int32 = 0
        Dim X As Int32
        If lWpnAmmoCnt Is Nothing Then ReDim lWpnAmmoCnt(-1)

        If iMsgCode = GlobalMessageCode.eRemoveObject Then
            ReDim yResp(69 + (Me.lWpnAmmoCnt.Length * 8))  '63
        Else
            ReDim yResp(76 + (Me.lWpnAmmoCnt.Length * 8))
        End If

        System.BitConverter.GetBytes(iMsgCode).CopyTo(yResp, 0) : lPos += 2
        GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6

        If iMsgCode = GlobalMessageCode.eRemoveObject Then
            yResp(lPos) = yRemoveType : lPos += 1
        End If

        System.BitConverter.GetBytes(LocX).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(LocZ).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(LocAngle).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(Fuel_Cap).CopyTo(yResp, lPos) : lPos += 4
        yResp(lPos) = Exp_Level : lPos += 1
        System.BitConverter.GetBytes(iTargetingTactics).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(iCombatTactics).CopyTo(yResp, lPos) : lPos += 4

        If iMsgCode <> GlobalMessageCode.eRemoveObject Then
            System.BitConverter.GetBytes(lExtendedCT_1).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lExtendedCT_2).CopyTo(yResp, lPos) : lPos += 4
        End If

        'To ensure that we don't error out
        If ParentEnvir Is Nothing = False Then
            System.BitConverter.GetBytes(ParentEnvir.ObjectID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(ParentEnvir.ObjTypeID).CopyTo(yResp, lPos) : lPos += 2
        Else
            System.BitConverter.GetBytes(CInt(-1)).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(CShort(-1)).CopyTo(yResp, lPos) : lPos += 2
        End If

        If Owner Is Nothing = False Then
            System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yResp, lPos) : lPos += 4
        Else : System.BitConverter.GetBytes(CInt(-1)).CopyTo(yResp, lPos) : lPos += 4
        End If

        'current status
        System.BitConverter.GetBytes(CurrentStatus).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(ShieldHP).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(ArmorHP(0)).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(ArmorHP(1)).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(ArmorHP(2)).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(ArmorHP(3)).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(StructuralHP).CopyTo(yResp, lPos) : lPos += 4

        System.BitConverter.GetBytes(CShort(Me.lWpnAmmoCnt.Length)).CopyTo(yResp, lPos) : lPos += 2
        If Me.lEntityDefServerIndex < 0 Then Return yResp
        With goEntityDefs(Me.lEntityDefServerIndex)
            For X = 0 To .WeaponDefUB
                System.BitConverter.GetBytes(.WeaponDefs(X).lEntityWpnDefID).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(Me.lWpnAmmoCnt(X)).CopyTo(yResp, lPos) : lPos += 4
            Next X
        End With

        Return yResp

    End Function

    Public Sub ApplyExpGain(ByVal lTargetsCombatRating As Int32, ByVal bKillingBlow As Boolean)
        'Ok, first, get our strength
        Dim lMyCombatRating As Int32
        Dim lDiffRating As Int32
        Dim fMultiplier As Single
        Dim lXPGain As Int32
        Try
            If Me.ObjTypeID = ObjectType.eFacility Then Return

            If Me.lEntityDefServerIndex > -1 AndAlso glEntityDefIdx(Me.lEntityDefServerIndex) > -1 Then
                lMyCombatRating = goEntityDefs(Me.lEntityDefServerIndex).CombatRating

                'Ok, here we go, get the difference in combat rating
                lDiffRating = lTargetsCombatRating - lMyCombatRating

                'Now, add half of my combat rating to that difference
                lDiffRating += lMyCombatRating \ 2

                'Now, do we reward experience?
                If lDiffRating > 0 Then
                    'Yes, we do... how many times over did this strength exceed mine?
                    fMultiplier = CSng(lDiffRating) / lMyCombatRating

                    'If we were the killing blow
                    'If bKillingBlow = True Then fMultiplier *= 2

                    'Now, determine our current experience level modifier to the multiplier
                    fMultiplier *= ((256.0F - Exp_Level) / 255.0F)

                    'Ok, get our experience gain
                    lXPGain = CInt(Math.Min(Math.Ceiling(fMultiplier * gl_EXP_POINT_BASE), gl_MAX_EXP_GAIN))

                    Dim lPrevLevel As Int32 = CInt(Exp_Level) \ 25
                    'Now, increase our experience by that, we increment lXPGain by Exp_level to avoid overflow
                    lXPGain += Exp_Level
                    If lXPGain > 255 Then lXPGain = 255
                    Exp_Level = CByte(lXPGain)

                    Dim lNewLevel As Int32 = CInt(Exp_Level) \ 25
                    Dim lOldCP As Int32 = CPUsage '+ Owner.BadWarDecCPIncrease
                    SetExpLevelMods()

                    If lPrevLevel <> lNewLevel Then
                        If Me.ParentEnvir.lPlayersInEnvirCnt > 0 Then
                            Dim yData(10) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eUpdateEntityAttrs).CopyTo(yData, 0)
                            Me.GetGUIDAsString.CopyTo(yData, 2)
                            yData(8) = Maneuver
                            yData(9) = MaxSpeed
                            yData(10) = Exp_Level
                            goMsgSys.BroadcastToEnvironmentClients(yData, Me.ParentEnvir)
                        End If

                        If Me.ObjTypeID <> ObjectType.eFacility Then Me.ParentEnvir.AdjustPlayerCommandPoints(Me.Owner.ObjectID, CPUsage - lOldCP)

                        Dim yPrimMsg(8) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eUpdateUnitExpLevel).CopyTo(yPrimMsg, 0)
                        Me.GetGUIDAsString.CopyTo(yPrimMsg, 2)
                        yPrimMsg(8) = Exp_Level
                        goMsgSys.SendToPrimary(yPrimMsg)
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    Public Sub SetExpLevelMods()
        Dim lTemp As Int32 = CInt(Exp_Level) \ 25

        If lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(lEntityDefServerIndex) <> -1 Then
            With goEntityDefs(lEntityDefServerIndex)
                Acceleration = .BaseAcceleration
                Maneuver = .BaseManeuver
                MaxSpeed = .BaseMaxSpeed
                TurnAmount = .BaseTurnAmount
                TurnAmountTimes100 = .BaseTurnAmountTimes100
            End With
        End If

        DamageMod = 0

        Dim yMaxSpeedMod As Byte = 0
        Dim yManeuverMod As Byte = 0

        Select Case lTemp
            Case 0
                ToHitMod = 0S
            Case 1
                ToHitMod = 5S
            Case 2
                ToHitMod = 8S
            Case 3
                ToHitMod = 8S : yMaxSpeedMod = 1
            Case 4
                ToHitMod = 8S : yMaxSpeedMod = 1 : yManeuverMod = 1
            Case 5
                ToHitMod = 10S : yMaxSpeedMod = 1 : yManeuverMod = 1
            Case 6
                ToHitMod = 10S : yMaxSpeedMod = 1 : yManeuverMod = 1 : DamageMod = 5I
            Case 7
                ToHitMod = 10S : yMaxSpeedMod = 3 : yManeuverMod = 1 : DamageMod = 5I
            Case 8
                ToHitMod = 10S : yMaxSpeedMod = 4 : yManeuverMod = 2 : DamageMod = 5I
            Case Else
                ToHitMod = 13S : yMaxSpeedMod = 5 : yManeuverMod = 3 : DamageMod = 10I
        End Select
        If CInt(MaxSpeed) + CInt(yMaxSpeedMod) > 255 Then MaxSpeed = 255 Else MaxSpeed = CByte(CInt(MaxSpeed) + CInt(yMaxSpeedMod))
        If CInt(Maneuver) + CInt(yManeuverMod) > 255 Then Maneuver = 255 Else Maneuver = CByte(CInt(Maneuver) + CInt(yManeuverMod))

        If Me.ObjTypeID = ObjectType.eFacility Then
            Me.CPUsage = 0
        Else : Me.CPUsage = Math.Max(10 - lTemp, 1) '+ Me.Owner.BadWarDecCPIncrease
        End If

        Acceleration = Maneuver * 0.01F ' / 100.0F
        TurnAmount = Maneuver
        TurnAmountTimes100 = TurnAmount * 100S
    End Sub

End Class

Public Enum eyCriticalHitType As Byte
    NoCriticalHit = 0
    NormalCriticalHit = 1
    FighterSubsystemEngine = 2
    FighterSubsystemHangar = 3
    FighterSubsystemRadar = 4
    FighterSubsystemShields = 5
    FighterSubsystemWeapons = 6
End Enum

Public Enum eyAOEDmgType As Byte
    eNone = 0
    eBeamDmg = 1
    eECMDmg = 2
    eFlameDmg = 3
    eChemicalDmg = 4
    eImpactDmg = 5
    ePierceDmg = 6
End Enum
Public Class WeaponEvent
    Public lFromIdx As Int32        'server index of attacker
    Public lToIdx As Int32          'server index of target

    Public yCritical As eyCriticalHitType

    Public lPierceDmg As Int32
    Public lImpactDmg As Int32
    Public lBeamDmg As Int32
    Public lECMDmg As Int32
    Public lFlameDmg As Int32
    Public lChemicalDmg As Int32

    Public lEventCycle As Int32     'when the event occurs

    Public ySideHit As Byte

    Public AOERange As Byte     'copied from weapon def
    Public yAOEDmgType As Byte = 0

    Public FromPlayerID As Int32 = -1
    Public yAttackerType As Byte

    'Location specific
    Public lFromX As Int32
    Public lFromZ As Int32
    Public fDegradation As Single = 0.0F
    Public Sub HandlePulseDegradation(ByVal lImpactX As Int32, ByVal lImpactZ As Int32)
        'fDegradation /= 25.0F
        fDegradation *= 0.04F

        Dim fX As Single = lFromX - lImpactX
        Dim fZ As Single = lFromZ - lImpactZ
        fX *= fX
        fZ *= fZ

        Dim fTotal As Single = CSng(Math.Sqrt(fX + fZ))

        fTotal *= fDegradation
        fTotal = lBeamDmg - fTotal
        If fTotal < 1 Then fTotal = 1
        Dim lNewVal As Int32 '= CInt(lBeamDmg - fTotal)
        If fTotal > Int32.MaxValue Then lNewVal = Int32.MaxValue Else lNewVal = CInt(fTotal)
        'If lNewVal < 1 Then lNewVal = 1
        lBeamDmg = lNewVal
    End Sub

    Public Sub SetForAOEDmg()
        Select Case yAOEDmgType
            Case eyAOEDmgType.eBeamDmg
                lPierceDmg = 0 : lImpactDmg = 0 : lECMDmg = 0 : lFlameDmg = 0 : lChemicalDmg = 0
            Case eyAOEDmgType.eChemicalDmg
                lPierceDmg = 0 : lImpactDmg = 0 : lBeamDmg = 0 : lECMDmg = 0 : lFlameDmg = 0
            Case eyAOEDmgType.eECMDmg
                lPierceDmg = 0 : lImpactDmg = 0 : lBeamDmg = 0 : lFlameDmg = 0 : lChemicalDmg = 0
            Case eyAOEDmgType.eFlameDmg
                lPierceDmg = 0 : lImpactDmg = 0 : lBeamDmg = 0 : lECMDmg = 0 : lChemicalDmg = 0
            Case eyAOEDmgType.eImpactDmg
                lPierceDmg = 0 : lBeamDmg = 0 : lECMDmg = 0 : lFlameDmg = 0 : lChemicalDmg = 0
            Case eyAOEDmgType.ePierceDmg
                lImpactDmg = 0 : lBeamDmg = 0 : lECMDmg = 0 : lFlameDmg = 0 : lChemicalDmg = 0
        End Select
    End Sub
End Class

Public Class Epica_Entity_Def
    Inherits Epica_GUID

    Public Enum eExtendedFlagValues As Int32
        'Bitwise
        ContainsMicroTech = 1
    End Enum

    Public DefName(19) As Byte

    'Both Facilities and Units use this class
    Private myManeuver As Byte

    'Public TurnAmount As Byte
    'Public TurnAmountTimes100 As Int16
    'Public MaxSpeed As Byte
    'Public Acceleration As Single   'based on Maneuver... equals Maneuver / 100
    Public BaseTurnAmount As Byte
    Public BaseTurnAmountTimes100 As Int16
    Public BaseMaxSpeed As Byte
    Public BaseAcceleration As Single

    Public FuelEfficiency As Byte

    'These values are the actual values obtained by the entity def
    Public yDefOptRadarRange As Byte
    Public yDefMaxRadarRange As Byte

    'These values are optimized and are to be used in engine code and range testing
    Public lOptRadarRange As Int32
    Public lMaxRadarRange As Int32
    Public fOneOverOptRadarRange As Single

    'These values are pulled from RS_Model.dat
    'TODO: we could make a Structure called ModelData that is preloaded and referenced from this class
    Public lModelRangeOffset As Int32
    '   Public lManeuverOffsetX_Max As Int32
    '   Public lManeuverOffsetX_Min As Int32
    '   Public lManeuverOffsetZ_Max As Int32
    'Public lManeuverOffsetZ_Min As Int32
    Public lModelSizeXZ As Int32

    Public ScanResolution As Byte
    Public DisruptionResist As Byte
    Public HullSize As Int32
    Public Cargo_Cap As Int32
    Public Hangar_Cap As Int32
    Public ModelID As Short

    Public HangarLaunchType As Byte     '0 = space, 1 = land, 2 = both
    Public CombatRating As Int32

    'From the Armor Tech
    Public PiercingResist As Byte
    Public ImpactResist As Byte
    Public BeamResist As Byte
    Public ECMResist As Byte
    Public FlameResist As Byte
    Public ChemicalResist As Byte
    Public ArmorIntegrity As Byte
    Public JamImmunity As Byte
    Public JamStrength As Byte
    Public JamTargets As Byte
    Public JamEffect As Byte

    Public FastPiercingResist As Single
    Public FastImpactResist As Single
    Public FastBeamResist As Single
    Public FastECMResist As Single
    Public FastFlameResist As Single
    Public FastChemicalResist As Single

    Public Armor_MaxHP(3) As Int32      'for each side
    Public DetectionResist As Byte

    'From the Shield Tech
    Public Shield_MaxHP As Int32
    Public ShieldRecharge As Int32
    Public ShieldRechargeFreq As Int32

    Public Structure_MaxHP As Int32

    Public yChassisType As Byte

    'Public MaxDoorSize As Int32
    'Public NumberOfDoors As Byte
    Public Weapon_Acc As Byte

    Public HullSizeTargetMod As Int32

    Public WeaponDefs() As WeaponDef
    Public WeaponDefUB As Int32 = -1

    Public SideWeaponPower(3) As Int32  'the sum of all WeaponDefs.lFirePowerRating for each side
    Public lSide1Crits() As Int32
    Public lSide2Crits() As Int32
    Public lSide3Crits() As Int32
    Public lSide4Crits() As Int32

    Private mlBaseCritLocs As Int32 = 0

    Public PlanetTacticalAnalysis As Int16 = 0
    Public SpaceTacticalAnalysis As Int16 = 0

    Public yFXColor As Byte = 0

    Public ProductionTypeID As Byte
    Public RequiredProductionTypeID As Byte

    Public lFighterCritDelay As Int32 = 30

    Public CriticalHitModifier As Byte = 0

    Public MaxCrew As Int32
    Public WorkerFactor As Int32
    Public ProdFactor As Int32

    Public lExtendedFlags As Int32 = 0
    'Public lWarpointValue As Int32
    Public yHullType As Epica_Entity.eyHullTypeDmgMod
    Public lMicroReliabilityTest As Int32 = 2000

    Public lAbsoluteTotalHP As Int32 = 1

    Public Structure CriticalHitChance
        Public ySide As UnitArcs
        Public lCritical As Int32
        Public yChance As Byte
    End Structure
    Public muCriticalHitChances() As CriticalHitChance
    Public mlCriticalHitChanceUB As Int32 = -1
    Public Sub AddCriticalHitChance(ByVal ySide As UnitArcs, ByVal lCritical As Int32, ByVal yChance As Byte)
        For X As Int32 = 0 To mlCriticalHitChanceUB
            With muCriticalHitChances(X)
                If .lCritical = lCritical AndAlso .ySide = ySide Then
                    .yChance = Math.Max(.yChance, yChance)
                    Return
                End If
            End With
        Next X
        mlCriticalHitChanceUB += 1
        ReDim Preserve muCriticalHitChances(mlCriticalHitChanceUB)
        With muCriticalHitChances(mlCriticalHitChanceUB)
            .lCritical = lCritical
            .yChance = yChance
            .ySide = ySide
        End With
    End Sub

    Public Function GetBaseCritLocs() As Int32
        If mlBaseCritLocs = 0 Then
            Dim X As Int32

            For X = 0 To lSide1Crits.Length - 1
                mlBaseCritLocs = mlBaseCritLocs Or lSide1Crits(X)
            Next X
            For X = 0 To lSide2Crits.Length - 1
                mlBaseCritLocs = mlBaseCritLocs Or lSide2Crits(X)
            Next X
            For X = 0 To lSide3Crits.Length - 1
                mlBaseCritLocs = mlBaseCritLocs Or lSide3Crits(X)
            Next X
            For X = 0 To lSide4Crits.Length - 1
                mlBaseCritLocs = mlBaseCritLocs Or lSide4Crits(X)
            Next X
        End If
        Return mlBaseCritLocs
    End Function

    Private Shared Function GetHullSizeTargetMod(ByVal lHullSize As Int32) As Int32
        Dim lBaseToHit As Int32

        If lHullSize < 100 Then
            lBaseToHit = CInt(-40.0F * (1.0F - (lHullSize * 0.01F))) ' / 100.0F)))
        Else : lBaseToHit = lHullSize \ 1000
        End If

        'If lHullSize < gl_MAX_FIGHTER_HULL Then
        '    lBaseToHit += CInt(-40 * (1 - (lHullSize / 100)))
        'ElseIf lHullSize < gl_MIN_BATTLESHIP_HULL Then
        '    lBaseToHit += 5
        'Else : lBaseToHit += 15
        'End If

        Return lBaseToHit
    End Function

    Public Property BaseManeuver() As Byte
        Get
            Return myManeuver
        End Get
        Set(ByVal Value As Byte)
            myManeuver = Value
            BaseTurnAmount = myManeuver
            BaseAcceleration = CSng(myManeuver) * 0.01F ' / 100.0F
            BaseTurnAmountTimes100 = BaseTurnAmount * 100S
        End Set
    End Property

    Public Sub OptimizeMe()
        Dim X As Int32
        '      Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        '      If sFile.EndsWith("\") = False Then sFile &= "\"
        '      Dim oINI As InitFile = New InitFile(sFile & "RS_Model.dat")
        '      Dim sModelHdr As String = "Model_" & ModelID
        '      Dim lTemp As Int32

        '      lModelRangeOffset = CInt(Val(oINI.GetString(sModelHdr, "RangeOffset", "0")))
        '      lTemp = lModelRangeOffset * gl_FINAL_GRID_SQUARE_SIZE
        'lModelSizeXZ = CInt(Val(oINI.GetString(sModelHdr, "ModelSizeXZ", lTemp.ToString)))
        '      oINI = Nothing

        Dim oModelDef As ModelDef = ModelDefs.GetModelDef(Me.ModelID)
        If oModelDef Is Nothing = False Then
            lModelRangeOffset = oModelDef.lModelRangeOffset
            lModelSizeXZ = oModelDef.lModelSizeXZ
            Select Case oModelDef.lSpecialTraitID
                Case elModelSpecialTrait.Critical1
                    CriticalHitModifier = 1
                Case elModelSpecialTrait.Critical2
                    CriticalHitModifier = 2
                Case elModelSpecialTrait.Maneuver10Critical1
                    CriticalHitModifier = 1
                Case elModelSpecialTrait.Maneuver10Critical2
                    CriticalHitModifier = 2
                Case elModelSpecialTrait.Maneuver5Critical2
                    CriticalHitModifier = 2
                Case elModelSpecialTrait.Speed5Critical2
                    CriticalHitModifier = 2
            End Select

            yHullType = oModelDef.yHullType
        End If

        lOptRadarRange = CInt(yDefOptRadarRange) + lModelRangeOffset
        lMaxRadarRange = CInt(yDefMaxRadarRange) + lModelRangeOffset

        '(1000 / Sqrt(hull)) * 30
        If HullSize > 10 Then lFighterCritDelay = CInt((1000.0F / CSng(Math.Sqrt(HullSize))) * 30.0F) Else lFighterCritDelay = 1200

        FastPiercingResist = PiercingResist * 0.01F '/ 100.0F
        FastImpactResist = ImpactResist * 0.01F '/ 100.0F
        FastBeamResist = BeamResist * 0.01F '/ 100.0F
        FastECMResist = ECMResist * 0.01F '/ 100.0F
        FastFlameResist = FlameResist * 0.01F '/ 100.0F
        FastChemicalResist = ChemicalResist * 0.01F '/ 100.0F

        fOneOverOptRadarRange = 0.0F
        If lOptRadarRange > 0 Then fOneOverOptRadarRange = 1.0F / lOptRadarRange

        HullSizeTargetMod = GetHullSizeTargetMod(HullSize)

        For X = 0 To 3
            SideWeaponPower(X) = 0
        Next X
        Dim lAllArcPower As Int32 = 0
        'lWarpointValue = 0

        Dim lFastestWeaponROF As Int32 = Int32.MaxValue

        For X = 0 To WeaponDefUB
            WeaponDefs(X).SetAOEDmgType()

            If WeaponDefs(X).ArcID <> UnitArcs.eAllArcs Then
                SideWeaponPower(WeaponDefs(X).ArcID) += WeaponDefs(X).lFirePowerRating
            Else : lAllArcPower += WeaponDefs(X).lFirePowerRating
            End If

            'Now, adjust the values...
            With WeaponDefs(X)
                If .yWeaponType >= WeaponType.eMissile_Color_1 AndAlso .yWeaponType <= WeaponType.eMissile_Color_9 Then
                    .lRange = lModelRangeOffset + CInt(CInt(.iDefRange) * CInt(.WeaponSpeed))
                Else
                    .lRange = lModelRangeOffset + CInt(.iDefRange)
                End If

                Dim lTempValue As Int32 = lOptRadarRange - .lRange
                If lTempValue > 0 Then .fOneOverOptMinusRange = 1.0F / lTempValue Else .fOneOverOptMinusRange = 0.0F
                If .lRange > 0 Then .fOneOverRange = 1.0F / .lRange Else .fOneOverRange = 0.0F

                If .WeaponSpeed > 0 Then .fOneOverWeaponSpeed = 1.0F / .WeaponSpeed Else .fOneOverWeaponSpeed = 0.0F
 
                '.MaxRangeAccuracy = (.lRange + lModelRangeOffset) * (.Accuracy / 100.0F)
                .lPD_AccVal = (CInt(Me.ScanResolution) + CInt(.Accuracy)) \ 2
                .lNrm_AccVal = (CInt(Me.Weapon_Acc) + CInt(.Accuracy)) \ 2

                If .ROF_Delay < lFastestWeaponROF Then lFastestWeaponROF = .ROF_Delay

                'Stored ranges for weapon hit rolls.
                .PiercingMaxMinRange = .PiercingMaxDmg - .PiercingMinDmg
                .ImpactMaxMinRange = .ImpactMaxDmg - .ImpactMinDmg
                .BeamMaxMinRange = .BeamMaxDmg - .BeamMinDmg
                .ECMMaxMinRange = .ECMMaxDmg - .ECMMinDmg
                .FlameMaxMinRange = .FlameMaxDmg - .FlameMinDmg
                .ChemicalMaxMinRange = .ChemicalMaxDmg - .ChemicalMinDmg
            End With
        Next X

        If lFastestWeaponROF = Int32.MaxValue OrElse lFastestWeaponROF = 0 Then lMicroReliabilityTest = 2000 Else lMicroReliabilityTest = 60000 \ lFastestWeaponROF
        lMicroReliabilityTest = 3

        'Now, calculate the combat rating
        Dim lTotalArmor As Int32 = Armor_MaxHP(0) + Armor_MaxHP(1) + Armor_MaxHP(2) + Armor_MaxHP(3)
        Dim lResists As Int32 = CInt(PiercingResist) + CInt(ImpactResist) + CInt(BeamResist) + CInt(ECMResist) + CInt(FlameResist) + CInt(ChemicalResist)
        lResists = CInt(lResists / 6.0F)
        Dim fShieldRegenPerSec As Single
        If Me.ShieldRechargeFreq > 0 Then
            fShieldRegenPerSec = Me.ShieldRecharge * (30.0F / Me.ShieldRechargeFreq)
        Else : fShieldRegenPerSec = 0
        End If
        Dim lWpnStr As Int32 = SideWeaponPower(0) + SideWeaponPower(1) + SideWeaponPower(2) + SideWeaponPower(3) + (lAllArcPower * 2)
        'CombatRating = CInt(((((lTotalArmor / ((lResists + 1) / 100)) + (Shield_MaxHP * fShieldRegenPerSec)) + (NumberOfDoors * (MaxDoorSize \ 100))) + (BaseMaxSpeed * (2 * BaseManeuver))) + (lWpnStr * 5))
        'CombatRating = CInt(((((lTotalArmor / ((lResists + 1) / 100.0F)) + (Shield_MaxHP * fShieldRegenPerSec))) + (BaseMaxSpeed * (2 * BaseManeuver))) + (lWpnStr * 5))
        CombatRating = CInt(((lTotalArmor * ((100 - ArmorIntegrity) / 100)) / ((lResists + 1) / 100.0F) + Shield_MaxHP))
        If Shield_MaxHP > 0 Then CombatRating = CInt(Math.Pow(CombatRating, (1 + (fShieldRegenPerSec / Shield_MaxHP))))
        CombatRating += (BaseMaxSpeed * 2 * BaseManeuver) + lWpnStr

        If WeaponDefUB = -1 Then CombatRating = 1
        If CombatRating < 1 Then CombatRating = 1

        'lWarpointValue = CInt(CombatRating / 100.0F)

        lAbsoluteTotalHP = Math.Max(1, lTotalArmor + Structure_MaxHP)

        SetTacticalAnalysis()
    End Sub

    Private Sub SetTacticalAnalysis()
        'this sub sets the TacticalAnalysis variable for us
        Dim lThirdHull As Int32

        PlanetTacticalAnalysis = 0
        SpaceTacticalAnalysis = 0
        lThirdHull = HullSize \ 3

        'Regardless of space or planet, these apply:
        If HullSize <= gl_MAX_FIGHTER_HULL Then
            SpaceTacticalAnalysis = SpaceTacticalAnalysis Or eiTacticalAttrs.eFighterClass
            If (yChassisType And ChassisType.eAtmospheric) <> 0 Then
                PlanetTacticalAnalysis = PlanetTacticalAnalysis Or eiTacticalAttrs.eFighterClass
            End If
        End If
        If Cargo_Cap > lThirdHull Then
            SpaceTacticalAnalysis = SpaceTacticalAnalysis Or eiTacticalAttrs.eCargoShipClass
            PlanetTacticalAnalysis = PlanetTacticalAnalysis Or eiTacticalAttrs.eCargoShipClass
        End If
        If Hangar_Cap > lThirdHull Then
            SpaceTacticalAnalysis = SpaceTacticalAnalysis Or eiTacticalAttrs.eCarrierClass
            PlanetTacticalAnalysis = PlanetTacticalAnalysis Or eiTacticalAttrs.eCarrierClass
        End If
        If WeaponDefUB > -1 Then
            SpaceTacticalAnalysis = SpaceTacticalAnalysis Or eiTacticalAttrs.eArmedUnit
            PlanetTacticalAnalysis = PlanetTacticalAnalysis Or eiTacticalAttrs.eArmedUnit
        Else
            SpaceTacticalAnalysis = SpaceTacticalAnalysis Or eiTacticalAttrs.eUnarmedUnit
            PlanetTacticalAnalysis = PlanetTacticalAnalysis Or eiTacticalAttrs.eUnarmedUnit
        End If
        If Me.ObjTypeID = ObjectType.eFacilityDef Then
            SpaceTacticalAnalysis = SpaceTacticalAnalysis Or eiTacticalAttrs.eFacility
            PlanetTacticalAnalysis = PlanetTacticalAnalysis Or eiTacticalAttrs.eFacility
        End If
        If HullSize > gl_MAX_FIGHTER_HULL AndAlso HullSize <= 30000 Then
            SpaceTacticalAnalysis = SpaceTacticalAnalysis Or eiTacticalAttrs.eEscortClass
            PlanetTacticalAnalysis = PlanetTacticalAnalysis Or eiTacticalAttrs.eEscortClass
        End If

        'Now, based on space/planet we do these special...
        If HullSize > 30000 AndAlso HullSize < gl_MIN_CAPITAL_HULL_SIZE Then
            SpaceTacticalAnalysis = SpaceTacticalAnalysis Or eiTacticalAttrs.eTroopClass        'troop doubles as heavy escort
        End If

        'TODO: Troops go here

        If HullSize > gl_MIN_CAPITAL_HULL_SIZE Then
            SpaceTacticalAnalysis = SpaceTacticalAnalysis Or eiTacticalAttrs.eCapitalClass
        End If

        'And now, tanks
        If HullSize <= gl_MAX_TANK_HULL AndAlso (yChassisType And ChassisType.eGroundBased) <> 0 Then
            PlanetTacticalAnalysis = PlanetTacticalAnalysis Or eiTacticalAttrs.eCapitalClass
        End If

    End Sub

End Class
