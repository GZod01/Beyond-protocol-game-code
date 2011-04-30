Option Strict On

Public Class MissileMgr
    '  Public Structure Missile
    '      Public lID As Int32
    '      Public lTargetIdx As Int32
    'Public lAttackerIdx As Int32
    'Public lTargetID As Int32

    '      Public TotalVelocity As Single
    '      Public VelX As Single
    '      Public VelZ As Single
    '      Public MoveLocX As Single
    '      Public MoveLocZ As Single
    '      Public LocX As Int32
    '      Public LocZ As Int32
    '      Public LocAngle As Int16

    '      Public DestX As Int32
    '      Public DestZ As Int32

    '      Public Maneuver As Byte
    '      Public TurnAmount As Int16
    '      Public MaxSpeed As Byte
    '      Public Acceleration As Single
    '      Public TurnAmountTimes100 As Int32

    '      Public DetonateTime As Int32

    '      Public HomingAccuracy As Byte

    '      Public LastCycleMoved As Int32

    '      Public bHit As Boolean

    '      Public oParentEnvir As Envir
    '      Public oWpnDef As WeaponDef

    '      Public lGridIndex As Int32
    '      Public lIncomingIndex As Int32
    '  End Structure

    '  Private myMissileIdx() As Byte
    '  Private mlMissileUB As Int32 = -1
    '  Private muMissile() As Missile

    '  Public Function AddMissileShot(ByVal lTargetIdx As Int32, ByVal lAttackerIdx As Int32, ByRef oMissile As WeaponDef) As Int32
    '      Dim lIdx As Int32 = -1

    '      For X As Int32 = 0 To mlMissileUB
    '          If myMissileIdx(X) = 0 Then
    '              lIdx = X
    '              Exit For
    '          End If
    '      Next X
    '      If lIdx = -1 Then
    '          mlMissileUB += 1
    '          ReDim Preserve myMissileIdx(mlMissileUB)
    '          ReDim Preserve muMissile(mlMissileUB)
    '          lIdx = mlMissileUB
    '      End If

    '      With muMissile(lIdx)
    '          .Acceleration = oMissile.Maneuver / 10.0F
    '          .DestX = goEntity(lTargetIdx).LocX
    '          .DestZ = goEntity(lTargetIdx).LocZ
    '          .DetonateTime = glCurrentCycle + oMissile.iDefRange
    '          .HomingAccuracy = oMissile.Accuracy
    '          .lID = lIdx
    '          .LastCycleMoved = glCurrentCycle
    '          .LocX = goEntity(lAttackerIdx).LocX
    '          .LocZ = goEntity(lAttackerIdx).LocZ
    '          .LocAngle = CShort(LineAngleDegrees(.LocX, .LocZ, goEntity(lTargetIdx).LocX, goEntity(lTargetIdx).LocZ) * 10)
    '          .lAttackerIdx = lAttackerIdx
    '	.lTargetIdx = lTargetIdx
    '	.lTargetID = goEntity(lTargetIdx).ObjectID
    '          .Maneuver = oMissile.Maneuver
    '          .MaxSpeed = oMissile.WeaponSpeed
    '          .MoveLocX = .LocX
    '          .MoveLocZ = .LocZ
    '          .oParentEnvir = goEntity(lAttackerIdx).ParentEnvir
    '          .oWpnDef = oMissile
    '          .TotalVelocity = 0
    '          .TurnAmount = .Maneuver
    '          .TurnAmountTimes100 = .TurnAmount * 100S
    '          .VelX = 0
    '          .VelZ = 0
    '      End With

    '      With goEntity(lTargetIdx)
    '          If (.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
    '              muMissile(lIdx).bHit = True
    '          Else : muMissile(lIdx).bHit = (Rnd() * 100) + .Maneuver < muMissile(lIdx).HomingAccuracy + muMissile(lIdx).Maneuver
    '          End If
    '      End With

    '      myMissileIdx(lIdx) = 255

    '      Return lIdx
    '  End Function

    '  Public Sub HandleMissileMovement()
    '      For X As Int32 = 0 To mlMissileUB
    '          If myMissileIdx(X) <> 0 Then
    '              With muMissile(X)
    '                  If .DetonateTime < glCurrentCycle Then
    '                      myMissileIdx(X) = 0

    '                      If .oWpnDef.AOERange <> 0 Then
    '                          Dim oEvent As New WeaponEvent()
    '                          oEvent.AOERange = .oWpnDef.AOERange
    '					oEvent.yCritical = eyCriticalHitType.NoCriticalHit
    '					oEvent.lPierceDmg = (CInt(Rnd() * (.oWpnDef.PiercingMaxDmg - .oWpnDef.PiercingMinDmg + 1)) + .oWpnDef.PiercingMinDmg)
    '					oEvent.lImpactDmg = (CInt(Rnd() * (.oWpnDef.ImpactMaxDmg - .oWpnDef.ImpactMinDmg + 1)) + .oWpnDef.ImpactMinDmg)
    '					oEvent.lBeamDmg = (CInt(Rnd() * (.oWpnDef.BeamMaxDmg - .oWpnDef.BeamMinDmg + 1)) + .oWpnDef.BeamMinDmg)
    '					oEvent.lECMDmg = (CInt(Rnd() * (.oWpnDef.ECMMaxDmg - .oWpnDef.ECMMinDmg + 1)) + .oWpnDef.ECMMinDmg)
    '					oEvent.lFlameDmg = (CInt(Rnd() * (.oWpnDef.FlameMaxDmg - .oWpnDef.FlameMinDmg + 1)) + .oWpnDef.FlameMinDmg)
    '					oEvent.lChemicalDmg = (CInt(Rnd() * (.oWpnDef.ChemicalMaxDmg - .oWpnDef.ChemicalMinDmg + 1)) + .oWpnDef.ChemicalMinDmg)
    '					oEvent.lFromIdx = -1
    '					oEvent.lToIdx = -1
    '					oEvent.lEventCycle = glCurrentCycle
    '					oEvent.ySideHit = 255
    '					ApplyAOEDamage(oEvent, CInt(.LocX), CInt(.LocZ), .oParentEnvir)
    '				End If

    '				'Notify connected clients that the missile is detonated
    '				If .oParentEnvir.lPlayersInEnvirCnt <> 0 Then
    '					Dim yMsg(13) As Byte
    '					System.BitConverter.GetBytes(GlobalMessageCode.eMissileDetonated).CopyTo(yMsg, 0)
    '					System.BitConverter.GetBytes(.lID).CopyTo(yMsg, 2)
    '					System.BitConverter.GetBytes(CInt(.LocX)).CopyTo(yMsg, 6)
    '					System.BitConverter.GetBytes(CInt(.LocZ)).CopyTo(yMsg, 10)
    '					If .oParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
    '						'40000 comes from max entity clip plane on client,the 125 comes from
    '						'  10,000,000 / 40000 = 250 / 2 = 125 to offset to 0
    '						Dim lCameraPosX As Int32 = (.LocX \ 40000) + 125
    '						Dim lCameraPosZ As Int32 = (.LocZ \ 40000) + 125
    '						goMsgSys.BroadcastToEnvironmentClients_Filter(yMsg, .oParentEnvir, lCameraPosX, lCameraPosZ)
    '					Else : goMsgSys.BroadcastToEnvironmentClients(yMsg, .oParentEnvir)
    '					End If
    '				End If


    '				Continue For
    '			End If


    '			Dim oTarget As Epica_Entity = Nothing
    '			Dim vecTarget As Vector3

    '			If .lTargetIdx <> -1 AndAlso glEntityIdx(.lTargetIdx) = .lTargetID Then
    '				oTarget = goEntity(.lTargetIdx)

    '				'Check homing for target acquisition
    '				If oTarget Is Nothing = False Then
    '					Dim lTargetX As Int32 = oTarget.LocX
    '					If .HomingAccuracy > 0 Then

    '						Dim fDist As Single
    '						Dim fDX As Single = 0 '= .oTarget.lX - .LocX
    '						Dim fDZ As Single = 0 '= .oTarget.lZ - .LocZ

    '						If oTarget Is Nothing = False Then

    '							'Calculate for MapWrap here.
    '							If .oParentEnvir.ObjTypeID = ObjectType.ePlanet Then
    '								If Math.Abs(lTargetX - .LocX) > (.oParentEnvir.lMapWrapAdjustX - Math.Abs(.LocX)) - Math.Abs(lTargetX) Then
    '									'ltargetx = 
    '									If .LocX < lTargetX Then
    '										lTargetX -= .oParentEnvir.lMapWrapAdjustX
    '									Else : lTargetX += .oParentEnvir.lMapWrapAdjustX
    '									End If
    '								End If
    '							End If

    '							fDX = goEntity(.lTargetIdx).LocX - .LocX
    '							fDZ = goEntity(.lTargetIdx).LocZ - .LocZ
    '							fDX *= fDX
    '							fDZ *= fDZ
    '							fDist = CSng(Math.Sqrt(fDX + fDZ))
    '							fDist /= (.TotalVelocity + 0.01F)
    '							.DestX = CInt(oTarget.LocX + (oTarget.VelX * fDist))
    '							.DestZ = CInt(oTarget.LocZ + (oTarget.VelZ * fDist))

    '							'Ok, see if we are moving the wrong way
    '							Dim bX_MD_LTZ As Boolean = (.LocX - .DestX) < 0
    '							Dim bY_MD_LTZ As Boolean = (.LocZ - .DestZ) < 0
    '							Dim bX_MT_LTZ As Boolean = (.LocX - oTarget.LocX) < 0
    '							Dim bY_MT_LTZ As Boolean = (.LocZ - oTarget.LocZ) < 0
    '							If bX_MD_LTZ <> bX_MT_LTZ AndAlso bY_MD_LTZ <> bY_MT_LTZ Then
    '								.DestX = oTarget.LocX
    '								.DestZ = oTarget.LocZ
    '							ElseIf Math.Abs(.LocX - .DestX) < Math.Abs(.LocX - oTarget.LocX) AndAlso Math.Abs(.LocZ - .DestZ) < Math.Abs(.LocZ - oTarget.LocZ) Then
    '								.DestX = oTarget.LocX
    '								.DestZ = oTarget.LocZ
    '							End If
    '						Else
    '							.DestX = .LocX
    '							.DestZ = .LocZ
    '						End If
    '					End If

    '					vecTarget.X = lTargetX
    '					vecTarget.Y = oTarget.LocY
    '					vecTarget.Z = oTarget.LocZ
    '				Else
    '					.DetonateTime = glCurrentCycle
    '					Continue For
    '				End If
    '			Else
    '				.DetonateTime = glCurrentCycle
    '				Continue For
    '			End If

    '			.TotalVelocity += .Acceleration
    '			If .TotalVelocity > .MaxSpeed Then .TotalVelocity = .MaxSpeed

    '			Dim vecMissile As Vector3 = New Vector3(.LocX, 0, .LocZ)
    '			Dim vecAcc As Vector3 = Vector3.Subtract(vecTarget, vecMissile)

    '			vecAcc.Y = 0
    '			vecAcc.Normalize()
    '			vecAcc.Multiply(.Acceleration)

    '			.VelX += vecAcc.X
    '			.VelZ += vecAcc.Z

    '			Dim fTotalSpeed As Single = Math.Abs(.VelX) + Math.Abs(.VelZ)
    '			If fTotalSpeed > .MaxSpeed Then
    '				Dim fTmp As Single = .MaxSpeed / fTotalSpeed
    '				.VelX *= fTmp
    '				.VelZ *= fTmp
    '			End If

    '			If .bHit = True OrElse oTarget Is Nothing Then
    '				If Math.Abs(.VelX) > Math.Abs(vecMissile.X - vecTarget.X) Then
    '					.VelX = 0
    '					vecMissile.X = vecTarget.X
    '				End If
    '				If Math.Abs(.VelZ) > Math.Abs(vecMissile.Z - vecTarget.Z) Then
    '					.VelZ = 0
    '					vecMissile.Z = vecTarget.Z
    '				End If
    '			End If

    '			'ok, move it or lose it
    '			If .VelX <> 0 OrElse .VelZ <> 0 Then
    '				.MoveLocX += .VelX
    '				.MoveLocZ += .VelZ
    '				.LocX = CInt(.MoveLocX)
    '				.LocZ = CInt(.MoveLocZ)
    '			End If

    '			'Ok, check for collision/explosion
    '			If oTarget Is Nothing = False Then
    '				If Math.Abs(oTarget.MoveLocX - .MoveLocX) + Math.Abs(oTarget.MoveLocZ - .MoveLocZ) < (.TotalVelocity / 2.0F) Then
    '					If .bHit = False Then
    '						'just off shoot it...
    '						.MoveLocX += 10

    '						If .HomingAccuracy <> 0 Then

    '							If (oTarget.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
    '								.bHit = True
    '							Else : .bHit = (Rnd() * 100) + oTarget.Maneuver < .HomingAccuracy + .Maneuver
    '							End If
    '						End If
    '					Else
    '						myMissileIdx(X) = 0

    '						'Ok, we hit...
    '						Dim oEvent As New WeaponEvent()
    '						oEvent.AOERange = .oWpnDef.AOERange
    '						oEvent.yCritical = eyCriticalHitType.NoCriticalHit
    '						oEvent.lPierceDmg = (CInt(Rnd() * (.oWpnDef.PiercingMaxDmg - .oWpnDef.PiercingMinDmg + 1)) + .oWpnDef.PiercingMinDmg)
    '						oEvent.lImpactDmg = (CInt(Rnd() * (.oWpnDef.ImpactMaxDmg - .oWpnDef.ImpactMinDmg + 1)) + .oWpnDef.ImpactMinDmg)
    '						oEvent.lBeamDmg = (CInt(Rnd() * (.oWpnDef.BeamMaxDmg - .oWpnDef.BeamMinDmg + 1)) + .oWpnDef.BeamMinDmg)
    '						oEvent.lECMDmg = (CInt(Rnd() * (.oWpnDef.ECMMaxDmg - .oWpnDef.ECMMinDmg + 1)) + .oWpnDef.ECMMinDmg)
    '						oEvent.lFlameDmg = (CInt(Rnd() * (.oWpnDef.FlameMaxDmg - .oWpnDef.FlameMinDmg + 1)) + .oWpnDef.FlameMinDmg)
    '						oEvent.lChemicalDmg = (CInt(Rnd() * (.oWpnDef.ChemicalMaxDmg - .oWpnDef.ChemicalMinDmg + 1)) + .oWpnDef.ChemicalMinDmg)
    '						oEvent.lFromIdx = .lAttackerIdx
    '						oEvent.lToIdx = .lTargetIdx
    '						oEvent.lEventCycle = glCurrentCycle
    '						If glEntityIdx(.lAttackerIdx) <> -1 Then
    '							Dim oAtkrEntity As Epica_Entity = goEntity(.lAttackerIdx)
    '							If oAtkrEntity Is Nothing = False Then
    '								oEvent.ySideHit = WhatSideDoIHit(oAtkrEntity, oTarget)
    '							Else
    '								oEvent.ySideHit = 255
    '							End If
    '						Else : oEvent.ySideHit = 255
    '						End If

    '						'Add the event
    '						Dim yTemp As Byte = 0
    '						For lTmp As Int32 = 0 To oTarget.lWpnEventUB
    '							If oTarget.yWpnEventUsed(lTmp) = 0 Then
    '								'add it
    '								yTemp = 1
    '								oTarget.oWpnEvents(lTmp) = oEvent
    '								oTarget.yWpnEventUsed(lTmp) = 255
    '								oTarget.lNextWpnEvent = Math.Min(oTarget.lNextWpnEvent, oEvent.lEventCycle)
    '								Exit For
    '							End If
    '						Next lTmp

    '						If yTemp = 0 Then
    '							oTarget.lWpnEventUB += 1
    '							ReDim Preserve oTarget.oWpnEvents(oTarget.lWpnEventUB)
    '							ReDim Preserve oTarget.yWpnEventUsed(oTarget.lWpnEventUB)
    '							oTarget.oWpnEvents(oTarget.lWpnEventUB) = oEvent
    '							oTarget.lNextWpnEvent = Math.Min(oTarget.lNextWpnEvent, oEvent.lEventCycle)
    '							oTarget.yWpnEventUsed(oTarget.lWpnEventUB) = 255
    '						End If

    '						If .oParentEnvir.lPlayersInEnvirCnt <> 0 Then
    '							Dim yMsg(6) As Byte
    '							System.BitConverter.GetBytes(GlobalMessageCode.eMissileImpact).CopyTo(yMsg, 0)
    '							System.BitConverter.GetBytes(.lID).CopyTo(yMsg, 2)
    '							If oTarget.ShieldHP > 0 AndAlso (oTarget.CurrentStatus And elUnitStatus.eShieldOperational) <> 0 Then
    '								yMsg(6) = 1
    '							Else : yMsg(6) = 0
    '							End If
    '							If .oParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
    '								'40000 comes from max entity clip plane on client,the 125 comes from
    '								'  10,000,000 / 40000 = 250 / 2 = 125 to offset to 0
    '								Dim lCameraPosX As Int32 = (.LocX \ 40000) + 125
    '								Dim lCameraPosZ As Int32 = (.LocZ \ 40000) + 125
    '								goMsgSys.BroadcastToEnvironmentClients_Filter(yMsg, .oParentEnvir, lCameraPosX, lCameraPosZ)
    '							Else : goMsgSys.BroadcastToEnvironmentClients(yMsg, .oParentEnvir)
    '							End If
    '						End If


    '						Exit For
    '					End If
    '				End If
    '			Else
    '				If Math.Abs(.DestX - .MoveLocX) + Math.Abs(.DestZ - .MoveLocZ) < (.TotalVelocity / 2.0F) Then
    '					'Here, we cheat, just set our detonate time to now
    '					.DetonateTime = glCurrentCycle
    '					Exit For
    '				End If
    '			End If

    '              End With
    '          End If
    '      Next X
    '  End Sub

    Public Enum eyDetonateType As Byte
        ''' <summary>
        ''' No AOE damage, no damage calculated... just peters out
        ''' </summary>
        ''' <remarks></remarks>
        BreakApart = 0

        ''' <summary>
        ''' Does not hit, but does AOE dmg, we send down a resulting locX and locZ so the client can try to render it
        ''' </summary>
        ''' <remarks></remarks>
        MissedShot = 1

        ''' <summary>
        ''' Impact with target. Does AOE dmg.
        ''' </summary>
        ''' <remarks></remarks>
        Impact = 2
    End Enum

    Public Class Missile
        Public lMissileID As Int32
        Public lMissileIdx As Int32 = -1

        Public lTargetIdx As Int32 = -1                 'index of the target in the glEntityIdx() array
        Public lTargetID As Int32 = -1                  'ID of the target for comparison purposes
        Public lAttackerIdx As Int32 = -1               'index of the attacker in the glEntityIdx() array
        Public lAttackerID As Int32 = -1                'ID of the attacker for comparison purposes
        Public lAttackerOwnerID As Int32 = -1           'ID of the player who owns the attacking entity
        Public yAttackingHullType As Byte = 0

        Public oParentEnvir As Envir
        Public lEnvirGridItemIndex As Int32 = -1        'index in the envirgrid's missile array that represents me
        Public lIncomingIndex As Int32 = -1             'index in the target's incoming missile array that represents me

        'For Movement engine
        Public TotalVelocity As Single
        Public VelX As Single
        Public VelZ As Single
        Public MoveLocX As Single
        Public MoveLocZ As Single
        Public LocX As Int32
        Public LocZ As Int32
        Public LocAngle As Int16
        Public DestX As Int32
        Public DestZ As Int32
        Public LastCycleMoved As Int32
        '--- end of movement engine ---

        'For Aggression engine
        Public lGridIndex As Int32         'the grid index of the environment object
        Public lSmallSectorID As Int32
        Public lTinyX As Int32
        Public lTinyZ As Int32
        '--- end of aggression engine

        'Items taken from the wpn def
        Public MaxSpeed As Int32
        Public Maneuver As Int32
        Public TurnAmount As Int16
        Public Acceleration As Single
        Public TurnAmountTimes100 As Int32
        Public HomingAccuracy As Byte
        Public lHitpoints As Int32
        'for damage purposes
        Public oWpnDef As WeaponDef
        '--- end of weapon def capture ---
        Public lCycleMissileFired As Int32

        Public bHit As Boolean = False      'indicates whether this missile will hit or not when it reaches the target

        Public Sub ProcessMissileHit(ByRef oTarget As Epica_Entity)
            'Ok, we hit...
            Dim oEvent As New WeaponEvent()
            With oWpnDef
                oEvent.AOERange = .AOERange
                oEvent.yAOEDmgType = .yAOEDmgType
                oEvent.yCritical = eyCriticalHitType.NoCriticalHit
                If .PiercingMaxDmg > 0 Then oEvent.lPierceDmg = (CInt(Rnd() * .PiercingMaxMinRange) + .PiercingMinDmg)
                If .ImpactMaxDmg > 0 Then oEvent.lImpactDmg = (CInt(Rnd() * .ImpactMaxMinRange) + .ImpactMinDmg)
                If .BeamMaxDmg > 0 Then oEvent.lBeamDmg = (CInt(Rnd() * .BeamMaxMinRange) + .BeamMinDmg)
                If .ECMMaxDmg > 0 Then oEvent.lECMDmg = (CInt(Rnd() * .ECMMaxMinRange) + .ECMMinDmg)
                If .FlameMaxDmg > 0 Then oEvent.lFlameDmg = (CInt(Rnd() * .FlameMaxMinRange) + .FlameMinDmg)
                If .ChemicalMaxDmg > 0 Then oEvent.lChemicalDmg = (CInt(Rnd() * .ChemicalMaxMinRange) + .ChemicalMinDmg)
                oEvent.FromPlayerID = Me.lAttackerOwnerID
            End With

            oEvent.lFromIdx = lAttackerIdx
            oEvent.lToIdx = lTargetIdx
            oEvent.lEventCycle = glCurrentCycle
            oEvent.yAttackerType = yAttackingHullType
            If glEntityIdx(lAttackerIdx) > 0 Then
                Dim oAtkrEntity As Epica_Entity = goEntity(lAttackerIdx)
                If oAtkrEntity Is Nothing = False Then
                    oEvent.ySideHit = WhatSideDoIHit(oAtkrEntity, oTarget)
                Else
                    oEvent.ySideHit = 0 'was 255
                End If
            Else : oEvent.ySideHit = 0  'was 255
            End If

            'Add the event
            Dim yTemp As Byte = 0
            For lTmp As Int32 = 0 To oTarget.lWpnEventUB
                If oTarget.yWpnEventUsed(lTmp) = 0 Then
                    'add it
                    yTemp = 1
                    oTarget.oWpnEvents(lTmp) = oEvent
                    oTarget.yWpnEventUsed(lTmp) = 255
                    oTarget.lNextWpnEvent = Math.Min(oTarget.lNextWpnEvent, oEvent.lEventCycle)
                    Exit For
                End If
            Next lTmp

            If yTemp = 0 Then
                oTarget.lWpnEventUB += 1
                ReDim Preserve oTarget.oWpnEvents(oTarget.lWpnEventUB)
                ReDim Preserve oTarget.yWpnEventUsed(oTarget.lWpnEventUB)
                oTarget.oWpnEvents(oTarget.lWpnEventUB) = oEvent
                oTarget.lNextWpnEvent = Math.Min(oTarget.lNextWpnEvent, oEvent.lEventCycle)
                oTarget.yWpnEventUsed(oTarget.lWpnEventUB) = 255
            End If

        End Sub
    End Class

    Public mlMissileIdx(-1) As Int32
    Public mlMissileUB As Int32 = -1
    Public moMissile() As Missile

    Private mlLastMissileID As Int32 = 1
    Private Function GetMissileID() As Int32
        SyncLock mlMissileIdx
            mlLastMissileID += 1
            If mlLastMissileID > 2000000000 Then
                mlLastMissileID = 1
            End If
            Return mlLastMissileID
        End SyncLock
    End Function

    Public Function AddMissileShot(ByVal oTarget As Epica_Entity, ByVal oAttacker As Epica_Entity, ByVal oAttackerDef As Epica_Entity_Def, ByRef oMissile As WeaponDef, ByVal oTargetDef As Epica_Entity_Def) As Int32
        Dim lIdx As Int32 = -1

        Dim lMissileID As Int32 = GetMissileID()

        'NOTE: if we ever put this into a multi-thread scenario, we will need to use an Add Queue similar to the primary's queue manager

        Dim oNewMissile As New Missile()
        With oNewMissile
            .Acceleration = oMissile.Maneuver / 10.0F
            .DestX = oTarget.LocX
            .DestZ = oTarget.LocZ
            .HomingAccuracy = oMissile.Accuracy
            .lMissileID = lMissileID
            .LastCycleMoved = glCurrentCycle
            .LocX = oAttacker.LocX
            .LocZ = oAttacker.LocZ
            .LocAngle = CShort(LineAngleDegrees(.LocX, .LocZ, oTarget.LocX, oTarget.LocZ) * 10)
            .lAttackerIdx = oAttacker.ServerIndex
            .lAttackerID = oAttacker.ObjectID
            .lAttackerOwnerID = oAttacker.lOwnerID
            .lHitpoints = oMissile.lStructHP
            .lTargetIdx = oTarget.ServerIndex
            .lTargetID = oTarget.ObjectID
            .Maneuver = oMissile.Maneuver
            .MaxSpeed = oMissile.WeaponSpeed
            .MoveLocX = .LocX
            .MoveLocZ = .LocZ
            .oParentEnvir = oAttacker.ParentEnvir
            .oWpnDef = oMissile
            .TotalVelocity = 0
            .TurnAmount = CShort(.Maneuver)
            .TurnAmountTimes100 = CInt(.TurnAmount) * 100
            .VelX = 0
            .VelZ = 0
            .yAttackingHullType = oAttackerDef.yHullType
            .lCycleMissileFired = glCurrentCycle

            'If (oTarget.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
            '    .bHit = True
            'Else
            If oMissile.blOverallMaxDmg > oTargetDef.HullSize * 3 Then
                .bHit = False
            Else
                Dim lTargetScore As Int32 = CInt(oTargetDef.ScanResolution) + CInt(oTarget.Maneuver) + CInt(oTarget.Exp_Level)
                Dim lMissileScore As Int32 = CInt(.HomingAccuracy) + CInt(.Maneuver) + CInt(Rnd() * 255)
                .bHit = lTargetScore < lMissileScore '(Rnd() * 100) + oTarget.Maneuver < .HomingAccuracy + .Maneuver
            End If
            'End If

            For X As Int32 = 0 To mlMissileUB
                If mlMissileIdx(X) < 1 Then
                    lIdx = X
                    Exit For
                End If
            Next X
            If lIdx = -1 Then
                mlMissileUB += 1
                ReDim Preserve mlMissileIdx(mlMissileUB)
                ReDim Preserve moMissile(mlMissileUB)
                lIdx = mlMissileUB
            End If

            .lMissileIdx = lIdx

            .lGridIndex = oAttacker.lGridIndex
            .lSmallSectorID = oAttacker.lSmallSectorID

            If .lSmallSectorID < 0 Then .lSmallSectorID = 0
            If .lSmallSectorID > gl_SMALL_GRID_SQUARE_SIZE - 1 Then .lSmallSectorID = gl_SMALL_GRID_SQUARE_SIZE - 1

            .lTinyX = oAttacker.lTinyX
            .lTinyX = oAttacker.lTinyZ

            Dim oGrid As EnvirGrid = .oParentEnvir.oGrid(.lGridIndex)
            If oGrid Is Nothing = False Then
                .lEnvirGridItemIndex = oGrid.AddMissile(lIdx)
            End If
        End With
        moMissile(lIdx) = oNewMissile
        oNewMissile.lIncomingIndex = oTarget.AddIncomingMissile(lIdx, oNewMissile.lMissileID)
        mlMissileIdx(lIdx) = oNewMissile.lMissileID

        'Ok, a weapon was fired, so set up our message to send...
        With oNewMissile
            If .oParentEnvir.lPlayersInEnvirCnt > 0 Then
                Dim yFireMsg(15) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eMissileFired).CopyTo(yFireMsg, 0)
                If .bHit = True Then
                    System.BitConverter.GetBytes(.lMissileID).CopyTo(yFireMsg, 2)
                Else
                    System.BitConverter.GetBytes(-.lMissileID).CopyTo(yFireMsg, 2)
                End If
                If oAttacker.ObjTypeID = ObjectType.eFacility Then
                    System.BitConverter.GetBytes(-oAttacker.ObjectID).CopyTo(yFireMsg, 6)
                Else
                    System.BitConverter.GetBytes(oAttacker.ObjectID).CopyTo(yFireMsg, 6)
                End If
                If oTarget.ObjTypeID = ObjectType.eFacility Then
                    System.BitConverter.GetBytes(-oTarget.ObjectID).CopyTo(yFireMsg, 10)
                Else
                    System.BitConverter.GetBytes(oTarget.ObjectID).CopyTo(yFireMsg, 10)
                End If
                yFireMsg(14) = .oWpnDef.yWeaponType
                yFireMsg(15) = .oWpnDef.WeaponSpeed

                If .oParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                    '40000 comes from max entity clip plane on client,the 125 comes from
                    '  10,000,000 / 40000 = 250 / 2 = 125 to offset to 0
                    Dim lCameraPosX As Int32 = (.LocX \ 40000) + 125
                    Dim lCameraPosZ As Int32 = (.LocZ \ 40000) + 125
                    goMsgSys.BroadcastToEnvironmentClients_Filter(yFireMsg, .oParentEnvir, lCameraPosX, lCameraPosZ)
                Else : goMsgSys.BroadcastToEnvironmentClients(yFireMsg, .oParentEnvir)
                End If
            End If
        End With


        Return lIdx
    End Function

    Public Sub HandleMissileMovement()
        For X As Int32 = 0 To mlMissileUB
            If mlMissileIdx(X) > 0 Then

                Dim oMissile As Missile = moMissile(X)
                If oMissile Is Nothing Then
                    mlMissileIdx(X) = -1
                    Continue For
                End If

                With oMissile

                    Dim oTarget As Epica_Entity = Nothing
                    Dim vecTarget As Vector3

                    'Do we have a target?
                    If .lTargetIdx <> -1 AndAlso glEntityIdx(.lTargetIdx) = .lTargetID Then
                        oTarget = goEntity(.lTargetIdx)

                        'check for our target... we are moving towards it
                        If oTarget Is Nothing = False Then
                            Dim lTargetX As Int32 = oTarget.LocX

                            'Calculate for MapWrap here.
                            If .oParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                If Math.Abs(lTargetX - .LocX) > (.oParentEnvir.lMapWrapAdjustX - Math.Abs(.LocX)) - Math.Abs(lTargetX) Then
                                    'ltargetx = 
                                    If .LocX < lTargetX Then
                                        lTargetX -= .oParentEnvir.lMapWrapAdjustX
                                    Else : lTargetX += .oParentEnvir.lMapWrapAdjustX
                                    End If
                                End If
                            End If

                            vecTarget.X = lTargetX
                            vecTarget.Y = oTarget.LocY
                            vecTarget.Z = oTarget.LocZ
                        Else
                            'TODO: could reacquire targets here...
                            DetonateMissile(oMissile, X, eyDetonateType.BreakApart)
                            Continue For
                        End If
                    Else
                        'TODO: could reacquire targets here...
                        DetonateMissile(oMissile, X, eyDetonateType.BreakApart)
                        Continue For
                    End If

                    If glCurrentCycle - .lCycleMissileFired > 900 Then      '30 seconds
                        DetonateMissile(oMissile, X, eyDetonateType.BreakApart)
                        Continue For
                    End If

                    'Now, move our missile....
                    .TotalVelocity += .Acceleration
                    If .TotalVelocity > .MaxSpeed Then .TotalVelocity = .MaxSpeed

                    Dim vecMissile As Vector3 = New Vector3(.LocX, 0, .LocZ)
                    Dim vecAcc As Vector3 = Vector3.Subtract(vecTarget, vecMissile)
                    vecAcc.Y = 0
                    vecAcc.Normalize()
                    vecAcc.Multiply(.Acceleration)

                    .VelX += vecAcc.X
                    .VelZ += vecAcc.Z

                    Dim fTotalSpeed As Single = Math.Abs(.VelX) + Math.Abs(.VelZ)
                    If fTotalSpeed > .MaxSpeed Then
                        Dim fTmp As Single = .MaxSpeed / fTotalSpeed
                        .VelX *= fTmp
                        .VelZ *= fTmp
                    End If

                    If .bHit = True OrElse oTarget Is Nothing Then
                        If Math.Abs(.VelX) > Math.Abs(vecMissile.X - vecTarget.X) Then
                            .VelX = 0
                            vecMissile.X = vecTarget.X
                        End If
                        If Math.Abs(.VelZ) > Math.Abs(vecMissile.Z - vecTarget.Z) Then
                            .VelZ = 0
                            vecMissile.Z = vecTarget.Z
                        End If
                    End If

                    'ok, move it or lose it
                    If .VelX <> 0 OrElse .VelZ <> 0 Then
                        .MoveLocX += .VelX
                        .MoveLocZ += .VelZ
                        .LocX = CInt(.MoveLocX)
                        .LocZ = CInt(.MoveLocZ)

                        'set the missiles grid index, small sectorid, tinyx, tinyz
                        Dim lPrevGridIndex As Int32 = .lGridIndex

                        Dim lTemp As Int32 = .LocZ + .oParentEnvir.lHalfEnvirSize
                        .lGridIndex = (lTemp \ .oParentEnvir.lGridSquareSize) * .oParentEnvir.lGridsPerRow
                        lTemp -= ((.lGridIndex \ .oParentEnvir.lGridsPerRow) * .oParentEnvir.lGridSquareSize)
                        .lSmallSectorID = ((lTemp \ gl_SMALL_GRID_SQUARE_SIZE) * gl_SMALL_PER_ROW)
                        lTemp -= ((.lSmallSectorID \ gl_SMALL_PER_ROW) * gl_SMALL_GRID_SQUARE_SIZE)
                        .lTinyZ = lTemp \ gl_FINAL_GRID_SQUARE_SIZE
                        Dim lTemp2 As Int32 = .LocX + .oParentEnvir.lHalfEnvirSize
                        lTemp = lTemp2 \ .oParentEnvir.lGridSquareSize

                        .lGridIndex += lTemp
                        lTemp2 -= (lTemp * .oParentEnvir.lGridSquareSize)
                        lTemp = lTemp2 \ gl_SMALL_GRID_SQUARE_SIZE
                        .lSmallSectorID += lTemp
                        lTemp2 -= (lTemp * gl_SMALL_GRID_SQUARE_SIZE)
                        .lTinyX = lTemp2 \ gl_FINAL_GRID_SQUARE_SIZE

                        If .lSmallSectorID < 0 Then .lSmallSectorID = 0
                        If .lSmallSectorID > gl_SMALL_GRID_SQUARE_SIZE - 1 Then .lSmallSectorID = gl_SMALL_GRID_SQUARE_SIZE - 1

                        If .lGridIndex <> lPrevGridIndex Then
                            .oParentEnvir.oGrid(lPrevGridIndex).RemoveMissile(.lEnvirGridItemIndex)
                            If .lGridIndex < 0 OrElse .lGridIndex > .oParentEnvir.GetGridUB Then
                                .lGridIndex = 0
                                DetonateMissile(oMissile, X, eyDetonateType.MissedShot)
                            End If
                            .lEnvirGridItemIndex = .oParentEnvir.oGrid(.lGridIndex).AddMissile(.lMissileIdx)
                        End If
                    End If

                    'Ok, check for collision/explosion.... 
                    If Math.Abs(oTarget.MoveLocX - .MoveLocX) + Math.Abs(oTarget.MoveLocZ - .MoveLocZ) < (.TotalVelocity * 0.5F) Then
                        If .bHit = True Then 'OrElse (.HomingAccuracy > 0 AndAlso (oTarget.CurrentStatus And elUnitStatus.eUnitMoving) = 0) Then
                            'hit
                            '.bHit = True
                            .ProcessMissileHit(oTarget)
                            DetonateMissile(oMissile, X, eyDetonateType.Impact)
                        Else
                            'just off shoot it...
                            Dim fRng As Single = ((256.0F - CSng(.HomingAccuracy)) + oTarget.lModelRangeOffset) * gl_FINAL_GRID_SQUARE_SIZE
                            If .VelX < 0 Then
                                .MoveLocX -= Rnd() * fRng
                            Else
                                .MoveLocX += Rnd() * fRng
                            End If
                            If .VelZ < 0 Then
                                .MoveLocZ -= Rnd() * fRng
                            Else
                                .MoveLocZ += Rnd() * fRng
                            End If
                            .LocX = CInt(.MoveLocX)
                            .LocZ = CInt(.MoveLocZ)
                            DetonateMissile(oMissile, X, eyDetonateType.MissedShot)
                        End If
                    End If

                End With
            End If
        Next X
    End Sub

    Public Sub DetonateMissile(ByRef oMissile As Missile, ByVal lMissileIndex As Int32, ByVal yDetonateType As eyDetonateType) '  ByVal bKilled As Boolean)
        'remove me from the target's Incoming Missile List
        Dim oTarget As Epica_Entity = Nothing
        If oMissile.lTargetIdx <> -1 AndAlso glEntityIdx(oMissile.lTargetIdx) = oMissile.lTargetID Then
            oTarget = goEntity(oMissile.lTargetIdx)
        End If
        If oTarget Is Nothing = False Then
            oTarget.RemoveIncomingMissile(oMissile.lIncomingIndex, oMissile.lMissileID)
        End If

        'remove me from the missile manager's myMissileIdx, etc...
        mlMissileIdx(lMissileIndex) = -1
        moMissile(lMissileIndex) = Nothing

        'remove me from the Envir Grid Missile List
        oMissile.oParentEnvir.oGrid(oMissile.lGridIndex).RemoveMissile(oMissile.lEnvirGridItemIndex)

        'Do an AOE event from the loc of the missile
        If yDetonateType <> eyDetonateType.BreakApart AndAlso oMissile.oWpnDef.AOERange > 0 AndAlso oMissile.oWpnDef.yAOEDmgType <> 0 Then
            Dim oEvent As WeaponEvent = New WeaponEvent()
            'Now, our damages
            oEvent.lToIdx = oTarget.ServerIndex
            Dim oWeapon As WeaponDef = oMissile.oWpnDef

            With oEvent
                .lBeamDmg = 0 : .lChemicalDmg = 0 : .lECMDmg = 0 : .lFlameDmg = 0 : .lImpactDmg = 0 : .lPierceDmg = 0
            End With

            Select Case oWeapon.yAOEDmgType
                Case eyAOEDmgType.eBeamDmg
                    If oWeapon.BeamMaxDmg > 0 Then oEvent.lBeamDmg = CInt(Rnd() * oWeapon.BeamMaxMinRange) + oWeapon.BeamMinDmg
                Case eyAOEDmgType.eChemicalDmg
                    If oWeapon.ChemicalMaxDmg > 0 Then oEvent.lChemicalDmg = CInt(Rnd() * oWeapon.ChemicalMaxMinRange) + oWeapon.ChemicalMinDmg
                Case eyAOEDmgType.eECMDmg
                    If oWeapon.ECMMaxDmg > 0 Then oEvent.lECMDmg = CInt(Rnd() * oWeapon.ECMMaxMinRange) + oWeapon.ECMMinDmg
                Case eyAOEDmgType.eFlameDmg
                    If oWeapon.FlameMaxDmg > 0 Then oEvent.lFlameDmg = CInt(Rnd() * oWeapon.FlameMaxMinRange) + oWeapon.FlameMinDmg
                Case eyAOEDmgType.eImpactDmg
                    If oWeapon.ImpactMaxDmg > 0 Then oEvent.lImpactDmg = CInt(Rnd() * oWeapon.ImpactMaxMinRange) + oWeapon.ImpactMinDmg
                Case eyAOEDmgType.ePierceDmg
                    If oWeapon.PiercingMaxDmg > 0 Then oEvent.lPierceDmg = CInt(Rnd() * oWeapon.PiercingMaxMinRange) + oWeapon.PiercingMinDmg
                Case Else
                    If oWeapon.PiercingMaxDmg > 0 Then oEvent.lPierceDmg = CInt(Rnd() * oWeapon.PiercingMaxMinRange) + oWeapon.PiercingMinDmg
                    If oWeapon.ImpactMaxDmg > 0 Then oEvent.lImpactDmg = CInt(Rnd() * oWeapon.ImpactMaxMinRange) + oWeapon.ImpactMinDmg
                    If oWeapon.BeamMaxDmg > 0 Then oEvent.lBeamDmg = CInt(Rnd() * oWeapon.BeamMaxMinRange) + oWeapon.BeamMinDmg
                    If oWeapon.ECMMaxDmg > 0 Then oEvent.lECMDmg = CInt(Rnd() * oWeapon.ECMMaxMinRange) + oWeapon.ECMMinDmg
                    If oWeapon.FlameMaxDmg > 0 Then oEvent.lFlameDmg = CInt(Rnd() * oWeapon.FlameMaxMinRange) + oWeapon.FlameMinDmg
                    If oWeapon.ChemicalMaxDmg > 0 Then oEvent.lChemicalDmg = CInt(Rnd() * oWeapon.ChemicalMaxMinRange) + oWeapon.ChemicalMinDmg
            End Select
            oEvent.FromPlayerID = oMissile.lAttackerOwnerID
            ApplyAOEDamage(oEvent, oMissile.LocX, oMissile.LocZ, oMissile.oParentEnvir, oMissile.lAttackerOwnerID)
        End If

        'Send msg to clients that missile is destroyed
        If oMissile.oParentEnvir.lPlayersInEnvirCnt > 0 Then
            Dim yMsg() As Byte

            If yDetonateType = eyDetonateType.MissedShot Then
                'missed shot... we send
                ReDim yMsg(14)
                System.BitConverter.GetBytes(GlobalMessageCode.eMissileDetonated).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(oMissile.lMissileID).CopyTo(yMsg, 2)
                yMsg(6) = oMissile.oWpnDef.AOERange
                System.BitConverter.GetBytes(oMissile.LocX).CopyTo(yMsg, 7)
                System.BitConverter.GetBytes(oMissile.LocZ).CopyTo(yMsg, 11)
            ElseIf yDetonateType = eyDetonateType.Impact Then
                'impact...  we send
                'ReDim yMsg(6)
                'System.BitConverter.GetBytes(GlobalMessageCode.eMissileImpact).CopyTo(yMsg, 0)
                'System.BitConverter.GetBytes(oMissile.lMissileID).CopyTo(yMsg, 2)
                'yMsg(6) = oMissile.oWpnDef.AOERange
                Return
            Else
                'break apart... we send...
                ReDim yMsg(6)
                System.BitConverter.GetBytes(GlobalMessageCode.eMissileDetonated).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(oMissile.lMissileID).CopyTo(yMsg, 2)
                yMsg(6) = 0
            End If

            With oMissile
                If .oParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                    '40000 comes from max entity clip plane on client,the 125 comes from
                    '  10,000,000 / 40000 = 250 / 2 = 125 to offset to 0
                    Dim lCameraPosX As Int32 = (.LocX \ 40000) + 125
                    Dim lCameraPosZ As Int32 = (.LocZ \ 40000) + 125
                    goMsgSys.BroadcastToEnvironmentClients_Filter(yMsg, .oParentEnvir, lCameraPosX, lCameraPosZ)
                Else : goMsgSys.BroadcastToEnvironmentClients(yMsg, .oParentEnvir)
                End If
            End With
        End If
        
    End Sub

End Class
