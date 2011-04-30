Public Class FacilityDef
    Inherits Epica_Entity_Def

    Private Const ml_MSG_BASE_UB As Int32 = 139 '119

    Public WorkerFactor As Int32
    Public MaxFacilitySize As Byte
    Public ProdFactor As Int32
    Public PowerFactor As Int32

    Public Overrides Function GetObjAsString() As Byte()
        Dim X As Int32
        Dim lPos As Int32

        'here we will return the entire object as a string
        If mbStringReady = False Then


            'this equation may be slightly off...
            ReDim mySendString(ml_MSG_BASE_UB + ((ml_WEAPON_DEF_LEN + 12) * (WeaponDefUB + 1)))  '6

            lPos = 0
            GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
            System.BitConverter.GetBytes(OwnerID).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Q1_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Q2_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Q3_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Q4_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            mySendString(lPos) = Maneuver : lPos += 1
            mySendString(lPos) = MaxSpeed : lPos += 1
            mySendString(lPos) = FuelEfficiency : lPos += 1
            System.BitConverter.GetBytes(Structure_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Hangar_Cap).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(HullSize).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Cargo_Cap).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Shield_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ShieldRecharge).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ShieldRechargeFreq).CopyTo(mySendString, lPos) : lPos += 4

            'System.BitConverter.GetBytes(Max_Door_Size).CopyTo(mySendString, lPos) : lPos += 4
			'mySendString(lPos) = NumberOfDoors : lPos += 1

			Dim lCrew As Int32 = 1
			If Me.ProductionRequirements Is Nothing = False Then
				lCrew += Me.ProductionRequirements.EnlistedCost + Me.ProductionRequirements.OfficerCost + Me.ProductionRequirements.ColonistCost
			End If
			System.BitConverter.GetBytes(lCrew).CopyTo(mySendString, lPos) : lPos += 4		  'was Fuel_Cap
            mySendString(lPos) = Weapon_Acc : lPos += 1
            mySendString(lPos) = ScanResolution : lPos += 1
            mySendString(lPos) = OptRadarRange : lPos += 1
            mySendString(lPos) = MaxRadarRange : lPos += 1
            mySendString(lPos) = DisruptionResistance : lPos += 1
            mySendString(lPos) = PiercingResist : lPos += 1
            mySendString(lPos) = ImpactResist : lPos += 1
            mySendString(lPos) = BeamResist : lPos += 1
            mySendString(lPos) = ECMResist : lPos += 1
            mySendString(lPos) = FlameResist : lPos += 1
            mySendString(lPos) = ChemicalResist : lPos += 1
            mySendString(lPos) = DetectionResist : lPos += 1
            System.BitConverter.GetBytes(ModelID).CopyTo(mySendString, lPos) : lPos += 2  'model id

            'Now, include the def name
            DefName.CopyTo(mySendString, lPos) : lPos += 20
            System.BitConverter.GetBytes(WorkerFactor).CopyTo(mySendString, lPos) : lPos += 4
            mySendString(lPos) = MaxFacilitySize : lPos += 1
            System.BitConverter.GetBytes(ProdFactor).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(PowerFactor).CopyTo(mySendString, lPos) : lPos += 4
            mySendString(lPos) = ProductionTypeID : lPos += 1
            mySendString(lPos) = RequiredProductionTypeID : lPos += 1
            mySendString(lPos) = yChassisType : lPos += 1
            mySendString(lPos) = yFXColors : lPos += 1
            mySendString(lPos) = ArmorIntegrity : lPos += 1
            mySendString(lPos) = JamImmunity : lPos += 1
            mySendString(lPos) = JamStrength : lPos += 1
            mySendString(lPos) = JamTargets : lPos += 1
            mySendString(lPos) = JamEffect : lPos += 1

            'side crits
            'lPos = 108  '102
            For X = 0 To 3
                System.BitConverter.GetBytes(lSideCrits(X)).CopyTo(mySendString, lPos)
                lPos += 4
            Next X

            'Now, we need to include the Unit Def's weapons...
            'next two bytes is the weapondef count
            System.BitConverter.GetBytes(CShort(WeaponDefUB + 1)).CopyTo(mySendString, lPos)
			lPos += 2

			''================== PUT IN FOR PRE-BALANCE FIXES
			'Dim blTotalDPS As Int64 = 0
			'For X = 0 To WeaponDefUB
			'	Dim blDPSTemp As Int64 = 0
			'	With WeaponDefs(X).oWeaponDef
			'		blDPSTemp += .BeamMaxDmg
			'		blDPSTemp += .ChemicalMaxDmg
			'		blDPSTemp += .ECMMaxDmg
			'		blDPSTemp += .FlameMaxDmg
			'		blDPSTemp += .ImpactMaxDmg
			'		blDPSTemp += .PiercingMaxDmg

			'		Dim fShotsPerSec As Single = 1.0F
			'		If .ROF <> 0 Then fShotsPerSec = 30.0F / .ROF
			'		blDPSTemp = CLng(fShotsPerSec * blDPSTemp)
			'		blTotalDPS += blDPSTemp
			'	End With
			'Next X
			'Dim fDPSExcess As Single = CSng(blTotalDPS / GetMaxDPS(Me.ModelID))
			'If fDPSExcess < 1 Then fDPSExcess = 1.0F 
			''================= END of Balancing Changes =====================

            For X = 0 To WeaponDefUB
                'now, we need to append the weapon def's message here...
                'first, append the Arc
                If WeaponDefs(X) Is Nothing OrElse WeaponDefs(X).oWeaponDef Is Nothing Then
                    lPos += 1 + ml_WEAPON_DEF_LEN - 2 + 4 + 4 + 4
                    Continue For
                End If

                mySendString(lPos) = WeaponDefs(X).yArcID
                lPos += 1
                'then, append the rest
				'WeaponDefs(X).oWeaponDef.GetObjAsStringEx(fDPSExcess).CopyTo(mySendString, lPos)
				WeaponDefs(X).oWeaponDef.GetObjAsString().CopyTo(mySendString, lPos)
                lPos += ml_WEAPON_DEF_LEN - 2

                System.BitConverter.GetBytes(WeaponDefs(X).lEntityStatusGroup).CopyTo(mySendString, lPos)
                lPos += 4

                System.BitConverter.GetBytes(WeaponDefs(X).mlAmmoCap).CopyTo(mySendString, lPos)
                lPos += 4

                System.BitConverter.GetBytes(WeaponDefs(X).ObjectID).CopyTo(mySendString, lPos)
                lPos += 4

            Next X

            System.BitConverter.GetBytes(lExtendedFlags).CopyTo(mySendString, lPos) : lPos += 4

            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Overrides Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'TODO: Each time a player adds a new UnitDef. that Unit Def needs to be transmitted to all domain servers.
        '  Immediately. That way, the domain servers do not require additional information regarding a unit other 
        '  than the unit's temporary attributes (hps, loc, etc...) and the unit's unitDefID.

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblStructureDef (PrototypeID, Q1_MaxHP, Q2_MaxHP, Q3_MaxHP, Q4_MaxHP, Maneuver, " & _
                  "MaxSpeed, FuelEfficiency, Structure_MaxHP, Hangar_Cap, " & _
                  "HullSize, Cargo_Cap, Fuel_Cap, Weapon_Acc, ScanResolution, OptRadarRange, MaxRadarRange, " & _
                  "DisruptionResistance, PiercingResist, ImpactResist, BeamResist, ECMResist, FlameResist, " & _
                  "ChemicalResist, DetectionResist, Shield_MaxHP, Shield_Recharge, ShieldRechargeFreq, ModelID, FacilityDefName, Side1Crits, " & _
                  "Side2Crits, Side3Crits, Side4Crits, ProductionTypeID, RequiredProductionTypeID, ChassisType, WorkerFactor, MaxFacilitySize, " & _
                  "ProdFactor, PowerFactor, OwnerID, FXColors, ArmorIntegrity, JamImmunity, JamStrength, " & _
                  "JamTargets, JamEffect, ExtendedFlags) VALUES ("
                If oPrototype Is Nothing = False Then
                    sSQL &= oPrototype.ObjectID
                Else : sSQL &= "0"
                End If
                sSQL &= ", " & Q1_MaxHP & ", " & Q2_MaxHP & ", " & Q3_MaxHP & ", " & Q4_MaxHP & _
                  ", " & Maneuver & ", " & MaxSpeed & ", " & FuelEfficiency & ", " & Structure_MaxHP & ", " & _
                  Hangar_Cap & ", " & HullSize & _
                  ", " & Cargo_Cap & ", " & Fuel_Cap & ", " & Weapon_Acc & ", " & ScanResolution & ", " & _
                  OptRadarRange & ", " & MaxRadarRange & ", " & DisruptionResistance & ", " & PiercingResist & _
                  ", " & ImpactResist & ", " & BeamResist & ", " & ECMResist & ", " & FlameResist & ", " & _
                  ChemicalResist & ", " & DetectionResist & ", " & Shield_MaxHP & ", " & ShieldRecharge & ", " & _
                  ShieldRechargeFreq & ", " & ModelID & ", '" & MakeDBStr(BytesToString(DefName)) & "', " & lSideCrits(0) & ", " & _
                  lSideCrits(1) & ", " & lSideCrits(2) & ", " & lSideCrits(3) & ", " & ProductionTypeID & ", " & RequiredProductionTypeID & _
                  ", " & yChassisType & ", " & WorkerFactor & ", " & MaxFacilitySize & ", " & ProdFactor & ", " & _
                  PowerFactor & ", " & Me.OwnerID & ", " & yFXColors & ", " & ArmorIntegrity & ", " & JamImmunity & ", " & _
                  JamStrength & ", " & JamTargets & ", " & JamEffect & ", " & lExtendedFlags & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblStructureDef SET PrototypeID = "
                If oPrototype Is Nothing = False Then
                    sSQL &= oPrototype.ObjectID
                Else : sSQL &= "0"
                End If
                sSQL &= ", Q1_MaxHP = " & Q1_MaxHP & _
                  ", Q2_MaxHP = " & Q2_MaxHP & ", Q3_MaxHP = " & Q3_MaxHP & ", Q4_MaxHP = " & Q4_MaxHP & ", Maneuver=" & _
                  Maneuver & ", MaxSpeed = " & MaxSpeed & ", FuelEfficiency = " & FuelEfficiency & ", Structure_MaxHP=" & _
                  Structure_MaxHP & ", Hangar_Cap = " & Hangar_Cap & ", HullSize = " & _
                  HullSize & ", Cargo_Cap = " & Cargo_Cap & ", Fuel_cap = " & Fuel_Cap & ", Weapon_acc = " & _
                  Weapon_Acc & ", ScanResolution = " & ScanResolution & ", OptRadarRange = " & OptRadarRange & _
                  ", MaxRadarRange = " & MaxRadarRange & ", DisruptionResistance = " & DisruptionResistance & _
                  ", PiercingResist = " & PiercingResist & ", ImpactResist = " & ImpactResist & ", BeamResist = " & _
                  BeamResist & ", ECMResist = " & ECMResist & ", FlameResist = " & FlameResist & ", ChemicalResist=" & _
                  ChemicalResist & ", DetectionResist = " & DetectionResist & ", Shield_MaxHP = " & Shield_MaxHP & _
                  ", Shield_Recharge = " & ShieldRecharge & ", ShieldRechargeFreq = " & ShieldRechargeFreq & _
                  ", ModelID = " & ModelID & ", FacilityDefName='" & MakeDBStr(BytesToString(DefName)) & "', Side1Crits = " & lSideCrits(0) & _
                  ", Side2Crits = " & lSideCrits(1) & ", Side3Crits = " & lSideCrits(2) & ", Side4Crits = " & lSideCrits(3) & ", ProductionTypeID = " & _
                  ProductionTypeID & ", RequiredProductionTypeID = " & RequiredProductionTypeID & ", ChassisType = " & yChassisType & _
                  ", WorkerFactor = " & WorkerFactor & ", MaxFacilitySize = " & MaxFacilitySize & ", ProdFactor = " & ProdFactor & _
                  ", PowerFactor = " & PowerFactor & ", OwnerID = " & Me.OwnerID & ", FXColors = " & yFXColors & ", ArmorIntegrity = " & _
                  ArmorIntegrity & ", JamImmunity = " & JamImmunity & ", JamStrength = " & JamStrength & ", JamTargets = " & JamTargets & _
                  ", JamEffect = " & JamEffect & ", ExtendedFlags = " & lExtendedFlags & _
                  " WHERE FacilityDefID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                'TODO: If oPrototype is nothing, then this will fail
                sSQL = "SELECT MAX(FacilityDefID) FROM tblStructureDef WHERE PrototypeID = " & oPrototype.ObjectID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            For X As Int32 = 0 To Me.WeaponDefUB
                Me.WeaponDefs(X).oUnitDef = Me
                'If Me.WeaponDefs(X).oWeaponDef.SaveObject() = True Then
                If Me.WeaponDefs(X).SaveObject() = False Then
                    LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save Entity Def's Weapon Def. Reason: " & Err.Description)
                End If
                'Else
                '    LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save Entity Def's Weapon Def. Reason: " & Err.Description)
                'End If
            Next X

            For X As Int32 = 0 To Me.lHangarDoorUB
                If moHangarDoors(X) Is Nothing = False Then
                    moHangarDoors(X).lEntityDefID = Me.ObjectID
                    moHangarDoors(X).iEntityDefTypeID = Me.ObjTypeID
                    moHangarDoors(X).SaveObject()
                End If
            Next X

            sSQL = "DELETE FROM tblEntityDefMineral WHERE EntityDefID = " & Me.ObjectID & " and EntityDefTypeID = " & Me.ObjTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            For X As Int32 = 0 To Me.lEntityDefMineralUB
                If Me.EntityDefMinerals(X) Is Nothing = False Then
                    Me.EntityDefMinerals(X).SaveObject()
                End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
	End Function

	Public Overrides Function GetSaveObjectText() As String
		Dim sSQL As String

		Try

			If ObjectID = -1 Then
				SaveObject()
				Return ""
			End If

			Dim oSB As New System.Text.StringBuilder

			'UPDATE
			sSQL = "UPDATE tblStructureDef SET PrototypeID = "
			If oPrototype Is Nothing = False Then
				sSQL &= oPrototype.ObjectID
			Else : sSQL &= "0"
			End If
            sSQL &= ", Q1_MaxHP = " & Q1_MaxHP & _
              ", Q2_MaxHP = " & Q2_MaxHP & ", Q3_MaxHP = " & Q3_MaxHP & ", Q4_MaxHP = " & Q4_MaxHP & ", Maneuver=" & _
              Maneuver & ", MaxSpeed = " & MaxSpeed & ", FuelEfficiency = " & FuelEfficiency & ", Structure_MaxHP=" & _
              Structure_MaxHP & ", Hangar_Cap = " & Hangar_Cap & ", HullSize = " & _
              HullSize & ", Cargo_Cap = " & Cargo_Cap & ", Fuel_cap = " & Fuel_Cap & ", Weapon_acc = " & _
              Weapon_Acc & ", ScanResolution = " & ScanResolution & ", OptRadarRange = " & OptRadarRange & _
              ", MaxRadarRange = " & MaxRadarRange & ", DisruptionResistance = " & DisruptionResistance & _
              ", PiercingResist = " & PiercingResist & ", ImpactResist = " & ImpactResist & ", BeamResist = " & _
              BeamResist & ", ECMResist = " & ECMResist & ", FlameResist = " & FlameResist & ", ChemicalResist=" & _
              ChemicalResist & ", DetectionResist = " & DetectionResist & ", Shield_MaxHP = " & Shield_MaxHP & _
              ", Shield_Recharge = " & ShieldRecharge & ", ShieldRechargeFreq = " & ShieldRechargeFreq & _
              ", ModelID = " & ModelID & ", FacilityDefName='" & MakeDBStr(BytesToString(DefName)) & "', Side1Crits = " & lSideCrits(0) & _
              ", Side2Crits = " & lSideCrits(1) & ", Side3Crits = " & lSideCrits(2) & ", Side4Crits = " & lSideCrits(3) & ", ProductionTypeID = " & _
              ProductionTypeID & ", RequiredProductionTypeID = " & RequiredProductionTypeID & ", ChassisType = " & yChassisType & _
              ", WorkerFactor = " & WorkerFactor & ", MaxFacilitySize = " & MaxFacilitySize & ", ProdFactor = " & ProdFactor & _
              ", PowerFactor = " & PowerFactor & ", OwnerID = " & Me.OwnerID & ", FXColors = " & yFXColors & ", ArmorIntegrity = " & _
              ArmorIntegrity & ", JamImmunity = " & JamImmunity & ", JamStrength = " & JamStrength & ", JamTargets = " & JamTargets & _
              ", JamEffect = " & JamEffect & ", ExtendedFlags = " & lExtendedFlags & _
              " WHERE FacilityDefID = " & ObjectID

			oSB.AppendLine(sSQL)

			For X As Int32 = 0 To Me.WeaponDefUB
				Me.WeaponDefs(X).oUnitDef = Me
				oSB.AppendLine(Me.WeaponDefs(X).GetSaveObjectText())
			Next X

			For X As Int32 = 0 To Me.lHangarDoorUB
				If moHangarDoors(X) Is Nothing = False Then
					moHangarDoors(X).lEntityDefID = Me.ObjectID
					moHangarDoors(X).iEntityDefTypeID = Me.ObjTypeID
					oSB.AppendLine(moHangarDoors(X).GetSaveObjectText())
				End If
			Next X

			Return oSB.ToString
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		End Try
		Return ""
    End Function

    Public Overrides Sub FillFromDataReader(ByVal oData As System.Data.OleDb.OleDbDataReader, ByVal iObjTypeID As Int16)
        MyBase.FillFromDataReader(oData, iObjTypeID)
        With Me
            .MaxFacilitySize = CByte(oData("MaxFacilitySize"))
            .PowerFactor = CInt(oData("PowerFactor"))
            .ProdFactor = CInt(oData("ProdFactor"))
            .WorkerFactor = CInt(oData("WorkerFactor"))
        End With
    End Sub

    Public Overrides Sub FillFromForwardedAddMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 0

        With Me
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q1_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q2_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q3_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q4_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Maneuver = yData(lPos) : lPos += 1
            .MaxSpeed = yData(lPos) : lPos += 1
            .FuelEfficiency = yData(lPos) : lPos += 1
            .Structure_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Hangar_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Cargo_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Shield_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ShieldRecharge = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ShieldRechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lPos += 4   'for crew
            Weapon_Acc = yData(lPos) : lPos += 1
            ScanResolution = yData(lPos) : lPos += 1
            OptRadarRange = yData(lPos) : lPos += 1
            MaxRadarRange = yData(lPos) : lPos += 1
            DisruptionResistance = yData(lPos) : lPos += 1
            PiercingResist = yData(lPos) : lPos += 1
            ImpactResist = yData(lPos) : lPos += 1
            BeamResist = yData(lPos) : lPos += 1
            ECMResist = yData(lPos) : lPos += 1
            FlameResist = yData(lPos) : lPos += 1
            ChemicalResist = yData(lPos) : lPos += 1
            DetectionResist = yData(lPos) : lPos += 1
            ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            ReDim .DefName(19)
            Array.Copy(yData, lPos, .DefName, 0, 20) : lPos += 20

            .WorkerFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .MaxFacilitySize = yData(lPos) : lPos += 1
            .ProdFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .PowerFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ProductionTypeID = yData(lPos) : lPos += 1
            .RequiredProductionTypeID = yData(lPos) : lPos += 1
            .yChassisType = yData(lPos) : lPos += 1
            .yFXColors = yData(lPos) : lPos += 1
            .ArmorIntegrity = yData(lPos) : lPos += 1
            .JamImmunity = yData(lPos) : lPos += 1
            .JamStrength = yData(lPos) : lPos += 1
            .JamTargets = yData(lPos) : lPos += 1
            .JamEffect = yData(lPos) : lPos += 1

            For X As Int32 = 0 To 3
                lSideCrits(X) = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
            Next X

            WeaponDefUB = System.BitConverter.ToInt16(yData, lPos) - 1 : lPos += 2
            ReDim WeaponDefs(WeaponDefUB)

            For X As Int32 = 0 To WeaponDefUB
                WeaponDefs(X) = New UnitWeaponDef

                WeaponDefs(X).yArcID = CType(yData(lPos), UnitArcs) : lPos += 1

                Dim lWpnDefID As Int32 = System.BitConverter.ToInt32(yData, lPos)
                Dim oWpnDef As WeaponDef = GetEpicaWeaponDef(lWpnDefID)
                If oWpnDef Is Nothing = True Then
                    oWpnDef = New WeaponDef()
                    lPos = oWpnDef.FillFromPrimaryMsg(yData, lPos)
                    AddWeaponDefToGlobalArray(oWpnDef)
                    WeaponDefs(X).oWeaponDef = oWpnDef
                Else
                    'adjust lPos by size of the weapondef
                    'create a temporary wpn def to do our position calcualtion
                    Dim oTmp As WeaponDef = New WeaponDef
                    lPos = oTmp.FillFromPrimaryMsg(yData, lPos)
                    'DO NOT USE oTmp!!!

                    WeaponDefs(X).oWeaponDef = oWpnDef  'use oWpnDef as it is the global version that already exists
                End If

                WeaponDefs(X).lEntityStatusGroup = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                WeaponDefs(X).mlAmmoCap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                WeaponDefs(X).ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Next X

            .lExtendedFlags = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        End With

        Me.ProductionRequirements = New ProductionCost
        With Me.ProductionRequirements
            lPos = .FillFromPrimaryAddMsg(yData, lPos)
        End With

    End Sub
End Class
