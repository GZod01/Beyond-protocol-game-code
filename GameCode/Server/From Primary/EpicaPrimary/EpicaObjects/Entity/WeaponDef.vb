Public Class WeaponDef
    Inherits Epica_GUID

    Public WeaponName(19) As Byte
    Public WeaponType As Byte
    Public ROF As Int16
    Public Range As Short
    Public Accuracy As Byte
    Public PiercingMinDmg As Int32
    Public PiercingMaxDmg As Int32
    Public ImpactMinDmg As Int32
    Public ImpactMaxDmg As Int32
    Public BeamMinDmg As Int32
    Public BeamMaxDmg As Int32
    Public ECMMinDmg As Int32
    Public ECMMaxDmg As Int32
    Public FlameMinDmg As Int32
    Public FlameMaxDmg As Int32
    Public ChemicalMinDmg As Int32
    Public ChemicalMaxDmg As Int32

    Public RelatedWeapon As BaseWeaponTech

    Public AOERange As Byte

    Public AmmoSize As Single       'NOTE: Not sent out with any messages
    Public AmmoReloadDelay As Int16 'NOTE: Not sent out with any messages

    Public MaxSpeed As Byte
    Public Maneuver As Byte

    Public WpnGroup As WeaponGroupType
    Public lFirePowerRating As Int32 

	Private mySendString() As Byte

    Public Function GetOverallDPS() As Decimal
        Dim decTemp As Decimal = 0D
        decTemp += PiercingMaxDmg
        decTemp += ImpactMaxDmg
        If WeaponType < EpicaPrimary.WeaponType.eMissile_Color_1 OrElse WeaponType > EpicaPrimary.WeaponType.eMissile_Color_9 Then
            decTemp += BeamMaxDmg
        End If
        decTemp += ECMMaxDmg
        decTemp += FlameMaxDmg
        decTemp += ChemicalMaxDmg

        Dim decROF As Decimal = Math.Max(1, ROF) / 30D
        decTemp /= decROF
        Return decTemp
    End Function

	'Public Function GetObjAsStringEx(ByVal fDPSExcess As Single) As Byte()
	'	'here we will return the entire object as a string
	'	If mbStringReady = False Then
	'		'NOTE: If I change the length of this array then I must also update the Unit Def's array system as well
	'		'ReDim mySendString(67)
	'		ReDim mySendString(91)

	'		Dim lPos As Int32 = 0

	'		GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
	'		WeaponName.CopyTo(mySendString, lPos) : lPos += 20
	'		mySendString(lPos) = WeaponType : lPos += 1
	'		'System.BitConverter.GetBytes(ROF).CopyTo(mySendString, lPos) : lPos += 2
	'		'================= for balancing issue =====================
	'		Dim iROFValue As Int16 = ROF
	'		If iROFValue < 30 Then iROFValue = 30
	'		System.BitConverter.GetBytes(iROFValue).CopyTo(mySendString, lPos) : lPos += 2
	'		'===========================================================

	'		System.BitConverter.GetBytes(Range).CopyTo(mySendString, lPos) : lPos += 2
	'		mySendString(lPos) = Accuracy : lPos += 1
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(PiercingMinDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(PiercingMaxDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(ImpactMinDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(ImpactMaxDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(BeamMinDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(BeamMaxDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(ECMMinDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(ECMMaxDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(FlameMinDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(FlameMaxDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(ChemicalMinDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4
	'		System.BitConverter.GetBytes(CInt(Math.Ceiling(ChemicalMaxDmg / fDPSExcess))).CopyTo(mySendString, lPos) : lPos += 4

	'		mySendString(lPos) = WpnGroup : lPos += 1
	'		System.BitConverter.GetBytes(lFirePowerRating).CopyTo(mySendString, lPos) : lPos += 4
	'		'System.BitConverter.GetBytes(lEntityStatusGroup).CopyTo(mySendString, 60)


	'		'TODO: I put this check in here for now, but eventually, it should no longer be needed
	'		If RelatedWeapon Is Nothing Then
	'			System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, lPos)
	'		Else
	'			System.BitConverter.GetBytes(RelatedWeapon.ObjectID).CopyTo(mySendString, lPos)
	'		End If
	'		lPos += 4

	'		'mySendString(lPos) = AOERange : lPos += 1
	'		'====================== AOE RANGE DISABLE ======================
	'		mySendString(lPos) = 0 : lPos += 1
	'		'===============================================================

	'		mySendString(lPos) = MaxSpeed : lPos += 1
	'		mySendString(lPos) = Maneuver : lPos += 1
	'		mbStringReady = True
	'	End If
	'	Return mySendString
	'End Function

    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        If mbStringReady = False Then
            'NOTE: If I change the length of this array then I must also update the Unit Def's array system as well
            'ReDim mySendString(67)
            ReDim mySendString(91)

            Dim lPos As Int32 = 0

            GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
            WeaponName.CopyTo(mySendString, lPos) : lPos += 20
            mySendString(lPos) = WeaponType : lPos += 1
            System.BitConverter.GetBytes(ROF).CopyTo(mySendString, lPos) : lPos += 2
            System.BitConverter.GetBytes(Range).CopyTo(mySendString, lPos) : lPos += 2
            mySendString(lPos) = Accuracy : lPos += 1
            System.BitConverter.GetBytes(PiercingMinDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(PiercingMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ImpactMinDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ImpactMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(BeamMinDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(BeamMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ECMMinDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ECMMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(FlameMinDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(FlameMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ChemicalMinDmg).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ChemicalMaxDmg).CopyTo(mySendString, lPos) : lPos += 4

            mySendString(lPos) = WpnGroup : lPos += 1
            System.BitConverter.GetBytes(lFirePowerRating).CopyTo(mySendString, lPos) : lPos += 4
            'System.BitConverter.GetBytes(lEntityStatusGroup).CopyTo(mySendString, 60)


            'TODO: I put this check in here for now, but eventually, it should no longer be needed
            If RelatedWeapon Is Nothing Then
                System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, lPos)
            Else
                System.BitConverter.GetBytes(RelatedWeapon.ObjectID).CopyTo(mySendString, lPos)
            End If
            lPos += 4

            mySendString(lPos) = AOERange : lPos += 1
            mySendString(lPos) = MaxSpeed : lPos += 1
            mySendString(lPos) = Maneuver : lPos += 1
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
                sSQL = "INSERT INTO tblWeaponDef (WeaponName, WeaponType, ROF, Range, Accuracy, PiercingMinDmg, " & _
                  "PiercingMaxDmg, ImpactMinDmg, ImpactMaxDmg, BeamMinDmg, BeamMaxDmg, ECMMinDmg, ECMMaxDMg, " & _
                  "FlameMinDmg, FlameMaxDmg, ChemicalMinDmg, ChemicalMaxDmg, AOERange, AmmoSize, AmmoReloadDelay," & _
                  "WpnGroup, FirePowerRating, MaxSpeed, Maneuver, WeaponID) VALUES ('" & MakeDBStr(BytesToString(WeaponName)) & _
                  "', " & WeaponType & ", " & ROF & ", " & Range & ", " & Accuracy & ", " & PiercingMinDmg & ", " & _
                  PiercingMaxDmg & ", " & ImpactMinDmg & ", " & ImpactMaxDmg & ", " & BeamMinDmg & ", " & _
                  BeamMaxDmg & ", " & ECMMinDmg & ", " & ECMMaxDmg & ", " & FlameMinDmg & ", " & FlameMaxDmg & ", " & _
                  ChemicalMinDmg & ", " & ChemicalMaxDmg & ", " & AOERange & ", " & _
                  AmmoSize & ", " & AmmoReloadDelay & ", " & CByte(WpnGroup) & ", " & lFirePowerRating & ", " & MaxSpeed & _
                  ", " & Maneuver
                If Me.RelatedWeapon Is Nothing = False Then
                    sSQL &= ", " & Me.RelatedWeapon.ObjectID
                Else : sSQL &= ", -1"
                End If
                sSQL &= ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblWeaponDef SET WeaponName = '" & MakeDBStr(BytesToString(WeaponName)) & "', WeaponType = " & WeaponType & _
                  ", ROF = " & ROF & ", Range = " & Range & ", Accuracy = " & Accuracy & ", PiercingMinDmg = " & _
                  PiercingMinDmg & ", PiercingMaxDmg = " & PiercingMaxDmg & ", ImpactMinDmg = " & ImpactMinDmg & _
                  ", ImpactMaxDmg = " & ImpactMaxDmg & ", BeamMinDmg = " & BeamMinDmg & ", BeamMaxDmg = " & _
                  BeamMaxDmg & ", ECMMinDmg = " & ECMMinDmg & ", ECMMaxDmg = " & ECMMaxDmg & ", FlameMinDmg = " & _
                  FlameMinDmg & ", FlameMaxDmg = " & FlameMaxDmg & ", ChemicalMinDmg = " & ChemicalMinDmg & _
                  ", ChemicalMaxDmg = " & ChemicalMaxDmg & ", WeaponID = "
                If RelatedWeapon Is Nothing = False Then
                    sSQL &= RelatedWeapon.ObjectID
                Else : sSQL &= "-1"
                End If
                sSQL &= ", AOERange = " & AOERange & ", AmmoSize = " & AmmoSize & ", AmmoReloadDelay = " & _
                  AmmoReloadDelay & ", WpnGroup = " & CByte(WpnGroup) & ", FirePowerRating = " & lFirePowerRating & _
                  ", MaxSpeed = " & MaxSpeed & ", Maneuver = " & Maneuver & _
                  " WHERE WeaponDefID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(WeaponDefID) FROM tblWeaponDef WHERE WeaponName = '" & MakeDBStr(BytesToString(WeaponName)) & "'"
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
		Dim sSQL As String

		Try
			If ObjectID = -1 Then
				'INSERT
				sSQL = "INSERT INTO tblWeaponDef (WeaponName, WeaponType, ROF, Range, Accuracy, PiercingMinDmg, " & _
				  "PiercingMaxDmg, ImpactMinDmg, ImpactMaxDmg, BeamMinDmg, BeamMaxDmg, ECMMinDmg, ECMMaxDMg, " & _
				  "FlameMinDmg, FlameMaxDmg, ChemicalMinDmg, ChemicalMaxDmg, AOERange, AmmoSize, AmmoReloadDelay," & _
				  "WpnGroup, FirePowerRating, MaxSpeed, Maneuver, WeaponID) VALUES ('" & MakeDBStr(BytesToString(WeaponName)) & _
				  "', " & WeaponType & ", " & ROF & ", " & Range & ", " & Accuracy & ", " & PiercingMinDmg & ", " & _
				  PiercingMaxDmg & ", " & ImpactMinDmg & ", " & ImpactMaxDmg & ", " & BeamMinDmg & ", " & _
				  BeamMaxDmg & ", " & ECMMinDmg & ", " & ECMMaxDmg & ", " & FlameMinDmg & ", " & FlameMaxDmg & ", " & _
				  ChemicalMinDmg & ", " & ChemicalMaxDmg & ", " & AOERange & ", " & _
				  AmmoSize & ", " & AmmoReloadDelay & ", " & CByte(WpnGroup) & ", " & lFirePowerRating & ", " & MaxSpeed & _
				  ", " & Maneuver
				If Me.RelatedWeapon Is Nothing = False Then
					sSQL &= ", " & Me.RelatedWeapon.ObjectID
				Else : sSQL &= ", -1"
				End If
				sSQL &= ")"
			Else
				'UPDATE
				sSQL = "UPDATE tblWeaponDef SET WeaponName = '" & MakeDBStr(BytesToString(WeaponName)) & "', WeaponType = " & WeaponType & _
				  ", ROF = " & ROF & ", Range = " & Range & ", Accuracy = " & Accuracy & ", PiercingMinDmg = " & _
				  PiercingMinDmg & ", PiercingMaxDmg = " & PiercingMaxDmg & ", ImpactMinDmg = " & ImpactMinDmg & _
				  ", ImpactMaxDmg = " & ImpactMaxDmg & ", BeamMinDmg = " & BeamMinDmg & ", BeamMaxDmg = " & _
				  BeamMaxDmg & ", ECMMinDmg = " & ECMMinDmg & ", ECMMaxDmg = " & ECMMaxDmg & ", FlameMinDmg = " & _
				  FlameMinDmg & ", FlameMaxDmg = " & FlameMaxDmg & ", ChemicalMinDmg = " & ChemicalMinDmg & _
				  ", ChemicalMaxDmg = " & ChemicalMaxDmg & ", WeaponID = "
				If RelatedWeapon Is Nothing = False Then
					sSQL &= RelatedWeapon.ObjectID
				Else : sSQL &= "-1"
				End If
				sSQL &= ", AOERange = " & AOERange & ", AmmoSize = " & AmmoSize & ", AmmoReloadDelay = " & _
				  AmmoReloadDelay & ", WpnGroup = " & CByte(WpnGroup) & ", FirePowerRating = " & lFirePowerRating & _
				  ", MaxSpeed = " & MaxSpeed & ", Maneuver = " & Maneuver & _
				  " WHERE WeaponDefID = " & ObjectID
			End If
			Return sSQL
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		End Try
		Return ""
    End Function

    Public Sub FillFromDataReader(ByVal oData As OleDb.OleDbDataReader)
        With Me
            .Accuracy = CByte(oData("Accuracy"))
            .AOERange = CByte(oData("AOERange"))
            .BeamMaxDmg = CInt(oData("BeamMaxDmg"))
            .BeamMinDmg = CInt(oData("BeamMinDmg"))
            .ChemicalMaxDmg = CInt(oData("ChemicalMaxDmg"))
            .ChemicalMinDmg = CInt(oData("ChemicalMinDmg"))
            .ECMMaxDmg = CInt(oData("ECMMaxDmg"))
            .ECMMinDmg = CInt(oData("ECMMinDmg"))
            .FlameMaxDmg = CInt(oData("FlameMaxDmg"))
            .FlameMinDmg = CInt(oData("FlameMinDmg"))
            .ImpactMaxDmg = CInt(oData("ImpactMaxDmg"))
            .ImpactMinDmg = CInt(oData("ImpactMinDmg"))
            .ObjectID = CInt(oData("WeaponDefID"))
            .ObjTypeID = ObjectType.eWeaponDef
            .PiercingMaxDmg = CInt(oData("PiercingMaxDmg"))
            .PiercingMinDmg = CInt(oData("PiercingMinDmg"))
            .Range = CShort(oData("Range"))
            .MaxSpeed = CByte(oData("MaxSpeed"))
            .Maneuver = CByte(oData("Maneuver"))

            If DBNull.Value Is oData("OwnerID") = False Then
                Dim lOwnerID As Int32 = CInt(oData("OwnerID"))
                Dim oPlayer As Player = Nothing
                If lOwnerID = 0 Then
                    oPlayer = goInitialPlayer
                Else : oPlayer = GetEpicaPlayer(lOwnerID)
                End If
                If oPlayer Is Nothing = False Then
                    .RelatedWeapon = CType(oPlayer.GetTech(CInt(oData("WeaponID")), ObjectType.eWeaponTech), BaseWeaponTech)
                End If
            End If

            .ROF = CShort(oData("ROF"))
            .WeaponName = StringToBytes(CStr(oData("WeaponName")))
            .WeaponType = CByte(oData("WeaponType"))
            .WpnGroup = CType(oData("WpnGroup"), WeaponGroupType)
            .lFirePowerRating = CInt(oData("FirePowerRating"))

            .AmmoSize = CSng(oData("AmmoSize"))
            .AmmoReloadDelay = CShort(oData("AmmoReloadDelay"))
        End With
    End Sub

    Public Function FillFromPrimaryMsg(ByVal yData() As Byte, ByVal lPos As Int32) As Int32
        'Dim lPos As Int32 = 0

        'GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
        'WeaponName.CopyTo(mySendString, lPos) : lPos += 20
        'mySendString(lPos) = WeaponType : lPos += 1
        'System.BitConverter.GetBytes(ROF).CopyTo(mySendString, lPos) : lPos += 2
        'System.BitConverter.GetBytes(Range).CopyTo(mySendString, lPos) : lPos += 2
        'mySendString(lPos) = Accuracy : lPos += 1
        'System.BitConverter.GetBytes(PiercingMinDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(PiercingMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(ImpactMinDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(ImpactMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(BeamMinDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(BeamMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(ECMMinDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(ECMMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(FlameMinDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(FlameMaxDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(ChemicalMinDmg).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(ChemicalMaxDmg).CopyTo(mySendString, lPos) : lPos += 4

        'mySendString(lPos) = WpnGroup : lPos += 1
        'System.BitConverter.GetBytes(lFirePowerRating).CopyTo(mySendString, lPos) : lPos += 4
        ''System.BitConverter.GetBytes(lEntityStatusGroup).CopyTo(mySendString, 60)


        ''TODO: I put this check in here for now, but eventually, it should no longer be needed
        'If RelatedWeapon Is Nothing Then
        '    System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, lPos)
        'Else
        '    System.BitConverter.GetBytes(RelatedWeapon.ObjectID).CopyTo(mySendString, lPos)
        'End If
        'lPos += 4

        'mySendString(lPos) = AOERange : lPos += 1
        'mySendString(lPos) = MaxSpeed : lPos += 1
        'mySendString(lPos) = Maneuver : lPos += 1
    End Function
End Class
