Public Class UnitWeaponDef
    Inherits Epica_GUID

    Public oUnitDef As Epica_Entity_Def
    Public oWeaponDef As WeaponDef
    Public yArcID As UnitArcs
    Public mlAmmoCap As Int32        '-1 if energy weapon or ammo-less weapon... otherwise, it is based on the space allocated to the hardpoint but not used by the weapon
    Public lEntityStatusGroup As Int32

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Dim sIDField As String
        Dim sDefField As String
        Dim sTableName As String

        If oUnitDef Is Nothing Then Return False
        If oWeaponDef Is Nothing Then Return False

        'if the definition hasn't been saved, and save fails
        If oUnitDef.ObjectID = -1 AndAlso oUnitDef.SaveObject() = False Then Return False
        'if the weapondef hasn't been saved, and save fails
        If oWeaponDef.ObjectID = -1 AndAlso oWeaponDef.SaveObject() = False Then Return False

        If oUnitDef.ObjTypeID = ObjectType.eFacilityDef Then
            sIDField = "SDW_ID"
            sDefField = "StructureDefID"
            sTableName = "tblStructureDefWeapon"
        Else
            sIDField = "UDW_ID"
            sDefField = "UnitDefID"
            sTableName = "tblUnitDefWeapon"
        End If


        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID < 0 Then
                'INSERT
                sSQL = "INSERT INTO " & sTableName & "(WeaponDefID, " & sDefField & ", Quadrant, EntityStatusGroup, iAmmoCap) VALUES (" & _
                    oWeaponDef.ObjectID & ", " & oUnitDef.ObjectID & ", " & CByte(yArcID) & ", " & lEntityStatusGroup & _
                    ", " & mlAmmoCap & ")"
            Else
                'UPDATE
                sSQL = "UPDATE " & sTableName & " SET WeaponDefID = " & oWeaponDef.ObjectID & ", " & sDefField & " = " & _
                oUnitDef.ObjectID & ", Quadrant = " & CByte(yArcID) & ", EntityStatusGroup = " & lEntityStatusGroup & _
                ", iAmmoCap = " & mlAmmoCap & " WHERE " & sIDField & " = " & Me.ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(" & sIDField & ") FROM " & sTableName & " WHERE WeaponDefID = " & oWeaponDef.ObjectID & " AND " & sDefField & " = " & oUnitDef.ObjectID
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

		Dim sIDField As String
		Dim sDefField As String
		Dim sTableName As String

		If oUnitDef Is Nothing Then Return ""
		If oWeaponDef Is Nothing Then Return ""

		'if the definition hasn't been saved, and save fails
		If oUnitDef.ObjectID = -1 AndAlso oUnitDef.SaveObject() = False Then Return ""
		'if the weapondef hasn't been saved, and save fails
		If oWeaponDef.ObjectID = -1 AndAlso oWeaponDef.SaveObject() = False Then Return ""

		If oUnitDef.ObjTypeID = ObjectType.eFacilityDef Then
			sIDField = "SDW_ID"
			sDefField = "StructureDefID"
			sTableName = "tblStructureDefWeapon"
		Else
			sIDField = "UDW_ID"
			sDefField = "UnitDefID"
			sTableName = "tblUnitDefWeapon"
		End If

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

		Try
			If ObjectID < 0 Then
				'INSERT
				sSQL = "INSERT INTO " & sTableName & "(WeaponDefID, " & sDefField & ", Quadrant, EntityStatusGroup, iAmmoCap) VALUES (" & _
				 oWeaponDef.ObjectID & ", " & oUnitDef.ObjectID & ", " & CByte(yArcID) & ", " & lEntityStatusGroup & _
				 ", " & mlAmmoCap & ")"
			Else
				'UPDATE
				sSQL = "UPDATE " & sTableName & " SET WeaponDefID = " & oWeaponDef.ObjectID & ", " & sDefField & " = " & _
				oUnitDef.ObjectID & ", Quadrant = " & CByte(yArcID) & ", EntityStatusGroup = " & lEntityStatusGroup & _
				", iAmmoCap = " & mlAmmoCap & " WHERE " & sIDField & " = " & Me.ObjectID
			End If
			Return sSQL
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		End Try
		Return ""
	End Function

End Class
