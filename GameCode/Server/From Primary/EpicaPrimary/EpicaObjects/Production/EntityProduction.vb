Public Class EntityProduction
    Public oParent As Epica_Entity
    Public OrderNumber As Byte
    Public PointsProduced As Int64      'points accumulated through production thus far

    Public ProductionTypeID As Int16    'what type of production is being made
    Public ProductionID As Int32        'in the event that the item is a Unit, Facility, etc... this indicates the Def
    Public ProdCost As ProductionCost

    'This is a duplicate of the ProdCost.PointsRequired, however, it will provide us the ability to modify the points required
    Public PointsRequired As Int64

    Public lFinishCycle As Int32
    Public lLastUpdateCycle As Int32

    Public lProdCount As Int32          'how many of these to produce... this value is decremented when
    'production is finished until the value = 0 (at which point, it is removed). If this value is -1, then the
    'production always stays up and the facility will always produce this production. (infinite production)

    Public lProdX As Int32 = Int32.MinValue
    Public lProdZ As Int32 = Int32.MinValue
    Public iProdA As Int16

    Public bPaidFor As Boolean = False      'indicates that the costs have been paid for this item

    'For Repair Items... and ammo
    Public lExtendedID As Int32
    Public iExtendedTypeID As Int16
    Public yExtendedType As Byte
    Public lAmmoQty As Int32        'used for the final production item when producing ammo for a reload
    Public fAmmoSize As Single
    Public lAmmoWpnTechID As Int32

    Private mbStringReady As Boolean = False
    Private mySendString() As Byte
    Public Sub DataChanged()    'call this sub when changing properties of this object
        mbStringReady = False
    End Sub

    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        If mbStringReady = False Then
            ReDim mySendString(20)      '0 to 16 = 17 bytes

            System.BitConverter.GetBytes(oParent.ObjectID).CopyTo(mySendString, 0)    '4
            mySendString(4) = OrderNumber                                               '1
            System.BitConverter.GetBytes(PointsProduced).CopyTo(mySendString, 5)        '4
            System.BitConverter.GetBytes(ProductionID).CopyTo(mySendString, 13)  '4
            System.BitConverter.GetBytes(ProductionTypeID).CopyTo(mySendString, 17) '4
            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Function SaveObject(ByVal bClearOld As Boolean) As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            'Ok, first, delete the object if it already exists
            If bClearOld = True Then
                sSQL = "DELETE FROM tblStructureProduction WHERE StructureID = " & oParent.ObjectID & " AND OrderNum = " & OrderNumber
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm = Nothing
            End If

            'Now, insert it

            sSQL = "INSERT INTO tblStructureProduction (StructureID, OrderNum, PointsProduced, ObjTypeID, ObjectID, ProdCount)" & _
              " VALUES (" & oParent.ObjectID & ", " & OrderNumber & ", " & PointsProduced & ", " & _
              ProductionTypeID & ", " & ProductionID & ", " & Me.lProdCount & ")"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object Facility Production " & OrderNumber & " for entity " & oParent.ObjectID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Public Function GetSaveObjectText() As String
		Dim sSQL As String

		Try
			'Ok, first, delete the object if it already exists
            sSQL = "" 'sSQL = "DELETE FROM tblStructureProduction WHERE StructureID = " & oParent.ObjectID & " AND OrderNum = " & OrderNumber

			'Now, insert it
			sSQL &= vbCrLf & "INSERT INTO tblStructureProduction (StructureID, OrderNum, PointsProduced, ObjTypeID, ObjectID, ProdCount)" & _
			 " VALUES (" & oParent.ObjectID & ", " & OrderNumber & ", " & PointsProduced & ", " & _
			 ProductionTypeID & ", " & ProductionID & ", " & Me.lProdCount & ")"
			Return sSQL
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object Facility Production " & OrderNumber & " for entity " & oParent.ObjectID & ". Reason: " & Err.Description)
		End Try
		Return ""
	End Function

    Private miProdItemModel As Int16 = Int16.MinValue
    Public ReadOnly Property ProductionItemModelID() As Int16
        Get
            If miProdItemModel = Int16.MinValue Then
                miProdItemModel = -1
                If ProductionTypeID = ObjectType.eFacilityDef Then
                    Dim oFacDef As FacilityDef = GetEpicaFacilityDef(ProductionID)
                    If oFacDef Is Nothing = False Then miProdItemModel = oFacDef.ModelID
                    oFacDef = Nothing
                End If
            End If
            Return miProdItemModel
        End Get
    End Property
End Class
