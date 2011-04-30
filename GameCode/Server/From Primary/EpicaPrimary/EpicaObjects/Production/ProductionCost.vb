Public Class ProductionCost
    Public PC_ID As Int32           'Unique ID

    Public ObjectID As Int32        'ID of the item type being produced
    Public ObjTypeID As Int16       'Type ID of the Item Type being produced

    Public CreditCost As Int64      'cost in credits
    Public ColonistCost As Int32    'cost in colonists
    Public EnlistedCost As Int32    'cost in Enlisted Personnel (created at barracks)
    Public OfficerCost As Int32     'cost in Officers (created at Officer Training Facility)

    Public PointsRequired As Int64  'number of points required (whether its research or production)

    Public ProductionCostType As Byte   '0 = Production, 1 = Research

    Public ItemCosts() As ProductionCostItem
    Public ItemCostUB As Int32 = -1

    Private mySendString() As Byte
    Private mbSendStringReady As Boolean = False

    Public Function GetObjAsString() As Byte()
        Dim X As Int32
        Dim lPos As Int32 = 0
        'If mbSendStringReady = False Then
        'ReDim mySendString(40 + ((ItemCostUB + 1) * 12))
        ReDim mySendString(44 + ((ItemCostUB + 1) * 10))

        System.BitConverter.GetBytes(PC_ID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(ObjectType.eProductionCost).CopyTo(mySendString, lPos) : lPos += 2
        System.BitConverter.GetBytes(ObjectID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(ObjTypeID).CopyTo(mySendString, lPos) : lPos += 2
        System.BitConverter.GetBytes(CreditCost).CopyTo(mySendString, lPos) : lPos += 8
        System.BitConverter.GetBytes(ColonistCost).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(EnlistedCost).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(OfficerCost).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(PointsRequired).CopyTo(mySendString, lPos) : lPos += 8
        mySendString(lPos) = ProductionCostType : lPos += 1     'NOTE: This is position sensitve on the client!!!
        System.BitConverter.GetBytes(CInt(ItemCostUB + 1)).CopyTo(mySendString, lPos) : lPos += 4

        For X = 0 To ItemCostUB
            'System.BitConverter.GetBytes(ItemCosts(X).PCM_ID).CopyTo(mySendString, lPos) : lPos += 4
            'If MineralCosts(X).oMineral Is Nothing = False Then
            '    System.BitConverter.GetBytes(MineralCosts(X).oMineral.ObjectID).CopyTo(mySendString, lPos) : lPos += 4
            'Else : System.BitConverter.GetBytes(MineralCosts(X).MineralID).CopyTo(mySendString, lPos) : lPos += 4
            'End If
            System.BitConverter.GetBytes(ItemCosts(X).ItemID).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ItemCosts(X).ItemTypeID).CopyTo(mySendString, lPos) : lPos += 2
            System.BitConverter.GetBytes(ItemCosts(X).QuantityNeeded).CopyTo(mySendString, lPos) : lPos += 4
        Next X

        mbSendStringReady = True
        'End If
        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If PC_ID < 1 Then
                'INSERT
                sSQL = "INSERT INTO tblProductionCost (ObjectID, ObjTypeID, Credits, " & _
                  "Colonists, Enlisted, Officers, PointsRequired, ProductionCostType) VALUES (" & Me.ObjectID & _
                  ", " & Me.ObjTypeID & ", " & Me.CreditCost & ", " & Me.ColonistCost & _
                  ", " & Me.EnlistedCost & ", " & Me.OfficerCost & ", " & Me.PointsRequired & ", " & Me.ProductionCostType & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblProductionCost SET ObjectID = " & Me.ObjectID & _
                  ", ObjTypeId = " & Me.ObjTypeID & ", Credits = " & Me.CreditCost & _
                  ", Colonists = " & Me.ColonistCost & ", Enlisted = " & Me.EnlistedCost & _
                  ", Officers = " & Me.OfficerCost & ", PointsRequired = " & _
                  Me.PointsRequired & ", ProductionCostType = " & Me.ProductionCostType & " WHERE PC_ID = " & Me.PC_ID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If PC_ID < 1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PC_ID) FROM tblProductionCost WHERE ObjectID = " & Me.ObjectID & " AND ObjTypeID = " & Me.ObjTypeID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    PC_ID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            For X As Int32 = 0 To Me.ItemCostUB
                Me.ItemCosts(X).PC_ID = PC_ID
                Me.ItemCosts(X).SaveObject()
            Next X

            oComm = Nothing

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & PC_ID & " of type ProductionCost. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Public Function GetSaveObjectText() As String
		Dim sSQL As String

		If PC_ID < 1 Then
			SaveObject()
			Return ""
		End If

		Try
			Dim oSB As New System.Text.StringBuilder

			'UPDATE
			sSQL = "UPDATE tblProductionCost SET ObjectID = " & Me.ObjectID & _
			  ", ObjTypeId = " & Me.ObjTypeID & ", Credits = " & Me.CreditCost & _
			  ", Colonists = " & Me.ColonistCost & ", Enlisted = " & Me.EnlistedCost & _
			  ", Officers = " & Me.OfficerCost & ", PointsRequired = " & _
			  Me.PointsRequired & ", ProductionCostType = " & Me.ProductionCostType & " WHERE PC_ID = " & Me.PC_ID
			oSB.AppendLine(sSQL)

			For X As Int32 = 0 To Me.ItemCostUB
				Me.ItemCosts(X).PC_ID = PC_ID
				oSB.AppendLine(Me.ItemCosts(X).GetSaveObjectText())
			Next X

			Return oSB.ToString
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & PC_ID & " of type ProductionCost. Reason: " & Err.Description)
		End Try
		Return ""
	End Function

	Public Sub AddProductionCostItem(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lQty As Int32)
		For X As Int32 = 0 To ItemCostUB
			If ItemCosts(X).ItemID = lID AndAlso ItemCosts(X).ItemTypeID = iTypeID Then
				ItemCosts(X).QuantityNeeded += lQty
				Return
			End If
		Next X

		ItemCostUB += 1
		ReDim Preserve ItemCosts(ItemCostUB)
		ItemCosts(ItemCostUB) = New ProductionCostItem
		With ItemCosts(ItemCostUB)
			.ItemID = lID
			.ItemTypeID = iTypeID
			.oProdCost = Me
			.PC_ID = Me.PC_ID
			.PCM_ID = -1
			.QuantityNeeded = lQty
		End With
	End Sub

    Public Sub MakeDirty()
        mbSendStringReady = False
    End Sub

	Public Function CreateClone(ByVal lMultiplyCost As Int32) As ProductionCost
		lMultiplyCost = Math.Max(lMultiplyCost, 1)
		Dim oResult As New ProductionCost
		With oResult
			.ColonistCost = ColonistCost * lMultiplyCost
			.CreditCost = CreditCost * lMultiplyCost
			.EnlistedCost = EnlistedCost * lMultiplyCost
			.ObjectID = ObjectID
			.ObjTypeID = ObjTypeID
			.OfficerCost = OfficerCost
			.PC_ID = -1
			.PointsRequired = PointsRequired * lMultiplyCost
			.ProductionCostType = ProductionCostType
			For X As Int32 = 0 To ItemCostUB
				.AddProductionCostItem(ItemCosts(X).ItemID, ItemCosts(X).ItemTypeID, ItemCosts(X).QuantityNeeded * lMultiplyCost)
			Next X
		End With
		Return oResult
	End Function

	Public Sub DeleteItems()
		Dim sSQL As String = "DELETE FROM tblProductionCostItem WHERE PC_ID = " & Me.PC_ID
		Dim oComm As OleDb.OleDbCommand = Nothing
		Try
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "Unable to Delete ProductionCostItems for PC_ID " & Me.PC_ID & ": " & ex.Message)
		Finally
			oComm = Nothing
		End Try
	End Sub

	Public Sub DeleteMe()

		DeleteItems()

		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand = Nothing
		Try
			sSQL = "DELETE FROM tblProductionCost WHERE PC_ID = " & Me.PC_ID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "Unable to Delete ProductionCost " & Me.PC_ID & ": " & ex.Message)
		Finally
			oComm = Nothing
		End Try
    End Sub

    Public Function FillFromPrimaryAddMsg(ByVal yData() As Byte, ByVal lPos As Int32) As Int32
        With Me
            .PC_ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lPos += 2
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .CreditCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .ColonistCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .EnlistedCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .OfficerCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .PointsRequired = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .ProductionCostType = yData(lPos) : lPos += 1
            .ItemCostUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
            ReDim .ItemCosts(.ItemCostUB)
            For X As Int32 = 0 To .ItemCostUB
                .ItemCosts(X) = New ProductionCostItem()
                With .ItemCosts(X)
                    .ItemID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ItemTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .QuantityNeeded = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .oProdCost = Me
                    .PC_ID = Me.PC_ID
                End With
            Next X
        End With

    End Function
End Class
