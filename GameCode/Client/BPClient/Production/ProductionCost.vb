'This class defines all of the requirements for a production item
Public Class ProductionCost
    Public PC_ID As Int32           'Unique ID

    Public ObjectID As Int32        'ID of the item type being produced
    Public ObjTypeID As Int16       'Type ID of the Item Type being produced

    Public CreditCost As Int64      'cost in credits
    Public ColonistCost As Int32    'cost in colonists
    Public EnlistedCost As Int32    'cost in Enlisted Personnel (created at barracks)
    Public OfficerCost As Int32     'cost in Officers (created at Officer Training Facility)

    Public PointsRequired As Int64  'number of points required (whether its research or production)

    Public ProductionCostType As Byte       '0 = ProductionCost, 1 = ResearchCost

    Public ItemCosts() As ProductionCostItem
    Public ItemCostUB As Int32 = -1

    Private mySendString() As Byte
    Private mbSendStringReady As Boolean = False

	Public Function GetBuildCostText(ByVal sPower As String) As String
		Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder

		If ProductionCostType = 0 Then
			oSB.AppendLine("BUILD COSTS")
		Else
			oSB.AppendLine("RESEARCH COSTS")
		End If

        If CreditCost <> 0 Then oSB.AppendLine("  Credits: " & CreditCost.ToString("#,##0"))
        If ColonistCost <> 0 Then oSB.AppendLine("  Colonists: " & ColonistCost.ToString("#,##0"))
		If ObjTypeID <> ObjectType.eArmorTech AndAlso ObjTypeID <> ObjectType.eEngineTech AndAlso ObjTypeID <> ObjectType.eHullTech _
		  AndAlso ObjTypeID <> ObjectType.eRadarTech AndAlso ObjTypeID <> ObjectType.eShieldTech AndAlso ObjTypeID <> ObjectType.eWeaponTech Then
            If EnlistedCost <> 0 Then oSB.AppendLine("  Enlisted: " & EnlistedCost.ToString("#,##0"))
            If OfficerCost <> 0 Then oSB.AppendLine("  Officers: " & OfficerCost.ToString("#,##0"))
		End If

		'Power requirement
		If sPower.Length <> 0 Then oSB.AppendLine(sPower)

		If ItemCostUB <> -1 Then oSB.AppendLine("MATERIALS")

		For X As Int32 = 0 To ItemCostUB
            oSB.AppendLine("  " & ItemCosts(X).QuantityNeeded.ToString("#,##0") & " " & ItemCosts(X).GetItemName())
		Next X

		Return oSB.ToString

	End Function
End Class
