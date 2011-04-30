Public Class ProductionCostItem
    'Only what we know about...
    'ItemID
    'ItemTypeID
    'Quantity

    Public ItemID As Int32       'FK to ObjectID
    Public ItemTypeID As Int16
    Public QuantityNeeded As Int32  'number of this mineral required

    Public Function GetItemName() As String
        If ItemTypeID = ObjectType.eMineralCache OrElse ItemTypeID = ObjectType.eMineral Then
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = ItemID Then
                    Return goMinerals(X).MineralName
                End If
            Next X
            Return "Unknown Mineral"
        Else
            'check for a tech
            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(ItemID, ItemTypeID)
            If oTech Is Nothing = False Then Return oTech.GetComponentName

            'if we are here, then try the cache
            Return GlobalVars.GetCacheObjectValue(ItemID, ItemTypeID)
        End If
    End Function
End Class
