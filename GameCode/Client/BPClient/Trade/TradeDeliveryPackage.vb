Option Strict On

Public Class TradeDeliveryPackage
    Public lTradePostID As Int32        'the tradepost receiving the items
    Public lSourcePlayerID As Int32     'the Player sending the items
    Public lDelivery As Int32           'seconds until the delivery is expected
    Public lStartedOn As Int32          'date/time the delivery was started
    Public lColonyID As Int32

    Public mdtReceived As Date = Date.MinValue

    Private Structure uTradeDeliveryItem
        Public lID1 As Int32                'typically the object ID
        Public iID2 As Int16                'typically the object type id
        Public lID3 As Int32                'for example, OwnerID of a componentcache's component
        Public blQty As Int64               'the quantity being purchased
    End Structure

    Private muItems() As uTradeDeliveryItem
    Private mlItemUB As Int32 = -1

    Public Sub ClearItems()
        mlItemUB = -1
        Erase muItems
    End Sub

    Public Function ItemContainsIntel() As Boolean
        For X As Int32 = 0 To mlItemUB
            If muItems(X).iID2 = ObjectType.ePlayerIntel OrElse muItems(X).iID2 = ObjectType.ePlayerItemIntel OrElse muItems(X).iID2 = ObjectType.ePlayerTechKnowledge Then
                Return True
            End If
        Next X
        Return False
    End Function

    Public Sub AddItem(ByVal lID1 As Int32, ByVal iID2 As Int16, ByVal lID3 As Int32, ByVal blQty As Int64)
        For X As Int32 = 0 To mlItemUB
            If muItems(X).lID1 = lID1 AndAlso muItems(X).iID2 = iID2 AndAlso muItems(X).lID3 = lID3 Then
                muItems(X).blQty = blQty
                Exit Sub
            End If
        Next X

        mlItemUB += 1
        ReDim Preserve muItems(mlItemUB)
        With muItems(mlItemUB)
            .lID1 = lID1
            .iID2 = iID2
            .lID3 = lID3
            .blQty = blQty
        End With
    End Sub

    Public Function GetListBoxText() As String
        Dim sPlayer As String = GetCacheObjectValue(lSourcePlayerID, ObjectType.ePlayer)
        Dim lDifference As Int32 = CInt(Now.Subtract(mdtReceived).TotalSeconds)

        For x As Int32 = 0 To mlItemUB
            If muItems(x).iID2 = ObjectType.ePlayerIntel OrElse muItems(x).iID2 = ObjectType.ePlayerItemIntel Then
                sPlayer = "Anonymous"
                Exit For
            End If
        Next

        Dim lTemp As Int32 = lDelivery - lDifference
        Dim sETA As String = "Delivery Imminent..."
        If lTemp > 0 Then
            sETA = GetDurationFromSeconds(lTemp, True)
        ElseIf lTemp < -5 Then
            sETA = "Delivered"
        End If
        Return sPlayer.PadRight(22, " "c) & sETA
    End Function

    Public Shared oTradeDeliveries() As TradeDeliveryPackage
    Public Shared yTradeDeliveryUsed() As Byte
    Public Shared lTradeDeliveryUB As Int32 = -1
    Public Shared Function AddOrGetTradeDelivery(ByVal lTPID As Int32, ByVal lSourceID As Int32, ByVal lDeliverVal As Int32, ByVal lStartVal As Int32) As TradeDeliveryPackage
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lTradeDeliveryUB
            If yTradeDeliveryUsed(X) <> 0 Then
                With oTradeDeliveries(X)
					If .lTradePostID = lTPID AndAlso .lSourcePlayerID = lSourceID AndAlso .lDelivery = lDeliverVal AndAlso .lStartedOn = lStartVal Then Return oTradeDeliveries(X)
                End With
            ElseIf lIdx = -1 Then
                lIdx = X
            End If
        Next X
        If lIdx = -1 Then
            ReDim Preserve yTradeDeliveryUsed(lTradeDeliveryUB + 1)
            ReDim Preserve oTradeDeliveries(lTradeDeliveryUB + 1)
            lTradeDeliveryUB += 1
            lIdx = lTradeDeliveryUB
        End If
        oTradeDeliveries(lIdx) = New TradeDeliveryPackage()
        With oTradeDeliveries(lIdx)
            .lTradePostID = lTPID
            .lSourcePlayerID = lSourceID
            .lDelivery = lDeliverVal
            .lStartedOn = lStartVal
            .ClearItems()
            .mdtReceived = Now
        End With
        yTradeDeliveryUsed(lIdx) = 255
        Return oTradeDeliveries(lIdx)
    End Function

    Public Function GetDeliveryText() As String
        If mdtReceived = Date.MinValue Then Return ""

        Dim lSecondsMod As Int32 = Math.Abs(CInt(Now.Subtract(mdtReceived).TotalSeconds))

        lSecondsMod = lDelivery - lSecondsMod
        If lSecondsMod < 5 Then
            Return "Delivered"
        ElseIf lSecondsMod < 0 Then
            Return "Imminent"
        End If

        Return GetDurationFromSeconds(lSecondsMod, True)
    End Function

    Public Function GetDestinationText() As String
        Dim sFacility As String = GetCacheObjectValue(lTradePostID, ObjectType.eFacility)
        Dim sColony As String = GetCacheObjectValue(lColonyID, ObjectType.eColony)
        Return sFacility & " (" & sColony & ")"
    End Function

    Public Sub SmartFillListBox(ByRef lstData As UIListBox)
        For X As Int32 = 0 To mlItemUB
            Dim iTempID As Int16 = muItems(X).iID2
            If iTempID = ObjectType.eMineralCache Then iTempID = ObjectType.eMineral
            Dim sTemp As String = ""
            If iTempID = ObjectType.ePlayerItemIntel OrElse iTempID = ObjectType.ePlayerTechKnowledge Then
                sTemp = "Intelligence"
            Else
                sTemp = GetCacheObjectValue(muItems(X).lID1, Math.Abs(iTempID))
                If iTempID <> ObjectType.eUnit Then sTemp &= " (" & muItems(X).blQty.ToString("#,##0") & ")"
            End If
            Dim bFound As Boolean = False

            For Y As Int32 = 0 To lstData.ListCount - 1
                If lstData.ItemData(Y) = muItems(X).lID1 AndAlso lstData.ItemData2(Y) = muItems(X).iID2 Then
                    If lstData.List(Y) <> sTemp Then lstData.List(Y) = sTemp
                    bFound = True
                    Exit For
                End If
            Next Y
            If bFound = False Then
                lstData.AddItem(sTemp)
                lstData.ItemData(lstData.NewIndex) = muItems(X).lID1
                lstData.ItemData2(lstData.NewIndex) = muItems(X).iID2
            End If
        Next X
    End Sub
End Class
