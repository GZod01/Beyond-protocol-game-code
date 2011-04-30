Option Strict On

'Actions to be taken at a transportroute waypoint
Public Class TransportRouteAction
    Public Enum TransportRouteActionType As Byte
        eUnload = 1
        eLoad = 2
        eQty = 4            'if this is set, extended3 is treated as an absolute quantity. -1 indicates all available.
        ePercentage = 8     'if this is set, extended3 is treated as a percentage (0 to 100). -1 or anything outside 0 to 100 indicates all available.
    End Enum

    Public oParentRoute As TransportRoute
    Public ActionOrderNum As Byte   'order of which to execute actions

    'BIT WISE
    Public ActionTypeID As TransportRouteActionType

    ''' <summary>
    ''' A specific ObjectID or -1. When -1, ObjectID is not used in comparisons.
    ''' </summary>
    ''' <remarks></remarks>
    Public Extended1 As Int32
    ''' <summary>
    ''' A specific ObjectType ID.
    ''' </summary>
    ''' <remarks></remarks>
    Public Extended2 As Int16
    ''' <summary>
    ''' A quantity identifier or -1. If Quantity is considered to be ALL or 100%.
    ''' </summary>
    ''' <remarks></remarks>
    Public Extended3 As Int32

    Public Sub Execute(ByRef oColony As Colony)
        'Ok, execute this action with the colony passed in
        'get a temp var to our transport
        Dim oTransport As Transport = oParentRoute.oTransport
        If oTransport Is Nothing Then Return
        'get a temp var to our unit def
        Dim oUD As Epica_Entity_Def = oTransport.oUnitDef
        If oUD Is Nothing Then Return

        'Do ActionTypeID on (Extended1,Extended2) by amt Extended3
        If (ActionTypeID And TransportRouteActionType.eLoad) <> 0 Then
            'Get our cargo cap at this time and store it
            Dim lCargoCap As Int32 = oTransport.Cargo_Cap

            'Loading... determine what we are loading...
            Select Case Extended2
                Case ObjectType.eMineral
                    'Minerals... so pulling minerals from the colony to the unit...
                    If Extended1 < 1 Then
                        'All minerals...
                        Dim lUB As Int32 = -1
                        If oColony.mlMineralCacheIdx Is Nothing = False Then lUB = Math.Min(oColony.mlMineralCacheUB, oColony.mlMineralCacheIdx.GetUpperBound(0))
                        If oColony.mlMineralCacheID Is Nothing = False Then lUB = Math.Min(lUB, oColony.mlMineralCacheID.GetUpperBound(0))
                        For X As Int32 = 0 To lUB
                            If oColony.mlMineralCacheIdx(X) > -1 Then
                                If glMineralCacheIdx(oColony.mlMineralCacheIdx(X)) = oColony.mlMineralCacheID(X) Then
                                    Dim oCache As MineralCache = goMineralCache(oColony.mlMineralCacheIdx(X))
                                    If oCache Is Nothing = False Then
                                        Dim lQty As Int32 = GetTransferQuantity(oCache.Quantity)
                                        'do the transfer
                                        lQty = Math.Min(lQty, lCargoCap)
                                        If lQty > 0 Then
                                            lCargoCap -= lQty
                                            oCache.Quantity -= lQty
                                            oTransport.AddCargo(oCache.oMineral.ObjectID, Extended2, oTransport.OwnerID, lQty)

                                            If lCargoCap < 1 Then Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next X
                    Else
                        'a specific mineral
                        Dim lQty As Int32 = GetTransferQuantity(oColony.ColonyMineralQuantity(Extended1))
                        lQty = Math.Min(lQty, lCargoCap)
                        If lQty > 0 Then
                            lCargoCap -= lQty
                            oColony.AdjustColonyMineralCache(Extended1, -lQty)
                            oTransport.AddCargo(Extended1, Extended2, oTransport.OwnerID, lQty)
                        End If
                    End If
                Case ObjectType.eEnlisted
                    Dim lQty As Int32 = GetTransferQuantity(oColony.ColonyEnlisted)
                    lQty = Math.Min(lQty, lCargoCap)
                    If lQty > 0 Then
                        lCargoCap -= lQty
                        oColony.ColonyEnlisted -= lQty
                        oTransport.AddCargo(0, Extended2, oTransport.OwnerID, lQty)
                    End If
                Case ObjectType.eOfficers
                    Dim lQty As Int32 = GetTransferQuantity(oColony.ColonyOfficers)
                    lQty = Math.Min(lQty, lCargoCap)
                    If lQty > 0 Then
                        lCargoCap -= lQty
                        oColony.ColonyOfficers -= lQty
                        oTransport.AddCargo(0, Extended2, oTransport.OwnerID, lQty)
                    End If
                Case ObjectType.eColonists
                    Dim lQty As Int32 = GetTransferQuantity(oColony.Population)
                    lQty = Math.Min(lQty, lCargoCap)
                    If lQty > 0 Then
                        lCargoCap -= lQty
                        oColony.Population -= lQty
                        oTransport.AddCargo(0, Extended2, oTransport.OwnerID, lQty)
                    End If
                Case -1
                    Dim lUB As Int32 = -1
                    If oColony.mlComponentCacheIdx Is Nothing = False Then lUB = Math.Min(oColony.mlComponentCacheUB, oColony.mlComponentCacheIdx.GetUpperBound(0))
                    If oColony.mlComponentCacheID Is Nothing = False Then lUB = Math.Min(lUB, oColony.mlComponentCacheID.GetUpperBound(0))
                    If oColony.mlComponentCacheCompID Is Nothing = False Then lUB = Math.Min(lUB, oColony.mlComponentCacheCompID.GetUpperBound(0))
                    For X As Int32 = 0 To lUB
                        If oColony.mlComponentCacheIdx(X) > -1 Then
                            If glComponentCacheIdx(oColony.mlComponentCacheIdx(X)) = oColony.mlComponentCacheID(X) Then
                                Dim oCache As ComponentCache = goComponentCache(oColony.mlComponentCacheIdx(X))
                                If oCache Is Nothing = False Then
                                    Dim lQty As Int32 = GetTransferQuantity(oCache.Quantity)
                                    lQty = Math.Min(lQty, lCargoCap)
                                    If lQty > 0 Then
                                        lCargoCap -= lQty
                                        oCache.Quantity -= lQty
                                        oTransport.AddCargo(oCache.ComponentID, oCache.ComponentTypeID, oCache.ComponentOwnerID, lQty)
                                        If lCargoCap < 1 Then Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next X

                    lUB = -1
                    If oColony.mlMineralCacheIdx Is Nothing = False Then lUB = Math.Min(oColony.mlMineralCacheUB, oColony.mlMineralCacheIdx.GetUpperBound(0))
                    If oColony.mlMineralCacheID Is Nothing = False Then lUB = Math.Min(lUB, oColony.mlMineralCacheID.GetUpperBound(0))
                    If lCargoCap < 1 Then lUB = -1
                    For X As Int32 = 0 To lUB
                        If oColony.mlMineralCacheIdx(X) > -1 Then
                            If glMineralCacheIdx(oColony.mlMineralCacheIdx(X)) = oColony.mlMineralCacheID(X) Then
                                Dim oCache As MineralCache = goMineralCache(oColony.mlMineralCacheIdx(X))
                                If oCache Is Nothing = False Then
                                    Dim lQty As Int32 = GetTransferQuantity(oCache.Quantity)
                                    'do the transfer
                                    lQty = Math.Min(lQty, lCargoCap)
                                    If lQty > 0 Then
                                        lCargoCap -= lQty
                                        oCache.Quantity -= lQty
                                        oTransport.AddCargo(oCache.oMineral.ObjectID, ObjectType.eMineral, oTransport.OwnerID, lQty)

                                        If lCargoCap < 1 Then Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next X
                Case Else
                    'Likely a component tech typeid
                    If Extended1 < 1 Then
                        'all components of this type (or all components if typeid = componentcache)
                        Dim lUB As Int32 = -1
                        If oColony.mlComponentCacheIdx Is Nothing = False Then lUB = Math.Min(oColony.mlComponentCacheUB, oColony.mlComponentCacheIdx.GetUpperBound(0))
                        If oColony.mlComponentCacheID Is Nothing = False Then lUB = Math.Min(lUB, oColony.mlComponentCacheID.GetUpperBound(0))
                        If oColony.mlComponentCacheCompID Is Nothing = False Then lUB = Math.Min(lUB, oColony.mlComponentCacheCompID.GetUpperBound(0))
                        For X As Int32 = 0 To lUB
                            If oColony.mlComponentCacheIdx(X) > -1 Then
                                If glComponentCacheIdx(oColony.mlComponentCacheIdx(X)) = oColony.mlComponentCacheID(X) Then
                                    Dim oCache As ComponentCache = goComponentCache(oColony.mlComponentCacheIdx(X))
                                    If oCache Is Nothing = False Then
                                        If Extended2 = ObjectType.eComponentCache OrElse oCache.ComponentTypeID = Extended2 Then
                                            Dim lQty As Int32 = GetTransferQuantity(oCache.Quantity)
                                            lQty = Math.Min(lQty, lCargoCap)
                                            If lQty > 0 Then
                                                lCargoCap -= lQty
                                                oCache.Quantity -= lQty
                                                oTransport.AddCargo(oCache.ComponentID, oCache.ComponentTypeID, oCache.ComponentOwnerID, lQty)
                                                If lCargoCap < 1 Then Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next X
                    Else
                        'a specific component
                        Dim oCC As ComponentCache = oColony.ColonyComponentCache(Extended1, Extended2)
                        If oCC Is Nothing = False Then
                            Dim lQty As Int32 = GetTransferQuantity(oCC.Quantity)
                            lQty = Math.Min(lQty, lCargoCap)
                            If lQty > 0 Then
                                lCargoCap -= lQty
                                oCC.Quantity -= lQty
                                oTransport.AddCargo(oCC.ComponentID, oCC.ComponentTypeID, oCC.ComponentOwnerID, lQty)
                            End If
                        End If
                    End If

            End Select


        ElseIf (ActionTypeID And TransportRouteActionType.eUnload) <> 0 Then
            'Unloading .... similar to loading - in fact, I copied and pasted... however, source is transport, dest is colony
            'Get our cargo cap at this time and store it
            Dim lCargoCap As Int32 = oColony.Cargo_Cap

            Select Case Extended2
                Case ObjectType.eMineral
                    'Minerals... so pulling minerals from the unit to the colony...
                    If Extended1 < 1 Then
                        'All minerals...
                        Dim lUB As Int32 = -1
                        If oTransport.oCargo Is Nothing = False Then lUB = Math.Min(oTransport.lCargoUB, oTransport.oCargo.GetUpperBound(0))

                        For X As Int32 = 0 To lUB
                            Dim oTmp As TransportCargo = oTransport.oCargo(X)
                            If oTmp Is Nothing = False Then
                                If oTmp.CargoTypeID = ObjectType.eMineral Then
                                    Dim lQty As Int32 = GetTransferQuantity(oTmp.Quantity)
                                    lQty = Math.Min(lQty, lCargoCap)
                                    If lQty > 0 Then
                                        lCargoCap -= lQty
                                        oTmp.Quantity -= lQty
                                        oTmp.SaveObject(False)
                                        oColony.AdjustColonyMineralCache(oTmp.CargoID, lQty)
                                        If lCargoCap < 1 Then Exit For
                                    End If
                                End If
                            End If
                        Next X
                    Else
                        'a specific mineral
                        Dim oTmp As TransportCargo = oTransport.GetCargo(Extended1, Extended2, oTransport.OwnerID)
                        If oTmp Is Nothing = False Then
                            Dim lQty As Int32 = GetTransferQuantity(oTmp.Quantity)
                            lQty = Math.Min(lQty, lCargoCap)
                            If lQty > 0 Then
                                lCargoCap -= lQty
                                oTmp.Quantity -= lQty
                                oTmp.SaveObject(False)
                                oColony.AdjustColonyMineralCache(Extended1, lQty)
                            End If
                        End If
                    End If
                Case ObjectType.eEnlisted
                    Dim oTmp As TransportCargo = oTransport.GetCargo(Extended1, Extended2, oTransport.OwnerID)
                    If oTmp Is Nothing = False Then
                        Dim lQty As Int32 = GetTransferQuantity(oTmp.Quantity)
                        lQty = Math.Min(lQty, lCargoCap)
                        If lQty > 0 Then
                            lCargoCap -= lQty
                            oTmp.Quantity -= lQty
                            oTmp.SaveObject(False)
                            oColony.ColonyEnlisted += lQty
                        End If
                    End If
                Case ObjectType.eOfficers
                    Dim oTmp As TransportCargo = oTransport.GetCargo(Extended1, Extended2, oTransport.OwnerID)
                    If oTmp Is Nothing = False Then
                        Dim lQty As Int32 = GetTransferQuantity(oTmp.Quantity)
                        lQty = Math.Min(lQty, lCargoCap)
                        If lQty > 0 Then
                            lCargoCap -= lQty
                            oTmp.Quantity -= lQty
                            oTmp.SaveObject(False)
                            oColony.ColonyOfficers += lQty
                        End If
                    End If
                Case ObjectType.eColonists
                    Dim oTmp As TransportCargo = oTransport.GetCargo(Extended1, Extended2, oTransport.OwnerID)
                    If oTmp Is Nothing = False Then
                        Dim lQty As Int32 = GetTransferQuantity(oTmp.Quantity)
                        lQty = Math.Min(lQty, lCargoCap)
                        If lQty > 0 Then
                            lCargoCap -= lQty
                            oTmp.Quantity -= lQty
                            oTmp.SaveObject(False)
                            oColony.Population += lQty
                        End If
                    End If
                Case -1
                    Dim lUB As Int32 = -1
                    If oTransport.oCargo Is Nothing = False Then lUB = Math.Min(oTransport.oCargo.GetUpperBound(0), oTransport.lCargoUB)
                    For X As Int32 = 0 To lUB
                        Dim oTmp As TransportCargo = oTransport.oCargo(X)
                        If oTmp Is Nothing = False Then
                            Dim lQty As Int32 = GetTransferQuantity(oTmp.Quantity)
                            lQty = Math.Min(lQty, lCargoCap)
                            If lQty > 0 Then
                                lCargoCap -= lQty
                                oTmp.Quantity -= lQty
                                oTmp.SaveObject(False)
                                If oTmp.CargoTypeID = ObjectType.eMineral Then
                                    oColony.AdjustColonyMineralCache(oTmp.CargoID, lQty)
                                ElseIf oTmp.CargoTypeID = ObjectType.eEnlisted Then
                                    oColony.ColonyEnlisted += lQty
                                ElseIf oTmp.CargoTypeID = ObjectType.eOfficers Then
                                    oColony.ColonyOfficers += lQty
                                ElseIf oTmp.CargoTypeID = ObjectType.eColonists Then
                                    oColony.Population += lQty
                                Else
                                    oColony.AdjustColonyComponentCache(oTmp.CargoID, oTmp.CargoTypeID, oTmp.OwnerID, lQty)
                                End If

                                If lCargoCap < 1 Then Exit For
                            End If
                        End If
                    Next X
                Case Else
                    'Likely a component tech typeid
                    If Extended1 < 1 Then
                        'all components of this type (or all components if typeid = componentcache)
                        Dim lUB As Int32 = -1
                        If oTransport.oCargo Is Nothing = False Then lUB = Math.Min(oTransport.oCargo.GetUpperBound(0), oTransport.lCargoUB)
                        For X As Int32 = 0 To lUB
                            Dim oTmp As TransportCargo = oTransport.oCargo(X)
                            If oTmp Is Nothing = False Then
                                'Is this a component?
                                If oTmp.CargoTypeID = ObjectType.eArmorTech OrElse oTmp.CargoTypeID = ObjectType.eEngineTech OrElse oTmp.CargoTypeID = ObjectType.eRadarTech OrElse oTmp.CargoTypeID = ObjectType.eShieldTech OrElse oTmp.CargoTypeID = ObjectType.eWeaponTech Then
                                    'yes, ok...
                                    If Extended2 = ObjectType.eComponentCache OrElse oTmp.CargoTypeID = Extended2 Then
                                        Dim lQty As Int32 = GetTransferQuantity(oTmp.Quantity)
                                        lQty = Math.Min(lQty, lCargoCap)
                                        If lQty > 0 Then
                                            lCargoCap -= lQty
                                            oTmp.Quantity -= lQty
                                            oTmp.SaveObject(False)
                                            oColony.AdjustColonyComponentCache(oTmp.CargoID, oTmp.CargoTypeID, oTmp.OwnerID, lQty)
                                            If lCargoCap < 1 Then Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next X
                    Else
                        'a specific component - we have to loop through the array because we do not know the owner of the component
                        Dim lUB As Int32 = -1
                        If oTransport.oCargo Is Nothing = False Then lUB = Math.Min(oTransport.oCargo.GetUpperBound(0), oTransport.lCargoUB)
                        For X As Int32 = 0 To lUB
                            Dim oTmp As TransportCargo = oTransport.oCargo(X)
                            If oTmp Is Nothing = False Then
                                'Is this what we're looking for?
                                If oTmp.CargoTypeID = Extended2 AndAlso oTmp.CargoID = Extended1 Then
                                    Dim lQty As Int32 = GetTransferQuantity(oTmp.Quantity)
                                    lQty = Math.Min(lQty, lCargoCap)
                                    If lQty > 0 Then
                                        lCargoCap -= lQty
                                        oTmp.Quantity -= lQty
                                        oTmp.SaveObject(False)
                                        oColony.AdjustColonyComponentCache(oTmp.CargoID, oTmp.CargoTypeID, oTmp.OwnerID, lQty)
                                    End If
                                    Exit For
                                End If
                            End If
                        Next X
                    End If

            End Select

        End If
    End Sub

    Private Function GetTransferQuantity(ByVal lBaseQty As Int32) As Int32
        Dim lQty As Int32 = 0
        If (ActionTypeID And TransportRouteActionType.ePercentage) <> 0 Then
            lQty = Extended3
            If lQty < 1 OrElse lQty > 100 Then Return lBaseQty
            lQty = CInt(lBaseQty * (lQty / 100))
        Else
            Return Math.Min(lBaseQty, Extended3)
        End If
        Return lQty
    End Function

End Class
