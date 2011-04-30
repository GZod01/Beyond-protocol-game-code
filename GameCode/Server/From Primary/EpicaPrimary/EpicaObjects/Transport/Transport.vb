Option Strict On

'should use the ObjectType of Transport if we can add one
Public Class Transport

    'THIS ENUM IS BITWISE
    Public Enum elTransportFlags As Byte
        eIdle = 1           'sitting on its thumb
        eEnRoute = 2        'moving from point A to point B
        eInTransfer = 4     'Transferring cargo - unit has arrived and is transferring - do not tell it to arrive again
        eBroken = 8         'Transport arrived at a destination but no colony was found, it then returned (
        ePaused = 16        'paused at the current location - pauses cannot happen mid flight
        eLoop = 32          'indicates whether to loop once at the end
    End Enum

    Public TransportID As Int32 = -1      'PK (identity)
    Public UnitName(19) As Byte
    Public OwnerID As Int32
    Private moOwner As Player = Nothing
    Public ReadOnly Property Owner() As Player
        Get
            If moOwner Is Nothing OrElse moOwner.ObjectID <> OwnerID Then
                moOwner = GetEpicaPlayer(OwnerID)
            End If
            Return moOwner
        End Get
    End Property
    Public LocationID As Int32
    Public LocationTypeID As Int16
    Public LocX As Int32
    Public LocZ As Int32
    Public DestinationID As Int32
    Public DestinationTypeID As Int16
    Public ETA As DateTime
    Public DepartureDate As DateTime
    Public UnitDefID As Int32
    Private moUnitDef As Epica_Entity_Def = Nothing
    Public ReadOnly Property oUnitDef() As Epica_Entity_Def
        Get
            If moUnitDef Is Nothing OrElse moUnitDef.ObjectID <> UnitDefID Then
                moUnitDef = GetEpicaUnitDef(UnitDefID)
            End If
            Return moUnitDef
        End Get
    End Property
    Public TransFlags As Byte           'BIT WISE of elTransportFlags

    Public CurrentWaypoint As Byte = 0          'Order number of the current route item that i am EXECUTING
    Public oRoute() As TransportRoute
    Public lRouteUB As Int32 = -1

    Public oCargo() As TransportCargo
    Public lCargoUB As Int32 = -1

    Public Sub TransportArrived()
        'We've arrived at our dest - remove the EnRoute flag
        If (TransFlags And elTransportFlags.eEnRoute) <> 0 Then TransFlags = TransFlags Xor elTransportFlags.eEnRoute
        'clear our ETA
        ETA = DateTime.MinValue
        'have an owner? if not, back out now
        If Me.Owner Is Nothing Then Return

        'We are transferring goods... set our transfer flag
        TransFlags = TransFlags Or elTransportFlags.eInTransfer

        'store our previous loc in case we need to reference it later
        Dim lPrevLocID As Int32 = LocationID
        Dim iPrevLocTypeID As Int16 = LocationTypeID

        'set our new location
        LocationID = DestinationID
        LocationTypeID = DestinationTypeID

        If (TransFlags And Transport.elTransportFlags.eBroken) <> 0 Then
            'Clear our route
            lRouteUB = -1

            'Clear our broken flag
            TransFlags = TransFlags Xor Transport.elTransportFlags.eBroken
            TransFlags = TransFlags Or Transport.elTransportFlags.eIdle
            'Save the Transport Object
            SaveObject()

            Return
        End If

        'OK... LocationTypeID can be a structure (ordering to a space station) or PLANET only or Colony specifically
        Dim oColony As Colony = Nothing
        If LocationTypeID = ObjectType.ePlanet Then
            LocX = 0
            LocZ = 0
            Dim lColonyIdx As Int32 = Owner.GetColonyFromParent(LocationID, LocationTypeID)
            If lColonyIdx > -1 AndAlso lColonyIdx <= glColonyUB Then
                oColony = goColony(lColonyIdx)
                If oColony Is Nothing = False Then
                    If oColony.Owner Is Nothing OrElse oColony.Owner.ObjectID <> Me.Owner.ObjectID Then
                        oColony = Nothing
                    Else
                        Dim oGuid As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                        If oGuid Is Nothing OrElse oGuid.ObjectID <> LocationID OrElse oGuid.ObjTypeID <> LocationTypeID Then
                            oColony = Nothing
                        End If
                    End If
                End If
            End If
        ElseIf LocationTypeID = ObjectType.eFacility Then
            Dim oFac As Facility = GetEpicaFacility(LocationID)
            If oFac Is Nothing = False Then
                Dim oRoot As Epica_GUID = oFac.GetRootParentEnvir()
                If oRoot Is Nothing = False Then
                    oColony = oFac.ParentColony
                    LocationID = oRoot.ObjectID
                    LocationTypeID = oRoot.ObjTypeID
                    LocX = oFac.LocX
                    LocZ = oFac.LocZ
                End If
            End If
        ElseIf LocationTypeID = ObjectType.eColony Then
            oColony = GetEpicaColony(LocationID)
            If oColony Is Nothing = False Then
                Dim oGuid As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                If oGuid Is Nothing Then
                    oColony = Nothing
                Else
                    If oGuid.ObjTypeID = ObjectType.eFacility Then
                        Dim lNewX As Int32 = CType(oGuid, Facility).LocX
                        Dim lNewZ As Int32 = CType(oGuid, Facility).LocZ
                        oGuid = CType(oGuid, Facility).GetRootParentEnvir()
                        If oGuid Is Nothing = False Then
                            LocationID = oGuid.ObjectID
                            LocationTypeID = oGuid.ObjTypeID
                            LocX = lNewX
                            LocZ = lNewZ
                        Else
                            oColony = Nothing
                        End If
                    ElseIf oGuid.ObjTypeID = ObjectType.ePlanet Then
                        LocationID = oGuid.ObjectID : LocationTypeID = oGuid.ObjTypeID
                        LocX = 0
                        LocZ = 0
                    Else
                        oColony = Nothing
                    End If
                End If
            End If
        End If

        If oColony Is Nothing = False Then
            'Ok, got a colony... execute our actions
            If oRoute Is Nothing = False AndAlso CurrentWaypoint <= oRoute.GetUpperBound(0) Then
                With oRoute(CurrentWaypoint)
                    For lAction As Int32 = 0 To .lActionUB
                        .oActions(lAction).Execute(oColony)
                    Next lAction
                End With
            End If

            'Now, are we paused?
            If (TransFlags And elTransportFlags.ePaused) <> 0 Then
                'Yes, so, set us to idle as well
                TransFlags = TransFlags Or elTransportFlags.eIdle
            Else
                'now, increment our waypoint
                Dim lTemp As Int32 = CurrentWaypoint
                lTemp += 1

                'Check for looping
                If lTemp > lRouteUB Then
                    lTemp = 0
                    If (TransFlags And elTransportFlags.eLoop) = 0 Then
                        'Ok, we're done here...
                        TransFlags = TransFlags Or elTransportFlags.eIdle
                    End If
                End If
                'set our current waypiont
                CurrentWaypoint = CByte(lTemp)
            End If
            'Now, are we not idle?
            If oRoute Is Nothing Then TransFlags = elTransportFlags.eIdle
            If (TransFlags And elTransportFlags.eIdle) = 0 Then
                'We are not idle - move to the next waypoint
                With oRoute(CurrentWaypoint)
                    DestinationID = .DestinationID
                    DestinationTypeID = .DestinationTypeID

                    Dim oFromSys As SolarSystem = Nothing
                    Dim lFromX As Int32 = LocX
                    Dim lFromZ As Int32 = LocZ
                    If oColony Is Nothing = False Then
                        Dim oPGuid As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                        If oPGuid Is Nothing = False Then
                            If oPGuid.ObjTypeID = ObjectType.eFacility Then
                                With CType(oPGuid, Facility)
                                    lFromX = .LocX
                                    lFromZ = .LocZ
                                End With
                                oPGuid = CType(CType(oPGuid, Facility).ParentObject, Epica_GUID)
                                If oPGuid.ObjTypeID = ObjectType.eSolarSystem Then
                                    oFromSys = CType(oPGuid, SolarSystem)
                                End If
                            ElseIf oPGuid.ObjTypeID = ObjectType.ePlanet Then
                                With CType(oPGuid, Planet)
                                    lFromX = .LocX
                                    lFromZ = .LocZ
                                    oFromSys = .ParentSystem
                                End With
                            End If
                        End If
                    End If

                    ETA = .GetRouteETA(oFromSys, lFromX, lFromZ)
                    DepartureDate = DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime
                    'Set our en route flag
                    TransFlags = TransFlags Or elTransportFlags.eEnRoute
                    'Now, instruct the loop to recalculate transport events
                    Me.Owner.ClearNextTransportEvent()
                End With
            End If
        Else
            'return us to the origin
            DestinationID = lPrevLocID
            DestinationTypeID = iPrevLocTypeID

            Dim dtNow As DateTime = DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime
            Dim oTS As TimeSpan = dtNow.Subtract(DepartureDate)
            ETA = dtNow.Add(oTS)
            DepartureDate = dtNow

            'should bit-wise with broken
            TransFlags = TransFlags Or elTransportFlags.eEnRoute Or elTransportFlags.eBroken
        End If

        'No longer in the Transfer process...
        If (TransFlags And elTransportFlags.eInTransfer) <> 0 Then TransFlags = TransFlags Xor elTransportFlags.eInTransfer
    End Sub

    Public ReadOnly Property Cargo_Cap() As Int32
        Get

            If oUnitDef Is Nothing Then Return 0
            Dim lCap As Int32 = oUnitDef.Cargo_Cap

            Dim lUB As Int32 = -1
            If oCargo Is Nothing = False Then lUB = Math.Min(lCargoUB, oCargo.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If oCargo(X) Is Nothing = False Then
                    If oCargo(X).CargoTypeID = ObjectType.eUnit Then
                        Dim oUnit As Unit = oCargo(X).oUnit
                        If oUnit Is Nothing = False Then
                            lCap -= oUnit.EntityDef.HullSize
                        End If
                    Else
                        lCap -= oCargo(X).Quantity
                    End If
                End If
            Next X
            Return lCap
        End Get
    End Property

    Public Sub AddCargo(ByVal lCargoID As Int32, ByVal iCargoTypeID As Int16, ByVal lOwnerID As Int32, ByVal lQty As Int32)
        Dim lUB As Int32 = -1
        If oCargo Is Nothing = False Then lUB = Math.Min(lCargoUB, oCargo.GetUpperBound(0))

        If iCargoTypeID <> ObjectType.eUnit Then
            For X As Int32 = 0 To lUB
                If oCargo(X) Is Nothing = False Then
                    If oCargo(X).CargoID = lCargoID AndAlso oCargo(X).CargoTypeID = iCargoTypeID AndAlso oCargo(X).OwnerID = lOwnerID Then
                        oCargo(X).Quantity += lQty
                        oCargo(X).SaveObject(False)
                        Return
                    End If
                End If
            Next X
        End If

        Dim oNew As New TransportCargo
        With oNew
            .CargoID = lCargoID
            .CargoTypeID = iCargoTypeID
            .OwnerID = lOwnerID
            .oParentTransport = Me
            .Quantity = lQty
            .SaveObject(True)
        End With

        For X As Int32 = 0 To lUB
            If oCargo(X) Is Nothing Then
                oCargo(X) = oNew
                Return
            End If
        Next X
        lUB += 1
        ReDim Preserve oCargo(lUB)
        oCargo(lUB) = oNew
        lCargoUB = lUB
    End Sub

    Public Function GetCargo(ByVal lCargoID As Int32, ByVal iCargoTypeID As Int16, ByVal lOwnerID As Int32) As TransportCargo
        Dim lUB As Int32 = -1
        If oCargo Is Nothing = False Then lUB = Math.Min(lCargoUB, oCargo.GetUpperBound(0))
        For X As Int32 = 0 To lUB
            If oCargo(X) Is Nothing = False Then
                With oCargo(X)
                    If .CargoID = lCargoID AndAlso .CargoTypeID = iCargoTypeID AndAlso .OwnerID = lOwnerID Then
                        Return oCargo(X)
                    End If
                End With
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetSaveObjectText() As String
        If TransportID < 1 Then
            SaveObject()
            Return ""
        End If

        Try
            Dim blETA As Int64 = 0
            Dim blDepartureDate As Int64 = 0
            If ETA <> DateTime.MinValue Then blETA = CLng(ETA.ToString("yyMMddHHmmss"))
            If DepartureDate <> DateTime.MinValue Then blDepartureDate = CLng(DepartureDate.ToString("yyMMddHHmmss"))

            Dim sSQL As String = "UPDATE tblTransport SET UnitName = '" & MakeDBStr(BytesToString(UnitName)) & "', OwnerID = " & OwnerID & ", LocationID = " & _
                    LocationID & ", LocationTypeID = " & LocationTypeID & ", DestinationID = " & DestinationID & ", DestinationTypeID = " & _
                    DestinationTypeID & ", ETA = " & blETA.ToString() & ", DepartureDate = " & blDepartureDate.ToString & _
                    ", UnitDefID = " & UnitDefID & ", RU_Flags = " & TransFlags & ", CurrentWaypoint = " & CurrentWaypoint & ", LocX = " & LocX & _
                    ", LocZ = " & LocZ & " WHERE TransportID = " & TransportID
            Return sSQL
        Catch
            SaveObject()
            Return ""
        End Try
    End Function

    Public Function SaveObject() As Boolean

        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand = Nothing

        If TransportID > 0 Then
            Try
                sSQL = "DELETE FROM tblTransportRoute WHERE TransportID = " & TransportID & vbCrLf
                sSQL &= "DELETE FROM tblTransportCargo WHERE TransportID = " & TransportID & vbCrLf
                sSQL &= "DELETE FROM tblTransportRouteAction WHERE TransportID = " & TransportID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm.Dispose()
                oComm = Nothing
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Transport.SaveObject.ClearAll(): " & ex.Message)
            End Try
        End If

        Try
            Dim blETA As Int64 = 0
            Dim blDepartureDate As Int64 = 0
            If ETA <> DateTime.MinValue Then blETA = CLng(ETA.ToString("yyMMddHHmmss"))
            If DepartureDate <> DateTime.MinValue Then blDepartureDate = CLng(DepartureDate.ToString("yyMMddHHmmss"))

            'Save the transport
            If TransportID > 0 Then
                sSQL = "UPDATE tblTransport SET UnitName = '" & MakeDBStr(BytesToString(UnitName)) & "', OwnerID = " & OwnerID & ", LocationID = " & _
                    LocationID & ", LocationTypeID = " & LocationTypeID & ", DestinationID = " & DestinationID & ", DestinationTypeID = " & _
                    DestinationTypeID & ", ETA = " & blETA.ToString() & ", DepartureDate = " & blDepartureDate.ToString & _
                    ", UnitDefID = " & UnitDefID & ", RU_Flags = " & TransFlags & ", CurrentWaypoint = " & CurrentWaypoint & ", LocX = " & LocX & _
                    ", LocZ = " & LocZ & " WHERE TransportID = " & TransportID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Throw New Exception("No Records Affected on Update!")
                End If
                oComm.Dispose()
                oComm = Nothing
            Else
                sSQL = "INSERT INTO tblTransport (UnitName, OwnerID, LocationID, LocationTypeID, DestinationID, DestinationTypeID, ETA, DepartureDate, " & _
                    "UnitDefID, RU_Flags, CurrentWaypoint, LocX, LocZ) VALUES ('" & MakeDBStr(BytesToString(UnitName)) & "', " & OwnerID & ", " & _
                    LocationID & ", " & LocationTypeID & ", " & DestinationID & ", " & DestinationTypeID & ", " & blETA.ToString & ", " & _
                    blDepartureDate.ToString() & ", " & UnitDefID & ", " & TransFlags & ", " & CurrentWaypoint & ", " & LocX & ", " & LocZ & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Throw New Exception("No Records Affected on Update!")
                End If
                oComm.Dispose()
                oComm = Nothing

                oComm = New OleDb.OleDbCommand("SELECT Max(TransportID) FROM tblTransport WHERE OwnerID = " & OwnerID & " AND UnitName = '" & MakeDBStr(BytesToString(UnitName)) & "'", goCN)
                Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader()
                If oData.Read = True Then
                    TransportID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
                oComm.Dispose()
                oComm = Nothing
            End If

            'Save tblTransportCargo
            Dim lUB As Int32 = -1
            If oCargo Is Nothing = False Then lUB = Math.Min(lCargoUB, oCargo.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If oCargo(X) Is Nothing = False Then
                    With oCargo(X)
                        .SaveObject(True)
                    End With
                End If
            Next X

            'Save tblTransportRoute
            lUB = -1
            If oRoute Is Nothing = False Then lUB = Math.Min(lRouteUB, oRoute.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If oRoute(X) Is Nothing = False Then
                    With oRoute(X)
                        .SaveObject(True)
                    End With
                End If
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Transport.SaveObject: " & ex.Message)
            Return False
        End Try

        Return True
    End Function

    Public Function HandleRequestTransportDetails() As Byte()

        Dim lLen As Int32 = 36
        Dim lRCnt As Int32 = 0
        Dim lCCnt As Int32 = 0
        For X As Int32 = 0 To lRouteUB
            If oRoute(X) Is Nothing = False Then
                lRCnt += 1
                lLen += 7
                For Y As Int32 = 0 To oRoute(X).lActionUB
                    If oRoute(X).oActions(Y) Is Nothing = False Then
                        lLen += 11
                    End If
                Next Y
            End If
        Next X
        For X As Int32 = 0 To lCargoUB
            If oCargo(X) Is Nothing = False AndAlso oCargo(X).Quantity > 0 Then
                lCCnt += 1
                lLen += 14
            End If
        Next X

        Dim yMsg(lLen - 1) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestTransportDetails).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(Me.TransportID).CopyTo(yMsg, lPos) : lPos += 4

        With oUnitDef
            System.BitConverter.GetBytes(.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = Me.oUnitDef.MaxSpeed : lPos += 1
            yMsg(lPos) = Me.oUnitDef.Maneuver : lPos += 1
            System.BitConverter.GetBytes(Me.Cargo_Cap).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(.Cargo_Cap).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(.ModelID).CopyTo(yMsg, lPos) : lPos += 2
        End With

        Dim lSecs As Int32 = 0
        If (TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 AndAlso ETA <> DateTime.MinValue Then
            lSecs = CInt(ETA.Subtract(DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime).TotalSeconds)
            If lSecs < 0 Then lSecs = 0
        End If
        System.BitConverter.GetBytes(lSecs).CopyTo(yMsg, lPos) : lPos += 4

        If (TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
            System.BitConverter.GetBytes(DestinationID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(DestinationTypeID).CopyTo(yMsg, lPos) : lPos += 2
        Else
            System.BitConverter.GetBytes(LocationID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(LocationTypeID).CopyTo(yMsg, lPos) : lPos += 2
        End If
        yMsg(lPos) = TransFlags : lPos += 1

        yMsg(lPos) = CByte(lRCnt) : lPos += 1
        'Now, the destinations...
        For X As Int32 = 0 To lRouteUB
            If oRoute(X) Is Nothing = False Then
                Dim lACnt As Int32 = 0
                Dim lACntPos As Int32 = 0
                With oRoute(X)
                    System.BitConverter.GetBytes(.DestinationID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.DestinationTypeID).CopyTo(yMsg, lPos) : lPos += 2
                    lACntPos = lPos : lPos += 1
                End With
                For Y As Int32 = 0 To oRoute(X).lActionUB
                    If oRoute(X).oActions(Y) Is Nothing = False Then
                        lACnt += 1
                        With oRoute(X).oActions(Y)
                            yMsg(lPos) = .ActionTypeID : lPos += 1
                            System.BitConverter.GetBytes(.Extended1).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.Extended2).CopyTo(yMsg, lPos) : lPos += 2
                            System.BitConverter.GetBytes(.Extended3).CopyTo(yMsg, lPos) : lPos += 4
                        End With
                    End If
                Next Y
                yMsg(lACntPos) = CByte(lACnt)
            End If
        Next X

        'Now, for the cargo
        System.BitConverter.GetBytes(CShort(lCCnt)).CopyTo(yMsg, lPos) : lPos += 2
        For X As Int32 = 0 To lCargoUB
            If oCargo(X) Is Nothing = False AndAlso oCargo(X).Quantity > 0 Then
                With oCargo(X)
                    System.BitConverter.GetBytes(.CargoID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.CargoTypeID).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(.OwnerID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.Quantity).CopyTo(yMsg, lPos) : lPos += 4
                End With
            End If
        Next X

        Return yMsg
    End Function

    Public Sub DeleteMe()
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand = Nothing

        If TransportID > 0 Then
            Try
                sSQL = "DELETE FROM tblTransportRoute WHERE TransportID = " & TransportID & vbCrLf
                sSQL &= "DELETE FROM tblTransportCargo WHERE TransportID = " & TransportID & vbCrLf
                sSQL &= "DELETE FROM tblTransportRouteAction WHERE TransportID = " & TransportID & vbCrLf
                sSQL &= "DELETE FROM tblTransport WHERE TransportID = " & TransportID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm.Dispose()
                oComm = Nothing
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Transport.DeleteMe(): " & ex.Message)
            End Try
        End If
    End Sub

    Public Sub HandleRecall()
        If (TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 AndAlso (TransFlags And Transport.elTransportFlags.ePaused) = 0 Then
            TransFlags = TransFlags Or Transport.elTransportFlags.ePaused
            DestinationID = LocationID
            DestinationTypeID = LocationTypeID
            Dim lTemp As Int32 = CurrentWaypoint
            lTemp -= 1
            If lTemp < 0 Then lTemp = 0
            CurrentWaypoint = CByte(lTemp)
            Dim dtNow As DateTime = DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime
            Dim oTS As TimeSpan = dtNow.Subtract(DepartureDate)
            ETA = dtNow.Add(oTS)
            DepartureDate = dtNow
        End If
    End Sub
 
    Public Sub HandleBegin()
        If (TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then Return
        'If (TransFlags And Transport.elTransportFlags.eInTransfer) <> 0 Then Return
        If (TransFlags And Transport.elTransportFlags.ePaused) <> 0 Then TransFlags = TransFlags Xor Transport.elTransportFlags.ePaused
        If (TransFlags And Transport.elTransportFlags.eIdle) <> 0 Then TransFlags = TransFlags Xor Transport.elTransportFlags.eIdle

        '1) Now, verify our route
        'If ValidateRoute(oRoute) = False Then Return

        CurrentWaypoint = 0
        If lRouteUB > -1 Then
            '2) Need to determine our first waypoint's ETA from where I am now...
            Dim oFromSys As SolarSystem = Nothing
            Dim lFromX As Int32 = LocX
            Dim lFromZ As Int32 = LocZ
            If LocationTypeID = ObjectType.ePlanet Then
                Dim oPlanet As Planet = GetEpicaPlanet(LocationID)
                If oPlanet Is Nothing Then Return
                oFromSys = oPlanet.ParentSystem
                lFromX = oPlanet.LocX
                lFromZ = oPlanet.LocZ
            ElseIf LocationTypeID = ObjectType.eSolarSystem Then
                oFromSys = GetEpicaSystem(LocationID)
            End If

            With oRoute(CurrentWaypoint)
                ETA = .GetRouteETA(oFromSys, lFromX, lFromZ)
                DestinationID = .DestinationID
                DestinationTypeID = .DestinationTypeID

                DepartureDate = DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime
                'Set our en route flag
                TransFlags = TransFlags Or elTransportFlags.eEnRoute
                'Now, instruct the loop to recalculate transport events
                Me.Owner.ClearNextTransportEvent()
            End With
        End If
    End Sub

    Public Sub HandleDeleteItem(ByVal lRouteNum As Int32, ByVal lActionNum As Int32)
        'ok, let's get the item we are moving
        If lRouteNum < 0 Then Return
        If lRouteNum > lRouteUB Then Return

        If oRoute(lRouteNum) Is Nothing = False Then
            If lActionNum > -1 Then
                With oRoute(lRouteNum)
                    If lActionNum <= .lActionUB Then
                        'Delete the action
                        If .oActions(lActionNum) Is Nothing = False Then
                            For X As Int32 = lActionNum To .lActionUB - 1
                                .oActions(X) = .oActions(X + 1)
                            Next X
                            .lActionUB -= 1
                            .SaveObject(False)
                        End If
                    End If
                End With
            Else
                'delete the route item
                For X As Int32 = lRouteNum To lRouteUB - 1
                    oRoute(X) = oRoute(X + 1)
                Next X
                lRouteUB -= 1
                SaveObject()
            End If
        End If
    End Sub

    Public Shared Function CreateFromUnit(ByRef oUnit As Unit) As Transport

        Dim oNew As New Transport
        With oNew
            .TransportID = -1
            ReDim .UnitName(19)
            oUnit.EntityName.CopyTo(.UnitName, 0)
            .OwnerID = oUnit.Owner.ObjectID

            Dim oPGuid As Epica_GUID = CType(oUnit.ParentObject, Epica_GUID)
            If oPGuid Is Nothing Then Return Nothing

            If oPGuid.ObjTypeID = ObjectType.eColony Then
                oPGuid = CType(CType(oPGuid, Colony).ParentObject, Epica_GUID)
                If oPGuid Is Nothing Then Return Nothing
            End If
            If oPGuid.ObjTypeID = ObjectType.eFacility Then
                .LocX = CType(oPGuid, Facility).LocX
                .LocZ = CType(oPGuid, Facility).LocZ

                oPGuid = CType(CType(oPGuid, Facility).ParentObject, Epica_GUID)
                If oPGuid Is Nothing Then Return Nothing
            End If
            If oPGuid.ObjTypeID = ObjectType.eSolarSystem Then
                .LocationID = oPGuid.ObjectID
                .LocationTypeID = oPGuid.ObjTypeID
            ElseIf oPGuid.ObjTypeID = ObjectType.ePlanet Then
                Dim oPlanet As Planet = CType(oPGuid, Planet)
                .LocX = oPlanet.LocX
                .LocZ = oPlanet.LocZ
                .LocationID = oPlanet.ParentSystem.ObjectID
                .LocationTypeID = oPlanet.ParentSystem.ObjTypeID
            End If

            .DestinationID = .LocationID
            .DestinationTypeID = .LocationTypeID
            .ETA = DateTime.MinValue
            .DepartureDate = DateTime.MinValue
            .UnitDefID = oUnit.EntityDef.ObjectID
            .CurrentWaypoint = 0

            .TransFlags = elTransportFlags.eIdle

            .lRouteUB = -1
            .lCargoUB = -1

            If .SaveObject() = False Then Return Nothing
        End With

        Return oNew
    End Function
End Class
