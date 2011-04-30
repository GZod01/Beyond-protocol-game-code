Option Strict On

'Contains data for a transport route waypoint
Public Class TransportRoute
    'Public TransportID As Int32 = -1          'which transport we refer to
    Public oTransport As Transport = Nothing
    Public OrderNum As Byte = 0         'Order for which this waypoint is to be executed

    Public DestinationID As Int32
    Public DestinationTypeID As Int16

    Public WaypointFlags As Byte = 0    'currently unused

    Public oActions() As TransportRouteAction
    Public lActionUB As Int32 = -1

    Private mlStoredX As Int32 = 0
    Private mlStoredZ As Int32 = 0
    Private moStoredSys As SolarSystem = Nothing
    Private mblStoredCycles As Int64 = Int64.MinValue
    Public Function GetRouteETA(ByRef oFromSys As SolarSystem, ByVal lFromX As Int32, ByVal lFromZ As Int32) As DateTime

        Dim dtNow As DateTime = DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime

        If mblStoredCycles <> Int64.MinValue Then
            If mlStoredX = lFromX AndAlso mlStoredZ = lFromZ AndAlso moStoredSys Is Nothing = False AndAlso oFromSys Is Nothing = False AndAlso moStoredSys.ObjectID = oFromSys.ObjectID Then
                Return dtNow.AddSeconds(mblStoredCycles)
            End If
        End If

        Dim oUD As Epica_Entity_Def = oTransport.oUnitDef
        If oUD Is Nothing Then Return Date.MaxValue

        Dim lLocX As Int32 = lFromX
        Dim lLocZ As Int32 = lFromZ
        Dim lDestX As Int32 = 0
        Dim lDestZ As Int32 = 0

        'oUD is our unit def and has the maxspeed and maneuver of the unit...
        '1) Get a position of the Colony (where I am now)
        If oFromSys Is Nothing Then Return DateTime.MaxValue

        '2) Get a position of the Destination (where I'm going)
        'Dest can be planet, facility (space station) or Colony
        Dim lDestEnvID As Int32 = DestinationID
        Dim iDestEnvTypeID As Int16 = DestinationTypeID

        If DestinationTypeID = ObjectType.eColony Then
            Dim oCol As Colony = GetEpicaColony(DestinationID)
            If oCol Is Nothing = False Then
                If oCol.Owner.ObjectID = Me.oTransport.OwnerID Then
                    Dim oTmp As Epica_GUID = CType(oCol.ParentObject, Epica_GUID)
                    lDestEnvID = oTmp.ObjectID
                    iDestEnvTypeID = oTmp.ObjTypeID
                End If
            End If
        End If

        Dim oDestSys As SolarSystem = Nothing
        Dim oDestPlanet As Planet = Nothing
        If iDestEnvTypeID = ObjectType.ePlanet Then
            'Ok, not same planet...
            oDestPlanet = GetEpicaPlanet(lDestEnvID)
            If oDestPlanet Is Nothing Then Return Date.MaxValue

            lDestX = oDestPlanet.LocX
            lDestZ = oDestPlanet.LocZ

            If oDestPlanet.ParentSystem Is Nothing Then Return Date.MaxValue
            oDestSys = oDestPlanet.ParentSystem
        ElseIf iDestEnvTypeID = ObjectType.eFacility Then
            'station
            Dim oFac As Facility = GetEpicaFacility(lDestEnvID)
            If oFac Is Nothing = False Then
                Dim oPGuid As Epica_GUID = CType(oFac.ParentObject, Epica_GUID)
                If oPGuid Is Nothing Then Return Date.MaxValue
                If oPGuid.ObjTypeID = ObjectType.ePlanet Then
                    oDestPlanet = CType(oPGuid, Planet)
                    lDestX = oDestPlanet.LocX
                    lDestZ = oDestPlanet.LocZ
                    oDestSys = oDestPlanet.ParentSystem
                ElseIf oPGuid.ObjTypeID = ObjectType.eSolarSystem Then
                    oDestSys = CType(oPGuid, SolarSystem)
                    lDestX = oFac.LocX
                    lDestZ = oFac.LocZ
                End If
            End If
        End If
        If oDestSys Is Nothing Then Return Date.MaxValue

        '3) Calculate path including Wormholes or BG movement
        Dim fTotalDist As Single = 0.0F
        If oDestSys.ObjectID <> oFromSys.ObjectID Then
            'System to system move
            'Call CalculatePath to get shortest distance space-to-space
            fTotalDist = CalculatePath(oFromSys.ObjectID, oDestSys.ObjectID, lLocX, lLocZ)
        End If
        'Within the system - simple, we have already reduced our locs to space coords
        fTotalDist += Distance(lDestX, lDestZ, lLocX, lLocZ)

        '4) Take distance and apply to the Unit Def to determine Cycles of movement
        Dim blCycles As Int64 = CLng(fTotalDist / oUD.MaxSpeed)

        '5) Divide cycles of movement by 30 to get Seconds
        blCycles = Math.Max(150, blCycles)
        blCycles \= 30

        '6) Add those seconds to Now()
        '7) Return result
        'store our result for optimization
        mblStoredCycles = blCycles
        moStoredSys = oFromSys
        mlStoredX = lFromX
        mlStoredZ = lFromZ

        Return dtNow.AddSeconds(blCycles)
    End Function

    Public Function SaveObject(ByVal bInsertOnly As Boolean) As Boolean
        'save the route actions too
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim bDoInsert As Boolean = bInsertOnly
        Try
            If bInsertOnly = False Then
                sSQL = "UPDATE tblTransportRoute SET DestinationID = " & DestinationID & ", destinationtypeid = " & DestinationTypeID & ", WaypointFlags = " & _
                    WaypointFlags & " WHERE TransportID = " & oTransport.TransportID & " AND OrderNum = " & OrderNum
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    bDoInsert = True
                End If
                oComm.Dispose()
                oComm = Nothing
            End If
            If bDoInsert = True Then
                sSQL = "INSERT INTO tblTransportRoute (TransportID, OrderNum, DestinationID, DestinationTypeID, WaypointFlags, TransOwnerID) VALUES (" & oTransport.TransportID & ", " & _
                    OrderNum & ", " & DestinationID & ", " & DestinationTypeID & ", " & WaypointFlags & ", " & oTransport.OwnerID & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Throw New Exception("No Records Affected!")
                End If
                oComm.Dispose()
                oComm = Nothing
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "TransportRoute.SaveObject: " & ex.Message)
            Return False
        End Try

        'Save tblTransportRouteAction
        'TransportID, OrderNum, ActionOrderNum, ActionTypeID, Extended1, Extended2, Extended3
        If bInsertOnly = False Then
            sSQL = "DELETE FROM tblTransportRouteAction WHERE TransportID = " & oTransport.TransportID & " AND OrderNum = " & OrderNum
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing
        End If
        Dim lUB As Int32 = -1
        lUB = Math.Min(lActionUB, oActions.GetUpperBound(0))
        For X As Int32 = 0 To lActionUB
            If oActions(X) Is Nothing = False Then
                With oActions(X)
                    Try
                        sSQL = "INSERT INTO tblTransportRouteAction (TransportID, OrderNum, ActionOrderNum, ActionTypeID, Extended1, Extended2, Extended3, TransOwnerID) VALUES (" & _
                            oTransport.TransportID & ", " & OrderNum & ", " & .ActionOrderNum & ", " & CByte(.ActionTypeID) & ", " & .Extended1 & ", " & _
                            .Extended2 & ", " & .Extended3 & ", " & oTransport.OwnerID & ")"
                        oComm = New OleDb.OleDbCommand(sSQL, goCN)
                        If oComm.ExecuteNonQuery() = 0 Then
                            Throw New Exception("No Records Affected!")
                        End If
                        oComm.Dispose()
                        oComm = Nothing
                    Catch ex As Exception
                        LogEvent(LogEventType.CriticalError, "TransportRoute.SaveObject of TransportRouteAction: " & ex.Message)
                    End Try
                End With
            End If
        Next

        Return True

    End Function


    Private Class WHPath
        Public oSys As SolarSystem
        Public oParent As SolarSystem
        Public oWH As Wormhole
        Public lJumpCnt As Int32
    End Class
    Private Function CalculatePath(ByVal mlFromSys As Int32, ByVal mlToSys As Int32, ByRef lFromX As Int32, ByRef lFromZ As Int32) As Single

        'Now, go and see if we can get from the FROM sys to the TO sys
        If mlFromSys > 0 AndAlso mlToSys > 0 Then
            Dim oPath(glSystemUB) As WHPath

            Dim oToPathObj As WHPath = Nothing
            For X As Int32 = 0 To glSystemUB
                oPath(X) = New WHPath
                With oPath(X)
                    .oSys = goSystem(X)
                    .oParent = Nothing
                    If .oSys.ObjectID = mlToSys Then
                        .lJumpCnt = 0
                        oToPathObj = oPath(X)
                    Else
                        .lJumpCnt = 9999999
                    End If
                End With
            Next X

            Dim colOpen As New Collection
            Dim colClosed As New Collection

            colOpen.Add(oToPathObj, "SYS" & mlToSys)

            Dim bDone As Boolean = False

            While bDone = False
                If colOpen.Count = 0 Then
                    bDone = True
                Else
                    Dim oCur As WHPath = Nothing
                    For Each oTmp As WHPath In colOpen
                        If oTmp Is Nothing = False Then
                            oCur = oTmp
                            Exit For
                        End If
                    Next
                    If oCur Is Nothing = False Then
                        Dim sCurKey As String = "SYS" & oCur.oSys.ObjectID
                        colOpen.Remove(sCurKey)
                        colClosed.Add(oCur, sCurKey)

                        With oCur.oSys

                            If oCur.oSys.ObjectID = mlFromSys Then
                                Dim oCurPath As WHPath = oCur
                                Dim fDist As Single = 0.0F
                                While oCurPath Is Nothing = False
                                    If oCur.oSys.ObjectID = mlToSys Then Exit While

                                    'Get to the wormhole in the current system from our previous position and then set our position to the wormhole on the other side
                                    If oCur.oWH.System1.ObjectID = oCur.oSys.ObjectID Then
                                        'sys 1
                                        fDist += Distance(oCur.oWH.LocX1, oCur.oWH.LocY1, lFromX, lFromZ)
                                        lFromX = oCur.oWH.LocX2
                                        lFromZ = oCur.oWH.LocY2
                                    Else
                                        'sys 2
                                        fDist += Distance(oCur.oWH.LocX2, oCur.oWH.LocY2, lFromX, lFromZ)
                                        lFromX = oCur.oWH.LocX1
                                        lFromZ = oCur.oWH.LocY1
                                    End If

                                    Dim oP As SolarSystem = oCur.oParent
                                    For X As Int32 = 0 To glSystemUB
                                        If goSystem(X).ObjectID = oP.ObjectID Then
                                            oCur = oPath(X)
                                            Exit For
                                        End If
                                    Next X
                                End While
                                Return fDist
                            End If

                            'Now, go through this system's wormholes
                            For X As Int32 = 0 To .mlWormholeUB
                                If .moWormholes(X).System1.ObjectID = .ObjectID Then
                                    If (.moWormholes(X).WormholeFlags And elWormholeFlag.eSystem1Jumpable) = 0 Then Continue For '  If wormhole isnt open continue
                                    If Me.oTransport.Owner.HasWormholeKnowledge(.moWormholes(X).ObjectID) = False Then Continue For

                                    'Ok, check if system2 is in the list
                                    Dim lOtherSysID As Int32 = .moWormholes(X).System2.ObjectID
                                    Dim sKey As String = "SYS" & lOtherSysID
                                    If colOpen.Contains(sKey) = False AndAlso colClosed.Contains(sKey) = False Then
                                        For Y As Int32 = 0 To glSystemUB
                                            If goSystem(Y).ObjectID = lOtherSysID Then
                                                oPath(Y).oParent = .moWormholes(X).System1

                                                Dim lJumpCnt As Int32 = oCur.lJumpCnt
                                                oPath(Y).lJumpCnt = lJumpCnt + 1
                                                oPath(Y).oWH = .moWormholes(X)
                                                colOpen.Add(oPath(Y), sKey)
                                                Exit For
                                            End If
                                        Next Y
                                    End If

                                Else
                                    If (.moWormholes(X).WormholeFlags And elWormholeFlag.eSystem2Jumpable) = 0 Then Continue For '  If wormhole isnt open continue
                                    If Me.oTransport.Owner.HasWormholeKnowledge(.moWormholes(X).ObjectID) = False Then Continue For

                                    'Ok, check if system1 is in the list
                                    Dim lOtherSysID As Int32 = .moWormholes(X).System1.ObjectID
                                    Dim sKey As String = "SYS" & lOtherSysID
                                    If colOpen.Contains(sKey) = False AndAlso colClosed.Contains(sKey) = False Then
                                        For Y As Int32 = 0 To glSystemUB
                                            If goSystem(Y).ObjectID = lOtherSysID Then
                                                oPath(Y).oParent = .moWormholes(X).System2

                                                Dim lJumpCnt As Int32 = oCur.lJumpCnt
                                                oPath(Y).lJumpCnt = lJumpCnt + 1
                                                oPath(Y).oWH = .moWormholes(X)
                                                colOpen.Add(oPath(Y), sKey)
                                                Exit For
                                            End If
                                        Next Y
                                    End If
                                End If
                            Next
                        End With
                    Else
                        bDone = True
                    End If
                End If
            End While

            'Ok, do the standard system to system movement...
            Dim oBGDest As SolarSystem = GetEpicaSystem(mlToSys)
            Dim oBGFrom As SolarSystem = GetEpicaSystem(mlFromSys)
            If oBGDest Is Nothing = False AndAlso oBGFrom Is Nothing = False Then
                Dim fX As Single = (oBGFrom.LocX - oBGDest.LocX)
                Dim fY As Single = (oBGFrom.LocY - oBGDest.LocY)
                Dim fZ As Single = (oBGFrom.LocZ - oBGDest.LocZ)
                fX *= fX : fY *= fY : fZ *= fZ
                Dim fDist As Single = CSng(Math.Sqrt(fX + fY + fZ))
                fDist *= 10000000

                Dim fAngle As Single
                Dim fX1 As Single = oBGDest.LocX
                Dim fY1 As Single = oBGDest.LocZ
                Dim fX2 As Single = oBGFrom.LocX
                Dim fY2 As Single = oBGFrom.LocZ

                'NOTE: This is essentially LineAngleDegrees but this app doesn't have it so I didn't add it, if we use it again, just copy this out of here
                Dim fDX As Single = fX2 - fX1
                Dim fDY As Single = fY2 - fY1
                If fDX = 0 Then 'vertical
                    If fDY < 0 Then
                        fAngle = Math.PI / 2.0F
                    Else : fAngle = Math.PI * 1.5F
                    End If
                ElseIf fDY = 0 Then 'horizontal
                    If fDX < 0 Then
                        fAngle = CSng(Math.PI)
                    Else : fAngle = 0.0F
                    End If
                Else
                    fAngle = CSng(Math.Atan(Math.Abs(fDY / fDX)))
                    If fDX > -1 AndAlso fDY > -1 Then
                        fAngle = CSng((Math.PI) * 2.0F) - fAngle
                    ElseIf fDX < 0 AndAlso fDY > -1 Then
                        fAngle = CSng(Math.PI) + fAngle
                    ElseIf fDX < 0 AndAlso fDY < 0 Then
                        fAngle = CSng(Math.PI) - fAngle
                    End If
                End If
                'Adjust the angle to degrees
                fAngle *= 57.2957764F           'gdDegreePerRad
                'Adjust for the weird stuff up above
                fAngle = 360.0F - fAngle
                'End of Line Angle Degrees

                'Now, we put the values in from Dest to Origin, so the angle is already good... let's make our values
                'Reuse some older values
                fX1 = 5000000
                fY1 = 0

                'NOTE: This is RotatePoint, but this app doesn't have it so I didn't add it, if we do, just copy this code
                Dim fRads As Single = fAngle * (CSng(Math.PI) / 180.0F)
                fDX = fX1   ' was originally lEndX - lAxisX
                fDY = fY1   ' was originally lEndY - lAxisY
                'The following two formulae should be lAxisX or lAxisY + the value
                fX1 = ((fDX * CSng(Math.Cos(fRads))) + (fDY * CSng(Math.Sin(fRads))))
                fY1 = -((fDX * CSng(Math.Sin(fRads))) + (fDY * CSng(Math.Cos(fRads))))
                'End of RotatePoint

                'fX1 and fY1 have our entry point X,Z in a circle... we now need to move that point to the edge
                'Gonna reuse X2 and Y2 as the ABSOLUTE values
                fX2 = Math.Abs(fX1) : fY2 = Math.Abs(fY1)
                If fX2 <> 5000000 AndAlso fY2 <> 5000000 Then
                    If fX2 > fY2 Then
                        'make X1 5000000 and adjust y
                        fY1 *= (5000000 / fX2)      'use the absolute value of X here
                        If fX1 < 0 Then fX1 = -5000000 Else fX1 = 5000000
                    Else
                        'Make Y1 5000000 and adjust x
                        fX1 *= (5000000 / fY2)      'use the absolute value of y here
                        If fY1 < 0 Then fY1 = -5000000 Else fY1 = 5000000
                    End If
                End If

                'Ok, fX1 and fY1 have our final coordinates
                lFromX = CInt(fX1)
                lFromZ = CInt(fY1)

                Return fDist
            End If
        End If

        'if we're here...
        Return Single.MaxValue
    End Function

End Class
