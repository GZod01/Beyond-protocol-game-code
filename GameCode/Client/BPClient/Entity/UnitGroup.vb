Option Strict On

Public Class UnitGroup
    Inherits Base_GUID

    Public sName As String

    Public Structure UnitGroupElement
        Public lUnitID As Int32
        Public lParentID As Int32
        Public iParentTypeID As Int16
    End Structure

    Public uUnitIDs() As UnitGroupElement
    Public lUnitUB As Int32 = -1

    Public lOwnerID As Int32

    Public lParentID As Int32
    Public iParentTypeID As Int16

    Public lLastInterSystemCycleUpdate As Int32
    Public lInterSystemTargetID As Int32 = -1
    Public lInterSystemOriginID As Int32 = -1
    Public yInterSystemMovementSpeed As Byte

    'As reported from the server
    Public lOriginX As Int32 = -1
    Public lOriginY As Int32 = -1
    Public lOriginZ As Int32 = -1

    Public lDefaultFormationID As Int32 = -1

    'Public lWarpointUpkeep As Int32 = -1

    Public LastMessageUpdate As Int32 = 0

#Region "  Reinforcer stuff  "
    Public muReinforcers() As UnitGroupElement
    Public mlReinforcerUB As Int32 = -1
    Public Sub AddReinforcer(ByVal lFacilityID As Int32, ByVal lParentID As Int32, ByVal iParentTypeID As Int16)
        For X As Int32 = 0 To mlReinforcerUB
            If muReinforcers(X).lUnitID = lFacilityID Then Return
        Next X

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlReinforcerUB
            If muReinforcers(X).lUnitID = -1 Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            mlReinforcerUB += 1
            ReDim Preserve muReinforcers(mlReinforcerUB)
            lIdx = mlReinforcerUB
        End If
        With muReinforcers(lIdx)
            .lUnitID = lFacilityID
            .lParentID = lParentID
            .iParentTypeID = iParentTypeID
        End With
    End Sub
    Public Sub RemoveReinforcer(ByVal lFacilityID As Int32)
        For X As Int32 = 0 To mlReinforcerUB
            If muReinforcers(X).lUnitID = lFacilityID Then
                muReinforcers(X).lUnitID = -1
            End If
        Next X
    End Sub
#End Region

    Private mlInterSystemTotalCycles As Int32 = -1
    Public ReadOnly Property InterSystemTotalCycles() As Int32
        Get
            Try
                If mlInterSystemTotalCycles = -1 Then
                    If lInterSystemOriginID = -1 OrElse lInterSystemTargetID = -1 Then Return 0

                    Dim lTX As Int32
                    Dim lTY As Int32
                    Dim lTZ As Int32
                    Dim lTargetID As Int32 = lInterSystemTargetID

                    For lSysIdx As Int32 = 0 To goGalaxy.mlSystemUB
                        With goGalaxy.moSystems(lSysIdx)
                            If .ObjectID = lTargetID Then
                                lTargetID = -1
                                lTX = .LocX : lTY = .LocY : lTZ = .LocZ
                                Exit For
                            End If
                        End With
                    Next lSysIdx

                    'Get distance in Systems between two systems
                    Dim fDX As Single = lOriginX - lTX : fDX *= fDX
                    Dim fDY As Single = lOriginY - lTY : fDY *= fDY
                    Dim fDZ As Single = lOriginZ - lTZ : fDZ *= fDZ
                    Dim fDist As Single = CSng(Math.Sqrt(fDX + fDY + fDZ))

                    'Multiply by 10 million for system size
                    fDist *= 10000000
                    fDist *= 0.5F

                    If yInterSystemMovementSpeed = 0 Then yInterSystemMovementSpeed = 1
                    'TODO: Put player NonGravityWellSpeed Multiplier here
                    Dim fPlayerMult As Single = 1.0F
                    fDist /= (yInterSystemMovementSpeed * fPlayerMult)
                    If fDist > Int32.MaxValue Then mlInterSystemTotalCycles = Int32.MaxValue Else mlInterSystemTotalCycles = CInt(fDist)
                End If
                Return mlInterSystemTotalCycles
            Catch
                Return 0
            End Try
        End Get
    End Property

    Private mlInterSystemMoveCyclesRemaining As Int32
    Public Property lInterSystemMoveCyclesRemaining() As Int32
        Get
            If lLastInterSystemCycleUpdate <> glCurrentCycle Then
                Dim lDiff As Int32 = glCurrentCycle - lLastInterSystemCycleUpdate
                mlInterSystemMoveCyclesRemaining -= lDiff
                lLastInterSystemCycleUpdate = glCurrentCycle
            End If
            Return mlInterSystemMoveCyclesRemaining
        End Get
        Set(ByVal value As Int32)
            mlInterSystemMoveCyclesRemaining = value
            lLastInterSystemCycleUpdate = glCurrentCycle
        End Set
    End Property

    Public Function GetInterSystemMovementETA() As String
        If lInterSystemTargetID = -1 OrElse lInterSystemOriginID = -1 Then Return ""

        Dim lSeconds As Int32 = CInt(Math.Ceiling(lInterSystemMoveCyclesRemaining / 33.3333F))
        Dim lMinutes As Int32 = lSeconds \ 60
        lSeconds -= (lMinutes * 60)
        Dim lHours As Int32 = lMinutes \ 60
        lMinutes -= (lHours * 60)
        Dim lDays As Int32 = lHours \ 24
        lHours -= (lDays * 24)

        Dim sreturn As String = ""
        Dim bForce As Boolean
        If lDays > 0 Then
            sreturn &= lDays.ToString("0#") & ":"
            bForce = True
        End If
        If lHours > 0 OrElse bForce = True Then
            sreturn &= lHours.ToString("0#") & ":"
            bForce = True
        End If
        If lMinutes > 0 OrElse bForce = True Then
            sreturn &= lMinutes.ToString("0#") & ":"
            bForce = True
        End If
        If lSeconds > 0 OrElse bForce = True Then
            sreturn &= lSeconds.ToString("0#")
        End If

        Return sreturn
        'Return lDays.ToString("0#") & ":" & lHours.ToString("0#") & ":" & lMinutes.ToString("0#") & ":" & lSeconds.ToString("0#")
    End Function

    Public Sub FillFromMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2

        With Me
            'My GUID
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            'Owner
            .lOwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'My name
            .sName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            'Parent
            .lParentID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .iParentTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            'Inter System Movement information
            mlInterSystemMoveCyclesRemaining = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lLastInterSystemCycleUpdate = glCurrentCycle
            .lInterSystemTargetID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lInterSystemOriginID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .yInterSystemMovementSpeed = yData(lPos) : lPos += 1
            .lOriginX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lOriginY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lOriginZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lDefaultFormationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            '.lWarpointUpkeep = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Erase uUnitIDs
            lUnitUB = lCnt - 1

            ReDim uUnitIDs(lUnitUB)

            For X As Int32 = 0 To lUnitUB
                With uUnitIDs(X)
                    .lUnitID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lParentID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .iParentTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                End With
            Next X

            mlReinforcerUB = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            mlReinforcerUB -= 1
            ReDim muReinforcers(mlReinforcerUB)
            For X As Int32 = 0 To mlReinforcerUB
                With muReinforcers(X)
                    .lUnitID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lParentID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .iParentTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                End With
            Next X

        End With
        mlInterSystemTotalCycles = -1
        LastMessageUpdate = glCurrentCycle
    End Sub

    Public Sub AddUnit(ByVal lUnitID As Int32, ByVal lParentID As Int32, ByVal iParentTypeID As Int16)
        Dim lIdx As Int32 = -1

        For X As Int32 = 0 To lUnitUB
            If uUnitIDs(X).lUnitID = lUnitID Then
                uUnitIDs(X).lParentID = lParentID
                uUnitIDs(X).iParentTypeID = iParentTypeID
                Return
            ElseIf lIdx = -1 AndAlso uUnitIDs(X).lUnitID = -1 Then
                lIdx = X
            End If
        Next X

        If lIdx = -1 Then
            lUnitUB += 1
            ReDim Preserve uUnitIDs(lUnitUB)
            lIdx = lUnitUB
        End If

        With uUnitIDs(lIdx)
            .lUnitID = lUnitID
            .lParentID = lParentID
            .iParentTypeID = iParentTypeID
        End With

        LastMessageUpdate = glCurrentCycle
    End Sub

    Public Sub RemoveUnit(ByVal lUnitID As Int32)
        For X As Int32 = 0 To lUnitUB
            If uUnitIDs(X).lUnitID = lUnitID Then
                uUnitIDs(X).lUnitID = -1
                LastMessageUpdate = glCurrentCycle
                Return
            End If
        Next X
    End Sub

    Public Function GetBoundingSphere() As BoundingSphere
        Dim lOX As Int32
        Dim lOY As Int32
        Dim lOZ As Int32
        Dim lTX As Int32
        Dim lTY As Int32
        Dim lTZ As Int32
        Dim lOriginID As Int32 = lInterSystemOriginID
        Dim lTargetID As Int32 = lInterSystemTargetID

        For lSysIdx As Int32 = 0 To goGalaxy.mlSystemUB
            With goGalaxy.moSystems(lSysIdx)
                If .ObjectID = lOriginID Then
                    lOriginID = -1
                    lOX = .LocX : lOY = .LocY : lOZ = .LocZ
                ElseIf .ObjectID = lTargetID Then
                    lTargetID = -1
                    lTX = .LocX : lTY = .LocY : lTZ = .LocZ
                End If
            End With
            If lOriginID = -1 AndAlso lTargetID = -1 Then Exit For
        Next lSysIdx

        'Now... find out the mult
        Dim fMult As Single = CSng(lInterSystemMoveCyclesRemaining / InterSystemTotalCycles)

        ''Now... find out the mult
        'Dim fMult As Single = CSng(.lInterSystemMoveCyclesRemaining / .InterSystemTotalCycles)
        'If fMult = Single.PositiveInfinity Then fMult = 0
        'Dim vecDiff As Vector3 = Vector3.Subtract(vecFrom, vecTo)
        'vecDiff.Multiply(fMult)
        'vecDiff = vecTo + vecDiff ' vecDiff.Add(vecFrom)


        Dim oResult As BoundingSphere
        oResult.SphereCenter.X = ((lOX - lTX) * fMult) + lTX
        oResult.SphereCenter.Y = ((lOY - lTY) * fMult) + 8 + lTY
        oResult.SphereCenter.Z = ((lOZ - lTZ) * fMult) + lTZ
        oResult.SphereRadius = 16

        Return oResult
    End Function

    Public Sub SetAllChildrenToParent()
        For X As Int32 = 0 To Me.lUnitUB
            Me.uUnitIDs(X).lParentID = Me.lParentID
            Me.uUnitIDs(X).iParentTypeID = Me.iParentTypeID
        Next X
    End Sub

End Class
