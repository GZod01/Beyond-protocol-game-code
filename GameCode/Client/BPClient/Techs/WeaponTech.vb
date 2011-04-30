Public MustInherit Class WeaponTech
    Inherits Base_Tech

    'Base attributes
    Public WeaponName As String
    Public WeaponClassTypeID As WeaponClassType
    Public WeaponTypeID As WeaponType
    Public PowerRequired As Int32
    Public mlHullRequired As Int32

    Public blAmmoProdCredits As Int64 = 0
    Public blAmmoProdPoints As Int64 = 0

    Public HullTypeID As Byte

    Public oResultWeaponDef As WeaponDef = Nothing

    Protected MustOverride Sub PopulateFromMsg(ByRef yData() As Byte, ByVal lPos As Int32)
    Public MustOverride Function GetPrototypeBuilderString(ByVal bPointDefense As Boolean) As String
    Public MustOverride Sub GetPrototypeBuilderAttributeString(ByRef oSB As System.Text.StringBuilder, ByVal bPointDefense As Boolean)

    Public Property HullRequired() As Int32
        Get
            If mlHullRequired = 0 Then
                If MyBase.oProductionCost Is Nothing = False Then
                    Dim lTemp As Int32 = 0
                    For X As Int32 = 0 To MyBase.oProductionCost.ItemCostUB
                        lTemp += MyBase.oProductionCost.ItemCosts(X).QuantityNeeded
                    Next X
                    Return lTemp
                End If
            End If

            Return mlHullRequired
        End Get
        Set(ByVal value As Int32)
            mlHullRequired = value
        End Set
    End Property

    Public Overrides Sub FillFromMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 0

        With Me
            'BASE
            '-----
            'MsgCode (2)
            lPos = 2

            'GUID (6)
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            'OwnerID (4)
            .OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            If .OwnerID = 0 Then .OwnerID = glPlayerID
            'DevPhase (4)
            .ComponentDevelopmentPhase = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'ErrorCodeReason (1)
            .ErrorCodeReason = yData(lPos) : lPos += 1
            'Researchers (1)
            .Researchers = yData(lPos) : lPos += 1
            'MajorDesignFlaw (1)
            .MajorDesignFlaw = yData(lPos) : lPos += 1

            .lSpecifiedHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedPower = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .blSpecifiedResCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .blSpecifiedResTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .blSpecifiedProdCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .blSpecifiedProdTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .lSpecifiedColonists = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedEnlisted = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedOfficers = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin4 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin5 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin6 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            'WeaponName (20)
            .WeaponName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            'WeaponClassTypeID (1)
            .WeaponClassTypeID = CType(yData(lPos), WeaponClassType) : lPos += 1
            'WeapontypeID (1)
            .WeaponTypeID = CType(yData(lPos), WeaponType) : lPos += 1
            'PowerRequired (4)
            .PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'HullRequired (4)
            .HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .HullTypeID = yData(lPos) : lPos += 1
        End With

        PopulateFromMsg(yData, lPos)
    End Sub

    Public Shared Function CreateWeaponClass(ByRef yData() As Byte) As WeaponTech
        Const lWpnClassTypePos As Int32 = 115 '39
        Select Case yData(lWpnClassTypePos)
			Case WeaponClassType.eBomb
				Return New BombWeaponTech
			Case WeaponClassType.eEnergyBeam
				Return New SolidBeamWeaponTech()
            Case WeaponClassType.eEnergyPulse
                Return New PulseWeaponTech()
            Case WeaponClassType.eMine
            Case WeaponClassType.eMissile
                Return New MissileWeaponTech()
            Case WeaponClassType.eProjectile
                Return New ProjectileWeaponTech()
        End Select
        Return Nothing
    End Function

    'Public Function GetAmmoSize() As Single
    '    Select Case WeaponClassTypeID
    '        Case WeaponClassType.eMissile
    '            Return CType(Me, MissileWeaponTech).MissileHullSize
    '        Case WeaponClassType.eProjectile
    '            Return CType(Me, ProjectileWeaponTech).fAmmoSize
    '        Case Else
    '            Return 0.0F
    '    End Select
    'End Function

    Protected Sub FillWeaponDefFromMsg(ByRef yData() As Byte, ByVal lPos As Int32)
        Try
            If lPos + 89 < yData.Length Then
                oResultWeaponDef = New WeaponDef()
                With oResultWeaponDef
                    .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .WeaponName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                    .yWeaponType = yData(lPos) : lPos += 1
                    .ROF_Delay = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .Range = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .Accuracy = yData(lPos) : lPos += 1
                    .PiercingMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .PiercingMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ImpactMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ImpactMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .BeamMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .BeamMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ECMMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ECMMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .FlameMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .FlameMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ChemicalMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ChemicalMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .WpnGroup = yData(lPos) : lPos += 1
                    .lFirePowerRating = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    lPos += 4

                    .AOERange = yData(lPos) : lPos += 1
                    .WeaponSpeed = yData(lPos) : lPos += 1
                    .Maneuver = yData(lPos) : lPos += 1
                End With
            End If
        Catch
        End Try
    End Sub
End Class
