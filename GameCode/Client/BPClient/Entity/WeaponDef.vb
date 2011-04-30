Public Class WeaponDef
    Inherits Base_GUID

    Public WeaponName As String
    Public yWeaponType As Byte       'indicates what type of weapon this is, how it fires, etc.
    Public WpnGroup As Byte         'indicates whether the weapon is Primary, Secondary, or Point Defense, etc...
    Public ROF_Delay As Int16       'Added to NextFireCycle to determine when the unit fires again
    Public Range As Short
    Public Accuracy As Byte

    Public ParentObject As Object

    Public ArcID As Byte            'what facing the weapon shoots from

    Public lEntityStatusGroup As Int32      'this is the Bit-Wise identifier of what group this weapon def belongs to in CurrentStatus (elUnitStatus)
    Public lFirePowerRating As Int32

    Public PiercingMinDmg As Int32
    Public PiercingMaxDmg As Int32
    Public ImpactMinDmg As Int32
    Public ImpactMaxDmg As Int32
    Public BeamMinDmg As Int32
    Public BeamMaxDmg As Int32
    Public ECMMinDmg As Int32
    Public ECMMaxDmg As Int32
    Public FlameMinDmg As Int32
    Public FlameMaxDmg As Int32
    Public ChemicalMinDmg As Int32
    Public ChemicalMaxDmg As Int32

    Public MaxRangeAccuracy As Single       'stored value to aid in computations

    Public WeaponSpeed As Byte = 100

    Public AOERange As Byte

    Public lAmmoCap As Int32

    Public WpnDefID As Int32

    Public Maneuver As Byte
    Public lStructHP As Int32

    'returns the new lpos value
    Public Function FillFromMsg(ByRef yData() As Byte, ByVal lPos As Int32) As Int32
        With Me
            .ArcID = yData(lPos) : lPos += 1
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            'Weapon Name (20)
            .WeaponName = GetStringFromBytes(yData, lPos, 20)
            lPos += 20

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
            'wpn tech id
            lPos += 4

            .AOERange = yData(lPos) : lPos += 1
            .WeaponSpeed = yData(lPos) : lPos += 1
            .Maneuver = yData(lPos) : lPos += 1

            .lEntityStatusGroup = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lAmmoCap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .WpnDefID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .MaxRangeAccuracy = .Range * (.Accuracy / 100.0F)

            If .yWeaponType >= WeaponType.eMissile_Color_1 AndAlso .yWeaponType <= WeaponType.eMissile_Color_9 Then
                .lStructHP = .BeamMinDmg
                .BeamMinDmg = 0
            End If
        End With
        Return lPos
    End Function

End Class
