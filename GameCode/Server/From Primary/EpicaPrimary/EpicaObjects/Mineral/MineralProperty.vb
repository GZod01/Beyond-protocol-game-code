Public Class MineralProperty        'lookup table defining the different types of minerals
    Inherits Epica_GUID

    Public MineralPropertyName(49) As Byte
    Public MinimumValue As Int32
    Public MaximumValue As Int32

    'Ordered from lowest to highest
    Private moValueRangeNames() As ValueRangeName
    Private mlValueRangeNameUB As Int32 = -1

    Private mySendString() As Byte      'never changes

    Public Sub AddValueRangeName(ByVal lMinVal As Int32, ByVal sName As String)
        'NOTE: IT IS ASSUMED THAT VALUE RANGES ARE ADDED IN ASCENDING ORDER!!!!
        mlValueRangeNameUB += 1
        ReDim Preserve moValueRangeNames(mlValueRangeNameUB)
        moValueRangeNames(mlValueRangeNameUB) = New ValueRangeName()
        With moValueRangeNames(mlValueRangeNameUB)
            .lMinValue = lMinVal
            .RangeName = StringToBytes(sName)
        End With
    End Sub

    Public Function GetValueRangeName(ByVal lValue As Int32) As Byte()
        Dim lIdx As Int32 = -1

        For X As Int32 = 0 To mlValueRangeNameUB
            If lValue >= moValueRangeNames(X).lMinValue Then
                lIdx = X
            Else : Exit For
            End If
        Next X

        If lIdx <> -1 Then
            Return moValueRangeNames(lIdx).RangeName
        Else : Return StringToBytes("Unknown")
        End If
    End Function

    Public Function GetObjAsString() As Byte()
        If mySendString Is Nothing Then
            ReDim mySendString(55)
            Me.GetGUIDAsString.CopyTo(mySendString, 0)
            MineralPropertyName.CopyTo(mySendString, 6)
        End If
        Return mySendString
    End Function
End Class

Public Enum eMinPropID As Integer
    Density = 1
    Hardness = 2
    Malleable = 3
    ElectricalResist = 4
    Compressibility = 5
    Reflection = 6
    Refraction = 7
    MeltingPoint = 8
    BoilingPoint = 9
    SuperconductivePoint = 10
    TemperatureSensitivity = 11
    ThermalConductance = 12
    ThermalExpansion = 13
    MagneticReaction = 14
    MagneticProduction = 15
    Toxicity = 16
    MineralColor = 17
    ChemicalReactance = 18
    Combustiveness = 19
    Complexity = 20
    Quantum = 21
    Psych = 22
End Enum