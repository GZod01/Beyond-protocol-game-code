Option Strict On

Public Class MapGeography
    Inherits Base_GUID

    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32
    Public sName As String

    Public lMinimumX As Int32
    Public lMaximumX As Int32
    Public lMinimumZ As Int32
    Public lMaximumZ As Int32

    'For Planets...
    Public PlanetSizeID As Byte
    Public PlanetRadius As Int16
    Public MapTypeID As Byte
    Public WaterHeight As Int32
    Public RotationDelay As Int16
    Public AxisAngle As Int32
    Public RotateAngle As Int32


End Class