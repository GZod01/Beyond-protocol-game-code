Option Strict On

Public Enum elWormholeFlag As Int32
    eSystem1Detectable = 1
    eSystem2Detectable = 2
    eSystem1Jumpable = 4
    eSystem2Jumpable = 8
End Enum

Public Class Wormhole
    Inherits Base_GUID

    Public System1 As SolarSystem
    Public System2 As SolarSystem
    Public LocX1 As Int32
    Public LocY1 As Int32
    Public LocX2 As Int32
    Public LocY2 As Int32

    Public StartCycle As Int32
    Public WormholeFlags As Int32

    Public Function FillFromAddMsg(ByRef yData() As Byte) As Boolean
        'msgcode = 2
        Dim lPos As Int32 = 2
        Me.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Me.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X).ObjectID = lTemp Then
                System1 = goGalaxy.moSystems(X)
                Exit For
            End If
        Next X
        lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X).ObjectID = lTemp Then
                System2 = goGalaxy.moSystems(X)
                Exit For
            End If
        Next X
        LocX1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        LocY1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        LocX2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        LocY2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        StartCycle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        WormholeFlags = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If System1 Is Nothing = False AndAlso System2 Is Nothing = False Then
            System1.AddWormhole(Me)
            System2.AddWormhole(Me)
            Return True
        Else : Return False
        End If

    End Function
End Class
