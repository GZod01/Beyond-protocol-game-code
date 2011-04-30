Option Strict On

Public Class Neighborhood
    Public lNeighborhoodID As Int32
    Public oSystems() As SolarSystem
    Public lSystemUB As Int32 = -1

    Public oPrimaryServer As ServerObject

    Public Sub AddSystem(ByRef oSystem As SolarSystem)
        'first, let's place the system
        For X As Int32 = 0 To lSystemUB
            If oSystems(X) Is Nothing = False AndAlso oSystems(X).ObjectID = oSystem.ObjectID Then
                'Already added, so return
                Return
            End If
        Next X

        lSystemUB += 1
        ReDim Preserve oSystems(lSystemUB)
        oSystems(lSystemUB) = oSystem

        'Now, check it for wormhole links
        For X As Int32 = 0 To oSystem.mlWormholeUB
            If oSystem.moWormholes(X) Is Nothing = False Then
                If oSystem.moWormholes(X).System1.ObjectID = oSystem.ObjectID Then
                    'Ok, add system2
                    AddSystem(oSystem.moWormholes(X).System2)
                Else
                    'ok, add system1
                    AddSystem(oSystem.moWormholes(X).System1)
                End If
            End If
        Next X
    End Sub

End Class
