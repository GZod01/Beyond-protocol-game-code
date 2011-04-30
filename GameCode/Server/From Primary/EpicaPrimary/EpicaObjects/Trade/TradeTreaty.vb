Option Strict On

'Public Class TradeTreaty
'    Public oPlayer1 As TradeTreatyPlayer
'    Public oPlayer2 As TradeTreatyPlayer

'    Private moRoutes() As TradeTreatyRoute

'    Private Class PiratePlayer
'        Public oPlayer As Player
'        Public lPirateIncome As Int32
'    End Class
'    'This is the total for all pirate players of this trade treaty
'    Private moPirates() As PiratePlayer
'    Private mlPiratePlayerUB As Int32 = -1

'    Public Sub ExecuteTreaty()
'        With oPlayer1
'            .oPlayer.oBudget.lTotalTradeIncome += .lTotalIncome
'        End With
'        With oPlayer2
'            .oPlayer.oBudget.lTotalTradeIncome += .lTotalIncome
'        End With

'        'Now, reward our pirate players
'        For X As Int32 = 0 To mlPiratePlayerUB
'            If moPirates(X) Is Nothing = False Then
'                'Finalizing hte budget will increase the pirate's credits
'                moPirates(X).oPlayer.oBudget.lPirateIncome += moPirates(X).lPirateIncome
'            End If
'        Next X
'    End Sub

'    Private Function GetTTKey() As String
'        Return "TT:" & oPlayer1.oPlayer.ObjectID.ToString & "," & oPlayer2.oPlayer.ObjectID.ToString
'    End Function

'    Public Sub RecalculateTreaty()
'        Dim lP1TotalIncome As Int32 = 0
'        Dim lP2TotalIncome As Int32 = 0

'        'Ok, let's get our route UB
'        Dim lUB As Int32 = oPlayer1.colSystems.Count + oPlayer2.colSystems.Count - 1
'        Dim oRoutes(lUB) As TradeTreatyRoute

'        Dim colPirates As New Collection
'        Dim sTTKey As String = GetTTKey()

'        'Now, go through our items
'        Dim lIdx As Int32 = -1
'        For Each oSys1 As SolarSystem In oPlayer1.colSystems
'            For Each oSys2 As SolarSystem In oPlayer2.colSystems
'                lIdx += 1
'                oRoutes(lIdx) = New TradeTreatyRoute()
'                With oRoutes(lIdx)
'                    .oSystemA = oSys1
'                    .oSystemB = oSys2

'                    'Now, determine the route between them
'                    .GenerateRoute(oPlayer1.oPlayer, oPlayer2.oPlayer)

'                    'Now, determine the route's income
'                    .CalculateIncomes(oPlayer1.oPlayer, oPlayer2.oPlayer)

'                    'Now, increment our total incomes for final storage
'                    lP1TotalIncome += .lPlayer1Income
'                    lP2TotalIncome += .lPlayer2Income

'                    'Finally, populate our pirates from the route
'                    For Z As Int32 = 0 To .lPiratePlayerUB
'                        Dim sKey As String = "P" & .oPiratePlayers(Z).ObjectID.ToString
'                        If colPirates.Contains(sKey) = False Then
'                            Dim oItem As New PiratePlayer()
'                            oItem.lPirateIncome = .lPiratePlayerIncome(Z)
'                        Else
'                            Dim oItem As PiratePlayer = CType(colPirates(sKey), PiratePlayer)
'                            If oItem Is Nothing = False Then oItem.lPirateIncome += .lPiratePlayerIncome(Z)
'                        End If
'                    Next Z

'                    'Register this route with the systems
'                    If .oSystemA.colTreaties.Contains(sTTKey) = False Then
'                        .oSystemA.colTreaties.Add(Me, sTTKey)
'                    End If
'                    If .oSystemB.colTreaties.Contains(sTTKey) = False Then
'                        .oSystemB.colTreaties.Add(Me, sTTKey)
'                    End If

'                    For Z As Int32 = 0 To .lRouteSystemUB
'                        If .oRouteSystems(Z).colTreaties.Contains(sTTKey) = False Then
'                            .oRouteSystems(Z).colTreaties.Add(Me, sTTKey)
'                        End If
'                    Next Z
'                End With
'            Next
'        Next

'        'Store our final income values
'        oPlayer1.lTotalIncome = lP1TotalIncome
'        oPlayer2.lTotalIncome = lP2TotalIncome

'        'And store our final pirate values
'        Dim oPirates(colPirates.Count - 1) As PiratePlayer
'        lIdx = -1
'        For Each oTmp As PiratePlayer In colPirates
'            lIdx += 1
'            oPirates(lIdx) = oTmp
'        Next
'        mlPiratePlayerUB = -1
'        moPirates = oPirates
'        mlPiratePlayerUB = moPirates.GetUpperBound(0)


'    End Sub

'    Public Sub PlayerAddSystem(ByVal lPlayerID As Int32, ByVal oSys As SolarSystem)
'        Dim sKey As String = "Sys" & oSys.ObjectID.ToString
'        If oPlayer1.oPlayer.ObjectID = lPlayerID Then
'            If oPlayer1.colSystems.Contains(sKey) = False Then
'                oPlayer1.colSystems.Add(oSys, sKey)
'                RecalculateTreaty()
'            End If
'        ElseIf oPlayer2.oPlayer.ObjectID = lPlayerID Then
'            If oPlayer2.colSystems.Contains(sKey) = False Then
'                oPlayer2.colSystems.Add(oSys, sKey)
'                RecalculateTreaty()
'            End If
'        End If
'    End Sub

'    Public Sub PlayerRemoveSystem(ByVal lPlayerID As Int32, ByVal oSys As SolarSystem)
'        Dim sKey As String = "Sys" & oSys.ObjectID.ToString
'        If oPlayer1.oPlayer.ObjectID = lPlayerID Then
'            If oPlayer1.colSystems.Contains(sKey) = True Then
'                oPlayer1.colSystems.Remove(sKey)
'                Dim sTTKey As String = GetTTKey()
'                If oSys.colTreaties.Contains(sTTKey) = True Then
'                    oSys.colTreaties.Remove(sTTKey)
'                End If

'                RecalculateTreaty()
'            End If
'        ElseIf oPlayer2.oPlayer.ObjectID = lPlayerID Then
'            If oPlayer2.colSystems.Contains(sKey) = True Then
'                oPlayer2.colSystems.Remove(sKey)
'                Dim sTTKey As String = GetTTKey()
'                If oSys.colTreaties.Contains(sTTKey) = True Then
'                    oSys.colTreaties.Remove(sTTKey)
'                End If

'                RecalculateTreaty()
'            End If
'        End If
'    End Sub
'End Class

'Public Class TradeTreatyPlayer
'    Public oPlayer As Player

'    'These systems are marked by the player as systems to be included in the trade
'    Public colSystems As New Collection

'    Public lTotalIncome As Int32 = 0        'total income for all routes in the trade treaty rewarded to this player
'End Class

'Public Class TradeTreatyRoute
'    Public oSystemA As SolarSystem
'    Public oSystemB As SolarSystem

'    Public oRouteSystems() As SolarSystem       'does not include A or B
'    Public lRouteSystemUB As Int32 = -1

'    Public lPiratePlanets As Int32 = 0          'number of planets along the route that are owned by pirate players

'    Public oPiratePlayers() As Player
'    Public lPiratePlayerIncome() As Int32       'income generated for this pirate player for this route
'    Public lPiratePlayerUB As Int32 = -1

'    Public lPlayer1Income As Int32 = 0          'income per cycle for this route rewarded to Player1
'    Public lPlayer2Income As Int32 = 0          'income per cycle for this route rewarded to Player2

'    Public Sub CalculateIncomes(ByVal oPlayerA As Player, ByVal oPlayerB As Player)
'        Dim lSystemAPlanets As Int32 = 0         'number of planets belonging to oPlayer1 in SystemA... AND are positive income
'        Dim lSystemBPlanets As Int32 = 0         'number of planets belonging to oPlayer2 in SystemB... AND are positive income

'        Dim lPlayerAID As Int32 = oPlayerA.ObjectID
'        Dim lPlayerBID As Int32 = oPlayerB.ObjectID

'        If lPiratePlanets > 9 Then
'            lPiratePlayerUB = -1
'            lPlayer1Income = 0
'            lPlayer2Income = 0
'            Return
'        End If

'        'Ok, check systemA
'        For X As Int32 = 0 To oSystemA.mlPlanetUB
'            Dim lPlanetIdx As Int32 = oSystemA.GetPlanetIdx(X)
'            Dim oPlanet As Planet = goPlanet(lPlanetIdx)
'            If oPlanet Is Nothing Then Exit For

'            If oPlayerA.oBudget.IsEnvironmentPositiveIncome(oPlanet.ObjectID, ObjectType.ePlanet) = True Then
'                lSystemAPlanets += 1
'            End If
'        Next X
'        For X As Int32 = 0 To oSystemB.mlPlanetUB
'            Dim lPlanetIdx As Int32 = oSystemB.GetPlanetIdx(X)
'            Dim oPlanet As Planet = goPlanet(lPlanetIdx)
'            If oPlanet Is Nothing Then Exit For

'            If oPlayerB.oBudget.IsEnvironmentPositiveIncome(oPlanet.ObjectID, ObjectType.ePlanet) = True Then
'                lSystemBPlanets += 1
'            End If
'        Next X

'        'Now, get the per planet income
'        Dim lPerPlanetIncome As Int32 = 10 - lPiratePlanets
'        'and assign it to the incomes
'        lPlayer1Income = lPerPlanetIncome * lSystemAPlanets
'        lPlayer2Income = lPerPlanetIncome * lSystemBPlanets
'    End Sub

'    Public Sub GenerateRoute(ByVal oPlayer1 As Player, ByVal oPlayer2 As Player)
'        'generate the route between system A and update oRouteSystems
'        Dim colPath As Collection = CalculatePath(oSystemA, oSystemB, oPlayer1, oPlayer2)

'        'as the route is generated, check how many planets are owned by pirate players and update the pirate array
'        For Each oSys As SolarSystem In colPath

'            'get how many planets are owned by pirate players
'            For X As Int32 = 0 To oSys.mlPlanetUB
'                Dim lPIdx As Int32 = oSys.GetPlanetIdx(X)
'                If lPIdx > -1 AndAlso lPIdx <= glPlanetUB Then
'                    Dim oPlanet As Planet = goPlanet(lPIdx)
'                    If oPlanet Is Nothing = False AndAlso oPlanet.OwnerID <> oPlayer1.ObjectID AndAlso oPlanet.OwnerID <> oPlayer2.ObjectID AndAlso oPlanet.OwnerID > 0 Then
'                        'lPiratePlayerUB += 1

'                    End If
'                End If
'            Next X



'            If oSys.ObjectID = oSystemA.ObjectID OrElse oSys.ObjectID = oSystemB.ObjectID Then Continue For

'        Next
'    End Sub


'    Private Class WHPath
'        Public oSys As SolarSystem
'        Public oParent As SolarSystem
'        Public oWH As Wormhole
'        Public lJumpCnt As Int32
'    End Class
'    Private Function CalculatePath(ByVal oFromSys As SolarSystem, ByVal oToSys As SolarSystem, ByVal oPlayerA As Player, ByVal oPlayerB As Player) As Collection

'        'Now, go and see if we can get from the FROM sys to the TO sys 
'        Dim oPath(glSystemUB) As WHPath

'        Dim lToSysID As Int32 = oToSys.ObjectID
'        Dim lFromSysID As Int32 = oFromSys.ObjectID

'        Dim oToPathObj As WHPath = Nothing
'        For X As Int32 = 0 To glSystemUB
'            oPath(X) = New WHPath
'            With oPath(X)
'                .oSys = goSystem(X)
'                .oParent = Nothing
'                If .oSys.ObjectID = lToSysID Then
'                    .lJumpCnt = 0
'                    oToPathObj = oPath(X)
'                Else
'                    .lJumpCnt = 9999999
'                End If
'            End With
'        Next X

'        Dim colOpen As New Collection
'        Dim colClosed As New Collection

'        colOpen.Add(oToPathObj, "SYS" & lToSysID)

'        Dim bDone As Boolean = False

'        Dim colResults As New Collection


'        While bDone = False
'            If colOpen.Count = 0 Then
'                bDone = True
'            Else
'                Dim oCur As WHPath = Nothing
'                For Each oTmp As WHPath In colOpen
'                    If oTmp Is Nothing = False Then
'                        oCur = oTmp
'                        Exit For
'                    End If
'                Next
'                If oCur Is Nothing = False Then
'                    Dim sCurKey As String = "SYS" & oCur.oSys.ObjectID
'                    colOpen.Remove(sCurKey)
'                    colClosed.Add(oCur, sCurKey)

'                    With oCur.oSys

'                        If oCur.oSys.ObjectID = lFromSysID Then

'                            'OK, create our results
'                            Dim oCurPath As WHPath = oCur
'                            Dim fDist As Single = 0.0F
'                            While oCurPath Is Nothing = False

'                                Dim sTmpKey As String = "SYS" & oCurPath.oSys.ObjectID
'                                If colResults.Contains(sTmpKey) = False Then colResults.Add(oCurPath.oSys, sTmpKey)

'                                If oCurPath.oSys.ObjectID = lToSysID Then Exit While

'                                Dim oP As SolarSystem = oCurPath.oParent
'                                For X As Int32 = 0 To glSystemUB
'                                    If goSystem(X).ObjectID = oP.ObjectID Then
'                                        oCurPath = oPath(X)
'                                        Exit For
'                                    End If
'                                Next X
'                            End While

'                            Return colResults
'                        End If

'                        'Now, go through this system's wormholes
'                        For X As Int32 = 0 To .mlWormholeUB
'                            If .moWormholes(X).System1.ObjectID = .ObjectID Then
'                                If (.moWormholes(X).WormholeFlags And elWormholeFlag.eSystem1Jumpable) = 0 Then Continue For '  If wormhole isnt open continue
'                                If oPlayerA.HasWormholeKnowledge(.moWormholes(X).ObjectID) = False OrElse oPlayerB.HasWormholeKnowledge(.moWormholes(X).ObjectID) = False Then Continue For

'                                'Ok, check if system2 is in the list
'                                Dim lOtherSysID As Int32 = .moWormholes(X).System2.ObjectID
'                                Dim sKey As String = "SYS" & lOtherSysID
'                                If colOpen.Contains(sKey) = False AndAlso colClosed.Contains(sKey) = False Then
'                                    For Y As Int32 = 0 To glSystemUB
'                                        If goSystem(Y).ObjectID = lOtherSysID Then
'                                            oPath(Y).oParent = .moWormholes(X).System1

'                                            Dim lJumpCnt As Int32 = oCur.lJumpCnt
'                                            oPath(Y).lJumpCnt = lJumpCnt + 1
'                                            oPath(Y).oWH = .moWormholes(X)
'                                            colOpen.Add(oPath(Y), sKey)
'                                            Exit For
'                                        End If
'                                    Next Y
'                                End If

'                            Else
'                                If (.moWormholes(X).WormholeFlags And elWormholeFlag.eSystem2Jumpable) = 0 Then Continue For '  If wormhole isnt open continue
'                                If oPlayerA.HasWormholeKnowledge(.moWormholes(X).ObjectID) = False OrElse oPlayerB.HasWormholeKnowledge(.moWormholes(X).ObjectID) = False Then Continue For

'                                'Ok, check if system1 is in the list
'                                Dim lOtherSysID As Int32 = .moWormholes(X).System1.ObjectID
'                                Dim sKey As String = "SYS" & lOtherSysID
'                                If colOpen.Contains(sKey) = False AndAlso colClosed.Contains(sKey) = False Then
'                                    For Y As Int32 = 0 To glSystemUB
'                                        If goSystem(Y).ObjectID = lOtherSysID Then
'                                            oPath(Y).oParent = .moWormholes(X).System2

'                                            Dim lJumpCnt As Int32 = oCur.lJumpCnt
'                                            oPath(Y).lJumpCnt = lJumpCnt + 1
'                                            oPath(Y).oWH = .moWormholes(X)
'                                            colOpen.Add(oPath(Y), sKey)
'                                            Exit For
'                                        End If
'                                    Next Y
'                                End If
'                            End If
'                        Next
'                    End With
'                Else
'                    bDone = True
'                End If
'            End If
'        End While

'        'if we're here... return nothing
'        colResults.Add(oFromSys, "SYS" & oFromSys.ObjectID)
'        colResults.Add(oToSys, "SYS" & oToSys.ObjectID)
'        Return colResults
'    End Function

'End Class
 