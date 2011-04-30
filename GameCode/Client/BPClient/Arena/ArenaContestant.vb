Option Strict On

Public Class ArenaContestant
    Public lID As Int32 = -1
    Public iTypeID As Int16 = -1
    Public lHullSize As Int32
    Public yHullType As Byte

    Public lRespawns As Int32 = 0
    Public lNextRespawnCycle As Int32 = -1

    Public Sub New(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16)
        lID = lEntityID
        iTypeID = iEntityTypeID
    End Sub

End Class
