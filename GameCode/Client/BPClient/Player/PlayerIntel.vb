Option Strict On

Public Class PlayerIntel
    Inherits Base_GUID

    Public PlayerName As String
    Public EmpireName As String
    Public RaceName As String

    Public lPlayerIcon As Int32

    Public lGuildID As Int32 = -1
    'Public lSenateID As Int32 = -1

    Public lStartID As Int32 = -1
    Public iStartTypeID As Int16 = -1

    Public yPlayerTitle As Byte
    Public yCustomTitle As Byte

    Public bIsMale As Boolean = True

    'These are the scores of the player...
    Public lTechnologyScore As Int32
    Public lTechnologyUpdate As Int32
    Public lDiplomacyScore As Int32
    Public lDiplomacyUpdate As Int32
    Public lMilitaryScore As Int32
    Public lMilitaryUpdate As Int32
    Public lPopulationScore As Int32
    Public lPopulationUpdate As Int32
    Public lProductionScore As Int32
    Public lProductionUpdate As Int32
    Public lWealthScore As Int32
    Public lWealthUpdate As Int32
    Public lTotalScore As Int32

    Public yRegardsCurrentPlayer As Byte

    Public lCelebrationEnds As Int32

    Public lTotalVoteStrength As Int32

    Public lLastUpdate As Int32

    Public Sub FillFromPlayerIntelMsg(ByRef yMsg() As Byte)
        Dim lPos As Int32 = 2       'for msgcode

        Me.ObjectID = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        Me.ObjTypeID = ObjectType.ePlayer ''TODO: Player or Player Intel?
        lPos += 2        'for the object typeid
        lPos += 4       'playerid
        lStartID = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        iStartTypeID = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
        lTechnologyScore = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4

        Dim lSeconds As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lTechnologyUpdate = GetDateAsNumber(Now.Subtract(New TimeSpan(0, 0, lSeconds)))

        lDiplomacyScore = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lSeconds = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lDiplomacyUpdate = GetDateAsNumber(Now.Subtract(New TimeSpan(0, 0, lSeconds)))

        lMilitaryScore = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lSeconds = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lMilitaryUpdate = GetDateAsNumber(Now.Subtract(New TimeSpan(0, 0, lSeconds)))

        lPopulationScore = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lSeconds = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lPopulationUpdate = GetDateAsNumber(Now.Subtract(New TimeSpan(0, 0, lSeconds)))

        lProductionScore = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lSeconds = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lProductionUpdate = GetDateAsNumber(Now.Subtract(New TimeSpan(0, 0, lSeconds)))

        lWealthScore = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lSeconds = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        lWealthUpdate = GetDateAsNumber(Now.Subtract(New TimeSpan(0, 0, lSeconds)))

        lTotalScore = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        yRegardsCurrentPlayer = yMsg(lPos) : lPos += 1

        bIsMale = yMsg(lPos) = 1 : lPos += 1
        lPlayerIcon = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        yPlayerTitle = yMsg(lPos) : lPos += 1
        lCelebrationEnds = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
        If lCelebrationEnds > 0 Then lCelebrationEnds += glCurrentCycle

        yCustomTitle = yMsg(lPos) : lPos += 1
        lGuildID = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4

        lTotalVoteStrength = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2

        lLastUpdate = glCurrentCycle
    End Sub

End Class
