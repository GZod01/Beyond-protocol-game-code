Option Strict On

'tracks events in the arena
Public Enum eyArenaEventType As Byte
    EntityKilled = 0
    FlagCaptured = 1
End Enum
Public Class ArenaEvent
    Public yTypeID As Byte
    Public oArena As Arena

    'Who did the action (the entity guid)
    Public lActorID As Int32
    Public iActorTypeID As Int16

    'related ID for this event
    Public lRelatedID As Int32
End Class

