Option Strict On

'This class is mainly used for storing data that would normally cause the servers to be inefficient
Public Class Colony
    Inherits Base_GUID

    Public sName As String
    Public sParentName As String

    'NOTE: This is not the colony's ACTUAL parent... only the parent environment. For example, for a Space Station's Colony these would be the parent solar system
    Public ParentEnvirID As Int32
    Public ParentEnvirTypeID As Int16

    Public LastTradeablesRequest As Int32

    'TRADEABLES DATA
    Public Colonists As Int32
    Public Enlisted As Int32
    Public Officers As Int32
    Public Food As Int32
    Public Ammo As Int32 
End Class
