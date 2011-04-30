''' <summary>
''' This class is a container class for objects that are not directly owned by this primary
''' This primary would be responsible for managing and transmitting changes to the objects
''' within this class. Furthermore, some data must be needed on certain servers even if the
''' server does not own the data. In those cases, this class is useful
''' </summary>
''' <remarks></remarks>
Public Class NonLocalData
    Public oTrades() As DirectTrade
    Public lTradeUB As Int32 = -1

    Public oAgents() As Agent
    Public lAgentIdx() As Int32
    Public lAgentUB As Int32 = -1

End Class
