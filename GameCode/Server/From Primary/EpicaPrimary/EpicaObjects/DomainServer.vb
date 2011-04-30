Public Class DomainServer
    'this class manages our domain servers
    Public DomainSocket As NetSock

    'Where the client connects to
    Public ClientListenPort As Int16
    Public ClientListenIP(3) As Byte

    'Public oGrandaddyObject As Object

    Public bRegistered As Boolean

    Public bReportingShutdown As Boolean = False
	Public bReceivedDefs As Boolean = False
End Class
