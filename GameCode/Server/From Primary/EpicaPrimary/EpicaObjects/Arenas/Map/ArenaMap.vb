Option Strict On

Imports System.Xml

'Referenced by the MapID
Public Class ArenaMap
    Public lMapID As Int32

    Public sMapName As String

    Public yGameMode As Byte
    Public lMinSideCnt As Int32
    Public lMaxSideCnt As Int32
    Public lMinPlayerCnt As Int32
    Public lMaxPlayerCnt As Int32
    Public lMinUnitCntPerSide As Int32
    Public lMaxUnitCntPerSide As Int32

    Public Sub LoadMapFromXML(ByRef oDoc As XmlDocument)

        Dim oNode As XmlNode = oDoc.FirstChild()
        While oNode Is Nothing = False
            If oNode.Name.ToUpper = "MAP" Then
                Exit While
            Else : oNode = oNode.NextSibling
            End If
        End While

        If oNode Is Nothing = False AndAlso oNode.Name.ToUpper = "MAP" Then
            While oNode Is Nothing = False
                ParseNode(oNode)
            End While
        End If
    End Sub

    Private Sub ParseNode(ByVal oRootNode As XmlNode)
        Dim oNode As XmlNode = oRootNode.FirstChild

        While oNode Is Nothing = False

            Select Case oNode.Name.ToUpper
                Case "MAPNAME"
                    sMapName = oNode.InnerText
                Case "GAMETYPE"
                    yGameMode = CByte(oNode.InnerText)
                Case "MINSIDECNT"
                    lMinSideCnt = CInt(oNode.InnerText)
                Case "MAXSIDECNT"
                    lMaxSideCnt = CInt(oNode.InnerText)
                Case "MINPLAYERCNT"
                    lMinPlayerCnt = CInt(oNode.InnerText)
                Case "MAXPLAYERCNT"
                    lMaxPlayerCnt = CInt(oNode.InnerText)
                Case "MINUNITCNTPERSIDE"
                    lMinUnitCntPerSide = CInt(oNode.InnerText)
                Case "MAXUNITCNTPERSIDE"
                    lMaxUnitCntPerSide = CInt(oNode.InnerText)
            End Select

            oNode = oNode.NextSibling
        End While
    End Sub

End Class