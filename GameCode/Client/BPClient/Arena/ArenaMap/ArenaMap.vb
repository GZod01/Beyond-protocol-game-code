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
    Public sThumbnail As String = ""

    Public oSpace() As MapGeography
    Public oPlanets() As MapGeography

    Private moCurrentParseObj As MapGeography
    Private miCurrentParseObjType As Int16

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
                oNode = oNode.NextSibling
            End While
        End If
    End Sub

    Private Sub ParseNode(ByVal oRootNode As XmlNode)
        Dim oNode As XmlNode = oRootNode.FirstChild

        While oNode Is Nothing = False

            Select Case oNode.Name.ToUpper
                Case "MAPID"
                    lMapID = CInt(oNode.InnerText)
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
                Case "MAPGEOGRAPHY"
                    ParseNode(oNode)
                Case "SPACE"
                    moCurrentParseObj = New MapGeography
                    miCurrentParseObjType = ObjectType.eSolarSystem
                    ParseNode(oNode)
                    'If oSpace Is Nothing Then ReDim oSpace(-1)
                    'ReDim Preserve oSpace(oSpace.GetUpperBound(0) + 1)
                    'oSpace(oSpace.GetUpperBound(0)) = moCurrentParseObj
                    moCurrentParseObj = Nothing
                Case "PLANET"
                    moCurrentParseObj = New MapGeography
                    miCurrentParseObjType = ObjectType.ePlanet
                    ParseNode(oNode)
                    'If oPlanets Is Nothing Then ReDim oPlanets(-1)
                    'ReDim Preserve oPlanets(oPlanets.GetUpperBound(0) + 1)
                    'oPlanets(oPlanets.GetUpperBound(0)) = moCurrentParseObj
                    moCurrentParseObj = Nothing
                Case "GEOGRAPHYID"
                    moCurrentParseObj.ObjectID = CInt(oNode.InnerText)
                Case "GEOGRAPHYNAME"
                    moCurrentParseObj.sName = oNode.InnerText
                Case "LOCX"
                    moCurrentParseObj.LocX = CInt(oNode.InnerText)
                Case "LOCY"
                    moCurrentParseObj.LocY = CInt(oNode.InnerText)
                Case "LOCZ"
                    moCurrentParseObj.LocZ = CInt(oNode.InnerText)
                Case "RESPAWNLOCS"
                    ParseNode(oNode)
                Case "EXTENDED"
                    ParseNode(oNode)
                Case "MINIMUMX"
                    moCurrentParseObj.lMinimumX = CInt(oNode.InnerText)
                Case "MAXIMUMX"
                    moCurrentParseObj.lMaximumX = CInt(oNode.InnerText)
                Case "MINIMUMZ"
                    moCurrentParseObj.lMinimumZ = CInt(oNode.InnerText)
                Case "MAXIMUMZ"
                    moCurrentParseObj.lMaximumZ = CInt(oNode.InnerText)
                Case "PLANETSIZEID"
                    moCurrentParseObj.PlanetSizeID = CByte(oNode.InnerText)
                    'moCurrentParseObj.SetPlanetSize()
                Case "PLANETRADIUS"
                    moCurrentParseObj.PlanetRadius = CShort(oNode.InnerText)
                Case "PLANETTYPEID"
                    moCurrentParseObj.MapTypeID = CByte(oNode.InnerText)
                Case "WATERHEIGHT"
                    moCurrentParseObj.WaterHeight = CInt(oNode.InnerText)
                Case "ROTATIONDELAY"
                    moCurrentParseObj.RotationDelay = CShort(oNode.InnerText)
                Case "AXISANGLE"
                    moCurrentParseObj.AxisAngle = CInt(oNode.InnerText)
                Case "ROTATEANGLE"
                    moCurrentParseObj.RotateAngle = CInt(oNode.InnerText)
                Case "TERRAINDATA"
                    'moCurrentParseObj.SetTerrainData(oNode.InnerText)
                Case "MAPTHUMBNAIL"
                    sThumbnail = oNode.InnerText
            End Select

            oNode = oNode.NextSibling
        End While
    End Sub

#Region "  Arena Map Manager  "
    Private Shared moAllMaps() As ArenaMap
    Private Shared mlAllMapUB As Int32 = -1
    Private Shared mbAllMapInitialized As Boolean = False

    Private Shared Sub InitializeAllMaps()
        If mbAllMapInitialized = False Then

            mlAllMapUB = -1

            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            sPath &= "Arenas\"

            Dim colResults As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(sPath, FileIO.SearchOption.SearchAllSubDirectories, "*.xml")
            For Each sValue As String In colResults
                Try
                    Dim oXML As New Xml.XmlDocument()
                    oXML.Load(sValue)
                    Dim oMap As New ArenaMap()
                    oMap.LoadMapFromXML(oXML)

                    mlAllMapUB += 1
                    ReDim Preserve moAllMaps(mlAllMapUB)
                    moAllMaps(mlAllMapUB) = oMap
                Catch
                End Try
            Next sValue

            mbAllMapInitialized = True
        End If

    End Sub
    Public Shared Function AllMaps(ByVal lIndex As Int32) As ArenaMap
        InitializeAllMaps()
        If lIndex < 0 OrElse lIndex > mlAllMapUB Then Return Nothing
        Return moAllMaps(lIndex)
    End Function
    Public Shared Function AllMapUB() As Int32
        InitializeAllMaps()
        Return mlAllMapUB
    End Function
    Public Shared Function AllMapByID(ByVal lID As Int32) As ArenaMap
        InitializeAllMaps()
        For X As Int32 = 0 To mlAllMapUB
            If moAllMaps(X) Is Nothing = False Then
                If moAllMaps(X).lMapID = lID Then
                    Return moAllMaps(X)
                End If
            End If
        Next X
        Return Nothing
    End Function
#End Region

End Class