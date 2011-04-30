''' <summary>
''' Contains a relationship between a mineral and a geography object for redistribution purposes
''' </summary>
''' <remarks></remarks>
Public Class MineralGeographyRel

    Public lMineralID As Int32 = -1
    Public lGeoID As Int32 = -1
    Public iGeoTypeID As Int16 = -1

    Public lRedistMaxQty As Int32 = 0
    Public lRedistMaxConc As Int32 = 0

    Private moMineral As Mineral = Nothing
    Public ReadOnly Property Mineral() As Mineral
        Get
            If moMineral Is Nothing Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lMineralID Then
                        moMineral = goMineral(X)
                        Exit For
                    End If
                Next X
            End If

            Return moMineral
        End Get
    End Property

    Private moGeoObject As Object = Nothing
    Public ReadOnly Property GeoObject() As Object
        Get
            If moGeoObject Is Nothing Then
                moGeoObject = GetEpicaObject(lGeoID, iGeoTypeID)
            End If

            Return moGeoObject
        End Get
    End Property

End Class
