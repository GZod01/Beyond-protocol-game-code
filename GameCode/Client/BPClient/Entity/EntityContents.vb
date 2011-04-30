
Public Class EntityContents
    Inherits Base_GUID
    Public lQuantity As Int32
    Public oDetailItem As Object

    'These are for tracking the collections
    Public colHangar As Collection
    Public colCargo As Collection
    Public ContentsLastUpdate As Int32 = 0      'cycle number of last contents update

    Public lDetailItemID As Int32 = -1
    Public iDetailItemTypeID As Int16 = -1
End Class
