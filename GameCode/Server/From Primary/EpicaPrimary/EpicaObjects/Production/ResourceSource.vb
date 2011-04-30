Public Structure ResourceSource
    ''' <summary>
    ''' The Facility Index within the colony's Children Array
    ''' </summary>
    ''' <remarks></remarks>
    Public lFacilityIdx As Int32

    ''' <summary>
    ''' The index of the cargo content item within the child Facility's cargo array
    ''' </summary>
    ''' <remarks></remarks>
    Public lCargoIdx As Int32

    ''' <summary>
    ''' The Item ID of the object I can find here
    ''' </summary>
    ''' <remarks></remarks>
    Public lItemID As Int32

    ''' <summary>
    ''' The Item TypeID of the object I can find here
    ''' </summary>
    ''' <remarks></remarks>
    Public iTypeID As Int16

    ''' <summary>
    ''' The Quantity of the Object I can find here
    ''' </summary>
    ''' <remarks></remarks>
    Public lQuantity As Int32

    ''' <summary>
    ''' The type of cache I can expect to find
    ''' </summary>
    ''' <remarks></remarks>
    Public iCacheTypeID As Int16

    ''' <summary>
    ''' Unit Index within the goUnit array... NOTE: lFacilityIdx must be set to int32.minvalue
    ''' </summary>
    ''' <remarks></remarks>
    Public lUnitIdx As Int32
End Structure
