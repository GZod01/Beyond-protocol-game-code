Option Strict On

Public Enum elModelSpecialTrait As int32
    NoSpecialTrait = 0
    ''' <summary>
    ''' Cargo and hangar capacity is +10%
    ''' </summary>
    ''' <remarks></remarks>
    CargoAndHangar10 = 1
    ''' <summary>
    ''' maneuver +2
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver2 = 2
    ''' <summary>
    ''' No jamming devices allowed
    ''' </summary>
    ''' <remarks></remarks>
    NoJammer = 3
    ''' <summary>
    ''' Cargo and hangar capacity +3%
    ''' </summary>
    ''' <remarks></remarks>
    CargoAndHangar3 = 4
    ''' <summary>
    ''' Maneuver +1
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver1 = 5
    ''' <summary>
    ''' Maneuver +4
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver4 = 6
    ''' <summary>
    ''' Speed +1
    ''' </summary>
    ''' <remarks></remarks>
    Speed1 = 7
    ''' <summary>
    ''' Cargo +5%
    ''' </summary>
    ''' <remarks></remarks>
    Cargo5 = 8
    ''' <summary>
    ''' Cargo and Hangar capacity +6%
    ''' </summary>
    ''' <remarks></remarks>
    CargoAndHangar6 = 9
    ''' <summary>
    ''' Maneuver +3
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver3 = 10
    ''' <summary>
    ''' Hangar Capacity +5%
    ''' </summary>
    ''' <remarks></remarks>
    Hangar5 = 11
    ''' <summary>
    ''' Speed and Maneuver increase +1
    ''' </summary>
    ''' <remarks></remarks>
    SpeedAndManeuver1 = 12
    ''' <summary>
    ''' Maneuver +5
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver5 = 13
    ''' <summary>
    ''' Hangar Capacity +10%
    ''' </summary>
    ''' <remarks></remarks>
    Hangar10 = 14
    ''' <summary>
    ''' Armor hitpoints +1000%
    ''' </summary>
    ''' <remarks></remarks>
    Armor1000 = 15
    ''' <summary>
    ''' Power Gen 2
    ''' </summary>
    ''' <remarks></remarks>
    PowerGen2 = 16
    ''' <summary>
    ''' Power Gen 3
    ''' </summary>
    ''' <remarks></remarks>
    PowerGen3 = 17
    ''' <summary>
    ''' Revenue from colony is increased 20%
    ''' </summary>
    ''' <remarks></remarks>
    Revenue20 = 18
    ''' <summary>
    ''' Launch speed is 6% faster
    ''' </summary>
    ''' <remarks></remarks>
    Launch6 = 19
    ''' <summary>
    ''' Scan Range +10
    ''' </summary>
    ''' <remarks></remarks>
    ScanRange10 = 20
    ''' <summary>
    ''' Scan Range +15
    ''' </summary>
    ''' <remarks></remarks>
    ScanRange15 = 21
    ''' <summary>
    ''' Colony Revenue +10%
    ''' </summary>
    ''' <remarks></remarks>
    Revenue10 = 22
    ''' <summary>
    ''' Launch Speed is 10% faster
    ''' </summary>
    ''' <remarks></remarks>
    Launch10 = 23
    ''' <summary>
    ''' Cargo capacity increased 20%
    ''' </summary>
    ''' <remarks></remarks>
    Cargo20 = 24
    ''' <summary>
    ''' Launch Speed is 20% faster
    ''' </summary>
    ''' <remarks></remarks>
    Launch20 = 25
    ''' <summary>
    ''' Maneuver +10
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver10 = 26
    ''' <summary>
    ''' +2 Critical hit chance
    ''' </summary>
    ''' <remarks></remarks>
    Critical2 = 27
    ''' <summary>
    ''' Maneuver +5 and +2 Critical Hit Chance
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver5Critical2 = 28
    ''' <summary>
    ''' Maneuver +10 and +2 Critical Hit Chance
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver10Critical2 = 29
    ''' <summary>
    ''' +1 Critical Hit Chance
    ''' </summary>
    ''' <remarks></remarks>
    Critical1 = 30
    ''' <summary>
    ''' Speed +5 and +2 Critical Hit Chance
    ''' </summary>
    ''' <remarks></remarks>
    Speed5Critical2 = 31
    ''' <summary>
    ''' Maneuver +10 and +1 Critical Hit Chance
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver10Critical1 = 32
    ''' <summary>
    ''' Speed +2
    ''' </summary>
    ''' <remarks></remarks>
    Speed2 = 33
    ''' <summary>
    ''' Cargo Capacity +10%
    ''' </summary>
    ''' <remarks></remarks>
    Cargo10 = 34
    ''' <summary>
    ''' Maneuver +30 
    ''' </summary>
    ''' <remarks></remarks>
    Maneuver30 = 35
    ''' <summary>
    ''' Cargo Capacity +200%
    ''' </summary>
    ''' <remarks></remarks>
    Cargo200 = 36
End Enum


''' <summary>
''' There should only be one instantiation of this class within the entire solution
''' </summary>
''' <remarks></remarks>
Public Class ModelDefs
    Private Shared moDefs() As ModelDef
    Private Shared ModelDefUB As Int32 = -1

    Public Shared Function GetModelDef(ByVal lModelID As Int32) As ModelDef
        'No values should ever be nothing... we load on demand...
        'MSC: we only care about the mesh portion of the modelid
        lModelID = (lModelID And 255)
        For X As Int32 = 0 To ModelDefUB
            If moDefs(X) Is Nothing = False AndAlso moDefs(X).ModelID = lModelID Then
                Return moDefs(X)
            End If
        Next X

        'If we are here, then we do not have it loaded
        ModelDefUB += 1
        ReDim Preserve moDefs(ModelDefUB)
        moDefs(ModelDefUB) = New ModelDef(lModelID)

        Return moDefs(ModelDefUB)
    End Function

    Public Shared Sub LoadAllModelDefs()
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        Dim oINI As InitFile = New InitFile(sFile & "RS_Model.dat")

        Dim lCnt As Int32 = CInt(oINI.GetString("MODEL_LIST", "ModelCnt", "0"))
        oINI = Nothing

        ModelDefUB = lCnt - 1
        ReDim moDefs(ModelDefUB)
        For X As Int32 = 1 To lCnt
            moDefs(X - 1) = New ModelDef(X)
        Next X
    End Sub
End Class

''' <summary>
''' This class should only be instantiated within the ModelDefs class
''' </summary>
''' <remarks></remarks>
Public Class ModelDef
    Private Shared msModelDefFile As String = String.Empty

    Public lModelRangeOffset As Int32
    Public lModelSizeXZ As Int32
    Public ModelID As Int32
    Public lSpecialTraitID As elModelSpecialTrait = 0

    Public yHullType As Epica_Entity.eyHullTypeDmgMod

    Public Sub New(ByVal lModelID As Int32)
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        Dim oINI As InitFile = New InitFile(sFile & "RS_Model.dat")
        Dim sModelHdr As String = "Model_" & lModelID

        'Now... load our model
        With Me
            .ModelID = lModelID
            .lModelRangeOffset = CInt(Val(oINI.GetString(sModelHdr, "RangeOffset", "0")))
            Dim lTemp As Int32 = lModelRangeOffset * gl_FINAL_GRID_SQUARE_SIZE
            .lModelSizeXZ = CInt(Val(oINI.GetString(sModelHdr, "ModelSizeXZ", lTemp.ToString)))
            .lSpecialTraitID = CType(CInt(Val(oINI.GetString(sModelHdr, "SpecTraitID", "0"))), elModelSpecialTrait)
            .yHullType = CType(CInt(Val(oINI.GetString(sModelHdr, "HullType", "0"))), Epica_Entity.eyHullTypeDmgMod)
        End With

        oINI = Nothing
    End Sub

 End Class
