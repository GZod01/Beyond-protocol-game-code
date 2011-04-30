
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
    Private moDefs() As ModelDef
    Public ModelDefUB As Int32 = -1

    Public Function GetModelDef(ByVal lModelID As Int32) As ModelDef
        'No values should ever be nothing... we load on demand...

        'MSC: We only care about the mesh portion of the modelid
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
End Class

''' <summary>
''' This class should only be instantiated within the ModelDefs class
''' </summary>
''' <remarks></remarks>
Public Class ModelDef
    Private Shared msModelDefFile As String = String.Empty

    Public ModelID As Int32
    Public TypeID As Byte
    Public SubTypeID As Byte
    Public FrameName(19) As Byte
    Public MinHull As Int32
    Public MaxHull As Int32
    Public FrameTypeID As Byte
    Public CrewReqPerc As Byte

    Public lSpecialTraitID As Int32 = 0

    Public rcSpacing As Rectangle

    Public FrontLocs() As Int32
    Public LeftLocs() As Int32
    Public RightLocs() As Int32
    Public RearLocs() As Int32
    Public AllArcLocs() As Int32

    Public bRequiresClaim As Boolean = False

    Public Sub New(ByVal lModelID As Int32)
        If msModelDefFile.Length = 0 Then
            msModelDefFile = AppDomain.CurrentDomain.BaseDirectory
            If msModelDefFile.EndsWith("\") = False Then msModelDefFile &= "\"
            msModelDefFile &= "ModelDef.dat"
        End If
        Dim oINI As InitFile = New InitFile(msModelDefFile)
        Dim sSection As String = "Model_" & lModelID
        Dim sTemp As String
        Dim sVals() As String

        'Now... load our model
        With Me
            .TypeID = CByte(Val(oINI.GetString(sSection, "TypeID", "0")))
            .SubTypeID = CByte(Val(oINI.GetString(sSection, "SubTypeID", "0")))
            .FrameName = StringToBytes(oINI.GetString(sSection, "FrameName", "Unnamed"))
            .MinHull = CInt(Val(oINI.GetString(sSection, "MinHull", "0")))
            .MaxHull = CInt(Val(oINI.GetString(sSection, "MaxHull", "0")))
            .FrameTypeID = CByte(Val(oINI.GetString(sSection, "FrameTypeID", "0")))
            .CrewReqPerc = CByte(Val(oINI.GetString(sSection, "CrewReq", "0")))
            .lSpecialTraitID = CInt(Val(oINI.GetString(sSection, "SpecTraitID", "0")))
            Dim lSpacing As Int32 = CInt(Val(oINI.GetString(sSection, "Spacing", "1")))
            rcSpacing = New Rectangle(0, 0, lSpacing, lSpacing)
            .ModelID = lModelID
            .bRequiresClaim = CInt(Val(oINI.GetString(sSection, "POffset", "0"))) <> 0
            If .ModelID >= 139 AndAlso .ModelID <= 147 Then
                .bRequiresClaim = True
            End If
        End With

        'Load our Forward...
        sTemp = oINI.GetString(sSection, "F", "")
        sVals = Split(sTemp, ",")
        ReDim FrontLocs(sVals.GetUpperBound(0))
        For lIdx As Int32 = 0 To sVals.GetUpperBound(0)
            If sVals(lIdx) <> "" Then
                FrontLocs(lIdx) = CInt(Val(sVals(lIdx)))
            Else
                ReDim Preserve FrontLocs(FrontLocs.GetUpperBound(0) - 1)
            End If
        Next lIdx

        'Load our Left...
        sTemp = oINI.GetString(sSection, "L", "")
        sVals = Split(sTemp, ",")
        ReDim LeftLocs(sVals.GetUpperBound(0))
        For lIdx As Int32 = 0 To sVals.GetUpperBound(0)
            If sVals(lIdx) <> "" Then
                LeftLocs(lIdx) = CInt(Val(sVals(lIdx)))
            Else
                ReDim Preserve LeftLocs(LeftLocs.GetUpperBound(0) - 1)
            End If
        Next lIdx

        'Load our Rear...
        sTemp = oINI.GetString(sSection, "B", "")
        sVals = Split(sTemp, ",")
        ReDim RearLocs(sVals.GetUpperBound(0))
        For lIdx As Int32 = 0 To sVals.GetUpperBound(0)
            If sVals(lIdx) <> "" Then
                RearLocs(lIdx) = CInt(Val(sVals(lIdx)))
            Else
                ReDim Preserve RearLocs(RearLocs.GetUpperBound(0) - 1)
            End If
        Next lIdx

        'Load our Right...
        sTemp = oINI.GetString(sSection, "R", "")
        sVals = Split(sTemp, ",")
        ReDim RightLocs(sVals.GetUpperBound(0))
        For lIdx As Int32 = 0 To sVals.GetUpperBound(0)
            If sVals(lIdx) <> "" Then
                RightLocs(lIdx) = CInt(Val(sVals(lIdx)))
            Else
                ReDim Preserve RightLocs(RightLocs.GetUpperBound(0) - 1)
            End If
        Next lIdx

        'Load our All Arc...
        sTemp = oINI.GetString(sSection, "A", "")
        sVals = Split(sTemp, ",")
        ReDim AllArcLocs(sVals.GetUpperBound(0))
        For lIdx As Int32 = 0 To sVals.GetUpperBound(0)
            If sVals(lIdx) <> "" Then
                AllArcLocs(lIdx) = CInt(Val(sVals(lIdx)))
            Else
                ReDim Preserve AllArcLocs(AllArcLocs.GetUpperBound(0) - 1)
            End If
        Next lIdx

        oINI = Nothing
    End Sub

    'Private mySendString() As Byte
    'Public Function GetModelDefAddString() As Byte()
    '    If mySendString Is Nothing Then

    '        Dim lF As Int32 = 0
    '        Dim lL As Int32 = 0
    '        Dim lB As Int32 = 0
    '        Dim lR As Int32 = 0
    '        Dim lA As Int32 = 0

    '        If FrontLocs Is Nothing = False Then lF = FrontLocs.GetUpperBound(0) + 1
    '        If LeftLocs Is Nothing = False Then lL = LeftLocs.GetUpperBound(0) + 1
    '        If RearLocs Is Nothing = False Then lB = RearLocs.GetUpperBound(0) + 1
    '        If RightLocs Is Nothing = False Then lR = RightLocs.GetUpperBound(0) + 1
    '        If AllArcLocs Is Nothing = False Then lA = AllArcLocs.GetUpperBound(0) + 1

    '        ReDim mySendString(((lF + lL + lB + lR + lA) * 2) + 51)

    '        Dim lPos As Int32 = 0
    '        System.BitConverter.GetBytes(EpicaMessageCode.eAddModelDef).CopyTo(mySendString, lPos) : lPos += 2
    '        System.BitConverter.GetBytes(ModelID).CopyTo(mySendString, lPos) : lPos += 4
    '        mySendString(lPos) = TypeID : lPos += 1
    '        mySendString(lPos) = SubTypeID : lPos += 1
    '        FrameName.CopyTo(mySendString, lPos) : lPos += 20
    '        System.BitConverter.GetBytes(MinHull).CopyTo(mySendString, lPos) : lPos += 4
    '        System.BitConverter.GetBytes(MaxHull).CopyTo(mySendString, lPos) : lPos += 4
    '        mySendString(lPos) = FrameTypeID : lPos += 1
    '        mySendString(lPos) = CrewReqPerc : lPos += 1
    '        Dim lMax As Int32 = Math.Max(rcSpacing.Width, rcSpacing.Height)
    '        System.BitConverter.GetBytes(lMax).CopyTo(mySendString, lPos) : lPos += 4

    '        System.BitConverter.GetBytes(CShort(lF)).CopyTo(mySendString, lPos) : lPos += 2
    '        System.BitConverter.GetBytes(CShort(lL)).CopyTo(mySendString, lPos) : lPos += 2
    '        System.BitConverter.GetBytes(CShort(lB)).CopyTo(mySendString, lPos) : lPos += 2
    '        System.BitConverter.GetBytes(CShort(lR)).CopyTo(mySendString, lPos) : lPos += 2
    '        System.BitConverter.GetBytes(CShort(lA)).CopyTo(mySendString, lPos) : lPos += 2

    '        If FrontLocs Is Nothing = False Then
    '            For X As Int32 = 0 To FrontLocs.GetUpperBound(0)
    '                System.BitConverter.GetBytes(CShort(FrontLocs(X))).CopyTo(mySendString, lPos) : lPos += 2
    '            Next X
    '        End If
    '        If LeftLocs Is Nothing = False Then
    '            For X As Int32 = 0 To LeftLocs.GetUpperBound(0)
    '                System.BitConverter.GetBytes(CShort(LeftLocs(X))).CopyTo(mySendString, lPos) : lPos += 2
    '            Next X
    '        End If
    '        If RearLocs Is Nothing = False Then
    '            For X As Int32 = 0 To RearLocs.GetUpperBound(0)
    '                System.BitConverter.GetBytes(CShort(RearLocs(X))).CopyTo(mySendString, lPos) : lPos += 2
    '            Next X
    '        End If
    '        If RightLocs Is Nothing = False Then
    '            For X As Int32 = 0 To RightLocs.GetUpperBound(0)
    '                System.BitConverter.GetBytes(CShort(RightLocs(X))).CopyTo(mySendString, lPos) : lPos += 2
    '            Next X
    '        End If
    '        If AllArcLocs Is Nothing = False Then
    '            For X As Int32 = 0 To AllArcLocs.GetUpperBound(0)
    '                System.BitConverter.GetBytes(CShort(AllArcLocs(X))).CopyTo(mySendString, lPos) : lPos += 2
    '            Next X
    '        End If
    '    End If
    '    Return mySendString
    'End Function
End Class
