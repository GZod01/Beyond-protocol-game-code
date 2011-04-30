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
    Private moDefs() As ModelDef
    Public ModelDefUB As Int32 = -1

    Public Function GetModelDef(ByVal lModelID As Int32) As ModelDef
        'MSC: we only care about the mesh portion of the model
        lModelID = (lModelID And 255)
        'No values should ever be nothing... we load on demand...
        For X As Int32 = 0 To ModelDefUB
            If moDefs(X).ModelID = lModelID Then
                Return moDefs(X)
            End If
        Next X

        'If we are here, then we do not have it loaded
        ModelDefUB += 1
        ReDim Preserve moDefs(ModelDefUB)
        moDefs(ModelDefUB) = New ModelDef(lModelID)

        If moDefs(ModelDefUB).bLoaded = False Then
            ModelDefUB -= 1
            ReDim Preserve moDefs(ModelDefUB)
            Return Nothing
        End If

        Return moDefs(ModelDefUB)
    End Function

    Public Sub AddLabelScrollVals(ByRef lscVals As UILabelScroller, ByVal yTypeID As Byte, ByVal ySubTypeID As Byte)
        Dim sFile As String = goResMgr.MeshPath
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "FrameDefs.dat"

        Dim oINI As InitFile = New InitFile(sFile)
        Dim sSection As String = "FRAME_" & yTypeID & "_" & ySubTypeID
        Dim lCnt As Int32 = CInt(Val(oINI.GetString(sSection, "Cnt", "0")))
        Dim sTemp As String
        Dim lTemp As Int32
        Dim lMaxHull As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)

        Dim bHasNavalUnit As Boolean = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 0
        Dim bHasNavalFac As Boolean = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 1

        Dim bHasOMP As Boolean = (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSuperSpecials) And 16) <> 0

        'lMaxHull = 10000000
        For X As Int32 = 0 To lCnt - 1
            sTemp = oINI.GetString(sSection, "Model" & X, "0")
			If CInt(Val(sTemp)) <> 0 Then ' AndAlso CInt(Val(sTemp)) <> 24 Then		'TODO: Remove the cint(val(stemp)) <> 24 portion
				lTemp = CInt(Val(sTemp))

				Dim oModelDef As ModelDef = GetModelDef(lTemp)
                If oModelDef Is Nothing = False Then
                    If (oModelDef.ModelID And 255) = 148 AndAlso bHasOMP = False Then Continue For
                    If oModelDef.TypeID = 2 OrElse oModelDef.MinHull < lMaxHull Then

                        If oModelDef.bRequiresClaim = True Then
                            Dim bGood As Boolean = False
                            If guClaimables Is Nothing = False Then
                                For Y As Int32 = 0 To guClaimables.GetUpperBound(0)
                                    If guClaimables(Y).iTypeID = ObjectType.eHullTech AndAlso guClaimables(Y).lID = oModelDef.ModelID Then
                                        If (guClaimables(Y).yClaimFlag And eyClaimFlags.eClaimed) <> 0 Then
                                            bGood = True
                                            Exit For
                                        End If
                                    End If
                                Next Y
                            End If
                            If bGood = False AndAlso oModelDef.ModelID = 141 Then
                                'ok, check if the player is in one of the 3 guilds
                                If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
                                    Dim lGuildID As Int32 = goCurrentPlayer.oGuild.ObjectID
                                    If lGuildID = 13 OrElse lGuildID = 14 OrElse lGuildID = 17 Then
                                        bGood = True
                                    End If
                                End If
                            End If
                            If bGood = False Then Continue For
                        End If

                        If oModelDef.FrameTypeID = 3 Then
                            'naval item.. is it a unit or facility?
                            If oModelDef.TypeID = 2 Then
                                If bHasNavalFac = False AndAlso (bHasNavalUnit = False OrElse oModelDef.SubTypeID <> 5) Then Continue For
                            ElseIf bHasNavalUnit = False Then
                                Continue For
                            End If
                        End If

                        sTemp = oINI.GetString(sSection, "Model" & X & "Name", "")
                        If sTemp <> "" Then
                            lscVals.AddItem(lTemp, sTemp)
                        End If
                    End If
                End If
			End If
        Next X

        oINI = Nothing
    End Sub

    Private mlMaxModelDefUB As Int32 = -1
    Public ReadOnly Property MaxModelDefUB() As Int32
        Get
            If mlMaxModelDefUB = -1 Then
                Dim sModelDefFile As String = goResMgr.MeshPath()
                If sModelDefFile.EndsWith("\") = False Then sModelDefFile &= "\"
                sModelDefFile &= "ModelDef.dat"
                Dim oINI As InitFile = New InitFile(sModelDefFile)
                Dim bDone As Boolean = False
                Dim lIdx As Int32 = 0
                While bDone = False
                    lIdx += 1
                    bDone = oINI.GetString("MODEL_" & lIdx, "FrameName", "") = ""
                End While
                oINI = Nothing
                mlMaxModelDefUB = lIdx - 1
            End If
            Return mlMaxModelDefUB
        End Get
    End Property

    Public Function GetMinHull(ByVal yTypeID As Byte, ByVal ySubTypeID As Byte) As Int32
        EnsureAllModelsLoaded()
        Dim lMin As Int32 = Int32.MaxValue
        For X As Int32 = 0 To ModelDefUB
            If moDefs(X).TypeID = yTypeID AndAlso (moDefs(X).SubTypeID = ySubTypeID OrElse ySubTypeID = 255) Then
                lMin = Math.Min(moDefs(X).MinHull, lMin)
            End If
        Next X
        Return lMin
    End Function
    Public Function GetMaxHull(ByVal yTypeID As Byte, ByVal ySubTypeID As Byte) As Int32
        EnsureAllModelsLoaded()
        Dim lMax As Int32 = 0
        For X As Int32 = 0 To ModelDefUB
            If moDefs(X).TypeID = yTypeID AndAlso (moDefs(X).SubTypeID = ySubTypeID OrElse ySubTypeID = 255) Then
                lMax = Math.Max(moDefs(X).MaxHull, lMax)
            End If
        Next X
        Return lMax
    End Function

    Private mbAllModelsLoaded As Boolean = False
    Private Sub EnsureAllModelsLoaded()
        If mbAllModelsLoaded = True Then Return
        mbAllModelsLoaded = True
        For X As Int32 = 1 To 138
            Dim oD As ModelDef = GetModelDef(X)
        Next X
    End Sub
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
    Public FrameName As String
    Public MinHull As Int32
    Public MaxHull As Int32
    Public FrameTypeID As Byte
    Public CrewReqPerc As Byte

    Public lSpecialTraitID As Int32 = 0

    Public FrontLocs() As Int32
    Public LeftLocs() As Int32
    Public RightLocs() As Int32
    Public RearLocs() As Int32
    Public AllArcLocs() As Int32

    Public bMTTexCapable As Boolean = False

    Public bRequiresClaim As Boolean = False

    Public bLoaded As Boolean = False

    Public Sub New(ByVal lModelID As Int32)
        'MSC: only the mesh portion of a model is important
        lModelID = (lModelID And 255)

        If msModelDefFile.Length = 0 Then
            msModelDefFile = goResMgr.MeshPath()
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
            .FrameName = oINI.GetString(sSection, "FrameName", "")
            .MinHull = CInt(Val(oINI.GetString(sSection, "MinHull", "0")))
            .MaxHull = CInt(Val(oINI.GetString(sSection, "MaxHull", "0")))
            .FrameTypeID = CByte(Val(oINI.GetString(sSection, "FrameTypeID", "0")))
            .CrewReqPerc = CByte(Val(oINI.GetString(sSection, "CrewReq", "0")))
            .lSpecialTraitID = CInt(Val(oINI.GetString(sSection, "SpecTraitID", "0")))
            .bMTTexCapable = CInt(Val(oINI.GetString(sSection, "MTTexCapable", "0"))) <> 0
            .ModelID = lModelID
            .bRequiresClaim = CInt(Val(oINI.GetString(sSection, "POffSet", "0"))) <> 0

            If .ModelID >= 139 AndAlso .ModelID <= 147 Then
                .bRequiresClaim = True
            End If
        End With

        If Me.FrameName = "" Then
            bLoaded = False
            Return
        Else : bLoaded = True
        End If

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

    Public Shared Function GetSpecialTraitText(ByVal lTraitID As Int32) As String
        Dim sResult As String = ""
        Select Case CType(lTraitID, elModelSpecialTrait)
            Case elModelSpecialTrait.Armor1000
                sResult = "Armor hitpoints +1000%"
            Case elModelSpecialTrait.Cargo10
                sResult = "Cargo Capacity +10%"
            Case elModelSpecialTrait.Cargo200
                sResult = "Cargo Capacity +200%"
            Case elModelSpecialTrait.Cargo20
                sResult = "Cargo Capacity +20%"
            Case elModelSpecialTrait.Cargo5
                sResult = "Cargo Capacity +5%"
            Case elModelSpecialTrait.CargoAndHangar10
                sResult = "Cargo and Hangar Capacity +10%"
            Case elModelSpecialTrait.CargoAndHangar3
                sResult = "Cargo and Hangar Capacity +3%"
            Case elModelSpecialTrait.CargoAndHangar6
                sResult = "Cargo and Hangar Capacity +6%"
            Case elModelSpecialTrait.Critical1
                sResult = "+1% Critical Hit Chance"
            Case elModelSpecialTrait.Critical2
                sResult = "+2% Critical Hit Chance"
            Case elModelSpecialTrait.Hangar10
                sResult = "Hangar Capacity +10%"
            Case elModelSpecialTrait.Hangar5
                sResult = "Hangar Capacity +5%"
            Case elModelSpecialTrait.Launch10
                sResult = "10% Faster Hangar Launches"
            Case elModelSpecialTrait.Launch20
                sResult = "20% Faster Hangar Launches"
            Case elModelSpecialTrait.Launch6
                sResult = "6% Faster Hangar Launches"
            Case elModelSpecialTrait.Maneuver1
                sResult = "Maneuver +1"
            Case elModelSpecialTrait.Maneuver10
                sResult = "Maneuver +10"
            Case elModelSpecialTrait.Maneuver10Critical1
                sResult = "Maneuver +10 and +1% Critical Hit Chance"
            Case elModelSpecialTrait.Maneuver10Critical2
                sResult = "Maneuver +10 and +2% Critical Hit Chance"
            Case elModelSpecialTrait.Maneuver2
                sResult = "Maneuver +2"
            Case elModelSpecialTrait.Maneuver3
                sResult = "Maneuver +3"
            Case elModelSpecialTrait.Maneuver30
                sResult = "Maneuver +30"
            Case elModelSpecialTrait.Maneuver4
                sResult = "Maneuver +4"
            Case elModelSpecialTrait.Maneuver5
                sResult = "Maneuver +5"
            Case elModelSpecialTrait.Maneuver5Critical2
                sResult = "Maneuver +5 and +2% Critical Hit Chance"
            Case elModelSpecialTrait.NoJammer
                sResult = "No Jamming Radar Capability"
            Case elModelSpecialTrait.PowerGen2
                sResult = "Engine Power Production is 200%"
            Case elModelSpecialTrait.PowerGen3
                sResult = "Engine Power Production is 300%"
            Case elModelSpecialTrait.Revenue10
                sResult = "Produces 10% More Revenue"
            Case elModelSpecialTrait.Revenue20
                sResult = "Produces 20% More Revenue"
            Case elModelSpecialTrait.ScanRange10
                sResult = "Radar Ranges +10"
            Case elModelSpecialTrait.ScanRange15
                sResult = "Radar Ranges +15"
            Case elModelSpecialTrait.Speed1
                sResult = "Speed +1"
            Case elModelSpecialTrait.Speed2
                sResult = "Speed +2"
            Case elModelSpecialTrait.Speed5Critical2
                sResult = "Speed +5 and +2% Critical Hit Chance"
            Case elModelSpecialTrait.SpeedAndManeuver1
                sResult = "Speed +1 and Maneuver +1"
        End Select
        Return sResult
    End Function
End Class
