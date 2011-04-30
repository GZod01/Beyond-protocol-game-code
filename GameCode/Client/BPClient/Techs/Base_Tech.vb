Public MustInherit Class Base_Tech
    Inherits Base_GUID

    Public Enum eyHullType As Byte
        BattleCruiser = 11
        Battleship = 12
        Corvette = 8
        Cruiser = 10
        Destroyer = 9
        Escort = 6
        Facility = 14
        Frigate = 7
        HeavyFighter = 4
        LightFighter = 0
        MediumFighter = 2
        NavalBattleship = 15
        NavalCarrier = 16
        NavalCruiser = 17
        NavalDestroyer = 18
        NavalFrigate = 19
        NavalSub = 20
        SmallVehicle = 1
        SpaceStation = 13
        Tank = 5
        Utility = 3
        NavalUtility = 21
        SpaceUtility = 22
    End Enum


    Public Enum eComponentDevelopmentPhase As Integer
        eInvalidDesign = -1
        eComponentDesign = 0        'initial phase, component is being designed
        eComponentResearching       'component is designed, now it is being researched
        eResearched                 'component was designed and was researched successfully 
    End Enum

    Public Enum eComponentDesignFlaw As Byte
        eNo_Particular_Design_Flaw = 0
        eMat1_Prop1 = 1
        eMat1_Prop2 = 2
        eMat1_Prop3 = 3
        eMat1_Prop4 = 4
        eMat1_Prop5 = 5
        eMat2_Prop1 = 6
        eMat2_Prop2 = 7
        eMat2_Prop3 = 8
        eMat2_Prop4 = 9
        eMat2_Prop5 = 10
        eMat3_Prop1 = 11
        eMat3_Prop2 = 12
        eMat3_Prop3 = 13
        eMat3_Prop4 = 14
        eMat3_Prop5 = 15
        eMat4_Prop1 = 16
        eMat4_Prop2 = 17
        eMat4_Prop3 = 18
        eMat4_Prop4 = 19
        eMat4_Prop5 = 20
        eMat5_Prop1 = 21
        eMat5_Prop2 = 22
        eMat5_Prop3 = 23
        eMat5_Prop4 = 24
        eMat5_Prop5 = 25
        eMat6_Prop1 = 26
        eMat6_Prop2 = 27
        eMat6_Prop3 = 28
        eMat6_Prop4 = 29
        eMat6_Prop5 = 30

		eGoodDesign = 32
		eShift_Should_Be_Higher = 64
		eShift_Not_Study = 128
    End Enum

    ''' <summary>
    ''' Cost to build the component physically
    ''' </summary>
    ''' <remarks></remarks>
    Public oProductionCost As ProductionCost = Nothing
    ''' <summary>
    ''' Cost to research the component
    ''' </summary>
    ''' <remarks></remarks>
    Public oResearchCost As ProductionCost = Nothing

    Public MustOverride Sub FillFromMsg(ByRef yData() As Byte)
    Public MustOverride Function GetDesignFlawText() As String
    Public MustOverride Sub FillFromPlayerTechKnowledge(ByRef yData() As Byte, ByVal lPos As Int32, ByVal yTechKnow As Byte, ByVal sName As String)
	Public MustOverride Function GetIntelReport(ByVal yLevel As PlayerTechKnowledge.KnowledgeType) As String

    Public OwnerID As Int32

    Public ErrorCodeReason As Byte
    Public Researchers As Byte              'number of entities that are currently researching this technology
    Public MajorDesignFlaw As Byte

    Public bArchived As Boolean = False

    Public lSpecifiedHull As Int32 = -1
    Public lSpecifiedPower As Int32 = -1
    Public lSpecifiedColonists As Int32 = -1
    Public lSpecifiedEnlisted As Int32 = -1
    Public lSpecifiedOfficers As Int32 = -1
    Public blSpecifiedProdCost As Int64 = -1
    Public blSpecifiedProdTime As Int64 = -1
    Public blSpecifiedResCost As Int64 = -1
    Public blSpecifiedResTime As Int64 = -1
    Public lSpecifiedMin1 As Int32 = -1
    Public lSpecifiedMin2 As Int32 = -1
    Public lSpecifiedMin3 As Int32 = -1
    Public lSpecifiedMin4 As Int32 = -1
    Public lSpecifiedMin5 As Int32 = -1
    Public lSpecifiedMin6 As Int32 = -1

    Private mlComponentDevelopmentPhase As Int32 = -1

    Public Property ComponentDevelopmentPhase() As Int32
        Get
            Return mlComponentDevelopmentPhase
        End Get
        Set(ByVal value As Int32)
            If mlComponentDevelopmentPhase <> value AndAlso mlComponentDevelopmentPhase <> -1 Then
                If mlComponentDevelopmentPhase = eComponentDevelopmentPhase.eComponentDesign Then
                    If goUILib Is Nothing = False Then
                        goUILib.AddNotification(GetComponentName() & " has been designed and is ready for research!", Color.Teal, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
                Else
                    If goUILib Is Nothing = False Then
                        goUILib.AddNotification(GetComponentName() & " has been successfully researched!", Color.Teal, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
                End If
            End If
            mlComponentDevelopmentPhase = value
        End Set
    End Property

    Public Function GetComponentName() As String
        Dim sTemp As String = ""
        Dim sDefault As String = ""
        Select Case Me.ObjTypeID
            Case ObjectType.eAlloyTech
                sTemp = CType(Me, AlloyTech).sAlloyName
                sDefault = "Alloy"
            Case ObjectType.eArmorTech
                sTemp = CType(Me, ArmorTech).sArmorName
                sDefault = "Armor"
            Case ObjectType.eEngineTech
                sTemp = CType(Me, EngineTech).sEngineName
                sDefault = "Engine"
                'Case ObjectType.eHangarTech
                '    sTemp = CType(Me, HangarTech).HangarName
                '    sDefault = "Hangar"
            Case ObjectType.eHullTech
                sTemp = CType(Me, HullTech).HullName
                sDefault = "Hull Design"
            Case ObjectType.ePrototype
                sTemp = CType(Me, PrototypeTech).PrototypeName
                sDefault = "Prototype"
            Case ObjectType.eRadarTech
                sTemp = CType(Me, RadarTech).sRadarName
                sDefault = "Radar"
            Case ObjectType.eShieldTech
                sTemp = CType(Me, ShieldTech).sShieldName
                sDefault = "Shield"
            Case ObjectType.eWeaponTech
                sTemp = CType(Me, WeaponTech).WeaponName
                sDefault = "Weapon"
            Case ObjectType.eSpecialTech
                sTemp = CType(Me, SpecialTech).sName
                sDefault = "Artifact"
            Case Else
                sTemp = "Unknown Component"
        End Select

        If sTemp = "" Then
            sTemp = sDefault & " " & Me.ObjectID
        End If
        Return sTemp
    End Function
    Public Sub SetComponentName(ByVal sNew As String)
        Dim sTemp As String = ""
        Dim sDefault As String = ""
        Select Case Me.ObjTypeID
            Case ObjectType.eAlloyTech
                CType(Me, AlloyTech).sAlloyName = sNew
            Case ObjectType.eArmorTech
                CType(Me, ArmorTech).sArmorName = sNew
            Case ObjectType.eEngineTech
                CType(Me, EngineTech).sEngineName = sNew
            Case ObjectType.eHullTech
                CType(Me, HullTech).HullName = sNew
            Case ObjectType.ePrototype
                CType(Me, PrototypeTech).PrototypeName = sNew
            Case ObjectType.eRadarTech
                CType(Me, RadarTech).sRadarName = sNew
            Case ObjectType.eShieldTech
                CType(Me, ShieldTech).sShieldName = sNew
            Case ObjectType.eWeaponTech
                CType(Me, WeaponTech).WeaponName = sNew
            Case ObjectType.eSpecialTech
                CType(Me, SpecialTech).sName = sNew
        End Select

    End Sub

    Public Shared Function GetComponentTypeName(ByVal iTypeID As Int16) As String
        Select Case Math.Abs(iTypeID)
            Case ObjectType.eAlloyTech
                Return "Alloy"
            Case ObjectType.eArmorTech
                Return "Armor"
            Case ObjectType.eEngineTech
                Return "Engine"
                'Case ObjectType.eHangarTech
                '    Return "Hangar"
            Case ObjectType.eHullTech
                Return "Hull"
            Case ObjectType.eRadarTech
                Return "Radar"
            Case ObjectType.eShieldTech
                Return "Shield"
            Case ObjectType.eWeaponTech
                Return "Weapon"
            Case ObjectType.ePrototype
                Return "Prototype"
            Case ObjectType.eSpecialTech
                Return "Special"
        End Select

        Return "Unknown"
    End Function

    Public Shared Function GetHullTypeName(ByVal yHullType As Byte) As String
        Select Case yHullType
            Case eyHullType.BattleCruiser
                Return "Battlecruiser"
            Case eyHullType.Battleship
                Return "Battleship"
            Case eyHullType.Corvette
                Return "Corvette"
            Case eyHullType.Cruiser
                Return "Cruiser"
            Case eyHullType.Destroyer
                Return "Destroyer"
            Case eyHullType.Escort
                Return "Escort"
            Case eyHullType.Facility
                Return "Facility"
            Case eyHullType.Frigate
                Return "Frigate"
            Case eyHullType.HeavyFighter
                Return "Heavy Fighter"
            Case eyHullType.LightFighter
                Return "Light Fighter"
            Case eyHullType.MediumFighter
                Return "Medium Fighter"
            Case eyHullType.NavalBattleship
                Return "Naval Battleship"
            Case eyHullType.NavalCarrier
                Return "Naval Carrier"
            Case eyHullType.NavalCruiser
                Return "Naval Cruiser"
            Case eyHullType.NavalDestroyer
                Return "Naval Destroyer"
            Case eyHullType.NavalFrigate
                Return "Naval Frigate"
            Case eyHullType.NavalSub
                Return "Naval Sub"
            Case eyHullType.SmallVehicle
                Return "Small Vehicle"
            Case eyHullType.SpaceStation
                Return "Space Station"
            Case eyHullType.Tank
                Return "Tank"
            Case eyHullType.Utility
                Return "Utility Vehicle"
            Case 255
                Return "Universal"
        End Select
        Return "Unknown"
    End Function

    Public Shared Function GetMinHullForType(ByVal lHullTypeID As Int32) As Int32
        Dim yTypeID As Byte = 0
        Dim ySubTypeID As Byte = 0
        Select Case lHullTypeID
            Case eyHullType.BattleCruiser
                yTypeID = 0 : ySubTypeID = 2
            Case eyHullType.Battleship
                yTypeID = 0 : ySubTypeID = 0
            Case eyHullType.Corvette
                yTypeID = 1 : ySubTypeID = 0
            Case eyHullType.Cruiser
                yTypeID = 1 : ySubTypeID = 1
            Case eyHullType.Destroyer
                yTypeID = 1 : ySubTypeID = 2
            Case eyHullType.Escort
                yTypeID = 1 : ySubTypeID = 4
            Case eyHullType.Facility
                yTypeID = 2 : ySubTypeID = 255
            Case eyHullType.Frigate
                yTypeID = 1 : ySubTypeID = 3
            Case eyHullType.HeavyFighter
                yTypeID = 3 : ySubTypeID = 2
            Case eyHullType.LightFighter
                yTypeID = 3 : ySubTypeID = 0
            Case eyHullType.MediumFighter
                yTypeID = 3 : ySubTypeID = 1
            Case eyHullType.NavalBattleship
                yTypeID = 8 : ySubTypeID = 0
            Case eyHullType.NavalCarrier
                yTypeID = 8 : ySubTypeID = 1
            Case eyHullType.NavalCruiser
                yTypeID = 8 : ySubTypeID = 2
            Case eyHullType.NavalDestroyer
                yTypeID = 8 : ySubTypeID = 3
            Case eyHullType.NavalFrigate
                yTypeID = 8 : ySubTypeID = 4
            Case eyHullType.NavalSub
                yTypeID = 8 : ySubTypeID = 5
            Case eyHullType.SmallVehicle
                yTypeID = 4 : ySubTypeID = 0
            Case eyHullType.SpaceStation
                yTypeID = 2 : ySubTypeID = 9
            Case eyHullType.Tank
                yTypeID = 5 : ySubTypeID = 1
            Case eyHullType.Utility
                yTypeID = 7 : ySubTypeID = 255
            Case 255
                Return 0
        End Select
        Return goModelDefs.GetMinHull(yTypeID, ySubTypeID)
    End Function

    Public Shared Function GetMaxHullForType(ByVal lHullTypeID As Int32) As Int32
        Dim yTypeID As Byte = 0
        Dim ySubTypeID As Byte = 0
        Select Case lHullTypeID
            Case eyHullType.BattleCruiser
                yTypeID = 0 : ySubTypeID = 2
            Case eyHullType.Battleship
                yTypeID = 0 : ySubTypeID = 0
            Case eyHullType.Corvette
                yTypeID = 1 : ySubTypeID = 0
            Case eyHullType.Cruiser
                yTypeID = 1 : ySubTypeID = 1
            Case eyHullType.Destroyer
                yTypeID = 1 : ySubTypeID = 2
            Case eyHullType.Escort
                yTypeID = 1 : ySubTypeID = 4
            Case eyHullType.Facility
                yTypeID = 2 : ySubTypeID = 255
            Case eyHullType.Frigate
                yTypeID = 1 : ySubTypeID = 3
            Case eyHullType.HeavyFighter
                yTypeID = 3 : ySubTypeID = 2
            Case eyHullType.LightFighter
                yTypeID = 3 : ySubTypeID = 0
            Case eyHullType.MediumFighter
                yTypeID = 3 : ySubTypeID = 1
            Case eyHullType.NavalBattleship
                yTypeID = 8 : ySubTypeID = 0
            Case eyHullType.NavalCarrier
                yTypeID = 8 : ySubTypeID = 1
            Case eyHullType.NavalCruiser
                yTypeID = 8 : ySubTypeID = 2
            Case eyHullType.NavalDestroyer
                yTypeID = 8 : ySubTypeID = 3
            Case eyHullType.NavalFrigate
                yTypeID = 8 : ySubTypeID = 4
            Case eyHullType.NavalSub
                yTypeID = 8 : ySubTypeID = 5
            Case eyHullType.SmallVehicle
                yTypeID = 4 : ySubTypeID = 0
            Case eyHullType.SpaceStation
                yTypeID = 2 : ySubTypeID = 9
            Case eyHullType.Tank
                yTypeID = 5 : ySubTypeID = 1
            Case eyHullType.Utility
                yTypeID = 7 : ySubTypeID = 255
            Case 255
                Return 0
        End Select
        Return goModelDefs.GetMaxHull(yTypeID, ySubTypeID)
    End Function

    Public Function GetResearchMainText() As String
        Dim sValue As String = GetComponentName()

        If sValue.Length > 30 Then sValue = sValue.Substring(0, 30)
        sValue = sValue.PadRight(32, " "c)

        If Researchers <> 0 Then
            sValue &= "Assigned (" & Researchers & ")"
        Else
            Select Case ComponentDevelopmentPhase
                Case Base_Tech.eComponentDevelopmentPhase.eComponentDesign
                    If ObjTypeID = ObjectType.eSpecialTech Then sValue &= "Unresearched" Else sValue &= "In Design"
                Case Base_Tech.eComponentDevelopmentPhase.eComponentResearching
                    If ObjTypeID = ObjectType.eSpecialTech Then sValue &= "Unresearched" Else sValue &= "Designed"
                Case Base_Tech.eComponentDevelopmentPhase.eResearched
                    sValue &= "Researched"
                Case Base_Tech.eComponentDevelopmentPhase.eInvalidDesign
                    sValue &= "Poor Design" ' Design"
            End Select
        End If
        Return sValue
    End Function

    Public Function GetTechHullTypeID() As Int32
        Select Case Me.ObjTypeID
            Case ObjectType.eEngineTech
                Return CType(Me, EngineTech).HullTypeID
            Case ObjectType.eHullTech
                With CType(Me, HullTech)
                    Return HullTech.GetHullTypeID(.yTypeID, .ySubTypeID)
                End With
            Case ObjectType.ePrototype
                Dim oHull As HullTech = CType(goCurrentPlayer.GetTech(CType(Me, PrototypeTech).HullID, ObjectType.eHullTech), HullTech)
                If oHull Is Nothing = False Then
                    Return HullTech.GetHullTypeID(oHull.yTypeID, oHull.ySubTypeID)
                Else
                    If Me.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        For X As Int32 = 0 To glEntityDefUB
                            If glEntityDefIdx(X) > -1 Then
                                Dim oDef As EntityDef = goEntityDefs(X)
                                If oDef Is Nothing = False AndAlso oDef.PrototypeID = Me.ObjectID Then
                                    Dim lModelID As Int32 = oDef.ModelID
                                    lModelID = lModelID And 255

                                    Dim oModelDef As ModelDef = goModelDefs.GetModelDef(lModelID)
                                    If oModelDef Is Nothing = False Then
                                        Return HullTech.GetHullTypeID(oModelDef.TypeID, oModelDef.SubTypeID)
                                    End If
                                    Exit For
                                End If
                            End If
                        Next X
                    End If
                End If
                Return 255
            Case ObjectType.eRadarTech
                Return CType(Me, RadarTech).HullTypeID
            Case ObjectType.eShieldTech
                Return CType(Me, ShieldTech).HullTypeID
            Case ObjectType.eWeaponTech
                Return CType(Me, WeaponTech).HullTypeID
        End Select

        Return 0
    End Function
End Class
