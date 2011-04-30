
Public Class EntityDef
    Inherits Base_GUID

    Public DefName As String

    Private myManeuver As Byte     '

	Public TurnAmount As Byte	'based on Maneuver... equals Maneuver - 5 with a minimum of 1
	Public MaxSpeed As Byte

	Public yFormationMaxSpeed As Byte = 0
	Public yFormationManeuver As Byte = 0
	Public fFormationAcceleration As Single = 0.0F

    Public OptRadarRange As Byte
    Public MaxRadarRange As Byte
    Public HullSize As Int32
    Public Cargo_Cap As Int32
    Public Hangar_Cap As Int32

    'Public ModelID As Short             'from hull tech
    Private miFullModelID As Int16
    Public Property ModelID() As Int16
        Get
            Return miFullModelID
        End Get
        Set(ByVal value As Int16)
            miFullModelID = value

            Dim lModelID As Int32 = (value And 255)
            If lModelID = 141 Then
                lTexNum = (value And 7936S) \ 256S  'fine
                lIllumNum = (value And 24576S) \ 8192S
                If (value And -32768) = 0 Then lNormalNum = 0 Else lNormalNum = 1
            Else
                lTexNum = (value And 7936S) \ 256S
                lNormalNum = (value And 24576S) \ 8192S
                If (value And -32768) = 0 Then lIllumNum = 0 Else lIllumNum = 1
            End If
            
        End Set
    End Property
    Public lTexNum As Int32
    Public lNormalNum As Int32
    Public lIllumNum As Int32

    'From the Armor Tech
    Public PiercingResist As Byte
    Public ImpactResist As Byte
    Public BeamResist As Byte
    Public ECMResist As Byte
    Public FlameResist As Byte
    Public ChemicalResist As Byte
    Public Armor_MaxHP(3) As Int32      'for each side
    Public yArmorIntegrityRoll As Byte
    Public ReadOnly Property lDisplayedIntegrity() As Int32
        Get
            Dim lTemp As Int32 = yArmorIntegrityRoll * 2
			If lTemp = 0 Then Return 1000
            Return CInt((1.0F / lTemp) * 1000)
        End Get
    End Property
    'From the Shield Tech
    Public Shield_MaxHP As Int32
    Public ShieldRecharge As Int32
    Public ShieldRechargeFreq As Int32

    Public Structure_MaxHP As Int32

    Public Acceleration As Single = 0.001F  'based on Maneuver... equals Maneuver / 100

    'FOR OPTIMIZATION PURPOSES
    Public FOW_OptRadarRange As Int32
    Public FOW_MaxRadarRange As Int32

    Public Weapon_Acc As Byte
    Public ScanResolution As Byte
    Public DisruptionResistance As Byte
    Public JamImmunity As Byte
    Public JamStrength As Byte
    Public JamTargets As Byte
    Public JamEffect As Byte

    Public WeaponDefs() As WeaponDef
    Public WeaponDefUB As Int32 = -1

    Public RequiredProductionTypeID As Byte     'in order to make me... this production type is required from the builder (includes Level)
    Public ProductionTypeID As Byte
    Public ProdFactor As Int32
    Public PowerFactor As Int32
    Public WorkerFactor As Int32

    Public yChassisType As Byte
    Public yFXColors As Byte

	Public ProductionCost As ProductionCost

    Public lExtendedFlags As Int32 = 0

	Private mlPrototypeID As Int32 = -1
	Public ReadOnly Property PrototypeID() As Int32
		Get
			If mlPrototypeID = -1 Then
				For X As Int32 = 0 To goCurrentPlayer.mlTechUB
					Dim oTech As Base_Tech = goCurrentPlayer.moTechs(X)
					If oTech Is Nothing = False Then
						If oTech.ObjTypeID = ObjectType.ePrototype Then
							Dim oPrototype As PrototypeTech = CType(oTech, PrototypeTech)
							If oPrototype Is Nothing = False Then
								If oPrototype.oActualResults Is Nothing = False Then
									If oPrototype.oActualResults.ObjectID = Me.ObjectID Then
										mlPrototypeID = oPrototype.ObjectID
										Exit For
									End If
								End If
							End If
						End If
					End If
				Next X
			End If
			Return mlPrototypeID
		End Get
	End Property

    Public Property Maneuver() As Byte
        Get
            Return myManeuver
        End Get
        Set(ByVal Value As Byte)
            myManeuver = Value
            ''Now, set our Turn Rate, Turn Amount and Acceleration values
            'If myManeuver < 5 Then
            '    Select Case myManeuver
            '        Case 1
            '            TurnRate = 10
            '        Case 2
            '            TurnRate = 4
            '        Case 3
            '            TurnRate = 3
            '        Case 4
            '            TurnRate = 2
            '    End Select
            '    TurnAmount = 0
            'Else
            '    TurnRate = 1
            '    TurnAmount = myManeuver - 4
            'End If
            'If TurnAmount < 1 Then TurnAmount = 1
            TurnAmount = myManeuver
            Acceleration = CSng(myManeuver / 100)
        End Set
    End Property
End Class
