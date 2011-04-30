Option Strict On

Public Class PlayerTechKnowledge
    Public Enum KnowledgeType As Byte
        eNameOnly = 0                   'indicates that only the name of the tech is known
        eSettingsLevel1 = 1             'indicates that only lvl 1 stats are known, this varies per component
        eSettingsLevel2 = 2             'indicates that all stats are known except minerals used
        eFullKnowledge = 3              'indicates all stats are known including minerals used
    End Enum

    Public oPlayer As Player                    'reference to the player who has the knowledge
    Public oTech As Base_Tech                  'the technology that the knowledge is about
	Public yKnowledgeType As KnowledgeType		'an indication of what knowledge level the player has about the tech

	Public bArchived As Boolean = False

    Public Function FillFromMsg(ByRef yData() As Byte, ByVal lPos As Int32) As Boolean
        'The default attributes (always applicable)
        Dim lTechID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sName As String '= GetStringFromBytes(yData, lPos, 20) : lPos += 20
        If iTypeID = ObjectType.eSpecialTech Then
            sName = GetStringFromBytes(yData, lPos, 50) : lPos += 50
        Else
            sName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        End If

        'Now, based on the typeid, we can determine what type of tech it is
        Select Case iTypeID
            Case ObjectType.eAlloyTech
                oTech = New AlloyTech
            Case ObjectType.eArmorTech
                oTech = New ArmorTech
            Case ObjectType.eEngineTech
                oTech = New EngineTech
            Case ObjectType.eHullTech
                oTech = New HullTech
            Case ObjectType.ePrototype
                oTech = New PrototypeTech
            Case ObjectType.eRadarTech
                oTech = New RadarTech
            Case ObjectType.eShieldTech
                oTech = New ShieldTech
            Case ObjectType.eSpecialTech
                oTech = New SpecialTech
            Case ObjectType.eWeaponTech
                Dim yWpnClassType As Byte = yData(lPos) : lPos += 1

                Select Case yWpnClassType
                    Case WeaponClassType.eBomb
                    Case WeaponClassType.eEnergyBeam
                        oTech = New SolidBeamWeaponTech()
                    Case WeaponClassType.eEnergyPulse
                        oTech = New PulseWeaponTech()
                    Case WeaponClassType.eMine
                    Case WeaponClassType.eMissile
                        oTech = New MissileWeaponTech()
                    Case WeaponClassType.eProjectile
                        oTech = New ProjectileWeaponTech()
                End Select
                If oTech Is Nothing = False Then CType(oTech, WeaponTech).WeaponClassTypeID = CType(yWpnClassType, WeaponClassType)
            Case Else
                Return False
        End Select

        If oTech Is Nothing = False Then
            oTech.ObjectID = lTechID
            oTech.ObjTypeID = iTypeID
            oTech.OwnerID = lOwnerID
            oTech.FillFromPlayerTechKnowledge(yData, lPos, yKnowledgeType, sName)
        Else : Return False
        End If

        Return True
	End Function

	Public Shared Function GetGroupValue(ByVal iTypeID As Int16) As Int32
		Select Case iTypeID
			Case ObjectType.eAlloyTech
				Return 101
			Case ObjectType.eArmorTech
				Return 102
			Case ObjectType.eEngineTech
				Return 103
			Case ObjectType.eHullTech
				Return 104
			Case ObjectType.ePrototype
				Return 105
			Case ObjectType.eRadarTech
				Return 106
			Case ObjectType.eShieldTech
				Return 107
			Case ObjectType.eSpecialTech
				Return 108
			Case ObjectType.eWeaponTech
				Return 109
			Case Else
				Return 300
		End Select
	End Function
	Public Shared Function GetGroupName(ByVal iTypeID As Int16) As String
		Select Case iTypeID
			Case ObjectType.eAlloyTech
				Return "ALLOY"
			Case ObjectType.eArmorTech
				Return "ARMOR"
			Case ObjectType.eEngineTech
				Return "ENGINE"
			Case ObjectType.eHullTech
				Return "HULL"
			Case ObjectType.ePrototype
				Return "PROTOTYPE"
			Case ObjectType.eRadarTech
				Return "RADAR"
			Case ObjectType.eShieldTech
				Return "SHIELD"
			Case ObjectType.eSpecialTech
				Return "SPECIAL"
			Case ObjectType.eWeaponTech
				Return "WEAPON"
			Case Else
				Return "OTHER"
		End Select
	End Function
End Class
