Option Strict On

Public Class ArenaContestant
    Public oEntity As Epica_Entity      'the entity (pointer)

    Public lRespawns As Int32 = 0
    Public lNextRespawnCycle As Int32 = -1

#Region "  Original Stats - when the arena is over, the unit gets returned to these stats  "

    Public OrigParentObject As Object
    Public OrigOwner As Player
    Public OrigLocX As Int32
    Public OrigLocZ As Int32
    Public OrigLocAngle As Int16

    Public OrigQ1_HP As Int32
    Public OrigQ2_HP As Int32
    Public OrigQ3_HP As Int32
    Public OrigQ4_HP As Int32

    Public OrigStructure_HP As Int32    'remaining structural hitpoints
    Public OrigShield_HP As Int32       'current shield hp
    Public OrigExpLevel As Byte = 0
    Public OrigCurrentAmmo() As Int32          'NOT REAL TIME, they are used to store the data over long term
    Public OrigCurrentStatus As Int32           'bit-wise representation of the ship's status and activity

    'Hangar and Cargo Holds... 
    Public OrigHangarContents() As Epica_GUID
    Public OrigHangarIdx() As Int32
    Public OrigHangarUB As Int32 = -1
    Public OrigCargoContents() As Epica_GUID
    Public OrigCargoIdx() As Int32
    Public OrigCargoUB As Int32 = -1

#Region "  Original Agent Effects on this entity go here  "
    '
#End Region
#End Region

    Public Sub New(ByRef oItem As Epica_Entity)
        oEntity = oItem


        With oEntity
            OrigParentObject = .ParentObject
            OrigOwner = .Owner
            OrigLocX = .LocX
            OrigLocZ = .LocZ
            OrigLocAngle = .LocAngle
            OrigQ1_HP = .Q1_HP
            OrigQ2_HP = .Q2_HP
            OrigQ3_HP = .Q3_HP
            OrigQ4_HP = .Q4_HP
            OrigStructure_HP = .Structure_HP
            OrigShield_HP = .Shield_HP
            OrigExpLevel = .ExpLevel
            If .lCurrentAmmo Is Nothing = False Then
                ReDim OrigCurrentAmmo(.lCurrentAmmo.GetUpperBound(0))
                For X As Int32 = 0 To .lCurrentAmmo.GetUpperBound(0)
                    OrigCurrentAmmo(X) = .lCurrentAmmo(X)
                Next X
            End If
            OrigCurrentStatus = .CurrentStatus

            OrigHangarUB = .lHangarUB
            ReDim OrigHangarContents(.lHangarUB)
            ReDim OrigHangarIdx(.lHangarUB)
            For X As Int32 = 0 To .lHangarUB
                OrigHangarContents(X) = .oHangarContents(X)
                OrigHangarIdx(X) = .lHangarIdx(X)
            Next X

            OrigCargoUB = .lCargoUB
            ReDim OrigCargoContents(.lCargoUB)
            ReDim OrigCargoIdx(.lCargoUB)
            For X As Int32 = 0 To .lCargoUB
                OrigCargoContents(X) = .oCargoContents(X)
                OrigCargoIdx(X) = .lCargoIdx(X)
            Next X

            'TODO: copy agent effects here
        End With
    End Sub

End Class
