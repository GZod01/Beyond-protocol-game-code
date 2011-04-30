Option Strict On

''' <summary>
''' This class contains the template definitions of the special tech table, it is NOT an instance of a relationship from player to special tech
''' </summary>
''' <remarks></remarks>
Public Class SpecialTech
    Inherits Epica_GUID

    'Special Techs themselves never save to the database as they cannot be changed in the game

    Public TechName(49) As Byte             '50 bytes
    Public InitialSuccessChance As Int32
    Public IncrementalSuccess As Int32
    Public FallOffSuccess As Int32
    Public MaxLinkChanceAttempts As Int32
    Public ProgramControl As Int32          'not sure if we'll use this

    Public RolePlayDesc() As Byte           'variable
    Public BenefitsDesc() As Byte           'variable

    Public oResearchCost As ProductionCost  'the cost to research this special tech (per attempt)

    Public bInGuaranteeList As Boolean = False

	Public fPercCostValue As Single = 0.0F

    Public lNewValue As Int32

    Public bHalfOwned As Boolean = False

    Private mySendString() As Byte

    Public bCanBeLinked As Boolean = True

    Public oPreqs() As SpecialTechPrerequisite
    Public lPreqUB As Int32 = -1

    ''' <summary>
    ''' This assumes that it is being called from PlayerSpecialTech, do NOT call this as a standard AddObjectMessage!
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSpecialTechSendMsg() As Byte()
        If mbStringReady = False Then
            ReDim mySendString(65 + RolePlayDesc.Length + BenefitsDesc.Length)
            Dim lPos As Int32 = 0

            TechName.CopyTo(mySendString, lPos) : lPos += 50
            System.BitConverter.GetBytes(ProgramControl).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(lNewValue).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(CInt(RolePlayDesc.Length)).CopyTo(mySendString, lPos) : lPos += 4
            If RolePlayDesc.Length <> 0 Then RolePlayDesc.CopyTo(mySendString, lPos) : lPos += RolePlayDesc.Length
            System.BitConverter.GetBytes(CInt(BenefitsDesc.Length)).CopyTo(mySendString, lPos) : lPos += 4
            If BenefitsDesc.Length <> 0 Then BenefitsDesc.CopyTo(mySendString, lPos) : lPos += BenefitsDesc.Length
            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Function GetMsgLen() As Int32
        Return GetSpecialTechSendMsg.Length
    End Function

    Public Function MeetsPreRequisites(ByVal oPlayer As Player) As Int32
        Dim lChance As Int32 = 0
        For X As Int32 = 0 To lPreqUB
            Dim lTemp As Int32 = oPreqs(X).MeetsPreRequisites(oPlayer)
            If lTemp = -1 Then
                If oPreqs(X).RequiredPrerequisite = True Then Return -1
            Else
                lChance += lTemp
            End If
        Next X
        Return lChance
    End Function
End Class

