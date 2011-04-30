Option Strict On

''' <summary>
''' Contains the data needed for special technology prerequisites (requirements for a tech to be linked)
''' </summary>
''' <remarks></remarks>
Public Class SpecialTechPrerequisite
    Inherits Epica_GUID

    Public TechID As Int32      'ID to which this prerequisite points to
    Public lPreqID As Int32     'ID to which this prerequisite relies
    Public iPreqTypeID As Int16 'type id of the object to which this prerequisite relies
    Public RequiredValue As Int32       'the value required of the Preq in order for this prerequisite to be filled (usually a 0 or 1 indicating researched/exists or not)
    Public ChanceToOpenLink As Int32            'cumulative
    Public RequiredPrerequisite As Boolean      'indicates that this prerequisite is required before ANY attempts can be made

    'Special Tech Prerequisites themselves never save as they never change inside the game

    Private moTech As SpecialTech = Nothing
    ''' <summary>
    ''' Gets the tech component to which this prerequisite points to
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property oTech() As SpecialTech
        Get
            If moTech Is Nothing Then moTech = GetEpicaSpecialTech(TechID)
            Return moTech
        End Get
    End Property

    Public Function MeetsPreRequisites(ByVal oPlayer As Player) As Int32
        Select Case iPreqTypeID
            Case 51 'ObjectType.eSpecialTech

                Dim lIdx As Int32 = -1
                For Y As Int32 = 0 To glSpecialTechUB
                    If TechID = glSpecialTechIdx(Y) Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y

                If lIdx <> -1 Then 'AndAlso lChances(lIdx) <> -1 Then
                    Dim bFound As Boolean = False
                    For Y As Int32 = 0 To oPlayer.oSpecials.mlSpecialTechUB
                        If oPlayer.oSpecials.mlSpecialTechIdx(Y) = lPreqID Then
                            'we found it in our linked list...
                            bFound = True

                            'check our required value
                            If RequiredValue = 0 Then
                                'Ok, add our chance to succeed ONLY if development phase <> researched
                                If oPlayer.oSpecials.moSpecialTech(Y).ComponentDevelopmentPhase <> 2 Then
                                    Return ChanceToOpenLink
                                End If
                            ElseIf oPlayer.oSpecials.moSpecialTech(Y).ComponentDevelopmentPhase = 2 Then
                                'ok, add our chance only if we have researched it
                                Return ChanceToOpenLink
                            Else : Return -1
                            End If
                        End If
                    Next Y

                    'If we didn't find it, and it is a required prerequisite...
                    If bFound = False AndAlso RequiredPrerequisite = True Then
                        'set our lchances to -1 so that it will never work
                        Return -1
                    End If
                End If
        End Select

    End Function

End Class
