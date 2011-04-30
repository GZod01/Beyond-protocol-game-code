Option Strict On

Public Class PlayerSpecialTech
    Inherits Epica_Tech

    Private Shared moEmptyCost As ProductionCost
    Public Enum eyThrowBackType As Byte
        ePushToBypass = 1
        eSwapFromBypass = 2
    End Enum

    'USES tblPlayerSpecialTech!!!

    Public Enum eySpecialTechFlags As Byte
        eLinked = 1
        eInTheTank = 2
        eTankThrowBackEvent = 4
    End Enum

    Public CreditResearchAttempts As Int32
    Public lThisPassLinkChance As Int32

    'Public bLinked As Boolean = False       'indicates that the tech is linked and available for research (or researched), without this, there is no technology to the player
    Public yFlags As Byte = 0
    Public LinkAttempts As Int32 = 0        'number of link attempts made so far (if this exceeds the SpecialTech's MaxLinkChanceAttempts, then it is a broken link
    Public ReadOnly Property bLinked() As Boolean
        Get
            Return (yFlags And eySpecialTechFlags.eLinked) <> 0
        End Get
    End Property
    Public ReadOnly Property bInTheTank() As Boolean
        Get
            Return (yFlags And eySpecialTechFlags.eInTheTank) <> 0
        End Get
    End Property

    'ObjectID = SpecialTechID
    'ObjTypeID = SpecialTech
    'Owner = PlayerID
    'CurrentSuccessChance
    'ResearchAttempts - number of times it was attempted research before succeeded (does not cause broken link)
    'ResPhase
    'SuccessChanceIncrement = SpecialTech's Success Chance Increment
    'moResearchCost - set from Special Tech's research cost

    'UNUSED
    '======
    'ErrorReasonCode - should always be 0
    'moDesignCost
    'moProductionCost 
    'RandomSeed

    Private mySendString() As Byte

    Private moSpecialTech As SpecialTech = Nothing
    ''' <summary>
    ''' Gets Special Tech related to this PlayerSpecialTech
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property oTech() As SpecialTech
        Get
            If moSpecialTech Is Nothing Then moSpecialTech = GetEpicaSpecialTech(Me.ObjectID)
            Return moSpecialTech
        End Get
    End Property

    Protected Overrides Sub CalculateBothProdCosts()
        'Ok, let's create our research cost
        If moResearchCost Is Nothing Then
            If oTech Is Nothing = False Then
                If moSpecialTech.oResearchCost Is Nothing = False Then
                    moResearchCost = moSpecialTech.oResearchCost.CreateClone(1)
                    Dim bCheckForMinValue As Boolean = (ResearchAttempts = 0)
                    moResearchCost.PointsRequired = GetResearchPointsCost(moResearchCost.PointsRequired)
                    If bCheckForMinValue = True Then CreditResearchAttempts = ResearchAttempts
                    moResearchCost.CreditCost *= CreditResearchAttempts

                    If moSpecialTech.bHalfOwned = True Then
                        If moResearchCost.PointsRequired > 6618240000 Then
                            moResearchCost.PointsRequired = 6618240000
                        End If
                    End If

                    If bCheckForMinValue = True Then
                        Dim blTest As Int64 = Owner.blCredits
                        If Owner.oGuild Is Nothing = False Then
                            Dim oMem As GuildMember = Owner.oGuild.GetMember(Owner.ObjectID)
                            If oMem Is Nothing = False Then
                                If (oMem.yMemberState And GuildMemberState.Approved) <> 0 Then blTest = Math.Max(blTest, Owner.oGuild.blTreasury)
                            End If
                        End If
                        blTest = Math.Max(blTest, Owner.blMaxPopulation * 1000)
                        Dim blMinCredits As Int64 = CLng(blTest * Me.oTech.fPercCostValue)
                        If moResearchCost.CreditCost < blMinCredits Then
                            'ok, determine how many attempts we will need...
                            Dim blAttempts As Int64 = blMinCredits \ moSpecialTech.oResearchCost.CreditCost
                            If blAttempts > Int32.MaxValue Then blAttempts = Int32.MaxValue
                            If blAttempts < 0 Then blAttempts = ResearchAttempts

                            'Now, use that...
                            'Me.ResearchAttempts = CInt(blAttempts)
                            Me.CreditResearchAttempts = CInt(blAttempts)
                            moResearchCost.CreditCost = moSpecialTech.oResearchCost.CreditCost * CLng(Me.CreditResearchAttempts)
                        End If
                    End If

                    moResearchCost.CreditCost = CLng(moResearchCost.CreditCost * Me.Owner.SpecTechCostMult)
                    moResearchCost.PointsRequired = CLng(moResearchCost.PointsRequired * Me.Owner.SpecTechCostMult)
                    'moResearchCost.PointsRequired \= 10
                Else
                    moResearchCost = New ProductionCost()
                    With moResearchCost
                        .ColonistCost = 0
                        .CreditCost = 1000000
                        .EnlistedCost = 0
                        .ItemCostUB = -1
                        .ObjectID = moSpecialTech.ObjectID
                        .ObjTypeID = ObjectType.eSpecialTech
                        .OfficerCost = 0
                        .PC_ID = -1
                        .PointsRequired = 350000
                        .ProductionCostType = 1
                        .PointsRequired = GetResearchPointsCost(.PointsRequired)
                        .CreditCost *= ResearchAttempts
                    End With
                End If
            End If
        End If

        If moEmptyCost Is Nothing Then
            moEmptyCost = New ProductionCost()
            With moEmptyCost
                .ColonistCost = 0
                .CreditCost = 0
                .EnlistedCost = 0
                .ItemCostUB = -1
                .OfficerCost = 0
                .PointsRequired = 0
                .ProductionCostType = 0
                .ObjTypeID = ObjectType.eSpecialTech
            End With
        End If
        moEmptyCost.ObjectID = Me.ObjectID
        moEmptyCost.MakeDirty()
        moProductionCost = moEmptyCost
    End Sub

    Public Overrides Sub ComponentDesigned()
        'This occurs on a successful link
        yFlags = yFlags Or eySpecialTechFlags.eLinked
        ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eComponentResearching
    End Sub

    Protected Overrides Sub FinalizeResearch()
        mbStringReady = False
        'This indicates the component has been researched completely... determine what action we need to make
        'NOTE: Some Hardcoding going on here. ;)
        Select Case Me.ObjectID
            Case 367
                Owner.AddKnownProperty(eMinPropID.ChemicalReactance, 2, True, True)
            Case 368
                Owner.AddKnownProperty(eMinPropID.Combustiveness, 2, True, True)
            Case 369
                Owner.AddKnownProperty(eMinPropID.Compressibility, 2, True, True)
                Owner.AddKnownProperty(eMinPropID.Malleable, 2, True, True)
            Case 370
                Owner.AddKnownProperty(eMinPropID.MagneticProduction, 2, True, True)
                Owner.AddKnownProperty(eMinPropID.MagneticReaction, 2, True, True)
            Case 371
                Owner.AddKnownProperty(eMinPropID.ElectricalResist, 2, True, True)
            Case 372
                Owner.AddKnownProperty(eMinPropID.Psych, 2, True, True)
            Case 373
                Owner.AddKnownProperty(eMinPropID.Quantum, 2, True, True)
            Case 374
                Owner.AddKnownProperty(eMinPropID.Reflection, 2, True, True)
            Case 375
                Owner.AddKnownProperty(eMinPropID.Refraction, 2, True, True)
            Case 376
                Owner.AddKnownProperty(eMinPropID.SuperconductivePoint, 2, True, True)
            Case 377
                Owner.AddKnownProperty(eMinPropID.TemperatureSensitivity, 2, True, True)
                Owner.AddKnownProperty(eMinPropID.ThermalConductance, 2, True, True)
                Owner.AddKnownProperty(eMinPropID.ThermalExpansion, 2, True, True)
            Case 378
                Owner.AddKnownProperty(eMinPropID.Toxicity, 2, True, True)
        End Select

        'Always call this
        Owner.oSpecials.ProcessTechResearched(Me.oTech)

        'Finally, we call the owner to perform link tests
        Owner.oSpecials.PerformLinkTest()
    End Sub

    Public Overrides Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            'Attempt to update this object first
            sSQL = "UPDATE tblPlayerSpecialTech SET CurrentSuccessChance = " & CurrentSuccessChance & ", ResearchAttempts = " & _
              ResearchAttempts & ", LinkAttempts = " & LinkAttempts & ", ResPhase = " & CInt(ComponentDevelopmentPhase).ToString & _
              ", CreditResearchAttempts = " & Me.CreditResearchAttempts & ", SuccessfulLink = " & yFlags
            'If bLinked = True Then sSQL &= "1" Else sSQL &= "0"
            sSQL &= ", bArchived = " & yArchived & ", VersionNumber = " & lVersionNum & " WHERE SpecialTechID = " & ObjectID & " AND PlayerID = " & Owner.ObjectID

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                'Ok, update did not affect anything, let's try to insert
                sSQL = "INSERT INTO tblPlayerSpecialTech (SpecialTechID, PlayerID, CurrentSuccessChance, ResearchAttempts, " & _
                  "LinkAttempts, ResPhase, CreditResearchAttempts, SuccessfulLink, bArchived, VersionNumber) VALUES (" & Me.ObjectID & ", " & Owner.ObjectID & ", " & _
                  Me.CurrentSuccessChance & ", " & Me.ResearchAttempts & ", " & Me.LinkAttempts & ", " & _
                  CInt(Me.ComponentDevelopmentPhase).ToString & ", " & Me.CreditResearchAttempts & ", " & yFlags & ", " & yArchived & ", " & lVersionNum & ")"
                'If bLinked = True Then sSQL &= "1, " Else sSQL &= "0, "
                'sSQL &= yArchived & ")"

                oComm.Dispose()
                oComm = Nothing
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If

            'MyBase.FinalizeSave()

            bResult = True
        Catch ex As Exception
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save PlayerSpecialTech (STID: " & Me.ObjectID & "). Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        'bNeedsSaved = False
        Return bResult

    End Function

    Public Overrides Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
        'This should never happen, report it if it does
        'TODO: Report who owns this...
        LogEvent(LogEventType.PossibleCheat, "Received SetFromDesignMsg in a PlayerSpecialTech!")
        Return False
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        ErrorReasonCode = TechBuilderErrorReason.eNoError
    End Function

    Public Function GetObjAsString() As Byte()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
            CalculateBothProdCosts()
            mbStringReady = False
        End If

        'here we will return the entire object as a string
        'If mbStringReady = False Then
        ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE_EXCLUDED + oTech.GetMsgLen + 1)
        Dim lPos As Int32
        MyBase.GetBaseObjAsString(True).CopyTo(mySendString, 0)
        lPos = Epica_Tech.BASE_OBJ_STRING_SIZE_EXCLUDED

        mySendString(lPos) = yFlags : lPos += 1

        oTech.GetSpecialTechSendMsg.CopyTo(mySendString, lPos) : lPos += oTech.GetMsgLen
        'mbStringReady = True
        'End If
        Return mySendString
    End Function

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(7 + oTech.GetMsgLen) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yResult, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yResult, lPos) : lPos += 6
        oTech.GetSpecialTechSendMsg.CopyTo(yResult, lPos) : lPos += oTech.GetMsgLen

        Return yResult
    End Function

    Public Overrides Function TechnologyScore() As Integer
        Return CInt(oTech.oResearchCost.CreditCost \ 1000I)
    End Function

    Public Overrides Function GetPlayerTechKnowledgeMsg(ByVal yTechLvl As PlayerTechKnowledge.KnowledgeType) As Byte()
        Dim yMsg() As Byte = Nothing
        Dim lPos As Int32 = 0
        Try
            'set up our msg size
            lPos = 60
            If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then lPos += (4 + Me.oTech.RolePlayDesc.Length)
            If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += (4 + Me.oTech.BenefitsDesc.Length)
            If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then lPos += 0
            ReDim yMsg(lPos - 1)
            lPos = 0

            'Default Attributes
            GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            Me.oTech.TechName.CopyTo(yMsg, lPos) : lPos += 50

            If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
                'at least settings 1 
                System.BitConverter.GetBytes(CInt(Me.oTech.RolePlayDesc.Length)).CopyTo(mySendString, lPos) : lPos += 4
                If Me.oTech.RolePlayDesc.Length <> 0 Then Me.oTech.RolePlayDesc.CopyTo(mySendString, lPos) : lPos += Me.oTech.RolePlayDesc.Length
            End If
            If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
                'at least settings 2  
                System.BitConverter.GetBytes(CInt(Me.oTech.BenefitsDesc.Length)).CopyTo(mySendString, lPos) : lPos += 4
                If Me.oTech.BenefitsDesc.Length <> 0 Then Me.oTech.BenefitsDesc.CopyTo(mySendString, lPos) : lPos += Me.oTech.BenefitsDesc.Length
            End If
            If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
                'Full Knowledge 
                'NOTE: AT FULL KNOWLEDGE THE PLAYER CAN PLACE THE COMPONENT ON A UNIT/FACILITY DESIGN
            End If
        Catch
        End Try

        Return yMsg
    End Function

    Public Overrides Sub FillFromPrimaryAddMsg(ByVal yData() As Byte)
        'TODO: What should we do here? 
    End Sub
End Class
