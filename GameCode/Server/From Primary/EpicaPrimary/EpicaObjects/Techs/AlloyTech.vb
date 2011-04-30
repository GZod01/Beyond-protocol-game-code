Public Class AlloyTech
    Inherits Epica_Tech

    Public AlloyName(19) As Byte
    'Public AlloyResult As Mineral
    Private moAlloyResult As Mineral = Nothing

    '=========== FROM COMPONENT DESIGN =============
    Public Mineral1ID As Int32 = 0
    Public Mineral2ID As Int32 = 0
    Public Mineral3ID As Int32 = 0
    Public Mineral4ID As Int32 = 0

    Public lPropertyID1 As Int32
    Public lPropertyID2 As Int32
    Public lPropertyID3 As Int32

    'Public bHigher1 As Boolean
    'Public bHigher2 As Boolean
    'Public bHigher3 As Boolean
    Public yNewVal1 As Byte = 255
    Public yNewVal2 As Byte = 255
    Public yNewVal3 As Byte = 255

    Public ResearchLevel As Byte
    '=================================================

    Private mySendString() As Byte

    Public Property AlloyResult() As Mineral
        Get
            If moAlloyResult Is Nothing AndAlso Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eComponentResearching Then
                'We should know this already...
                ComponentDesigned()
                'now we do
            End If
            Return moAlloyResult
        End Get
        Set(ByVal value As Mineral)
            moAlloyResult = value
        End Set
    End Property

    Public Function GetObjAsString() As Byte()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
            CalculateBothProdCosts()
            mbStringReady = False
        End If

        'here we will return the entire object as a string
        'If mbStringReady = False Then

        If AlloyResult Is Nothing OrElse moAlloyResult.ObjectID > 0 Then
            ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE_EXCLUDED + 55)
        Else
            ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE_EXCLUDED + 59 + ((Owner.mlMinPropertyUB + 1) * 5))
        End If

        Dim lPos As Int32
        MyBase.GetBaseObjAsString(True).CopyTo(mySendString, 0)
        lPos = Epica_Tech.BASE_OBJ_STRING_SIZE_EXCLUDED

        AlloyName.CopyTo(mySendString, lPos) : lPos += 20

        System.BitConverter.GetBytes(Mineral1ID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(Mineral2ID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(Mineral3ID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(Mineral4ID).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(lPropertyID1).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPropertyID2).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPropertyID3).CopyTo(mySendString, lPos) : lPos += 4

        'If bHigher1 = True Then mySendString(lPos) = 255 Else mySendString(lPos) = 0
        'lPos += 1
        'If bHigher2 = True Then mySendString(lPos) = 255 Else mySendString(lPos) = 0
        'lPos += 1
        'If bHigher3 = True Then mySendString(lPos) = 255 Else mySendString(lPos) = 0
        'lPos += 1
        mySendString(lPos) = yNewVal1 : lPos += 1
        mySendString(lPos) = yNewVal2 : lPos += 1
        mySendString(lPos) = yNewVal3 : lPos += 1

        mySendString(lPos) = ResearchLevel
        lPos += 1

        If moAlloyResult Is Nothing Then
            System.BitConverter.GetBytes(CInt(0)).CopyTo(mySendString, lPos) : lPos += 4
        ElseIf moAlloyResult.ObjectID > 0 Then
            System.BitConverter.GetBytes(moAlloyResult.ObjectID).CopyTo(mySendString, lPos) : lPos += 4
        Else
            System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, lPos) : lPos += 4
            'Put our count in there...
            System.BitConverter.GetBytes(Owner.mlMinPropertyUB + 1).CopyTo(mySendString, lPos) : lPos += 4
            'Now... loop through and put our stuff in there...
            For X As Int32 = 0 To Owner.mlMinPropertyUB
                'Property ID (4)
                Dim lPropID As Int32 = Owner.moMinProperties(X).lPropertyID
                System.BitConverter.GetBytes(lPropID).CopyTo(mySendString, lPos) : lPos += 4
                'Percentage... (1)
                mySendString(lPos) = PlayerMineral.GetClientDisplayedPropertyValue(moAlloyResult.GetPropertyValue(lPropID)) : lPos += 1
            Next X
        End If

        'mbStringReady = True
        'End If
        Return mySendString
    End Function

    Public Overrides Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                'Property1Higher, Property2Higher, Property3Higher, ResearchLevel" & _
                sSQL = "INSERT INTO tblAlloy (AlloyName, ResultMineralID, OwnerID, " & _
                  "Mineral1ID, Mineral2ID, Mineral3ID, Mineral4ID, MineralProperty1ID, MineralProperty2ID, " & _
                  "MineralProperty3ID, Property1Value, Property2Value, Property3Value, ResearchLevel" & _
                  ", CurrentSuccessChance, SuccessChanceIncrement, ResearchAttempts, RandomSeed, ResPhase, MajorDesignFlaw" & _
                  ", PopIntel, bArchived, VersionNumber) VALUES ('" & MakeDBStr(BytesToString(AlloyName)) & "', "
                If AlloyResult Is Nothing Then sSQL &= "0, " Else sSQL &= AlloyResult.ObjectID & ", "
                sSQL &= Owner.ObjectID & ", " & Mineral1ID & _
                  ", " & Mineral2ID & ", " & Mineral3ID & ", " & Mineral4ID & ", " & lPropertyID1 & ", " & _
                  lPropertyID2 & ", " & lPropertyID3 & ", " '& bHigher1 & ", " & bHigher2 & ", " & bHigher3 & _
                'If bHigher1 = True Then sSQL &= "1, " Else sSQL &= "0, "
                'If bHigher2 = True Then sSQL &= "1, " Else sSQL &= "0, "
                'If bHigher3 = True Then sSQL &= "1, " Else sSQL &= "0, "
                sSQL &= yNewVal1 & ", " & yNewVal2 & ", " & yNewVal3 & ", " & ResearchLevel & ", " & CurrentSuccessChance & ", " & _
                  SuccessChanceIncrement & ", " & ResearchAttempts & ", " & RandomSeed & ", " & ComponentDevelopmentPhase & ", " & _
                  MajorDesignFlaw & ", " & PopIntel & ", " & yArchived & ", " & lVersionNum & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblAlloy SET AlloyName='" & MakeDBStr(BytesToString(AlloyName)) & _
                  "', ResultMineralID = "
                If AlloyResult Is Nothing Then sSQL &= "0" Else sSQL &= AlloyResult.ObjectID
                sSQL &= ", OwnerID = " & Owner.ObjectID & ", Mineral1ID = " & Mineral1ID & _
                  ", Mineral2ID = " & Mineral2ID & ", Mineral3ID = " & Mineral3ID & ", Mineral4ID = " & _
                  Mineral4ID & ", MineralProperty1ID = " & lPropertyID1 & ", MineralProperty2ID = " & _
                  lPropertyID2 & ", MineralProperty3ID = " & lPropertyID3 '& ", Property1Higher = " & _
                sSQL &= ", Property1Value = " & yNewVal1 & ", Property2Value = " & yNewVal2 & ", Property3Value = " & yNewVal3
                'If bHigher1 = True Then sSQL &= ", Property1Higher = 1" Else sSQL &= ", Property1Higher = 0"
                'If bHigher2 = True Then sSQL &= ", Property2Higher = 1" Else sSQL &= ", Property2Higher = 0"
                'If bHigher3 = True Then sSQL &= ", Property3Higher = 1" Else sSQL &= ", Property3Higher = 0"
                sSQL &= ", ResearchLevel = " & ResearchLevel & _
                  ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
                  SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
                  RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", MajorDesignFlaw = " & MajorDesignFlaw & _
                  ", PopIntel = " & PopIntel & ", bArchived = " & yArchived & ", VersionNumber = " & lVersionNum & " WHERE AlloyID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(AlloyID) FROM tblAlloy WHERE AlloyName = '" & MakeDBStr(BytesToString(AlloyName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            'Now, save the alloy result
            If Me.ComponentDevelopmentPhase <> eComponentDevelopmentPhase.eComponentDesign AndAlso moAlloyResult Is Nothing = False Then
                sSQL = "DELETE FROM tblAlloyResultProperty WHERE AlloyID = " & Me.ObjectID
                oComm = Nothing
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm = Nothing

                Try
                    For X As Int32 = 0 To glMineralPropertyUB
                        Dim lPropID As Int32 = glMineralPropertyIdx(X)
                        If lPropID <> -1 Then
                            Dim lValue As Int32 = moAlloyResult.GetPropertyValue(lPropID)
                            sSQL = "INSERT INTO tblAlloyResultProperty (AlloyID, MineralPropertyID, PropertyValue) VALUES (" & _
                              Me.ObjectID & ", " & lPropID & ", " & lValue & ")"
                            oComm = New OleDb.OleDbCommand(sSQL, goCN)

                            oComm.ExecuteNonQuery()
                            oComm = Nothing
                        End If
                    Next X
                Catch
                    LogEvent(LogEventType.CriticalError, "Saving the alloy result's mineral properties resulted in an error: " & Err.Description)
                    sSQL = "DELETE FROM tblAlloyResultProperty WHERE AlloyID = " & Me.ObjectID
                    oComm = Nothing
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oComm.ExecuteNonQuery()
                    oComm = Nothing
                End Try
            End If

            MyBase.FinalizeSave()

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        'bNeedsSaved = False
        Return bResult
    End Function

    Public Function GetSaveObjectText() As String
        Dim bResult As Boolean = False
        Dim sSQL As String

        If ObjectID = -1 Then
            SaveObject()
            Return ""
        End If

        Try

            Dim oSB As New System.Text.StringBuilder

            'UPDATE
            sSQL = "UPDATE tblAlloy SET AlloyName='" & MakeDBStr(BytesToString(AlloyName)) & _
              "', ResultMineralID = "
            If AlloyResult Is Nothing Then sSQL &= "0" Else sSQL &= AlloyResult.ObjectID
            sSQL &= ", OwnerID = " & Owner.ObjectID & ", Mineral1ID = " & Mineral1ID & _
              ", Mineral2ID = " & Mineral2ID & ", Mineral3ID = " & Mineral3ID & ", Mineral4ID = " & _
              Mineral4ID & ", MineralProperty1ID = " & lPropertyID1 & ", MineralProperty2ID = " & _
              lPropertyID2 & ", MineralProperty3ID = " & lPropertyID3 '& ", Property1Higher = " & _
            sSQL &= ", Property1Value = " & yNewVal1 & ", Property2Value = " & yNewVal2 & ", Property3Value = " & yNewVal3
            'If bHigher1 = True Then sSQL &= ", Property1Higher = 1" Else sSQL &= ", Property1Higher = 0"
            'If bHigher2 = True Then sSQL &= ", Property2Higher = 1" Else sSQL &= ", Property2Higher = 0"
            'If bHigher3 = True Then sSQL &= ", Property3Higher = 1" Else sSQL &= ", Property3Higher = 0"
            sSQL &= ", ResearchLevel = " & ResearchLevel & _
              ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
              SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
              RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", MajorDesignFlaw = " & MajorDesignFlaw & _
              ", PopIntel = " & PopIntel & ", bArchived = " & yArchived & " WHERE AlloyID = " & ObjectID

            oSB.AppendLine(sSQL)

            'Now, save the alloy result
            If Me.ComponentDevelopmentPhase <> eComponentDevelopmentPhase.eComponentDesign AndAlso moAlloyResult Is Nothing = False Then
                sSQL = "DELETE FROM tblAlloyResultProperty WHERE AlloyID = " & Me.ObjectID
                oSB.AppendLine(sSQL)

                Try
                    For X As Int32 = 0 To glMineralPropertyUB
                        Dim lPropID As Int32 = glMineralPropertyIdx(X)
                        If lPropID <> -1 Then
                            Dim lValue As Int32 = moAlloyResult.GetPropertyValue(lPropID)
                            sSQL = "INSERT INTO tblAlloyResultProperty (AlloyID, MineralPropertyID, PropertyValue) VALUES (" & _
                              Me.ObjectID & ", " & lPropID & ", " & lValue & ")"
                            oSB.AppendLine(sSQL)
                        End If
                    Next X
                Catch
                    LogEvent(LogEventType.CriticalError, "Saving the alloy result's mineral properties resulted in an error: " & Err.Description)
                    sSQL = "DELETE FROM tblAlloyResultProperty WHERE AlloyID = " & Me.ObjectID
                    oSB.AppendLine(sSQL)
                End Try
            End If

            oSB.AppendLine(MyBase.GetFinalizeSaveText())

            Return oSB.ToString
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Private Function GetPropertyMax(ByVal lPropID As Int32) As Int32
        Dim lVal1 As Int32 = 0
        Dim lVal2 As Int32 = 0
        Dim lVal3 As Int32 = 0
        Dim lVal4 As Int32 = 0

        If Mineral1 Is Nothing = False Then lVal1 = Mineral1.GetPropertyValue(lPropID)
        If Mineral2 Is Nothing = False Then lVal2 = Mineral2.GetPropertyValue(lPropID)
        If Mineral3 Is Nothing = False Then lVal3 = Mineral3.GetPropertyValue(lPropID)
        If Mineral4 Is Nothing = False Then lVal4 = Mineral4.GetPropertyValue(lPropID)

        Return Math.Max(Math.Max(Math.Max(lVal1, lVal2), lVal3), lVal4)
    End Function

    Private Function GetPropertyMin(ByVal lPropID As Int32) As Int32
        Dim lVal1 As Int32 = Int32.MaxValue
        Dim lVal2 As Int32 = Int32.MaxValue
        Dim lVal3 As Int32 = Int32.MaxValue
        Dim lVal4 As Int32 = Int32.MaxValue

        If Mineral1 Is Nothing = False Then lVal1 = Mineral1.GetPropertyValue(lPropID)
        If Mineral2 Is Nothing = False Then lVal2 = Mineral2.GetPropertyValue(lPropID)
        If Mineral3 Is Nothing = False Then lVal3 = Mineral3.GetPropertyValue(lPropID)
        If Mineral4 Is Nothing = False Then lVal4 = Mineral4.GetPropertyValue(lPropID)

        Return Math.Min(Math.Min(Math.Min(lVal1, lVal2), lVal3), lVal4)
    End Function

    Private Function GetPropertySum(ByVal lPropID As Int32) As Int32
        Dim lVal1 As Int32 = 0
        Dim lVal2 As Int32 = 0
        Dim lVal3 As Int32 = 0
        Dim lVal4 As Int32 = 0

        If Mineral1 Is Nothing = False Then lVal1 = Mineral1.GetPropertyValue(lPropID)
        If Mineral2 Is Nothing = False Then lVal2 = Mineral2.GetPropertyValue(lPropID)
        If Mineral3 Is Nothing = False Then lVal3 = Mineral3.GetPropertyValue(lPropID)
        If Mineral4 Is Nothing = False Then lVal4 = Mineral4.GetPropertyValue(lPropID)

        Return (lVal1 + lVal2 + lVal3 + lVal4)
    End Function

    Private moMineral1 As Mineral = Nothing
    Private ReadOnly Property Mineral1() As Mineral
        Get
            If Mineral1ID < 1 AndAlso moMineral1 Is Nothing = False Then moMineral1 = Nothing
            If (moMineral1 Is Nothing OrElse moMineral1.ObjectID <> Mineral1ID) AndAlso Mineral1ID > 0 Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = Mineral1ID Then
                        moMineral1 = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moMineral1
        End Get
    End Property

    Private moMineral2 As Mineral
    Private ReadOnly Property Mineral2() As Mineral
        Get
            If Mineral2ID < 1 AndAlso moMineral2 Is Nothing = False Then moMineral2 = Nothing
            If (moMineral2 Is Nothing OrElse moMineral2.ObjectID <> Mineral2ID) AndAlso Mineral2ID > 0 Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = Mineral2ID Then
                        moMineral2 = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moMineral2
        End Get
    End Property

    Private moMineral3 As Mineral
    Private ReadOnly Property Mineral3() As Mineral
        Get
            If Mineral3ID < 1 AndAlso moMineral3 Is Nothing = False Then moMineral3 = Nothing
            If (moMineral3 Is Nothing OrElse moMineral3.ObjectID <> Mineral3ID) AndAlso Mineral3ID > 0 Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = Mineral3ID Then
                        moMineral3 = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moMineral3
        End Get
    End Property

    Private moMineral4 As Mineral
    Private ReadOnly Property Mineral4() As Mineral
        Get
            If Mineral4ID < 1 AndAlso moMineral4 Is Nothing = False Then moMineral4 = Nothing
            If (moMineral4 Is Nothing OrElse moMineral4.ObjectID <> Mineral4ID) AndAlso Mineral4ID > 0 Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = Mineral4ID Then
                        moMineral4 = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moMineral4
        End Get
    End Property

    Public ReadOnly Property SuccessDivider() As Single
        Get
            Select Case ResearchLevel       '0 to 4
                Case 1
                    Return 4
                Case 2
                    Return 9
                Case 3
                    Return 16
                Case 4
                    Return 25
                Case Else       '0 included
                    Return 2
            End Select
        End Get
    End Property

    Private Function GetMineralCost(ByRef oMin As Mineral) As Int32
        Dim lQty As Int32 = 1
        Dim lProp1Act As Int32
        Dim lProp2Act As Int32
        Dim lProp3Act As Int32

        Dim lActVal As Int32
        Dim lPropID As Int32

        If lPropertyID1 > 0 Then lProp1Act = oMin.GetPropertyValue(lPropertyID1) Else lProp1Act = 0
        If lPropertyID2 > 0 Then lProp2Act = oMin.GetPropertyValue(lPropertyID2) Else lProp2Act = 0
        If lPropertyID3 > 0 Then lProp3Act = oMin.GetPropertyValue(lPropertyID3) Else lProp3Act = 0
        If lProp1Act > lProp2Act AndAlso lProp1Act > lProp3Act Then
            'Property 1
            lPropID = lPropertyID1
            lActVal = lProp1Act
        ElseIf lProp2Act > lProp3Act Then
            'Property 2
            lPropID = lPropertyID2
            lActVal = lProp2Act
        Else
            'Property 3
            lPropID = lPropertyID3
            lActVal = lProp3Act
        End If
        Dim lKnowVal As Int32 = Owner.GetMineralPropertyKnowledge(oMin.ObjectID, lPropertyID1)
        If lKnowVal = 0 Then
            lQty = 0
        Else : lQty = CInt(Math.Ceiling(lActVal / lKnowVal))
        End If

        Return lQty
    End Function

    Private Function ComputeSuccessChance() As Int32

        Dim fProp1Val As Single = 0
        Dim fProp2Val As Single = 0
        Dim fProp3Val As Single = 0

        Dim lTemp As Int32

        If Mineral1 Is Nothing = False Then
            With Mineral1
                If lPropertyID1 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID1)
                    If lTemp <> 0 Then
                        fProp1Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID1) / lTemp)
                    End If
                End If
                If lPropertyID2 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID2)
                    If lTemp <> 0 Then
                        fProp2Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID2) / lTemp)
                    End If
                End If
                If lPropertyID3 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID3)
                    If lTemp <> 0 Then
                        fProp3Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID3) / lTemp)
                    End If
                End If
            End With
        End If
        If Mineral2 Is Nothing = False Then
            With Mineral2
                If lPropertyID1 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID1)
                    If lTemp <> 0 Then
                        fProp1Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID1) / lTemp)
                    End If
                End If
                If lPropertyID2 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID2)
                    If lTemp <> 0 Then
                        fProp2Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID2) / lTemp)
                    End If
                End If
                If lPropertyID3 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID3)
                    If lTemp <> 0 Then
                        fProp3Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID3) / lTemp)
                    End If
                End If
            End With
        End If
        If Mineral3 Is Nothing = False Then
            With Mineral3
                If lPropertyID1 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID1)
                    If lTemp <> 0 Then
                        fProp1Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID1) / lTemp)
                    End If
                End If
                If lPropertyID2 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID2)
                    If lTemp <> 0 Then
                        fProp2Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID2) / lTemp)
                    End If
                End If
                If lPropertyID3 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID3)
                    If lTemp <> 0 Then
                        fProp3Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID3) / lTemp)
                    End If
                End If
            End With
        End If
        If Mineral4 Is Nothing = False Then
            With Mineral4
                If lPropertyID1 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID1)
                    If lTemp <> 0 Then
                        fProp1Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID1) / lTemp)
                    End If
                End If
                If lPropertyID2 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID2)
                    If lTemp <> 0 Then
                        fProp2Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID2) / lTemp)
                    End If
                End If
                If lPropertyID3 > 0 Then
                    lTemp = .GetPropertyValue(lPropertyID3)
                    If lTemp <> 0 Then
                        fProp3Val += CSng(Owner.GetMineralPropertyKnowledge(.ObjectID, lPropertyID3) / lTemp)
                    End If
                End If
            End With
        End If

        If lPropertyID1 > 0 Then
            fProp1Val /= 4.0F
            fProp1Val /= Me.SuccessDivider
            fProp1Val *= 100.0F
        End If
        If lPropertyID2 > 0 Then
            fProp2Val /= 4.0F
            fProp2Val /= Me.SuccessDivider
            fProp2Val *= 100.0F
        End If
        If lPropertyID3 > 0 Then
            fProp3Val /= 4.0F
            fProp3Val /= Me.SuccessDivider
            fProp3Val *= 100.0F
        End If

        Dim lResult As Int32 = Int32.MaxValue

        'Now, get our minimum value
        If lPropertyID1 > 0 Then lResult = Math.Min(CInt(fProp1Val), lResult)
        If lPropertyID2 > 0 Then lResult = Math.Min(CInt(fProp2Val), lResult)
        If lPropertyID3 > 0 Then lResult = Math.Min(CInt(fProp3Val), lResult)

        If lResult = Int32.MaxValue Then Return 1I Else Return lResult
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        'the player needs to have a player mineral entry for each mineral id used... and must be "discovered"
        Dim bMin1Good As Boolean = (Mineral1ID < 1) OrElse Owner.IsMineralDiscovered(Mineral1ID)
        Dim bMin2Good As Boolean = (Mineral2ID < 1) OrElse Owner.IsMineralDiscovered(Mineral2ID)
        Dim bMin3Good As Boolean = (Mineral3ID < 1) OrElse Owner.IsMineralDiscovered(Mineral3ID)
        Dim bMin4Good As Boolean = (Mineral4ID < 1) OrElse Owner.IsMineralDiscovered(Mineral4ID)

        If bMin1Good = False OrElse bMin2Good = False OrElse bMin3Good = False OrElse bMin4Good = False Then
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
            LogEvent(LogEventType.PossibleCheat, "AlloyTech.ValidateDesign mineral selected is bad: " & Me.Owner.ObjectID)
            Return False
        End If

        Dim bProp1Good As Boolean = (lPropertyID1 < 1) OrElse Owner.IsMineralPropertyKnown(lPropertyID1)
        Dim bProp2Good As Boolean = (lPropertyID2 < 1) OrElse Owner.IsMineralPropertyKnown(lPropertyID2)
        Dim bProp3Good As Boolean = (lPropertyID3 < 1) OrElse Owner.IsMineralPropertyKnown(lPropertyID3)

        If bProp1Good = False OrElse bProp2Good = False OrElse bProp3Good = False Then
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            LogEvent(LogEventType.PossibleCheat, "AlloyTech.ValidateDesign property selected is bad: " & Me.Owner.ObjectID)
            Return False
        End If

        If yNewVal1 = 255 AndAlso yNewVal2 = 255 AndAlso yNewVal3 = 255 Then
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            LogEvent(LogEventType.PossibleCheat, "AlloyTech.ValidateDesign selected new val is bad: " & Me.Owner.ObjectID)
            Return False
        End If

        ''Now, check our special techs...
        'Dim lMaxPropCnt As Int32 = Owner.oSpecials.yAlloyBuilderImprovements

        ''Now, check how many properties we are using
        'If lPropertyID1 > 0 Then lMaxPropCnt -= 1
        'If lPropertyID2 > 0 Then lMaxPropCnt -= 1
        'If lPropertyID3 > 0 Then lMaxPropCnt -= 1

        'If lMaxPropCnt < 0 Then
        '    Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
        '    Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
        '    Return False
        'End If

        'Finally, because this should only be called once...
        'oTech = Owner.GetTech(146, ObjectType.eSpecialTech)
        'If oTech Is Nothing = False AndAlso oTech.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eResearched Then
        '    Dim fTempVal As Single = Rnd()
        '    If fTempVal > RandomSeed Then RandomSeed = fTempVal
        'End If

        Return True
    End Function

    Public Overrides Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
        Dim lPos As Int32 = 2
        Dim bRes As Boolean = False
        Try
            With Me
                .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                'for the ID of the tech
                lPos += 4

                'Plus 6 for the researcher ID and researcher type ID
                lPos += 6
                .Mineral1ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Mineral2ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Mineral3ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Mineral4ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lPropertyID1 = yData(lPos) : lPos += 1
                .lPropertyID2 = yData(lPos) : lPos += 1
                .lPropertyID3 = yData(lPos) : lPos += 1

                '.bHigher1 = (yData(lPos) And 1) <> 0
                '.bHigher2 = (yData(lPos) And 2) <> 0
                '.bHigher3 = (yData(lPos) And 4) <> 0
                'lPos += 1
                .yNewVal1 = yData(lPos) : lPos += 1
                .yNewVal2 = yData(lPos) : lPos += 1
                .yNewVal3 = yData(lPos) : lPos += 1

                .ResearchLevel = yData(lPos) : lPos += 1

                ReDim .AlloyName(19)
                Array.Copy(yData, lPos, .AlloyName, 0, 20)
                lPos += 20
            End With
            bRes = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "ArmorTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bRes
    End Function

    Protected Overrides Sub CalculateBothProdCosts()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
            ComponentDesigned()
        End If

        'If moResearchCost Is Nothing Then moResearchCost = New ProductionCost
        'moResearchCost.ObjectID = Me.ObjectID
        'moResearchCost.ObjTypeID = Me.ObjTypeID
        'moResearchCost.ColonistCost = 0 : moResearchCost.EnlistedCost = 0 : moResearchCost.OfficerCost = 0

        'Dim lMaxVal As Int32 = Int32.MinValue
        'Dim lPopIntVal As Int32 = (255 - PopIntel)
        'For lProp As Int32 = 0 To glMineralPropertyUB
        '	Dim lPropID As Int32 = goMineralProperty(lProp).ObjectID

        '	If lPropID = lPropertyID1 OrElse lPropID = lPropertyID2 OrElse lPropID = lPropertyID3 Then
        '		Dim lKnown As Int32 = 0
        '		Dim lActual As Int32 = GetPropertySum(lPropID)

        '		If Mineral1 Is Nothing = False Then lKnown += Owner.GetMineralPropertyKnowledge(Mineral1.ObjectID, lPropID)
        '		If Mineral2 Is Nothing = False Then lKnown += Owner.GetMineralPropertyKnowledge(Mineral2.ObjectID, lPropID)
        '		If Mineral3 Is Nothing = False Then lKnown += Owner.GetMineralPropertyKnowledge(Mineral3.ObjectID, lPropID)
        '		If Mineral4 Is Nothing = False Then lKnown += Owner.GetMineralPropertyKnowledge(Mineral4.ObjectID, lPropID)

        '		Dim lVal As Int32 = Math.Abs(lKnown - lActual)
        '		lVal *= (ResearchLevel + 1)
        '		lVal *= lPopIntVal

        '		lMaxVal = Math.Max(lVal, lMaxVal)
        '	End If
        'Next lProp

        ''Now, lMaxVal has our max value... which is our production time...
        ''moResearchCost.PointsRequired = lMaxVal * 100   '100 production per cycle normal
        'moResearchCost.PointsRequired = GetResearchPointsCost((lMaxVal + 1) * 10) '100?
        'Dim lCost As Int32 = (lMaxVal \ PopIntel) '* Result.Complexity
        'moResearchCost.CreditCost = lCost * Me.ResearchAttempts
        ''TODO: Reset ResearchAttempts?

        'With moResearchCost
        '	.ItemCostUB = -1
        '	ReDim .ItemCosts(-1)
        '	If Mineral1 Is Nothing = False Then .AddProductionCostItem(Mineral1.ObjectID, ObjectType.eMineral, GetMineralCost(Mineral1))
        '	If Mineral2 Is Nothing = False Then .AddProductionCostItem(Mineral2.ObjectID, ObjectType.eMineral, GetMineralCost(Mineral2))
        '	If Mineral3 Is Nothing = False Then .AddProductionCostItem(Mineral3.ObjectID, ObjectType.eMineral, GetMineralCost(Mineral3))
        '	If Mineral4 Is Nothing = False Then .AddProductionCostItem(Mineral4.ObjectID, ObjectType.eMineral, GetMineralCost(Mineral4))
        'End With
        'Me.CurrentSuccessChance = ComputeSuccessChance()


        'If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
        'With moProductionCost
        '	.ColonistCost = moResearchCost.ColonistCost * 10
        '	.CreditCost = Math.Max(1, moResearchCost.CreditCost \ 10)	 '100
        '	.EnlistedCost = moResearchCost.EnlistedCost * 10
        '	.ObjectID = moResearchCost.ObjectID
        '	.ObjTypeID = moResearchCost.ObjTypeID
        '	.OfficerCost = moResearchCost.OfficerCost * 10
        '	'.PC_ID = -1
        '	.PointsRequired = Math.Max(1, moResearchCost.PointsRequired \ 100)	 '1000

        '	Dim lCnt As Int32 = 0
        '	If Mineral1 Is Nothing = False Then lCnt += 1
        '	If Mineral2 Is Nothing = False Then lCnt += 1
        '	If Mineral3 Is Nothing = False Then lCnt += 1
        '	If Mineral4 Is Nothing = False Then lCnt += 1
        '	.ItemCostUB = -1
        '	Dim lMinPer As Int32 = 12 \ lCnt
        '	If Mineral1 Is Nothing = False Then .AddProductionCostItem(Mineral1.ObjectID, ObjectType.eMineral, lMinPer)
        '	If Mineral2 Is Nothing = False Then .AddProductionCostItem(Mineral2.ObjectID, ObjectType.eMineral, lMinPer)
        '	If Mineral3 Is Nothing = False Then .AddProductionCostItem(Mineral3.ObjectID, ObjectType.eMineral, lMinPer)
        '	If Mineral4 Is Nothing = False Then .AddProductionCostItem(Mineral4.ObjectID, ObjectType.eMineral, lMinPer)

        '	'.ItemCostUB = moResearchCost.ItemCostUB
        '	'ReDim .ItemCosts(.ItemCostUB)
        '	'For X As Int32 = 0 To .ItemCostUB
        '	'	.ItemCosts(X) = New ProductionCostItem()
        '	'	'.MineralCosts(X).MineralID = moResearchCost.MineralCosts(X).MineralID
        '	'	.ItemCosts(X).ItemID = moResearchCost.ItemCosts(X).ItemID
        '	'	.ItemCosts(X).ItemTypeID = moResearchCost.ItemCosts(X).ItemTypeID

        '	'	.ItemCosts(X).oProdCost = moProductionCost
        '	'	.ItemCosts(X).QuantityNeeded = moResearchCost.ItemCosts(X).QuantityNeeded * 10
        '	'Next X
        'End With
    End Sub

    '  Public Overrides Sub ComponentDesigned()
    '      Dim lPropID As Int32

    '      'Make the alloy result but don't save or send it... it is like an extra property of this tech right now...

    '      moAlloyResult = New Mineral()
    '      moAlloyResult.MineralName = AlloyName
    '      moAlloyResult.lAlloyTechID = Me.ObjectID
    '      moAlloyResult.ObjTypeID = ObjectType.eMineral

    '      For lProp As Int32 = 0 To glMineralPropertyUB
    '          lPropID = goMineralProperty(lProp).ObjectID

    '          Dim lNewHigh As Int32 = 0
    '          Dim lNewLow As Int32 = 0
    '          Dim lNewVal As Int32 = 0

    '          SetHighLowValue(lPropID, lNewHigh, lNewLow)

    '          If lNewHigh > 0 Then
    '              lNewVal = lNewHigh
    '          ElseIf lNewLow > 0 Then
    '              lNewVal = lNewLow
    '          Else : lNewVal = CInt(0.8F * GetAverageAKScore(lPropID))
    '          End If

    '          'Now, our new val is set...
    '          moAlloyResult.SetPropertyValue(-1, lPropID, lNewVal)
    'Next lProp

    ''Ok, now, get our complexity score
    'Dim lMaxComp As Int32 = 0
    'If Mineral1 Is Nothing = False Then
    '	lMaxComp = Math.Max(lMaxComp, Mineral1.GetPropertyValue(eMinPropID.Complexity))
    'End If
    'If Mineral2 Is Nothing = False Then
    '	lMaxComp = Math.Max(lMaxComp, Mineral2.GetPropertyValue(eMinPropID.Complexity))
    'End If
    'If Mineral3 Is Nothing = False Then
    '	lMaxComp = Math.Max(lMaxComp, Mineral3.GetPropertyValue(eMinPropID.Complexity))
    'End If
    'If Mineral4 Is Nothing = False Then
    '	lMaxComp = Math.Max(lMaxComp, Mineral4.GetPropertyValue(eMinPropID.Complexity))
    'End If
    'lMaxComp = CInt(Math.Ceiling(lMaxComp * 1.1F))
    'If lMaxComp > 255 Then lMaxComp = 255
    'If lMaxComp < 0 Then lMaxComp = 0
    'moAlloyResult.SetPropertyValue(-1, eMinPropID.Complexity, lMaxComp)

    'CalculateBothProdCosts()
    '  End Sub

    Public Overrides Sub ComponentDesigned()

        Try
            If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
            With moResearchCost
                .ObjectID = Me.ObjectID
                .ObjTypeID = Me.ObjTypeID
                Erase .ItemCosts
                .ItemCostUB = -1
            End With
            If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
            With moProductionCost
                .ObjectID = Me.ObjectID
                .ObjTypeID = Me.ObjTypeID
                Erase .ItemCosts
                .ItemCostUB = -1
            End With

            'Make the alloy result but don't save or send it... it is like an extra property of this tech right now...
            moAlloyResult = New Mineral()
            moAlloyResult.MineralName = AlloyName
            moAlloyResult.lAlloyTechID = Me.ObjectID
            moAlloyResult.ObjTypeID = ObjectType.eMineral

            Dim lMinAmts(3) As Int32
            Dim lMineralIDs(3) As Int32
            For X As Int32 = 0 To 3
                lMinAmts(X) = 0
            Next X
            lMineralIDs(0) = Mineral1ID : lMineralIDs(1) = Mineral2ID : lMineralIDs(2) = Mineral3ID : lMineralIDs(3) = Mineral4ID

            Dim lCombineCnt As Int32 = 0
            Dim fResearchTime As Single = 0.0F
            Dim lMineralsUsedCnt As Int32 = 0
            For X As Int32 = 0 To 3
                If lMineralIDs(X) > 0 Then lMineralsUsedCnt += 1
            Next X

            Dim lAdjResLvl As Int32 = 5 - CInt(Me.ResearchLevel)

            'ok, we have 3 values that are set to desireds...
            If yNewVal1 <> 255 Then
                If lPropertyID1 > 0 Then
                    Dim lValue As Int32 = DoDesiredValueWork(lPropertyID1, yNewVal1, lMineralsUsedCnt, lMineralIDs, lMinAmts, fResearchTime, lCombineCnt, lAdjResLvl)
                    moAlloyResult.SetPropertyValue(-1, lPropertyID1, lValue)
                End If
            End If
            If yNewVal2 <> 255 Then
                If lPropertyID2 > 0 Then
                    Dim lValue As Int32 = DoDesiredValueWork(lPropertyID2, yNewVal2, lMineralsUsedCnt, lMineralIDs, lMinAmts, fResearchTime, lCombineCnt, lAdjResLvl)
                    moAlloyResult.SetPropertyValue(-1, lPropertyID2, lValue)
                End If
            End If
            If yNewVal3 <> 255 Then
                If lPropertyID3 > 0 Then
                    Dim lValue As Int32 = DoDesiredValueWork(lPropertyID3, yNewVal3, lMineralsUsedCnt, lMineralIDs, lMinAmts, fResearchTime, lCombineCnt, lAdjResLvl)
                    moAlloyResult.SetPropertyValue(-1, lPropertyID3, lValue)
                End If
            End If

            Dim lTmp As Int32 = Me.Owner.oSpecials.yAlloyBuilderImprovements
            lTmp -= 1
            lCombineCnt = Math.Max(1, lCombineCnt - lTmp)

            'Now, the remainder of the properties
            For lProp As Int32 = 0 To glMineralPropertyUB
                Dim lPropID As Int32 = goMineralProperty(lProp).ObjectID
                If lPropID <> lPropertyID1 AndAlso lPropID <> lPropertyID2 AndAlso lPropID <> lPropertyID3 AndAlso lPropID <> eMinPropID.Complexity Then
                    Dim lNewVal As Int32 = 0
                    If Mineral1 Is Nothing = False Then lNewVal += Mineral1.GetPropertyValue(lPropID)
                    If Mineral2 Is Nothing = False Then lNewVal += Mineral2.GetPropertyValue(lPropID)
                    If Mineral3 Is Nothing = False Then lNewVal += Mineral3.GetPropertyValue(lPropID)
                    If Mineral4 Is Nothing = False Then lNewVal += Mineral4.GetPropertyValue(lPropID)
                    lNewVal = CInt(Math.Floor((lNewVal / lMineralsUsedCnt) * Math.Pow(0.9F, lAdjResLvl + lCombineCnt)))
                    moAlloyResult.SetPropertyValue(-1, lPropID, lNewVal)
                End If
            Next lProp

            'Now, complexity
            Dim fComplexity As Single = 1.0F + (0.1F * lAdjResLvl)
            fComplexity *= CSng(GetPropertySum(eMinPropID.Complexity) / lMineralsUsedCnt)
            moAlloyResult.SetPropertyValue(-1, eMinPropID.Complexity, CInt(Math.Min(fComplexity, (26 - lAdjResLvl) * 10.0F)))

            Dim blResearchPoints As Int64 = CLng(fResearchTime * (2760 * (6 - lAdjResLvl)) * lCombineCnt)
            Dim blResearchCredits As Int64 = 5000000 \ lAdjResLvl
            Dim blProductionPoints As Int64 = 400 * (6 - lAdjResLvl) * lCombineCnt
            Dim blProductionCredits As Int64 = 100 * (6 - lAdjResLvl) * lCombineCnt

            Me.CurrentSuccessChance = PopIntel - 110
            Me.SuccessChanceIncrement = 1

            Dim bValid As Boolean = True
            Try
                moResearchCost.ObjectID = Me.ObjectID
                moResearchCost.ObjTypeID = Me.ObjTypeID
                moResearchCost.CreditCost = blResearchCredits
                moResearchCost.ColonistCost = 0 : moResearchCost.EnlistedCost = 0 : moResearchCost.OfficerCost = 0
                moResearchCost.PointsRequired = GetResearchPointsCost(blResearchPoints)
                moResearchCost.CreditCost *= Me.ResearchAttempts

                'Minerals Required
                With moResearchCost
                    .ItemCostUB = -1
                    Erase .ItemCosts

                    For X As Int32 = 0 To 3
                        If lMineralIDs(X) > 0 Then
                            .AddProductionCostItem(lMineralIDs(X), ObjectType.eMineral, Math.Max(ResearchLevel, lMinAmts(X)))
                        End If
                    Next X
                End With
            Catch
                ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
                bValid = False
            End Try

            Try

                Dim blTotal As Int64 = 0
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To 3
                    If lMineralIDs(X) > 0 Then
                        lCnt += 1
                        Dim blTemp As Int64 = Math.Max(ResearchLevel, lMinAmts(X))
                        blTotal += (blTemp * blTemp)
                    End If
                Next X

                Dim lFinalMinAmt As Int32 = 5
                If lCnt = 2 Then
                    lFinalMinAmt = 6
                ElseIf lCnt = 3 Then
                    lFinalMinAmt = 4
                ElseIf lCnt = 4 Then
                    lFinalMinAmt = 3
                End If

                'Dim dblTemp As Double = Math.Sqrt(blTotal) / lCnt
                'blTotal = Math.Max(5, CLng(dblTemp / 10))
                'If blTotal > Int32.MaxValue Then lFinalMinAmt = Int32.MaxValue Else lFinalMinAmt = CInt(blTotal)

                moProductionCost.ObjectID = Me.ObjectID
                moProductionCost.ObjTypeID = Me.ObjTypeID
                moProductionCost.ColonistCost = 0
                moProductionCost.EnlistedCost = 0
                moProductionCost.OfficerCost = 0
                moProductionCost.CreditCost = blProductionCredits
                moProductionCost.PointsRequired = blProductionPoints

                With moProductionCost
                    .ItemCostUB = -1
                    Erase .ItemCosts

                    For X As Int32 = 0 To 3
                        If lMineralIDs(X) > 0 Then
                            '.AddProductionCostItem(lMineralIDs(X), ObjectType.eMineral, Math.Max(ResearchLevel, lFinalMinAmt)) ' lMinAmts(X)))
                            .AddProductionCostItem(lMineralIDs(X), ObjectType.eMineral, lFinalMinAmt) ' lMinAmts(X)))
                        End If
                    Next X
                End With
            Catch
                ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
                bValid = False
            End Try
            If bValid = False Then Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
        Catch
            ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
        End Try
    End Sub
    Private Function DoDesiredValueWork(ByVal lPropID As Int32, ByVal yNewVal As Byte, ByVal lMineralsUsedCnt As Int32, ByRef lMineralIDs() As Int32, ByRef lMinAmts() As Int32, ByRef fResearchTime As Single, ByRef lMaxCombineCnt As Int32, ByVal lAdjResLvl As Int32) As Int32
        'ok, here we go... get the first max value we find...
        Dim lMainValue As Int32 = -1
        Dim lMineralID As Int32 = -1
        Dim lCombineCnt As Int32 = 0

        Dim lFirstActVal As Int32 = 0
        Dim lSecondActVal As Int32 = 0

        FillFirstMaxMin(lPropID, lMineralID, lMainValue, -1, lFirstActVal)

        Dim lSecondMineralID As Int32 = -1
        Dim lValueDifference As Int32 = 0

        'Ok, now, check it...
        If yNewVal >= lMainValue Then
            'ok, find second max value
            Dim lSecondMax As Int32 = -1
            FillFirstMaxMin(lPropID, lSecondMineralID, lSecondMax, lMineralID, lSecondActVal)

            Dim lMineralMaxAverage As Int32 = CInt(Math.Ceiling((lMainValue + lSecondMax) / 2.0F))
            lValueDifference = CInt(yNewVal) - lMineralMaxAverage
        Else
            'find first min value
            FillFirstMinMin(lPropID, lMineralID, lMainValue, -1, lFirstActVal)
            Dim lSecondMin As Int32 = -1
            FillFirstMinMin(lPropID, lSecondMineralID, lSecondMin, lMineralID, lSecondActVal)

            Dim lMineralMinAverage As Int32 = CInt(Math.Floor((lMainValue + lSecondMin) / 2.0F))
            lValueDifference = CInt(yNewVal) - lMineralMinAverage
        End If

        If lValueDifference = 0 Then
            For X As Int32 = 0 To 3
                If lMineralIDs(X) = lMineralID OrElse lMineralIDs(X) = lSecondMineralID Then lMinAmts(X) += lAdjResLvl
            Next X
            lCombineCnt += 1
        Else
            For X As Int32 = 0 To 3
                If lMineralIDs(X) = lMineralID OrElse lMineralIDs(X) = lSecondMineralID Then lMinAmts(X) += CInt(((12.0F / lMineralsUsedCnt) + lAdjResLvl) * (lValueDifference * lValueDifference))
            Next X
            fResearchTime = Math.Max(lValueDifference * lValueDifference, fResearchTime)
            lCombineCnt += (lValueDifference * lValueDifference)
        End If

        lMaxCombineCnt = Math.Max(lMaxCombineCnt, lCombineCnt)

        'Ok, now determine our resulting value
        Dim lValue As Int32 = CInt((lFirstActVal + lSecondActVal) / 2.0F)
        Dim yActVal As Byte = PlayerMineral.GetClientDisplayedPropertyValue(lValue)
        If yActVal <> yNewVal Then
            lValue = PlayerMineral.GetRandomValueWithinRange(yNewVal, yActVal)
        End If

        Return lValue
    End Function
    Private Sub FillFirstMaxMin(ByVal lPropertyID As Int32, ByRef lMineralID As Int32, ByRef lMaxValue As Int32, ByVal lExcludeMineralID As Int32, ByRef lActualValue As Int32)
        lMaxValue = -1
        lMineralID = -1
        If Mineral1 Is Nothing = False Then
            With Mineral1
                If .ObjectID <> lExcludeMineralID Then
                    Dim lValue As Int32 = .GetPropertyValue(lPropertyID)
                    Dim lTemp As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lValue)
                    If lTemp > lMaxValue Then
                        lActualValue = lValue
                        lMaxValue = lTemp
                        lMineralID = .ObjectID
                    End If
                End If
            End With
        End If
        If Mineral2 Is Nothing = False Then
            With Mineral2
                If .ObjectID <> lExcludeMineralID Then
                    Dim lValue As Int32 = .GetPropertyValue(lPropertyID)
                    Dim lTemp As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lValue)
                    If lTemp > lMaxValue Then
                        lActualValue = lValue
                        lMaxValue = lTemp
                        lMineralID = .ObjectID
                    End If
                End If
            End With
        End If
        If Mineral3 Is Nothing = False Then
            With Mineral3
                If .ObjectID <> lExcludeMineralID Then
                    Dim lValue As Int32 = .GetPropertyValue(lPropertyID)
                    Dim lTemp As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lValue)
                    If lTemp > lMaxValue Then
                        lActualValue = lValue
                        lMaxValue = lTemp
                        lMineralID = .ObjectID
                    End If
                End If
            End With
        End If
        If Mineral4 Is Nothing = False Then
            With Mineral4
                If .ObjectID <> lExcludeMineralID Then
                    Dim lValue As Int32 = .GetPropertyValue(lPropertyID)
                    Dim lTemp As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lValue)
                    If lTemp > lMaxValue Then
                        lActualValue = lValue
                        lMaxValue = lTemp
                        lMineralID = .ObjectID
                    End If
                End If
            End With
        End If
    End Sub
    Private Sub FillFirstMinMin(ByVal lPropertyID As Int32, ByRef lMineralID As Int32, ByRef lMinValue As Int32, ByVal lExcludeMineralID As Int32, ByRef lActualValue As Int32)
        lMinValue = 255
        lMineralID = -1
        If Mineral1 Is Nothing = False Then
            With Mineral1
                If .ObjectID <> lExcludeMineralID Then
                    Dim lValue As Int32 = .GetPropertyValue(lPropertyID)
                    Dim lTemp As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lValue)
                    If lTemp < lMinValue Then
                        lActualValue = lValue
                        lMinValue = lTemp
                        lMineralID = .ObjectID
                    End If
                End If
            End With
        End If
        If Mineral2 Is Nothing = False Then
            With Mineral2
                If .ObjectID <> lExcludeMineralID Then
                    Dim lValue As Int32 = .GetPropertyValue(lPropertyID)
                    Dim lTemp As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lValue)
                    If lTemp < lMinValue Then
                        lActualValue = lValue
                        lMinValue = lTemp
                        lMineralID = .ObjectID
                    End If
                End If
            End With
        End If
        If Mineral3 Is Nothing = False Then
            With Mineral3
                If .ObjectID <> lExcludeMineralID Then
                    Dim lValue As Int32 = .GetPropertyValue(lPropertyID)
                    Dim lTemp As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lValue)
                    If lTemp < lMinValue Then
                        lActualValue = lValue
                        lMinValue = lTemp
                        lMineralID = .ObjectID
                    End If
                End If
            End With
        End If
        If Mineral4 Is Nothing = False Then
            With Mineral4
                If .ObjectID <> lExcludeMineralID Then
                    Dim lValue As Int32 = .GetPropertyValue(lPropertyID)
                    Dim lTemp As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lValue)
                    If lTemp < lMinValue Then
                        lActualValue = lValue
                        lMinValue = lTemp
                        lMineralID = .ObjectID
                    End If
                End If
            End With
        End If
    End Sub

    Private Function GetAverageAKScore(ByVal lPropID As Int32) As Int32
        Dim lCnt As Int32 = 0
        Dim lVal As Int32 = 0
        On Error Resume Next
        If Mineral1 Is Nothing = False Then
            lCnt += 1
            Dim lActual As Int32 = Mineral1.GetPropertyValue(lPropID)
            Dim lKnown As Int32 = Owner.GetMineralPropertyKnowledge(Mineral1.ObjectID, lPropID)
            Dim lScore As Int32 = CInt(Math.Floor(lActual * (lKnown / (lActual * 0.95F))))
            lVal += lScore
        End If
        If Mineral2 Is Nothing = False Then
            lCnt += 1
            Dim lActual As Int32 = Mineral2.GetPropertyValue(lPropID)
            Dim lKnown As Int32 = Owner.GetMineralPropertyKnowledge(Mineral2.ObjectID, lPropID)
            Dim lScore As Int32 = CInt(Math.Floor(lActual * (lKnown / (lActual * 0.95F))))
            lVal += lScore
        End If
        If Mineral3 Is Nothing = False Then
            lCnt += 1
            Dim lActual As Int32 = Mineral3.GetPropertyValue(lPropID)
            Dim lKnown As Int32 = Owner.GetMineralPropertyKnowledge(Mineral3.ObjectID, lPropID)
            Dim lScore As Int32 = CInt(Math.Floor(lActual * (lKnown / (lActual * 0.95F))))
            lVal += lScore
        End If
        If Mineral4 Is Nothing = False Then
            lCnt += 1
            Dim lActual As Int32 = Mineral4.GetPropertyValue(lPropID)
            Dim lKnown As Int32 = Owner.GetMineralPropertyKnowledge(Mineral4.ObjectID, lPropID)
            Dim lScore As Int32 = CInt(Math.Floor(lActual * (lKnown / (lActual * 0.95F))))
            lVal += lScore
        End If
        Return lVal \ lCnt
    End Function

    Private Sub SetHighLowValue(ByVal lPropID As Int32, ByRef lHigh As Int32, ByRef lLow As Int32)
        lHigh = 0
        lLow = 0

        Dim yType As Byte = 0       '0 unused, 1 higher, 2 lower

        'If lPropID = lPropertyID1 Then
        '    If bHigher1 = True Then
        '        yType = 1
        '    Else : yType = 2
        '    End If
        'ElseIf lPropID = lPropertyID2 Then
        '    If bHigher2 = True Then
        '        yType = 1
        '    Else : yType = 2
        '    End If
        'ElseIf lPropID = lPropertyID3 Then
        '    If bHigher3 = True Then
        '        yType = 1
        '    Else : yType = 2
        '    End If
        'Else : Return
        'End If

        'Now, determine our score
        If yType = 1 Then
            '=IF(F15="H",RandomNumber*2*AVERAGE(B15:E15),0)
            lHigh = CInt((1 + (RandomSeed / fHighLookupVal)) * GetAverageAKScore(lPropID))
        ElseIf yType = 2 Then
            '=IF(F12="L",0.4/RandomNumber*AVERAGE(B12:E12),0)
            'lLow = CInt(0.4F / RandomSeed * GetAverageAKScore(lPropID))
            'lLow = CInt((0.4F + (RandomSeed / 2.0F)) * GetAverageAKScore(lPropID))
            lLow = CInt(((RandomSeed * fLowLookupVal) + 0.1F) * GetAverageAKScore(lPropID))
        End If
    End Sub

    Private ReadOnly Property fLowLookupVal() As Single
        Get
            Select Case ResearchLevel
                Case 0 : Return 0.95F
                Case 1 : Return 0.9F
                Case 2 : Return 0.8F
                Case 3 : Return 0.6F
                Case Else : Return 0.4F
            End Select
        End Get
    End Property

    Private ReadOnly Property fHighLookupVal() As Single
        Get
            Select Case ResearchLevel
                Case 0 : Return 8.0F
                Case 1 : Return 6.0F
                Case 2 : Return 4.0F
                Case 3 : Return 3.0F
                Case Else : Return 2.0F
            End Select
        End Get
    End Property

    Protected Overrides Sub FinalizeResearch()
        Dim lPropID As Int32

        'moAlloyResult = New Mineral()

        'moAlloyResult.MineralName = AlloyName
        'moAlloyResult.lAlloyTechID = Me.ObjectID
        'moAlloyResult.ObjTypeID = ObjectType.eMineral

        'For lProp As Int32 = 0 To glMineralPropertyUB
        '    lPropID = goMineralProperty(lProp).ObjectID

        '    Dim X As Int32 = 0
        '    Dim fBase As Single
        '    Dim fResult As Single

        '    If Mineral1 Is Nothing = False Then
        '        X += Owner.GetMineralPropertyKnowledge(Mineral1.ObjectID, lPropID) - Mineral1.GetPropertyValue(lPropID)
        '    End If
        '    If Mineral2 Is Nothing = False Then
        '        X += Owner.GetMineralPropertyKnowledge(Mineral2.ObjectID, lPropID) - Mineral2.GetPropertyValue(lPropID)
        '    End If
        '    If Mineral3 Is Nothing = False Then
        '        X += Owner.GetMineralPropertyKnowledge(Mineral3.ObjectID, lPropID) - Mineral3.GetPropertyValue(lPropID)
        '    End If
        '    If Mineral4 Is Nothing = False Then
        '        X += Owner.GetMineralPropertyKnowledge(Mineral4.ObjectID, lPropID) - Mineral4.GetPropertyValue(lPropID)
        '    End If
        '    fBase = CInt(Math.Ceiling(X)) * PointAward

        '    If ((lPropID = lPropertyID1) AndAlso bHigher1) OrElse ((lPropID = lPropertyID2) AndAlso bHigher2) OrElse ((lPropID = lPropertyID3) AndAlso bHigher3) Then
        '        'Make the property Higher
        '        fResult = X + GetPropertyMax(lPropID)
        '    ElseIf (lPropID = lPropertyID1) OrElse (lPropID = lPropertyID2) OrElse (lPropID = lPropertyID3) Then
        '        'Make the property lower
        '        fResult = GetPropertyMin(lPropID) - X
        '    Else
        '        fResult = GetPropertyMin(lPropID) + ((GetPropertySum(lPropID) / (8 * Rnd())) / (26 - SuccessDivider))
        '    End If

        '    fResult = Math.Min(Math.Max(0, fResult), 255)

        '    'store the result
        '    moAlloyResult.SetPropertyValue(-1, lPropID, CInt(fResult))
        'Next lProp

        If moAlloyResult Is Nothing Then ComponentDesigned()
        moAlloyResult.SaveObject()

        'Set the player mineral
        Dim lTempIdx As Int32 = Owner.AddPlayerMineral(moAlloyResult.ObjectID, True, -1, False)
        Dim lPlayerMineralID As Int32
        If lTempIdx <> -1 Then
            Owner.oPlayerMinerals(lTempIdx).SaveObject(Owner.ObjectID)
            lPlayerMineralID = Owner.oPlayerMinerals(lTempIdx).PlayerMineralID
        End If

        'Set the Player's initial KNOWN values
        'Dim lMinVal1, lMinVal2, lMinVal3, lMinVal4 As Int32
        Dim lValResult As Int32
        'lMinVal1 = Int32.MaxValue : lMinVal2 = Int32.MaxValue : lMinVal3 = Int32.MaxValue : lMinVal4 = Int32.MaxValue
        For X As Int32 = 0 To Me.Owner.mlMinPropertyUB
            lPropID = Me.Owner.moMinProperties(X).lPropertyID
            'If Mineral1ID > 0 Then lMinVal1 = Owner.GetMineralPropertyKnowledge(Mineral1ID, lPropID)
            'If Mineral2ID > 0 Then lMinVal2 = Owner.GetMineralPropertyKnowledge(Mineral2ID, lPropID)
            'If Mineral3ID > 0 Then lMinVal3 = Owner.GetMineralPropertyKnowledge(Mineral3ID, lPropID)
            'If Mineral4ID > 0 Then lMinVal4 = Owner.GetMineralPropertyKnowledge(Mineral4ID, lPropID)

            'lValResult = Math.Min(Math.Min(Math.Min(lMinVal1, lMinVal2), lMinVal3), lMinVal4)
            lValResult = moAlloyResult.GetPropertyValue(lPropID)

            Owner.SetPlayerMineralProperty(-1, lPlayerMineralID, lPropID, lValResult)
        Next X

        'put the moAlloyResult into goMinerals
        glMineralUB += 1
        ReDim Preserve goMineral(glMineralUB)
        ReDim Preserve glMineralIdx(glMineralUB)
        goMineral(glMineralUB) = moAlloyResult
        goMineral(glMineralUB).ServerIndex = glMineralUB
        glMineralIdx(glMineralUB) = moAlloyResult.ObjectID

        'Now, send the playermineral to the player
        If Owner.lConnectedPrimaryID > -1 OrElse Owner.HasOnlineAliases(AliasingRights.eAddProduction Or AliasingRights.eViewResearch Or AliasingRights.eAddResearch Or AliasingRights.eViewTechDesigns) = True Then
            Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Owner.oPlayerMinerals(lTempIdx), GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddProduction Or AliasingRights.eViewResearch Or AliasingRights.eAddResearch Or AliasingRights.eViewTechDesigns)
        End If

        Me.SaveObject()

        For X As Int32 = 0 To Me.Owner.lPlayerMineralUB
            If Me.Owner.oPlayerMinerals(X) Is Nothing = False AndAlso Me.Owner.oPlayerMinerals(X).lMineralID = moAlloyResult.ObjectID Then
                Me.Owner.oPlayerMinerals(X).SaveObject(Me.Owner.ObjectID)
                Exit For
            End If
        Next X

        'ok, now, we force all primaries to sync the alloy result mineral
        If moAlloyResult Is Nothing = False Then
            goMsgSys.SendMsgToOperator(moAlloyResult.GetForcePrimarySyncMsg)
        End If

    End Sub

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(27) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yResult, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yResult, lPos) : lPos += 6
        Me.AlloyName.CopyTo(yResult, lPos) : lPos += 20
        Return yResult
    End Function

    Public Sub SetLoadedAlloyResultProperty(ByVal lPropID As Int32, ByVal lValue As Int32)
        If moAlloyResult Is Nothing Then moAlloyResult = New Mineral
        moAlloyResult.MineralName = AlloyName
        moAlloyResult.lAlloyTechID = Me.ObjectID
        moAlloyResult.ObjTypeID = ObjectType.eMineral
        moAlloyResult.SetPropertyValue(-1, lPropID, lValue)
    End Sub

    Public Overrides Function TechnologyScore() As Integer
        Return 1000I
    End Function

    Public Overrides Function GetPlayerTechKnowledgeMsg(ByVal yTechLvl As PlayerTechKnowledge.KnowledgeType) As Byte()
        'SHOULD NEVER BE CALLED!!!
        Return Nothing
    End Function

    Public Overrides Sub FillFromPrimaryAddMsg(ByVal yData() As Byte)

        Dim lPos As Int32 = 2   'for msgcode
        With Me
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = ObjectType.eAlloyTech : lPos += 2
            .Owner = GetEpicaPlayer(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
            .ComponentDevelopmentPhase = CType(System.BitConverter.ToInt32(yData, lPos), Epica_Tech.eComponentDevelopmentPhase) : lPos += 4
            .ErrorReasonCode = yData(lPos) : lPos += 1
            lPos += 1   'researchercnt
            .MajorDesignFlaw = yData(lPos) : lPos += 1

            ReDim .AlloyName(19)
            Array.Copy(yData, lPos, .AlloyName, 0, 20) : lPos += 2

            .Mineral1ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Mineral2ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Mineral3ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Mineral4ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .lPropertyID1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lPropertyID2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lPropertyID3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .yNewVal1 = yData(lPos) : lPos += 1
            .yNewVal2 = yData(lPos) : lPos += 1
            .yNewVal3 = yData(lPos) : lPos += 1

            .ResearchLevel = yData(lPos) : lPos += 1

            Dim lAlloyResult As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            If lAlloyResult > 0 Then
                moAlloyResult = GetEpicaMineral(lAlloyResult)
            End If

            If Me.moResearchCost Is Nothing Then
                Me.moResearchCost = New ProductionCost
                lPos = Me.moResearchCost.FillFromPrimaryAddMsg(yData, lPos)
            Else
                Dim oTmp As New ProductionCost
                lPos = oTmp.FillFromPrimaryAddMsg(yData, lPos)
            End If

            If Me.moProductionCost Is Nothing Then
                Me.moProductionCost = New ProductionCost
                lPos = Me.moProductionCost.FillFromPrimaryAddMsg(yData, lPos)
            Else
                Dim oTmp As New ProductionCost
                lPos = oTmp.FillFromPrimaryAddMsg(yData, lPos)
            End If

            If .Owner Is Nothing = False Then
                Dim lFirstIdx As Int32 = -1
                For X As Int32 = 0 To .Owner.mlAlloyUB
                    If .Owner.mlAlloyIdx(X) = .ObjectID Then
                        .Owner.moAlloy(X) = Me
                        Return
                    ElseIf lFirstIdx = -1 AndAlso .Owner.mlAlloyIdx(X) = -1 Then
                        lFirstIdx = X
                    End If
                Next X

                If lFirstIdx = -1 Then
                    lFirstIdx = .Owner.mlAlloyUB + 1
                    ReDim Preserve .Owner.mlAlloyIdx(lFirstIdx)
                    ReDim Preserve .Owner.moAlloy(lFirstIdx)
                    .Owner.mlAlloyUB = lFirstIdx
                End If
                .Owner.moAlloy(lFirstIdx) = Me
                .Owner.mlAlloyIdx(lFirstIdx) = Me.ObjectID
            End If
        End With
        
    End Sub
End Class
