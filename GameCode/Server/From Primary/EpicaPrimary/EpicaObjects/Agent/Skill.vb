Public Enum lSkillHardcodes As Int32
    eBattleLanguage = 4
    eCryptography = 9
    eDeductiveReasoning = 11
    eDisguises = 14
    eElectronics = 15
    eEndurance = 16
    eEscapeArtist = 18
    eEtiquetteGeneral = 19
    eEtiquetteMilitary = 20
    eEtiquetteFederal = 21
    eEtiquetteColony = 22
    eEtiquetteWorker = 23
    eForger = 25
    eHonor = 28
    eInterrogation = 29
    eLaw = 30
    eLeadership = 31
    eLegitimateCover = 32
    eNaturallyTalented = 36
    eParanoid = 38
    ePerceptive = 39
    ePersuasive = 40
    ePoisons = 43
    eSecurity = 48
    eSpyGames = 51
    eStealth = 52
    eTinyElectronics = 35
    eTorture = 55
    eTracking = 56
    eTreachery = 57
    eWeaponSpecialist = 58
End Enum

Public Class Skill
    Inherits Epica_GUID

    Public Const SKILL_MSG_LEN As Int32 = 284

    Public SkillName(19) As Byte
    Public MinVal As Byte
    Public MaxVal As Byte
    Public SkillType As Byte
    Public SkillDesc(254) As Byte

    Public RelatedSkills() As SkillRel
    Public lRelatedSkillUB As Int32 = -1

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        If mbStringReady = False Then
            ReDim mySendString(SKILL_MSG_LEN - 1)

            GetGUIDAsString.CopyTo(mySendString, 0)
            SkillName.CopyTo(mySendString, 6)
            mySendString(26) = MinVal
            mySendString(27) = MaxVal
            mySendString(28) = SkillType

            SkillDesc.CopyTo(mySendString, 29)

            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblSkill (SkillName, MinVal, MaxVal, SkillType, SkillDesc) VALUES ('" & _
                  MakeDBStr(BytesToString(SkillName)) & "', " & MinVal & ", " & MaxVal & ", " & SkillType & ", '" & _
                  MakeDBStr(BytesToString(SkillDesc)) & "')"
            Else
                'UPDATE
                sSQL = "UPDATE tblSkill SET SkillName = '" & MakeDBStr(BytesToString(SkillName)) & "', MinVal = " & MinVal & ", MaxVal = " & _
                  MaxVal & ", SkillType = " & SkillType & ", SkillDesc = '" & MakeDBStr(BytesToString(SkillDesc)) & _
                  "' WHERE SkillID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(SkillID) FROM tblSkill WHERE SkillName = '" & MakeDBStr(BytesToString(SkillName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function
End Class

Public Structure SkillRel
    Public oToSkill As Skill
    Public lToHitNumber As Int32
End Structure