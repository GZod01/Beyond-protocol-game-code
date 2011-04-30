Option Strict On

Public Class RouteTemplate
    Public lTemplateID As Int32 = -1
    Public lPlayerID As Int32

    Public TemplateName(19) As Byte

    Public uItems() As RouteItem
    Public lItemUB As Int32 = -1

    Public Function SaveObject() As Boolean
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim bResult As Boolean = False
        Try
            If lTemplateID > 0 Then
                oComm = New OleDb.OleDbCommand("DELETE FROM tblRouteTemplateItem WHERE TemplateID = " & Me.lTemplateID, goCN)
                oComm.ExecuteNonQuery()
                oComm.Dispose()
                oComm = Nothing
            End If

            Dim sSQL As String

            Dim sName As String = MakeDBStr(BytesToString(TemplateName))

            If lTemplateID < 1 Then
                'insert
                sSQL = "INSERT INTO tblRouteTemplate (PlayerID, TemplateName) VALUES (" & Me.lPlayerID & ", '" & sName & "')"
            Else
                'update
                sSQL = "UPDATE tblRouteTemplate SET PlayerID = " & lPlayerID & ", TemplateName = '" & sName & "' WHERE TemplateID = " & Me.lTemplateID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Throw New Exception("Statement returned 0 results: " & sSQL)
            End If
            oComm.Dispose()
            oComm = Nothing

            If lTemplateID < 1 Then
                oComm = New OleDb.OleDbCommand("SELECT Max(TemplateID) FROM tblRouteTemplate WHERE TemplateName = '" & sName & "'", goCN)
                Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader()
                If oData Is Nothing = False AndAlso oData.Read = True Then
                    lTemplateID = CInt(oData(0))
                Else
                    oData = Nothing
                    Throw New Exception("Could not get template ID from table.")
                End If
                oData.Close()
                oData = Nothing
                oComm.Dispose()
                oComm = Nothing
            End If

            For X As Int32 = 0 To lItemUB
                With uItems(X)
                    sSQL = "INSERT INTO tblRouteTemplateItem (TemplateID, OrderNum, DestID, DestTypeID, LocX, LocZ, LoadItemID, LoadItemTypeID, ExtraFlags) VALUES ("
                    sSQL &= lTemplateID.ToString & ", " & X.ToString & ", "
                    If .oDest Is Nothing = False Then
                        sSQL &= .oDest.ObjectID & ", " & .oDest.ObjTypeID & ", "
                    Else
                        sSQL &= "-1, -1, "
                    End If
                    sSQL &= .lLocX.ToString & ", " & .lLocZ.ToString & ", "
                    If .oLoadItem Is Nothing = False Then
                        sSQL &= .oLoadItem.ObjectID & ", " & .oLoadItem.ObjTypeID & ", "
                    Else
                        sSQL &= "-1, -1, "
                    End If
                    sSQL &= CByte(.yExtraFlags).ToString & ")"

                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        Throw New Exception("Statement returned 0 results: " & sSQL)
                    End If
                    oComm.Dispose()
                    oComm = Nothing
                End With
            Next X

            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Unable to save route template: " & ex.Message)
        End Try
        If oComm Is Nothing = False Then oComm.Dispose()
        oComm = Nothing

        Return bResult
    End Function

    Public Sub AssignRouteToUnit(ByRef oUnit As Unit)
        With oUnit
            .lRouteUB = -1
            .bRoutePaused = False
            .bRunRouteOnce = False
            .lCurrentRouteIdx = -1
            For X As Int32 = 0 To lItemUB
                .AddRouteItem(uItems(X))
            Next X
        End With
    End Sub

    Public Sub DeleteMe()
        If Me.lTemplateID < 1 Then Return

        Dim oComm As OleDb.OleDbCommand = Nothing
        Try
            oComm = New OleDb.OleDbCommand("DELETE FROM tblRouteTemplateItem WHERE TemplateID = " & Me.lTemplateID, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            oComm = New OleDb.OleDbCommand("DELETE FROM tblRouteTemplate WHERE TemplateID = " & Me.lTemplateID, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Unable to delete template: " & lTemplateID & ". " & ex.Message)
        End Try
        If oComm Is Nothing = False Then oComm.Dispose()
        oComm = Nothing

    End Sub

    'Public Shared Sub test()
    '    If True = True Then
    '        Dim sSQL As String
    '        Dim oComm As OleDb.OleDbCommand
    '        Dim oData As OleDb.OleDbDataReader
    '        Dim sPlayerInStr As String = "OwnerID = 3200"
    '        Dim oPlayer As Player = GetEpicaPlayer(3200)

    '        oPlayer.mlAlloyUB = -1

    '        'Alloy
    '        LogEvent(LogEventType.Informational, "Loading Alloys Techs...")
    '        sSQL = "SELECT * FROM tblAlloy WHERE " & sPlayerInStr
    '        oComm = New OleDb.OleDbCommand(sSQL, goCN)
    '        oData = oComm.ExecuteReader(CommandBehavior.Default)
    '        While oData.Read

    '            If oPlayer Is Nothing = False Then
    '                oPlayer.mlAlloyUB += 1
    '                ReDim Preserve oPlayer.mlAlloyIdx(oPlayer.mlAlloyUB)
    '                ReDim Preserve oPlayer.moAlloy(oPlayer.mlAlloyUB)

    '                oPlayer.moAlloy(oPlayer.mlAlloyUB) = New AlloyTech
    '                With oPlayer.moAlloy(oPlayer.mlAlloyUB)
    '                    .AlloyName = StringToBytes(CStr(oData("AlloyName")))
    '                    .ObjectID = CInt(oData("AlloyID"))
    '                    .ObjTypeID = ObjectType.eAlloyTech
    '                    .Owner = oPlayer

    '                    If CInt(oData("ResultMineralID")) > 0 Then .AlloyResult = GetEpicaMineral(CInt(oData("ResultMineralID")))
    '                    .Mineral1ID = CInt(oData("Mineral1ID"))
    '                    .Mineral2ID = CInt(oData("Mineral2ID"))
    '                    .Mineral3ID = CInt(oData("Mineral3ID"))
    '                    .Mineral4ID = CInt(oData("Mineral4ID"))
    '                    .lPropertyID1 = CInt(oData("MineralProperty1ID"))
    '                    .lPropertyID2 = CInt(oData("MineralProperty2ID"))
    '                    .lPropertyID3 = CInt(oData("MineralProperty3ID")) 
    '                    .yNewVal1 = CByte(oData("Property1Value"))
    '                    .yNewVal2 = CByte(oData("Property2Value"))
    '                    .yNewVal3 = CByte(oData("Property3Value"))
    '                    .ResearchLevel = CByte(oData("ResearchLevel"))
    '                    .MajorDesignFlaw = CByte(oData("MajorDesignFlaw"))
    '                    .PopIntel = CInt(oData("PopIntel"))
    '                    '.lVersionNum = CInt(oData("VersionNumber"))
    '                    If oData("VersionNumber") Is DBNull.Value Then
    '                        .lVersionNum = Epica_Tech.TechVersionNum
    '                    Else
    '                        .lVersionNum = CInt(oData("VersionNumber"))
    '                    End If

    '                    .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
    '                    .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
    '                    .ResearchAttempts = CInt(oData("ResearchAttempts"))
    '                    .RandomSeed = CSng(oData("RandomSeed"))
    '                    .ComponentDevelopmentPhase = CType(oData("ResPhase"), Epica_Tech.eComponentDevelopmentPhase)

    '                    .yArchived = CByte(oData("bArchived"))

    '                    oPlayer.mlAlloyIdx(oPlayer.mlAlloyUB) = .ObjectID

    '                    AddQuickTechnologyLookup(.ObjectID, .ObjTypeID, oPlayer.ObjectID)
    '                End With
    '            End If
    '        End While
    '        oData.Close()
    '        oData = Nothing
    '        oComm = Nothing

    '        'Alloy Result Properties
    '        LogEvent(LogEventType.Informational, "Loading Alloy Result Properties...")
    '        sSQL = "SELECT arp.AlloyID, arp.MineralPropertyID, arp.PropertyValue, a.OwnerID FROM tblAlloyResultProperty arp " & _
    '          "LEFT OUTER JOIN tblAlloy a ON a.AlloyID = arp.AlloyID " & _
    '          " WHERE " & sPlayerInStr & " ORDER BY arp.AlloyID"
    '        oComm = New OleDb.OleDbCommand(sSQL, goCN)
    '        oData = oComm.ExecuteReader(CommandBehavior.Default)
    '        While oData.Read

    '            If oPlayer Is Nothing = False Then
    '                Dim oAlloy As AlloyTech = CType(oPlayer.GetTech(CInt(oData("AlloyID")), ObjectType.eAlloyTech), AlloyTech)
    '                If oAlloy Is Nothing = False Then
    '                    oAlloy.SetLoadedAlloyResultProperty(CInt(oData("MineralPropertyID")), CInt(oData("PropertyValue")))
    '                End If
    '            End If
    '        End While
    '        oData.Close()
    '        oData = Nothing
    '        oComm = Nothing

    '        sSQL = "SELECT pc.PC_ID, pc.ProductionCostType, pc.PointsRequired, pc.Officers, pc.Colonists, pc.Credits, pc.Enlisted, pc.ObjectID, pc.ObjTypeID, t.OwnerID FROM "
    '        sSQL &= "tblAlloy t LEFT OUTER JOIN tblProductionCost pc ON pc.ObjectID = t.AlloyID "
    '        sSQL &= "WHERE pc.ObjTypeID = " & ObjectType.eAlloyTech & " and t.OwnerID = 3200"

    '        oComm = New OleDb.OleDbCommand(sSQL, goCN)
    '        oData = oComm.ExecuteReader(CommandBehavior.Default)

    '        Dim oProdCosts() As ProductionCost
    '        Dim lProdCostUB As Int32 = -1

    '        Dim lIdx As Int32 = -1
    '        While oData.Read
    '            lIdx += 1

    '            lProdCostUB += 1
    '            ReDim Preserve oProdCosts(lProdCostUB)

    '            oProdCosts(lIdx) = New ProductionCost
    '            With oProdCosts(lIdx)
    '                .ColonistCost = CInt(oData("Colonists"))
    '                .CreditCost = CLng(oData("Credits"))
    '                .EnlistedCost = CInt(oData("Enlisted"))
    '                .ObjectID = CInt(oData("ObjectID"))
    '                .ObjTypeID = CShort(oData("ObjTypeID"))
    '                .OfficerCost = CInt(oData("Officers"))
    '                .PC_ID = CInt(oData("PC_ID"))
    '                .PointsRequired = CLng(oData("PointsRequired"))
    '                .ProductionCostType = CByte(oData("ProductionCostType"))
    '                .ItemCostUB = -1
    '            End With

    '            If oPlayer Is Nothing = False AndAlso (oPlayer.InMyDomain = True OrElse oPlayer.ObjectID = 0) Then
    '                Dim oTech As Epica_Tech = oPlayer.GetTech(oProdCosts(lIdx).ObjectID, oProdCosts(lIdx).ObjTypeID)
    '                If oTech Is Nothing = False Then
    '                    oTech.SetTechCost(oProdCosts(lIdx))
    '                End If
    '            End If
    '        End While
    '        oData.Close()
    '        oData = Nothing
    '        oComm = Nothing


    '        'Now, do one call to the database for all productioncostitems
    '        LogEvent(LogEventType.Informational, "Loading ProductionCostItems...")
    '        sSQL = "select * from tblproductioncostitem where pc_id in (select pc_id from tblproductioncost where objtypeid = " & CInt(ObjectType.eAlloyTech) & " AND ObjectID IN (SELECT AlloyID FROM tblAlloy WHERE OwnerID = 3200))"
    '        oComm = New OleDb.OleDbCommand(sSQL, goCN)
    '        oData = oComm.ExecuteReader(CommandBehavior.Default)
    '        While oData.Read = True
    '            Dim lPCID As Int32 = CInt(oData("PC_ID"))

    '            For X As Int32 = 0 To lProdCostUB
    '                If oProdCosts(X) Is Nothing = False AndAlso oProdCosts(X).PC_ID = lPCID Then

    '                    oProdCosts(X).ItemCostUB += 1
    '                    ReDim Preserve oProdCosts(X).ItemCosts(oProdCosts(X).ItemCostUB)
    '                    oProdCosts(X).ItemCosts(oProdCosts(X).ItemCostUB) = New ProductionCostItem()
    '                    With oProdCosts(X).ItemCosts(oProdCosts(X).ItemCostUB)
    '                        .ItemID = CInt(oData("ItemID"))
    '                        .ItemTypeID = CShort(oData("ItemTypeID"))
    '                        .oProdCost = oProdCosts(X)
    '                        .PC_ID = lPCID
    '                        .PCM_ID = CInt(oData("PCM_ID"))
    '                        .QuantityNeeded = CInt(oData("Quantity"))
    '                    End With

    '                    Exit For
    '                End If
    '            Next X

    '        End While
    '        oData.Close()
    '        oData = Nothing
    '        oComm = Nothing

    '    End If
    'End Sub

End Class
