Option Strict On

Public Class TradePostContents
    Public lTradePostID As Int32
    Public lColonists As Int32
    Public lEnlisted As Int32
    Public lOfficers As Int32

    Public ySellSlotsUsed As Byte
    Public yBuySlotsUsed As Byte

    Private Structure CompCache
        Public lObjectID As Int32
        Public iObjTypeID As Int16
        Public lComponentID As Int32
        Public iComponentTypeID As Int16
        Public lComponentOwnerID As Int32
        Public lQuantity As Int32
        Public lForSaleQty As Int32
        Public blPrice As Int64
    End Structure
    Private Structure UnitItem
        Public lObjectID As Int32
        Public iObjTypeID As Int16 
        Public lForSaleQty As Int32
        Public blPrice As Int64
        Public yChassisType As Byte
    End Structure
    Private Structure MinCache
        Public lObjectID As Int32
        Public iObjTypeID As Int16
        Public lMineralID As Int32
        Public lQty As Int32
        Public lForSaleQty As Int32
        Public blPrice As Int64
	End Structure
	Private Structure IntelItem
		Public lObjectID As Int32
		Public iObjTypeID As Int16
		Public iExtTypeID As Int16
		Public blPrice As Int64
	End Structure

    Private muComponents() As CompCache
    Private mlComponentUB As Int32 = -1
    Private muMinerals() As MinCache
    Private mlMineralUB As Int32 = -1
    Private muUnits() As UnitItem
	Private mlUnitUB As Int32 = -1
	'Only to be used for trackign items for sale
	Private muIntel() As IntelItem
	Private mlIntelUB As Int32 = -1

    Public Shared oTradePostContents() As TradePostContents
    Public Shared lTradePostContentsIdx() As Int32
    Public Shared lTradePostContentsUB As Int32 = -1

    Public lLastUpdate As Int32 = -1

    Public Sub AddComponentCache(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lComponentID As Int32, ByVal iComponentTypeID As Int16, ByVal lComponentOwnerID As Int32, ByVal lQty As Int32, ByVal lForSale As Int32, ByVal blPrice As Int64)
        mlComponentUB += 1
        ReDim Preserve muComponents(mlComponentUB)
        With muComponents(mlComponentUB)
            .lObjectID = lID
            .iObjTypeID = iTypeID
            .lComponentID = lComponentID
            .iComponentTypeID = iComponentTypeID
            .lComponentOwnerID = lComponentOwnerID
            .lQuantity = lQty
            .lForSaleQty = lForSale
            .blPrice = blPrice
        End With
    End Sub

    Public Sub AddMineralCache(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lMineralID As Int32, ByVal lQty As Int32, ByVal lForSale As Int32, ByVal blPrice As Int64)
        mlMineralUB += 1
        ReDim Preserve muMinerals(mlMineralUB)
        With muMinerals(mlMineralUB)
            .lObjectID = lID
            .iObjTypeID = iTypeID
            .lMineralID = lMineralID
            .lQty = lQty
            .lForSaleQty = lForSale
            .blPrice = blPrice
        End With
    End Sub

    Public Sub AddUnit(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lForSale As Int32, ByVal blPrice As Int64, ByVal yChassisType As Byte)
        mlUnitUB += 1
        ReDim Preserve muUnits(mlUnitUB)
        With muUnits(mlUnitUB)
            .lObjectID = lID
            .iObjTypeID = iTypeID
            .lForSaleQty = lForSale
            .blPrice = blPrice
            .yChassisType = yChassisType
        End With
	End Sub

	Public Sub AddIntel(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal iExtTypeID As Int16, ByVal blPrice As Int64)
		mlIntelUB += 1
		ReDim Preserve muIntel(mlIntelUB)
		With muIntel(mlIntelUB)
			.lObjectID = lID
			.iObjTypeID = iTypeID
			.iExtTypeID = iExtTypeID
			.blPrice = blPrice
		End With
	End Sub

    Public Sub PopulateList(ByRef lstData As UIListBox, ByVal yType As Byte)
		Dim sName As String = ""

		Dim clrSpecial As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 0)

        lstData.Clear()
        'Now, fill our list
        If lstData.oIconTexture Is Nothing OrElse lstData.oIconTexture.Disposed = True Then
            lstData.oIconTexture = goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
        End If
        lstData.RenderIcons = True

        '  Credits
		If yType = 1 OrElse yType = 255 Then
			Dim bPlayer As Boolean = True
			If yType = 255 Then
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					If goCurrentPlayer.oGuild.lGuildHallID = Me.lTradePostID Then
						bPlayer = False
					End If
				End If
			End If
			If bPlayer = True Then
				lstData.AddItem("Credits (" & goCurrentPlayer.blCredits.ToString("#,##0") & ")", False)
			Else
				lstData.AddItem("Credits (" & goCurrentPlayer.oGuild.blTreasury.ToString("#,##0") & ")", False)
			End If
			lstData.ItemData(lstData.NewIndex) = 1 : lstData.ItemData2(lstData.NewIndex) = ObjectType.eCredits
			lstData.ApplyIconOffset(lstData.NewIndex) = False

			If yType = 1 Then
				lstData.AddItem("AGENTS", True)
				lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
				lstData.ApplyIconOffset(lstData.NewIndex) = False
				Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName)
				If lSorted Is Nothing = False Then
					For X As Int32 = 0 To lSorted.GetUpperBound(0)
						Dim oAgent As Agent = goCurrentPlayer.Agents(lSorted(X))
						If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.OnAMission)) = 0 Then
							lstData.AddItem(oAgent.sAgentName, False)
							lstData.ItemData(lstData.NewIndex) = oAgent.ObjectID : lstData.ItemData2(lstData.NewIndex) = ObjectType.eAgent
							lstData.ApplyIconOffset(lstData.NewIndex) = True
						End If
					Next X
				End If
			End If
		End If

		'Components
		lstData.AddItem("COMPONENTS", True)
		lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
		lstData.ApplyIconOffset(lstData.NewIndex) = False
		For X As Int32 = 0 To mlComponentUB
			With muComponents(X)
				sName = GetCacheObjectValue(.lComponentID, .iComponentTypeID)
                lstData.AddItem(sName & " (" & .lQuantity.ToString("#,##0") & ")", (.lForSaleQty <> 0 AndAlso .blPrice <> 0))
				If lstData.ItemBold(lstData.NewIndex) = True Then lstData.ItemCustomColor(lstData.NewIndex) = clrSpecial
				lstData.ItemData(lstData.NewIndex) = .lObjectID : lstData.ItemData2(lstData.NewIndex) = -.iComponentTypeID

				Select Case Math.Abs(.iComponentTypeID)
					Case ObjectType.eArmorTech
						lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(0, 0, 16, 16)
						lstData.IconForeColor(lstData.NewIndex) = Color.White
					Case ObjectType.eEngineTech
						lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(0, 16, 16, 16)
						lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
					Case ObjectType.eRadarTech
						lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(48, 0, 16, 16)
						lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
					Case ObjectType.eShieldTech
						lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 16, 16, 16)
						lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 255)
					Case ObjectType.eWeaponTech
						lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(32, 32, 16, 16)
						lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
				End Select
				lstData.ApplyIconOffset(lstData.NewIndex) = True
			End With
		Next X

		If yType = 1 Then
			'Component Designs
			'lstData.AddItem("COMPONENT DESIGNS", True)
			'lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
			'lstData.ApplyIconOffset(lstData.NewIndex) = False

			'If goCurrentPlayer Is Nothing = False Then
			'    For X As Int32 = 0 To goCurrentPlayer.mlTechUB
			'        If goCurrentPlayer.moTechs(X) Is Nothing = False Then
			'            With goCurrentPlayer.moTechs(X)
			'                If .ObjTypeID <> ObjectType.eSpecialTech AndAlso .ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
			'                    sName = .GetComponentName()
			'                    lstData.AddItem(sName, False)
			'                    lstData.ItemData(lstData.NewIndex) = .ObjectID : lstData.ItemData2(lstData.NewIndex) = .ObjTypeID

			'                    Select Case Math.Abs(.ObjTypeID)
			'                        Case ObjectType.eArmorTech
			'                            lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(0, 0, 16, 16)
			'                            lstData.IconForeColor(lstData.NewIndex) = Color.White
			'                        Case ObjectType.eEngineTech
			'                            lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(0, 16, 16, 16)
			'                            lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
			'                        Case ObjectType.eRadarTech
			'                            lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(48, 0, 16, 16)
			'                            lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
			'                        Case ObjectType.eShieldTech
			'                            lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 16, 16, 16)
			'                            lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 255)
			'                        Case ObjectType.eWeaponTech
			'                            lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(32, 32, 16, 16)
			'                            lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
			'                    End Select
			'                    lstData.ApplyIconOffset(lstData.NewIndex) = True
			'                End If
			'            End With
			'        End If
			'    Next X
			'End If

			'Facilities
			lstData.AddItem("FACILITIES", True)
			lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
			lstData.ApplyIconOffset(lstData.NewIndex) = False
			Dim oEnvir As BaseEnvironment = goCurrentEnvir
			If oEnvir Is Nothing = False Then
                For X As Int32 = 0 To oEnvir.lEntityUB
                    If oEnvir.lEntityIdx(X) <> -1 AndAlso oEnvir.oEntity(X).ObjTypeID = ObjectType.eFacility AndAlso oEnvir.oEntity(X).OwnerID = glPlayerID AndAlso (oEnvir.oEntity(X).oUnitDef Is Nothing = False AndAlso (oEnvir.oEntity(X).oUnitDef.ModelID And 255) <> 148) Then
                        With oEnvir.oEntity(X)
                            If .yProductionType = ProductionType.eTradePost OrElse .yProductionType = ProductionType.eCommandCenterSpecial Then Continue For
                            sName = GetCacheObjectValue(.ObjectID, .ObjTypeID) ' .EntityName
                            lstData.AddItem(sName, False)
                            lstData.ItemData(lstData.NewIndex) = .ObjectID : lstData.ItemData2(lstData.NewIndex) = .ObjTypeID
                            lstData.ApplyIconOffset(lstData.NewIndex) = True
                        End With
                    End If
                Next X
			End If
		End If

		'Materials
		lstData.AddItem("MATERIALS", True)
		lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
		lstData.ApplyIconOffset(lstData.NewIndex) = False
		For X As Int32 = 0 To mlMineralUB
			With muMinerals(X)
				For Y As Int32 = 0 To glMineralUB
					If glMineralIdx(Y) = .lMineralID Then
						If goMinerals(Y).bDiscovered = True Then
							sName = goMinerals(Y).MineralName
                            lstData.AddItem(sName & " (" & .lQty.ToString("#,##0") & ")", (.lForSaleQty <> 0 AndAlso .blPrice <> 0))
							If lstData.ItemBold(lstData.NewIndex) = True Then lstData.ItemCustomColor(lstData.NewIndex) = clrSpecial
							lstData.ItemData(lstData.NewIndex) = .lObjectID : lstData.ItemData2(lstData.NewIndex) = .iObjTypeID : lstData.ItemData3(lstData.NewIndex) = .lMineralID
							lstData.ApplyIconOffset(lstData.NewIndex) = True
						End If
						Exit For
					End If
				Next Y
			End With
		Next X

		If yType = 1 Then
			'Personnel
			lstData.AddItem("PERSONNEL", True)
			lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
			lstData.ApplyIconOffset(lstData.NewIndex) = False

            lstData.AddItem("Colonists (" & lColonists.ToString("#,##0") & ")", False)
			lstData.ItemData(lstData.NewIndex) = 1 : lstData.ItemData2(lstData.NewIndex) = ObjectType.eColonists
			lstData.ApplyIconOffset(lstData.NewIndex) = True
			lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 0, 16, 16)
			lstData.IconForeColor(lstData.NewIndex) = Color.White

            lstData.AddItem("Enlisted (" & lEnlisted.ToString("#,##0") & ")", False)
			lstData.ItemData(lstData.NewIndex) = 1 : lstData.ItemData2(lstData.NewIndex) = ObjectType.eEnlisted
			lstData.ApplyIconOffset(lstData.NewIndex) = True
			lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 0, 16, 16)
			lstData.IconForeColor(lstData.NewIndex) = Color.FromArgb(255, 64, 192, 64)

            lstData.AddItem("Officers (" & lOfficers.ToString("#,##0") & ")", False)
			lstData.ItemData(lstData.NewIndex) = 1 : lstData.ItemData2(lstData.NewIndex) = ObjectType.eOfficers
			lstData.ApplyIconOffset(lstData.NewIndex) = True
			lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 0, 16, 16)
			lstData.IconForeColor(lstData.NewIndex) = Color.FromArgb(255, 192, 192, 0)
		End If

		'Player Intel
		If yType <> 255 Then
			lstData.AddItem("PLAYER STATS INTEL", True)
			lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
			lstData.ApplyIconOffset(lstData.NewIndex) = False
			If goCurrentPlayer Is Nothing = False Then
				For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
					Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(X)
					If oRel Is Nothing = False Then
						If oRel.oPlayerIntel Is Nothing = False Then
							With oRel.oPlayerIntel
								If .lDiplomacyScore <> Int32.MinValue OrElse .lMilitaryScore <> Int32.MinValue OrElse _
								   .lPopulationScore <> Int32.MinValue OrElse .lProductionScore <> Int32.MinValue OrElse _
								   .lTechnologyScore <> Int32.MinValue OrElse .lWealthScore <> Int32.MinValue Then

									Dim bBold As Boolean = False
									For Z As Int32 = 0 To mlIntelUB
										If muIntel(Z).lObjectID = oRel.lThisPlayer AndAlso muIntel(Z).iExtTypeID = ObjectType.ePlayer AndAlso muIntel(Z).iObjTypeID = ObjectType.ePlayerIntel Then
											If muIntel(Z).blPrice > 0 Then
												bBold = True
												Exit For
											End If
										End If
									Next Z
									lstData.AddItem(GetCacheObjectValue(oRel.lThisPlayer, ObjectType.ePlayer) & " Scores", bBold)
									If lstData.ItemBold(lstData.NewIndex) = True Then lstData.ItemCustomColor(lstData.NewIndex) = clrSpecial
									lstData.ItemData(lstData.NewIndex) = oRel.lThisPlayer : lstData.ItemData2(lstData.NewIndex) = ObjectType.ePlayerIntel : lstData.ItemData3(lstData.NewIndex) = ObjectType.ePlayer
									lstData.ApplyIconOffset(lstData.NewIndex) = True
								End If
							End With
						End If
					End If
				Next X
			End If

			'Now, component design intel
			lstData.AddItem("COMPONENT DESIGN INTEL", True)
			lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
			lstData.ApplyIconOffset(lstData.NewIndex) = False
			For X As Int32 = 0 To glPlayerTechKnowledgeUB
				If glPlayerTechKnowledgeIdx(X) <> -1 Then
					Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(X)
					If oPTK Is Nothing = False AndAlso oPTK.yKnowledgeType >= PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
						If oPTK.oTech Is Nothing = False Then
							Dim bBold As Boolean = False
							For Z As Int32 = 0 To mlIntelUB
								If muIntel(Z).lObjectID = oPTK.oTech.ObjectID AndAlso muIntel(Z).iObjTypeID = ObjectType.ePlayerTechKnowledge AndAlso muIntel(Z).iExtTypeID = oPTK.oTech.ObjTypeID Then
									If muIntel(Z).blPrice > 0 Then
										bBold = True
										Exit For
									End If
								End If
							Next Z

							lstData.AddItem(oPTK.oTech.GetComponentName, bBold)
							If lstData.ItemBold(lstData.NewIndex) = True Then lstData.ItemCustomColor(lstData.NewIndex) = clrSpecial
							lstData.ItemData(lstData.NewIndex) = oPTK.oTech.ObjectID : lstData.ItemData2(lstData.NewIndex) = ObjectType.ePlayerTechKnowledge : lstData.ItemData3(lstData.NewIndex) = oPTK.oTech.ObjTypeID
							lstData.ApplyIconOffset(lstData.NewIndex) = True
						End If
					End If
				End If
			Next X

			'Now, component design intel
			lstData.AddItem("COLONY LOCATION INTEL", True)
			lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
			lstData.ApplyIconOffset(lstData.NewIndex) = False
			For X As Int32 = 0 To glItemIntelUB
				If glItemIntelIdx(X) <> -1 Then
					Dim oItemIntel As PlayerItemIntel = goItemIntel(X)
					If oItemIntel Is Nothing = False AndAlso oItemIntel.yIntelType >= PlayerItemIntel.PlayerItemIntelType.eLocation Then
						If oItemIntel.iItemTypeID = ObjectType.eColony Then

							Dim bBold As Boolean = False
							For Z As Int32 = 0 To mlIntelUB
								If muIntel(Z).lObjectID = oItemIntel.lItemID AndAlso muIntel(Z).iObjTypeID = ObjectType.ePlayerItemIntel AndAlso muIntel(Z).iExtTypeID = oItemIntel.iItemTypeID Then
									If muIntel(Z).blPrice > 0 Then
										bBold = True
										Exit For
									End If
								End If
							Next Z

							lstData.AddItem(GetCacheObjectValue(oItemIntel.lItemID, oItemIntel.iItemTypeID) & " (" & GetCacheObjectValue(oItemIntel.lOtherPlayerID, ObjectType.ePlayer) & ")", bBold)
							If lstData.ItemBold(lstData.NewIndex) = True Then lstData.ItemCustomColor(lstData.NewIndex) = clrSpecial
							lstData.ItemData(lstData.NewIndex) = oItemIntel.lItemID : lstData.ItemData2(lstData.NewIndex) = ObjectType.ePlayerItemIntel : lstData.ItemData3(lstData.NewIndex) = oItemIntel.iItemTypeID
							lstData.ApplyIconOffset(lstData.NewIndex) = True
						End If
					End If
				End If
			Next X

            'Units
            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To mlUnitUB
                Dim lIdx As Int32 = -1
                'TODO: Switch to alpha numeric sort, but only once the Transfer window and contents window do the same: for uniformnity
                'sName = GetAlphaNumericSortable(GetCacheObjectValue(muUnits(X).lObjectID, muUnits(X).iObjTypeID))
                sName = GetCacheObjectValue(muUnits(X).lObjectID, muUnits(X).iObjTypeID)
                For Y As Int32 = 0 To lSortedUB
                    'Dim sOtherName As String = GetAlphaNumericSortable(GetCacheObjectValue(muUnits(lSorted(Y)).lObjectID, muUnits(lSorted(Y)).iObjTypeID))
                    Dim sOtherName As String = GetCacheObjectValue(muUnits(lSorted(Y)).lObjectID, muUnits(lSorted(Y)).iObjTypeID)
                    If sOtherName > sName OrElse (sOtherName = sName AndAlso muUnits(X).lObjectID < muUnits(Y).lObjectID) Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                ReDim Preserve lSorted(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                End If
            Next X
			lstData.AddItem("UNITS", True)
			lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
			lstData.ApplyIconOffset(lstData.NewIndex) = False
			For X As Int32 = 0 To mlUnitUB
                With muUnits(lSorted(X))
                    sName = GetCacheObjectValue(.lObjectID, .iObjTypeID)
                    lstData.AddItem(sName, (.lForSaleQty <> 0 AndAlso .blPrice <> 0))
                    If lstData.ItemBold(lstData.NewIndex) = True Then lstData.ItemCustomColor(lstData.NewIndex) = clrSpecial
                    lstData.ItemData(lstData.NewIndex) = .lObjectID : lstData.ItemData2(lstData.NewIndex) = .iObjTypeID
                    lstData.ApplyIconOffset(lstData.NewIndex) = True
                End With
			Next X
		End If
	End Sub

    ''' <summary>
    ''' Updates the listbox smartly by check whether an entry needs updating
    ''' </summary>
    ''' <param name="lstData"></param>
    ''' <param name="yType"> Pass 0 for Sell Screen, 1 for Direct Trades </param>
    ''' <remarks></remarks>
    Public Sub SmartPopulateList(ByRef lstData As UIListBox, ByVal yType As Byte)
        Try
            For X As Int32 = 0 To lstData.ListCount - 1
                Dim lID As Int32 = lstData.ItemData(X)
                Dim iTypeID As Int16 = Math.Abs(CShort(lstData.ItemData2(X)))
                If lID <> -1 Then
                    Select Case iTypeID
                        Case ObjectType.eColonists, ObjectType.eEnlisted, ObjectType.eOfficers, ObjectType.eMineralCache
                            'Do nothing
                        Case ObjectType.eCredits
                            Dim sName As String = "Credits (" & goCurrentPlayer.blCredits.ToString("#,##0") & ")"
							If lstData.List(X) <> sName Then lstData.List(X) = sName
						Case ObjectType.ePlayerIntel
							Dim sName As String = GetCacheObjectValue(lID, ObjectType.ePlayer) & " Scores"
							If lstData.List(X) <> sName Then lstData.List(X) = sName
						Case ObjectType.ePlayerItemIntel
							Dim iTempTypeID As Int16 = CShort(lstData.ItemData3(X))
							For Z As Int32 = 0 To glItemIntelUB
								If glItemIntelIdx(Z) <> -1 Then
									Dim oPII As PlayerItemIntel = goItemIntel(Z)
									If oPII Is Nothing = False AndAlso oPII.lItemID = lID AndAlso oPII.iItemTypeID = iTempTypeID Then
										Dim sName As String = GetCacheObjectValue(oPII.lItemID, oPII.iItemTypeID) & " (" & GetCacheObjectValue(oPII.lOtherPlayerID, ObjectType.ePlayer) & ")"
										If lstData.List(X) <> sName Then lstData.List(X) = sName
										Exit For
									End If
								End If
							Next Z
						Case ObjectType.ePlayerTechKnowledge
							Dim iTempTypeID As Int16 = CShort(lstData.ItemData3(X))
							For Z As Int32 = 0 To glPlayerTechKnowledgeUB
								If glPlayerTechKnowledgeIdx(Z) <> -1 Then
									Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(Z)
									If oPTK Is Nothing = False AndAlso oPTK.oTech Is Nothing = False Then
										If oPTK.oTech.ObjectID = lID AndAlso oPTK.oTech.ObjTypeID = iTempTypeID Then
											Dim sName As String = oPTK.oTech.GetComponentName()
											If lstData.List(X) <> sName Then lstData.List(X) = sName
											Exit For
										End If
									End If
								End If
							Next Z
						Case ObjectType.eAgent
							For Y As Int32 = 0 To goCurrentPlayer.AgentUB
								If goCurrentPlayer.AgentIdx(Y) = lID Then
									If lstData.List(X) <> goCurrentPlayer.Agents(Y).sAgentName Then lstData.List(X) = goCurrentPlayer.Agents(Y).sAgentName
									Exit For
								End If
							Next Y

						Case Else
							Dim lTempID As Int32 = lID
							For Y As Int32 = 0 To mlComponentUB
								If muComponents(Y).lObjectID = lID AndAlso muComponents(Y).iComponentTypeID = iTypeID Then
									lTempID = muComponents(Y).lComponentID
									Exit For
								End If
							Next Y
							Dim lParen As Int32 = lstData.List(X).LastIndexOf("(")
                            Dim sName As String = "  " & GetCacheObjectValue(lTempID, iTypeID)
							If lParen <> -1 Then
								sName &= lstData.List(X).Substring(lParen)
							End If
							If lstData.List(X) <> sName Then lstData.List(X) = sName
					End Select
                End If
            Next X
        Catch
            'do nothing
        End Try
    End Sub

	Public Sub GetItemsDetails(ByVal lID As Int32, ByVal iTypeID As Int16, ByRef blQty As Int64, ByRef blCurrPrice As Int64, ByRef blCurrQty As Int64, ByVal iExtTypeID As Int16)
		Select Case iTypeID
			Case ObjectType.eMineralCache, ObjectType.eMineral
				For X As Int32 = 0 To mlMineralUB
					If muMinerals(X).lObjectID = lID AndAlso muMinerals(X).iObjTypeID = iTypeID Then
						blQty = muMinerals(X).lQty
						blCurrPrice = muMinerals(X).blPrice
						blCurrQty = muMinerals(X).lForSaleQty
						Exit For
					End If
				Next X
			Case ObjectType.eUnit
				For X As Int32 = 0 To mlUnitUB
					If muUnits(X).lObjectID = lID AndAlso muUnits(X).iObjTypeID = iTypeID Then
						blQty = 1
						blCurrPrice = muUnits(X).blPrice
						blCurrQty = muUnits(X).lForSaleQty
						Exit For
					End If
				Next X
			Case ObjectType.eEnlisted
				blQty = Me.lEnlisted
				blCurrPrice = 0
				blCurrQty = 0
			Case ObjectType.eOfficers
				blQty = Me.lOfficers
				blCurrPrice = 0
				blCurrQty = 0
			Case ObjectType.eColonists
				blQty = Me.lColonists
				blCurrPrice = 0
				blCurrQty = 0
			Case ObjectType.ePlayerIntel, ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
				For X As Int32 = 0 To mlIntelUB
                    If muIntel(X).lObjectID = lID AndAlso muIntel(X).iObjTypeID = iTypeID Then ' AndAlso muIntel(X).iExtTypeID = ObjectType.ePlayer Then
                        blQty = 1
                        blCurrPrice = muIntel(X).blPrice
                        blCurrQty = 1
                        Exit For
                    End If
				Next X
			Case Else
				For X As Int32 = 0 To mlComponentUB
					If muComponents(X).lObjectID = lID AndAlso -muComponents(X).iComponentTypeID = iTypeID Then
						blQty = muComponents(X).lQuantity
						blCurrPrice = muComponents(X).blPrice
						blCurrQty = muComponents(X).lForSaleQty
						Exit For
					End If
				Next X
		End Select
	End Sub

	Public Sub ClearPrice(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal iExtTypeID As Int16)
		Select Case iTypeID
			Case ObjectType.eMineralCache, ObjectType.eMineral
				For X As Int32 = 0 To mlMineralUB
					If muMinerals(X).lObjectID = lID AndAlso muMinerals(X).iObjTypeID = iTypeID Then
						muMinerals(X).blPrice = 0
						Exit For
					End If
				Next X
			Case ObjectType.eUnit
				For X As Int32 = 0 To mlUnitUB
					If muUnits(X).lObjectID = lID AndAlso muUnits(X).iObjTypeID = iTypeID Then
						muUnits(X).blPrice = 0
						Exit For
					End If
				Next X
			Case ObjectType.ePlayerIntel, ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
				For X As Int32 = 0 To mlIntelUB
					If muIntel(X).lObjectID = lID AndAlso muIntel(X).iObjTypeID = iTypeID AndAlso muIntel(X).iExtTypeID = iExtTypeID Then
						muIntel(X).blPrice = 0
						Exit For
					End If
				Next X
			Case Else
				For X As Int32 = 0 To mlComponentUB
					If muComponents(X).lObjectID = lID AndAlso -muComponents(X).iComponentTypeID = iTypeID Then
						muComponents(X).blPrice = 0
						Exit For
					End If
				Next X
		End Select
	End Sub

    Public Function GetMineralIDFromCache(ByVal lCacheID As Int32) As Int32
        For X As Int32 = 0 To mlMineralUB
            If muMinerals(X).lObjectID = lCacheID Then Return muMinerals(X).lMineralID
        Next X
        Return -1
    End Function

    Public Sub FillDeliverList(ByVal iTypeID As Int16, ByVal blMinQty As Int64, ByRef lstData As UIListBox)
        lstData.Clear()
        Select Case iTypeID
            Case ObjectType.eMineral, ObjectType.eMineralCache
                For X As Int32 = 0 To mlMineralUB
                    If muMinerals(X).lQty >= blMinQty Then
                        For Y As Int32 = 0 To glMineralUB
                            If glMineralIdx(Y) = muMinerals(X).lMineralID Then
								lstData.AddItem(goMinerals(Y).MineralName & " (" & muMinerals(X).lQty & ")", False)
                                lstData.ItemData(lstData.NewIndex) = muMinerals(X).lObjectID
                                lstData.ItemData2(lstData.NewIndex) = muMinerals(X).iObjTypeID

								Exit For
                            End If
                        Next Y
                    End If
                Next X
            Case Else
                If iTypeID < 0 Then iTypeID = Math.Abs(iTypeID)
                For X As Int32 = 0 To mlComponentUB
                    If muComponents(X).iComponentTypeID = iTypeID AndAlso muComponents(X).lQuantity >= blMinQty Then
                        Dim oTech As Base_Tech = goCurrentPlayer.GetTech(muComponents(X).lComponentID, muComponents(X).iComponentTypeID)
                        If oTech Is Nothing = False Then
                            lstData.AddItem(oTech.GetComponentName)
                        Else : lstData.AddItem(GetCacheObjectValue(muComponents(X).lComponentID, muComponents(X).iComponentTypeID))
                        End If
                        lstData.ItemData(lstData.NewIndex) = muComponents(X).lObjectID
                        lstData.ItemData2(lstData.NewIndex) = muComponents(X).iObjTypeID
                    End If
                Next X
        End Select
    End Sub

    Public Sub PopulateComponentCacheProperties(ByVal lCacheID As Int32, ByRef lCompID As Int32, ByRef iCompTypeID As Int16, ByRef lCompOwnerID As Int32)
        For X As Int32 = 0 To mlComponentUB
            If muComponents(X).lObjectID = lCacheID Then
                lCompID = muComponents(X).lComponentID
                iCompTypeID = muComponents(X).iComponentTypeID
                lCompOwnerID = muComponents(X).lComponentOwnerID
                Exit For
            End If
        Next X
    End Sub

    Public Function GetUnitItemChassisType(ByVal lID As Int32, ByVal lTypeID As Int32) As Byte
        For X As Int32 = 0 To mlUnitUB
            If muUnits(X).lObjectID = lID AndAlso muUnits(X).iObjTypeID = lTypeID Then Return muUnits(X).yChassisType
        Next X
        Return ChassisType.eGroundBased
	End Function

	Public Shared Sub HandleGetGuildAssets(ByRef yData() As Byte)
		Dim lPos As Int32 = 2		'msgcode

		Dim oTPC As New TradePostContents()

		With oTPC
			.lTradePostID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If .lTradePostID = -1 Then
				With goCurrentPlayer.oGuild
					.lGuildHallLocID = -1
					.lGuildHallID = -1
					.lGuildHallLocX = -1
					.lGuildHallLocZ = -1
					.iGuildHallLocTypeID = -1
					Return
				End With
			End If

			Dim lItems As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			For X As Int32 = 0 To lItems - 1
				Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

				Select Case iTypeID
					Case ObjectType.eComponentCache, Is < 0
						Dim lCompID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim iCompTypeID As Int16 = Math.Abs(iTypeID)
						Dim lCompOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						iTypeID = ObjectType.eComponentCache
						Dim lForSale As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim blPrice As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
						oTPC.AddComponentCache(lID, iTypeID, lCompID, iCompTypeID, lCompOwnerID, lQty, lForSale, blPrice)
					Case ObjectType.eMineralCache
						Dim lMineralID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lForSale As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim blPrice As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
						oTPC.AddMineralCache(lID, iTypeID, lMineralID, lQty, lForSale, blPrice)
					Case ObjectType.eUnit
						Dim lForSale As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim blPrice As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
						Dim yChassisType As Byte = yData(lPos) : lPos += 1
						oTPC.AddUnit(lID, iTypeID, lForSale, blPrice, yChassisType)
					Case ObjectType.ePlayerIntel, ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
						Dim iExtTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
						Dim blPrice As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
						oTPC.AddIntel(lID, iTypeID, iExtTypeID, blPrice)
					Case Else
						'TODO: what else?
				End Select
			Next X
		End With

		'Now, refresh the list....
		oTPC.lLastUpdate = glCurrentCycle
		Dim lIdx As Int32 = -1
		Dim lFirstIdx As Int32 = -1
		For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
			If TradePostContents.lTradePostContentsIdx(X) = oTPC.lTradePostID Then
				lIdx = X
				Exit For
			ElseIf lFirstIdx = -1 AndAlso TradePostContents.lTradePostContentsIdx(X) = -1 Then
				lFirstIdx = X
			End If
		Next X
		If lIdx = -1 Then
			If lFirstIdx = -1 Then
				lIdx = TradePostContents.lTradePostContentsUB + 1
				ReDim Preserve TradePostContents.oTradePostContents(lIdx)
				ReDim Preserve TradePostContents.lTradePostContentsIdx(lIdx)
				TradePostContents.lTradePostContentsIdx(lIdx) = -1
				TradePostContents.lTradePostContentsUB += 1
			Else : lIdx = lFirstIdx
			End If
		End If
		TradePostContents.oTradePostContents(lIdx) = oTPC
		TradePostContents.lTradePostContentsIdx(lIdx) = oTPC.lTradePostID
	End Sub

    Private Sub SortItems()
        Try

            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            Dim sSortedVal() As String = Nothing

            For X As Int32 = 0 To mlComponentUB
                Dim lIdx As Int32 = -1
                
                Dim sName As String = GetCacheObjectValue(muComponents(X).lComponentID, muComponents(X).iComponentTypeID)

                For Y As Int32 = 0 To lSortedUB
                    If sSortedVal(Y).ToUpper > sName.ToUpper Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                ReDim Preserve lSorted(lSortedUB)
                ReDim Preserve sSortedVal(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                    sSortedVal(lSortedUB) = sName
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                        sSortedVal(Y) = sSortedVal(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                    sSortedVal(lIdx) = sName
                End If
            Next X

            Dim uNewList(mlComponentUB) As CompCache
            For X As Int32 = 0 To mlComponentUB
                uNewList(X) = muComponents(lSorted(X))
            Next X
            muComponents = uNewList

            lSorted = Nothing
            lSortedUB = -1
            sSortedVal = Nothing
            For X As Int32 = 0 To mlMineralUB
                Dim lIdx As Int32 = -1

                Dim sName As String = "" '= GetCacheObjectValue(muMinerals(X).lMineralID, ObjectType.eMineral)
                For Z As Int32 = 0 To glMineralUB
                    If glMineralIdx(Z) = muMinerals(X).lMineralID Then
                        sName = goMinerals(Z).MineralName
                        Exit For
                    End If
                Next Z
                If sName = "" Then sName = GetCacheObjectValue(muMinerals(X).lMineralID, ObjectType.eMineral)

                For Y As Int32 = 0 To lSortedUB
                    If sSortedVal(Y).ToUpper > sName.ToUpper Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                ReDim Preserve lSorted(lSortedUB)
                ReDim Preserve sSortedVal(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                    sSortedVal(lSortedUB) = sName
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                        sSortedVal(Y) = sSortedVal(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                    sSortedVal(lIdx) = sName
                End If
            Next X

            Dim uMinList(mlMineralUB) As MinCache
            For X As Int32 = 0 To mlMineralUB
                uMinList(X) = muMinerals(lSorted(X))
            Next X
            muMinerals = uMinList
        Catch
            'do nothing
        End Try
    End Sub

	Public Shared Sub HandleTradepostTradeables(ByRef yData() As Byte)
		Dim lPos As Int32 = 2		'msgcode

		Dim oTPC As New TradePostContents()

		With oTPC
			.lTradePostID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			.ySellSlotsUsed = yData(lPos) : lPos += 1
			.yBuySlotsUsed = yData(lPos) : lPos += 1

			.lColonists = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lEnlisted = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lOfficers = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			While lPos + 6 < yData.Length
				Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

				Select Case iTypeID
					Case ObjectType.eComponentCache, Is < 0
						Dim lCompID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim iCompTypeID As Int16 = Math.Abs(iTypeID)
						Dim lCompOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						iTypeID = ObjectType.eComponentCache
						Dim lForSale As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim blPrice As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
						oTPC.AddComponentCache(lID, iTypeID, lCompID, iCompTypeID, lCompOwnerID, lQty, lForSale, blPrice)
					Case ObjectType.eMineralCache
						Dim lMineralID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lForSale As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim blPrice As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
						oTPC.AddMineralCache(lID, iTypeID, lMineralID, lQty, lForSale, blPrice)
					Case ObjectType.eUnit
						Dim lForSale As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim blPrice As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
						Dim yChassisType As Byte = yData(lPos) : lPos += 1
						oTPC.AddUnit(lID, iTypeID, lForSale, blPrice, yChassisType)
					Case ObjectType.ePlayerIntel, ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
						Dim iExtTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
						Dim blPrice As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
						oTPC.AddIntel(lID, iTypeID, iExtTypeID, blPrice)
					Case Else
						'TODO: what else?
				End Select
			End While
        End With

        oTPC.SortItems()

		'Now, refresh the list....
		oTPC.lLastUpdate = glCurrentCycle
		Dim lIdx As Int32 = -1
		Dim lFirstIdx As Int32 = -1
		For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
			If TradePostContents.lTradePostContentsIdx(X) = oTPC.lTradePostID Then
				lIdx = X
				Exit For
			ElseIf lFirstIdx = -1 AndAlso TradePostContents.lTradePostContentsIdx(X) = -1 Then
				lFirstIdx = X
			End If
		Next X
		If lIdx = -1 Then
			If lFirstIdx = -1 Then
				lIdx = TradePostContents.lTradePostContentsUB + 1
				ReDim Preserve TradePostContents.oTradePostContents(lIdx)
				ReDim Preserve TradePostContents.lTradePostContentsIdx(lIdx)
				TradePostContents.lTradePostContentsIdx(lIdx) = -1
				TradePostContents.lTradePostContentsUB += 1
			Else : lIdx = lFirstIdx
			End If
		End If
		TradePostContents.oTradePostContents(lIdx) = oTPC
		TradePostContents.lTradePostContentsIdx(lIdx) = oTPC.lTradePostID
    End Sub
    Public Sub SetItemPrice(ByVal blQty As Int64, ByVal blPrice As Int64, ByVal lItemID As Int32, ByVal iTypeID As Int16)
        Select Case iTypeID
            Case ObjectType.eComponentCache, Is < 0
                For X As Int32 = 0 To mlComponentUB
                    If muComponents(X).lComponentID = lItemID AndAlso muComponents(X).iComponentTypeID = Math.Abs(iTypeID) Then
                        muComponents(X).blPrice = blPrice
                        muComponents(X).lForSaleQty = CInt(blQty)
                        Exit For
                    End If
                Next X
            Case ObjectType.eMineralCache
                For X As Int32 = 0 To mlMineralUB
                    If muMinerals(X).lObjectID = lItemID Then
                        muMinerals(X).lForSaleQty = CInt(blQty)
                        muMinerals(X).blPrice = blPrice
                        Exit For
                    End If
                Next X
            Case ObjectType.eUnit
                For X As Int32 = 0 To mlUnitUB
                    If muUnits(X).lObjectID = lItemID Then
                        muUnits(X).blPrice = blPrice
                        muUnits(X).lForSaleQty = CInt(blQty)
                        Exit For
                    End If
                Next X
            Case ObjectType.ePlayerIntel, ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
                For X As Int32 = 0 To mlIntelUB
                    If muIntel(X).lObjectID = lItemID AndAlso muIntel(X).iObjTypeID = iTypeID Then
                        muIntel(X).blPrice = blPrice
                        Exit For
                    End If
                Next X
        End Select
    End Sub
End Class