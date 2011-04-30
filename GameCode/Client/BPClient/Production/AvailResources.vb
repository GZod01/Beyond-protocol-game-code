Option Strict On

Public Class AvailResources
    Private Structure CacheItem
        Public ObjectID As Int32
        Public ObjTypeID As Int16
        Public Quantity As Int32
        Public sName As String
    End Structure

    Public lColonyParentID As Int32         'the colony's parent's id
    Public iColonyParentTypeID As Int16     'the colony's parent's type id 
    Public lColonyID As Int32               'the colony ID

    Public lColonists As Int32 = 0
    Public lEnlisted As Int32 = 0
    Public lOfficers As Int32 = 0

    Public blTotalCargoCap As Int64 = 0
    Public blAvailCargoCap As Int64 = 0

    Private muItems() As CacheItem
    Private mlItemUB As Int32 = -1

    Private mlLastUpdate As Int32 = -1

    Public Sub ClearResources()
        mlItemUB = -1
        Erase muItems
        mlLastUpdate = -1
    End Sub

    Public ReadOnly Property LastUpdate() As Int32
        Get
			Dim sTemp As String
			Dim lCurUB As Int32 = -1
			If muItems Is Nothing = False Then lCurUB = Math.Min(muItems.GetUpperBound(0), mlItemUB)
			Try
				For X As Int32 = 0 To lCurUB
					'check our name values to see if a name changed
					If muItems(X).sName = "Unknown" Then
						sTemp = GlobalVars.GetCacheObjectValue(muItems(X).ObjectID, muItems(X).ObjTypeID)
						If sTemp <> muItems(X).sName Then
							muItems(X).sName = sTemp
							mlLastUpdate = glCurrentCycle
						End If
					End If
				Next X
			Catch
			End Try
			Return mlLastUpdate
		End Get
    End Property

    Public Sub FillFromMsg(ByRef yData() As Byte)
        lColonyParentID = System.BitConverter.ToInt32(yData, 2)
        iColonyParentTypeID = System.BitConverter.ToInt16(yData, 6)

        Dim lPos As Int32 = 8

        lColonyID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        lColonists = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lEnlisted = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lOfficers = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        blTotalCargoCap = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
        blAvailCargoCap = System.BitConverter.ToInt64(yData, lPos) : lPos += 8

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Now, let's go through it...
        mlItemUB = lCnt - 1
        ReDim muItems(mlItemUB)
        For X As Int32 = 0 To lCnt - 1
            With muItems(X)
                'ID
                .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                'TypeID
                .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                'Qty
                .Quantity = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                'set hte name to unknown so the update will pick it up
                .sName = "Unknown"

                If .ObjTypeID = ObjectType.eMineralCache OrElse .ObjTypeID = ObjectType.eMineral Then
                    'Find it locally first...
                    For Y As Int32 = 0 To glMineralUB
                        If glMineralIdx(Y) = .ObjectID Then
                            .sName = goMinerals(Y).MineralName
                            Exit For
                        End If
                    Next Y
                End If

            End With
        Next X

        Me.mlLastUpdate = glCurrentCycle
    End Sub

    Public Function GetTextBoxText() As String
        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder

        Dim lSpcTotal As Int32 = 0
        Dim lSpcAvailable As Int32 = 0
        Dim sTotalCargoCap As String = blTotalCargoCap.ToString("#,##0")
        Dim sAvailableCargoCap As String = blAvailCargoCap.ToString("#,##0")

        If sTotalCargoCap.Length > sAvailableCargoCap.Length Then
            lSpcAvailable = sTotalCargoCap.Length - sAvailableCargoCap.Length
        ElseIf sAvailableCargoCap.Length > sTotalCargoCap.Length Then
            lSpcTotal = sAvailableCargoCap.Length - sTotalCargoCap.Length
        End If

        oSB.AppendLine("Total Capacity : " & Space(lSpcTotal) & blTotalCargoCap.ToString("#,##0"))
        If blAvailCargoCap < 0 Then
            If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
                blAvailCargoCap = 0
            End If
        End If

        oSB.AppendLine("     Available : " & Space(lSpcAvailable) & blAvailCargoCap.ToString("#,##0"))
        oSB.AppendLine()

        'Skip Mineral caches in the first pass
        Dim bFirst As Boolean = True
        Dim lCompSorted() As Int32 = Nothing
        Dim lCompSortedUB As Int32 = -1

        For X As Int32 = 0 To mlItemUB
            If muItems(X).ObjTypeID <> ObjectType.eMineralCache AndAlso muItems(X).ObjTypeID <> ObjectType.eMineral Then
                If bFirst = True Then
                    oSB.AppendLine("COMPONENTS: ")
                    bFirst = False
                End If

                Dim lIdx As Int32 = -1
                For Y As Int32 = 0 To lCompSortedUB
                    If muItems(lCompSorted(Y)).ObjTypeID > muItems(X).ObjTypeID OrElse (muItems(lCompSorted(Y)).ObjTypeID = muItems(X).ObjTypeID AndAlso muItems(lCompSorted(Y)).sName > muItems(X).sName) Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lCompSortedUB += 1
                ReDim Preserve lCompSorted(lCompSortedUB)
                If lIdx = -1 Then
                    lCompSorted(lCompSortedUB) = X
                Else
                    For Y As Int32 = lCompSortedUB To lIdx + 1 Step -1
                        lCompSorted(Y) = lCompSorted(Y - 1)
                    Next
                    lCompSorted(lIdx) = X
                End If
                'oSB.AppendLine("  " & muItems(X).sName & ": " & muItems(X).Quantity)
            End If
        Next X

        'oSB.AppendLine("  MMMMMMMMMMMMMMMMMMMM #######")
        Dim bAlloys As Boolean = False
        Dim bArmor As Boolean = False
        Dim bEngine As Boolean = False
        Dim bRadar As Boolean = False
        Dim bShield As Boolean = False
        Dim bWeapon As Boolean = False
        
        Dim lSpcComponent As Int32 = 0
        Dim sCargoQuantity As String = ""
        For X As Int32 = 0 To lCompSortedUB
            sCargoQuantity = muItems(lCompSorted(X)).Quantity.ToString("#,##0")
            'Debug.Print(CType(muItems(lCompSorted(X)).ObjTypeID, ObjectType).ToString & " " & muItems(lCompSorted(X)).sName)
            Select Case muItems(lCompSorted(X)).ObjTypeID
                Case ObjectType.eAlloyTech
                    If bAlloys = False Then
                        'Debug.Print(" ... Alloys ...")
                        oSB.AppendLine(" Alloys")
                        bAlloys = True
                    End If
                Case ObjectType.eArmorTech
                    If bArmor = False Then
                        'Debug.Print(" ... Armor ...")
                        oSB.AppendLine(" Armor")
                        bArmor = True
                    End If
                Case ObjectType.eEngineTech
                    If bEngine = False Then
                        'Debug.Print(" ... Engines ...")
                        oSB.AppendLine(" Engines")
                        bEngine = True
                    End If
                Case ObjectType.eRadarTech
                    If bRadar = False Then
                        'Debug.Print(" ... Radars ...")
                        oSB.AppendLine(" Radars")
                        bRadar = True
                    End If
                Case ObjectType.eShieldTech
                    If bShield = False Then
                        'Debug.Print(" ... Shields ...")
                        oSB.AppendLine(" Shields")
                        bShield = True
                    End If
                Case ObjectType.eWeaponTech
                    If bWeapon = False Then
                        'Debug.Print(" ... Weapons ...")
                        oSB.AppendLine(" Weapons")
                        bWeapon = True
                    End If
            End Select
            If muItems(lCompSorted(X)).Quantity < 999999 Then
                lSpcComponent = 30 - 2 - muItems(lCompSorted(X)).sName.Length - sCargoQuantity.Length
                oSB.AppendLine("  " & muItems(lCompSorted(X)).sName & Space(lSpcComponent) & sCargoQuantity)
            Else
                If 2 + muItems(lCompSorted(X)).sName.Length + sCargoQuantity.Length < 30 Then
                    lSpcComponent = 30 - 2 - muItems(lCompSorted(X)).sName.Length - sCargoQuantity.Length
                    oSB.AppendLine("  " & muItems(lCompSorted(X)).sName & Space(lSpcComponent) & sCargoQuantity)
                Else
                    lSpcComponent = 26 - sCargoQuantity.Length
                    oSB.AppendLine("  " & muItems(lCompSorted(X)).sName)
                    oSB.AppendLine("    " & Space(lSpcComponent) & sCargoQuantity)
                End If
            End If
        Next X


        'Go back through and do only mineral caches
        Dim lSorted() As Int32 = GetSortedMineralIdxArray(True)

        bFirst = True

        If lSorted Is Nothing = False Then
            For lMinIdx As Int32 = 0 To lSorted.GetUpperBound(0)
                Dim lMineralID As Int32 = glMineralIdx(lSorted(lMinIdx))

                For X As Int32 = 0 To mlItemUB
                    If (muItems(X).ObjTypeID = ObjectType.eMineralCache OrElse muItems(X).ObjTypeID = ObjectType.eMineral) AndAlso muItems(X).ObjectID = lMineralID Then
                        If bFirst = True Then
                            oSB.AppendLine("MATERIALS: ")
                            bFirst = False
                        End If
                        sCargoQuantity = muItems(X).Quantity.ToString("#,##0")
                        If muItems(X).Quantity < 999999 Then
                            lSpcComponent = 30 - 2 - muItems(X).sName.Length - sCargoQuantity.Length
                            oSB.AppendLine("  " & muItems(X).sName & Space(lSpcComponent) & sCargoQuantity)
                        Else
                            If 2 + muItems(X).sName.Length + sCargoQuantity.Length < 30 Then
                                lSpcComponent = 30 - 2 - muItems(X).sName.Length - sCargoQuantity.Length
                                oSB.AppendLine("  " & muItems(X).sName & Space(lSpcComponent) & sCargoQuantity)
                            Else
                                lSpcComponent = 26 - sCargoQuantity.Length
                                oSB.AppendLine("  " & muItems(X).sName)
                                oSB.AppendLine("    " & Space(lSpcComponent) & sCargoQuantity)
                            End If
                        End If
                        Exit For
                    End If
                Next X

            Next lMinIdx
        End If



        Return oSB.ToString
    End Function

	Public Sub FillListBox(ByRef lstData As UIListBox) ', ByVal bSpaceTradepost As Boolean)

		If lstData.oIconTexture Is Nothing OrElse lstData.oIconTexture.Disposed = True Then
			lstData.oIconTexture = goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
		End If
		lstData.RenderIcons = True

		lstData.Clear()

		'If bSpaceTradepost = False Then lstData.AddItem("COLONY", True) Else lstData.AddItem("TRADEPOST", True)
		lstData.AddItem("COLONY", True)
		lstData.ItemData(lstData.NewIndex) = lColonyID
		lstData.ItemData2(lstData.NewIndex) = ObjectType.eColony
		lstData.ApplyIconOffset(lstData.NewIndex) = False

		'Now, the personnel
		'If bSpaceTradepost = False Then
        'AddListBoxItem(lstData, "1234567890123456789012345678", 0, 0)
        If lColonists > 0 Then
            Dim sText As String = "Colonists : " & lColonists.ToString("#,##0")
            AddListBoxItem(lstData, sText, lColonists, ObjectType.eColonists)
        End If
		If lEnlisted > 0 Then
            Dim sText As String = "Enlisted  : " & lEnlisted.ToString("#,##0")
			AddListBoxItem(lstData, sText, lEnlisted, ObjectType.eEnlisted)
		End If
		If lOfficers > 0 Then
            Dim sText As String = "Officers  : " & lOfficers.ToString("#,##0")
            AddListBoxItem(lstData, sText, lOfficers, ObjectType.eOfficers)
		End If
		'End If

        'Do a sort
        Try
            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To mlItemUB
                Dim lIdx As Int32 = -1

                For Y As Int32 = 0 To lSortedUB
                    If muItems(lSorted(Y)).sName > muItems(X).sName Then
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
            If mlItemUB <> lSortedUB Then
                For lType As Int32 = 0 To 1
                    If lType = 0 Then
                        lstData.AddItem("MINERALS", True)
                        lstData.ItemLocked(lstData.NewIndex) = True
                        lstData.ItemData(lstData.NewIndex) = -1
                        lstData.ItemData2(lstData.NewIndex) = -1
                        lstData.ApplyIconOffset(lstData.NewIndex) = False
                    Else
                        lstData.AddItem("COMPONENTS", True)
                        lstData.ItemLocked(lstData.NewIndex) = True
                        lstData.ItemData(lstData.NewIndex) = -1
                        lstData.ItemData2(lstData.NewIndex) = -1
                        lstData.ApplyIconOffset(lstData.NewIndex) = False
                    End If
                    'AddListBoxItem(lstData, "MMMMMMMMMMMMMMMMMMMM #######", 0, 0)
                    For X As Int32 = 0 To mlItemUB
                        With muItems(X)
                            Dim sText As String = .sName & ": " & .Quantity.ToString("#,##0")

                            Dim bFound As Boolean = False
                            For Y As Int32 = 0 To lstData.ListCount - 1
                                If lstData.ItemData(Y) = .ObjectID AndAlso lstData.ItemData2(Y) = .ObjTypeID Then
                                    If lstData.List(Y) <> sText Then lstData.List(Y) = sText
                                    bFound = True
                                    Exit For
                                End If
                            Next Y
                            If bFound = False Then
                                If (.ObjTypeID = ObjectType.eMineralCache OrElse .ObjTypeID = ObjectType.eMineral) Then
                                    If lType = 0 Then AddListBoxItem(lstData, sText, .ObjectID, .ObjTypeID)
                                Else
                                    If lType = 1 Then AddListBoxItem(lstData, sText, .ObjectID, .ObjTypeID)
                                End If
                            End If
                        End With
                    Next X
                Next lType
            Else
                For lType As Int32 = 0 To 1
                    If lType = 0 Then
                        lstData.AddItem("MINERALS", True)
                        lstData.ItemLocked(lstData.NewIndex) = True
                        lstData.ItemData(lstData.NewIndex) = -1
                        lstData.ItemData2(lstData.NewIndex) = -1
                    Else
                        lstData.AddItem("COMPONENTS", True)
                        lstData.ItemLocked(lstData.NewIndex) = True
                        lstData.ItemData(lstData.NewIndex) = -1
                        lstData.ItemData2(lstData.NewIndex) = -1
                    End If
                    'AddListBoxItem(lstData, "MMMMMMMMMMMMMMMMMMMM #######", 0, 0)

                    For X As Int32 = 0 To mlItemUB
                        With muItems(lSorted(X))
                            Dim sText As String = .sName & ": " & .Quantity.ToString("#,##0")

                            Dim bFound As Boolean = False
                            For Y As Int32 = 0 To lstData.ListCount - 1
                                If lstData.ItemData(Y) = .ObjectID AndAlso lstData.ItemData2(Y) = .ObjTypeID Then
                                    If lstData.List(Y) <> sText Then lstData.List(Y) = sText
                                    bFound = True
                                    Exit For
                                End If
                            Next Y
                            If bFound = False Then
                                If (.ObjTypeID = ObjectType.eMineralCache OrElse .ObjTypeID = ObjectType.eMineral) Then
                                    If lType = 0 Then AddListBoxItem(lstData, sText, .ObjectID, .ObjTypeID)
                                Else
                                    If lType = 1 Then AddListBoxItem(lstData, sText, .ObjectID, .ObjTypeID)
                                End If
                            End If
                        End With
                    Next X
                Next lType
            End If
        Catch
            mlLastUpdate += 1
        End Try

        lstData.RenderIcons = True
    End Sub

    Private Sub AddListBoxItem(ByRef lstData As UIListBox, ByVal sText As String, ByVal lID1 As Int32, ByVal lID2 As Int32)
        lstData.AddItem(sText, False)
        lstData.ItemData(lstData.NewIndex) = lID1
        lstData.ItemData2(lstData.NewIndex) = lID2
        lstData.ApplyIconOffset(lstData.NewIndex) = True

        lstData.rcIconRectangle(lstData.NewIndex) = GetTypeIDIconRect(CShort(Math.Abs(lID2)))
        lstData.IconForeColor(lstData.NewIndex) = GetTypeIDIconColor(CShort(Math.Abs(lID2)))
    End Sub

    Public Function GetObjectQuantity(ByVal lObjID As Int32, ByVal iObjTypeID As Int16) As Int32
        For X As Int32 = 0 To mlItemUB
            If muItems(X).ObjectID = lObjID AndAlso muItems(X).ObjTypeID = iObjTypeID Then Return muItems(X).Quantity
        Next X
        Return 0
    End Function

    Public Sub AdjustObjectQuantity(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal lAddVal As Int32)
        For X As Int32 = 0 To mlItemUB
            If muItems(X).ObjectID = lObjID AndAlso muItems(X).ObjTypeID = iObjTypeID Then
                muItems(X).Quantity += lAddVal
                mlLastUpdate = glCurrentCycle
                Return
            End If
        Next X
    End Sub

    Public Shared Function GetTypeIDIconRect(ByVal iTypeID As Int16) As Rectangle
        Select Case iTypeID
            Case ObjectType.eArmorTech
                Return New Rectangle(0, 0, 16, 16)
            Case ObjectType.eEngineTech
                Return New Rectangle(0, 16, 16, 16)
            Case ObjectType.eRadarTech
                Return New Rectangle(48, 0, 16, 16)
            Case ObjectType.eShieldTech
                Return New Rectangle(16, 16, 16, 16)
            Case ObjectType.eWeaponTech
                Return New Rectangle(32, 32, 16, 16)
            Case ObjectType.eOfficers
                Return New Rectangle(16, 0, 16, 16)
            Case ObjectType.eEnlisted
                Return New Rectangle(16, 0, 16, 16)
            Case ObjectType.eColonists
                Return New Rectangle(16, 0, 16, 16)
            Case ObjectType.eUnit, ObjectType.eUnitDef
                Return New Rectangle(48, 32, 16, 16)
            Case ObjectType.eFacility, ObjectType.eFacilityDef
                Return New Rectangle(0, 48, 16, 16)
        End Select
        Return Rectangle.Empty
    End Function
    Public Shared Function GetTypeIDIconColor(ByVal iTypeID As Int16) As System.Drawing.Color
        Select Case iTypeID
            Case ObjectType.eArmorTech
                Return Color.White
            Case ObjectType.eEngineTech
                Return System.Drawing.Color.FromArgb(255, 255, 255, 0)
            Case ObjectType.eRadarTech
                Return System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Case ObjectType.eShieldTech
                Return System.Drawing.Color.FromArgb(255, 0, 255, 255)
            Case ObjectType.eWeaponTech
                Return System.Drawing.Color.FromArgb(255, 255, 0, 0)
            Case ObjectType.eOfficers
                Return Color.FromArgb(255, 192, 192, 0)
            Case ObjectType.eEnlisted
                Return Color.FromArgb(255, 64, 192, 64)
            Case ObjectType.eColonists
                Return Color.White
            Case ObjectType.eUnit, ObjectType.eUnitDef
                Return Color.FromArgb(255, 64, 192, 64)
            Case ObjectType.eFacility, ObjectType.eFacilityDef
                Return Color.FromArgb(255, 255, 128, 64)
        End Select
    End Function
End Class
