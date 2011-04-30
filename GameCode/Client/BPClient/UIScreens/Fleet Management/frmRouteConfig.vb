Option Strict On

'Clear - clears all route steps
'Force Next Step - makes the selected step the next step in the route
'Pause - pauses the route
'Add Step - prompts for location/entity to add as a step
'  - if a location is selected, the unit will move to that location
'  - if an entity is selected, the unit will dock with that entity
'  - if a cache is selected, the unit will set mining loc to that cache (if it has mining)
'Remove Step - removes the selected step from the route
'Begin Route - starts executing at the selected step
'Fill Up With Mineral - need to be able to select a mineral

'Destination(21)      Located(21)          Pickup Mineral
'Destination(21)      Destination(21)      Destination(21)

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmRouteConfig
	Inherits UIWindow

	Private Const ml_ROUTE_ITEMS As Int32 = 7
 
	Private vscScroll As UIScrollBar
	Private WithEvents btnBegin As UIButton
	Private WithEvents btnClose As UIButton
	Private WithEvents btnPause As UIButton
	Private WithEvents btnClear As UIButton
	Private WithEvents btnForceNext As UIButton

	Private WithEvents btnRunOnce As UIButton

    Private lblTemplates As UILabel
    Private WithEvents cboTemplates As UIComboBox
    Private WithEvents btnEditTemplate As UIButton

	Private lnDiv1 As UILine
	Private lblTitle As UILabel
	Private lblDestTitle As UILabel
	Private lblLocTitle As UILabel
	Private lblMineralTitle As UILabel

	Private mlEntityIndex As Int32 = -1

	Private msCurrentSetMineralControl As String = ""

	Private msCurrentSetLocControl As String = ""

	Private moPopup As fraSelectMin = Nothing
    Private mbInPopup As Boolean = False
    Private mbScrollOne As Boolean = False

    Public Shared sParms() As String = Nothing

#Region "  Route Config Lines  "
	Private lblDestination() As UILabel
	Private lblMineral() As UILabel
	Private btnRemoveItem() As UIButton
	Private btnSetMineral() As UIButton
	Private btnSetLocation() As UIButton
	Private lnDivs() As UILine
#End Region

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenRoutesWindow)

		'frmRouteConfig initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eRouteConfig
            .ControlName = "frmRouteConfig"
            .Left = 293
            .Top = 160
            .Width = 511
            .Height = 255
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 1
        End With

		'vscScroll initial props
		vscScroll = New UIScrollBar(oUILib, True)
		With vscScroll
			.ControlName = "vscScroll"
			.Left = 485
			.Top = 50
			.Width = 24
			.Height = 175
			.Enabled = True
			.Visible = True
			.Value = 0
			.MaxValue = 100
			.MinValue = 0
			.SmallChange = 1
			.LargeChange = 1
			.ReverseDirection = True
		End With
		Me.AddChild(CType(vscScroll, UIControl))

		'btnBegin initial props
		btnBegin = New UIButton(oUILib)
		With btnBegin
			.ControlName = "btnBegin"
			.Left = 410
			.Top = 228
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Begin Route"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnBegin, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 485
			.Top = 1
			.Width = 24
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "X"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnClose, UIControl))

		'btnPause initial props
		btnPause = New UIButton(oUILib)
		With btnPause
			.ControlName = "btnPause"
			.Left = 305
			.Top = 228
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Pause"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnPause, UIControl))

		'btnClear initial props
		btnClear = New UIButton(oUILib)
		With btnClear
			.ControlName = "btnClear"
			.Left = 5
			.Top = 228
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Clear"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnClear, UIControl))

		'btnForceNext initial props
		btnForceNext = New UIButton(oUILib)
		With btnForceNext
			.ControlName = "btnForceNext"
			.Left = 110
			.Top = 228
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Force Next"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnForceNext, UIControl))

		'btnRunOnce initial props
		btnRunOnce = New UIButton(oUILib)
		With btnRunOnce
			.ControlName = "btnRunOnce"
			.Left = btnForceNext.Left + btnForceNext.Width + 5
			.Top = 228
			.Width = 85
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Run Once"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRunOnce, UIControl))

		'lnDiv1 initial props
		lnDiv1 = New UILine(oUILib)
		With lnDiv1
			.ControlName = "lnDiv1"
			.Left = 1
			.Top = 25
			.Width = 511
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 0
			.Width = 144
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Route Configuration"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'lblDest initial props
		lblDestTitle = New UILabel(oUILib)
		With lblDestTitle
			.ControlName = "lblDest"
			.Left = 29
			.Top = 25
			.Width = 150
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Destination Name"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
        Me.AddChild(CType(lblDestTitle, UIControl))

        'lblTemplates initial props
        lblTemplates = New UILabel(oUILib)
        With lblTemplates
            .ControlName = "lblTemplates"
            .Left = Me.Width \ 2 - 80
            .Top = 1
            .Width = 80
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Templates:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTemplates, UIControl))

		'lblMineral initial props
		lblMineralTitle = New UILabel(oUILib)
		With lblMineralTitle
			.ControlName = "lblMineral"
			.Left = 280
			.Top = 25
			.Width = 150
			.Height = 24
			.Enabled = True
			.Visible = True
			'.Caption = "Pickup Mineral"
			.Caption = "Waypoint Action"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMineralTitle, UIControl))

		'Now, our line items, we have 7 of them
		ReDim lblDestination(ml_ROUTE_ITEMS - 1)
		ReDim lblMineral(ml_ROUTE_ITEMS - 1)
		ReDim btnRemoveItem(ml_ROUTE_ITEMS - 1)
		ReDim btnSetMineral(ml_ROUTE_ITEMS - 1)
		ReDim btnSetLocation(ml_ROUTE_ITEMS - 1)
		ReDim lnDivs(ml_ROUTE_ITEMS)

		For X As Int32 = 0 To ml_ROUTE_ITEMS - 1
			Dim lItemTop As Int32 = 50 + (X * 25)
			Dim lItemLeft As Int32 = 2

			lnDivs(X) = New UILine(oUILib)
			With lnDivs(X)
				.ControlName = "lnDiv1(" & X & ")"
				.Left = lItemLeft
				.Top = lItemTop - 1
				.Width = vscScroll.Left - lItemLeft
				.Height = 0
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
			End With
			Me.AddChild(CType(lnDivs(X), UIControl))

			btnSetLocation(X) = New UIButton(oUILib)
			With btnSetLocation(X)
				.ControlName = "btnSetLocation(" & X & ")"
				.Left = lItemLeft + 3
				.Top = lItemTop
				.Width = 24
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "+"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
				.ToolTipText = "Click to set the location of this item adding it to the list of the route."
			End With
			Me.AddChild(CType(btnSetLocation(X), UIControl))

			lblDestination(X) = New UILabel(oUILib)
			With lblDestination(X)
				.ControlName = "lblDestination(" & X & ")"
				.Left = btnSetLocation(X).Left + btnSetLocation(X).Width
				.Top = lItemTop
				.Width = 250
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Unassigned"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblDestination(X), UIControl))


			lblMineral(X) = New UILabel(oUILib)
			With lblMineral(X)
				.ControlName = "lblMineral(" & X & ")"
				.Left = lblDestination(X).Left + lblDestination(X).Width
				.Top = lItemTop
				.Width = 150
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Unload"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblMineral(X), UIControl))

			btnRemoveItem(X) = New UIButton(oUILib)
			With btnRemoveItem(X)
				.ControlName = "btnRemoveItem(" & X & ")"
				.Left = vscScroll.Left - 25 ' btnSetMineral(X).Left + btnSetMineral(X).Width 
				.Top = lItemTop
				.Width = 24
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "X"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
				.ToolTipText = "Click to remove this item from the route list."
			End With
			Me.AddChild(CType(btnRemoveItem(X), UIControl))

			btnSetMineral(X) = New UIButton(oUILib)
			With btnSetMineral(X)
				.ControlName = "btnSetMineral(" & X & ")"
				.Left = btnRemoveItem(X).Left - 25
				.Top = lItemTop
				.Width = 24
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "M"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
				.ToolTipText = "Click to change the mineral to be picked up upon arrival at destination." & vbCrLf & "This requires that destination is a facility or unit."
			End With
			Me.AddChild(CType(btnSetMineral(X), UIControl))
		Next X

		lnDivs(ml_ROUTE_ITEMS) = New UILine(oUILib)
		With lnDivs(ml_ROUTE_ITEMS)
			.ControlName = "lnDiv1(" & ml_ROUTE_ITEMS & ")"
			.Left = 2
			.Top = 49 + (ml_ROUTE_ITEMS * 25)
			.Width = vscScroll.Left - 2
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
        Me.AddChild(CType(lnDivs(ml_ROUTE_ITEMS), UIControl))

        cboTemplates = New UIComboBox(oUILib)
        With cboTemplates
            .ControlName = "cboTemplates"
            .Left = Me.Width \ 2
            .Top = 3 'lblTitle.Top
            .Width = 173
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = Me.Height - .Top - .Height - 4 '128
        End With
        Me.AddChild(CType(cboTemplates, UIControl))


        btnEditTemplate = New UIButton(oUILib)
        With btnEditTemplate
            .ControlName = "btnEditTemplate"
            .Left = cboTemplates.Left + cboTemplates.Width + 5
            .Top = 1 'lblTitle.Top
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "..."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to add or edit route templates."
        End With
        Me.AddChild(CType(btnEditTemplate, UIControl))

		For X As Int32 = 0 To ml_ROUTE_ITEMS - 1
			AddHandler btnRemoveItem(X).Click, AddressOf HandleRemoveItemButtonClick
			AddHandler btnSetMineral(X).Click, AddressOf HandleSetMineralButtonClick
			AddHandler btnSetLocation(X).Click, AddressOf HandleSetLocationButtonClick
		Next X

        frmRouteTemplate.RequestTemplates()

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub FillTemplateList()
        Dim lID As Int32 = -1
        If cboTemplates.ListIndex > -1 Then lID = cboTemplates.ItemData(cboTemplates.ListIndex)

        cboTemplates.Clear()
        Try
            For X As Int32 = 0 To frmRouteTemplate.lTemplateUB
                If frmRouteTemplate.oTemplates(X) Is Nothing = False Then
                    cboTemplates.AddItem(frmRouteTemplate.oTemplates(X).sTemplateName)
                    cboTemplates.ItemData(cboTemplates.NewIndex) = frmRouteTemplate.oTemplates(X).lTemplateID
                End If
            Next X
        Catch
        End Try

        If lID > -1 Then cboTemplates.FindComboItemData(lID)
    End Sub

	Public Sub SetFromEntity(ByVal lEntityIndex As Int32)
		mlEntityIndex = lEntityIndex

		vscScroll.MaxValue = 0
		vscScroll.Value = 0

		Dim yMsg(5) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eGetRouteList).CopyTo(yMsg, 0)
		If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.lEntityIdx(mlEntityIndex) <> -1 Then
			System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, 2)
		Else : Return
		End If
		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Private Sub frmRouteConfig_OnNewFrame() Handles Me.OnNewFrame
		If mbInPopup = True Then Return
        Try
            If cboTemplates Is Nothing = False Then
                If cboTemplates.ListCount <> frmRouteTemplate.lTemplateUB + 1 Then
                    FillTemplateList()
                End If
            End If
            If mlEntityIndex = -1 OrElse goCurrentEnvir Is Nothing OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then
                mlEntityIndex = -1
                For X As Int32 = 0 To ml_ROUTE_ITEMS - 1
                    ConfigureItemAsUnused(X)
                Next X
            Else

                'If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso glCurrentEnvirView <> CurrentView.ePlanetView Then
                '    MyBase.moUILib.RemoveWindow(Me.ControlName)
                '    If goUILib.lUISelectState = UILib.eSelectState.eSelectRouteLoc Then goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                '    Return
                'ElseIf goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso glCurrentEnvirView <> CurrentView.eSystemView Then
                '    MyBase.moUILib.RemoveWindow(Me.ControlName)
                '    If goUILib.lUISelectState = UILib.eSelectState.eSelectRouteLoc Then goUILib.lUISelectState = UILib.eSelectState.eNoSelectState
                '    Return
                'End If

                'ok, get our entity
                Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(mlEntityIndex)
                If oEntity Is Nothing = False Then
                    'ok, now, check our route
                    If oEntity.bRoutePaused = True Then
                        If btnPause.Caption = "Pause" Then btnPause.Caption = "Unpause"
                    ElseIf btnPause.Caption = "Unpause" Then
                        btnPause.Caption = "Pause"
                    End If
                    If oEntity.lRouteUB = -1 Then
                        ConfigureItemAsAdd(0)
                        For X As Int32 = 1 To ml_ROUTE_ITEMS - 1
                            ConfigureItemAsUnused(X)
                        Next X
                    Else
                        'ok, the person has items...
                        Dim lScrMax As Int32 = Math.Max((oEntity.lRouteUB + 1) - ml_ROUTE_ITEMS + 1, 0)
                        If vscScroll.MaxValue <> lScrMax Then vscScroll.MaxValue = lScrMax
                        If vscScroll.Value > vscScroll.MaxValue Then vscScroll.Value = vscScroll.MaxValue

                        'Now, go through our items
                        For X As Int32 = vscScroll.Value To vscScroll.Value + ml_ROUTE_ITEMS - 1
                            Dim lListIdx As Int32 = X - vscScroll.Value
                            If X > oEntity.lRouteUB Then
                                If X = oEntity.lRouteUB + 1 Then
                                    ConfigureItemAsAdd(lListIdx)
                                Else : ConfigureItemAsUnused(lListIdx)
                                End If
                                Continue For
                            End If
                            With oEntity.uRoute(X)
                                ConfigureItemAsUsed(lListIdx)
                                Dim sTemp As String = ""
                                If .lDestID > -1 AndAlso .iDestTypeID > -1 Then
                                    If .iDestTypeID = ObjectType.eWormhole Then
                                        sTemp = "Wormhole"
                                        If goGalaxy Is Nothing = False Then
                                            Dim bFound As Boolean = False
                                            For iSystem As Int32 = 0 To goGalaxy.mlSystemUB
                                                For iWormhole As Int32 = 0 To goGalaxy.moSystems(iSystem).WormholeUB
                                                    If goGalaxy.moSystems(iSystem).moWormholes(iWormhole).ObjectID = .lDestID Then
                                                        If .lLocX = goGalaxy.moSystems(iSystem).moWormholes(iWormhole).LocX1 AndAlso .lLocZ = goGalaxy.moSystems(iSystem).moWormholes(iWormhole).LocY1 Then
                                                            sTemp = "Wormhole To " & goGalaxy.moSystems(iSystem).moWormholes(iWormhole).System2.SystemName
                                                        ElseIf .lLocX = goGalaxy.moSystems(iSystem).moWormholes(iWormhole).LocX2 AndAlso .lLocZ = goGalaxy.moSystems(iSystem).moWormholes(iWormhole).LocY2 Then
                                                            sTemp = "Wormhole To " & goGalaxy.moSystems(iSystem).moWormholes(iWormhole).System1.SystemName
                                                        End If
                                                        bFound = True
                                                        Exit For
                                                    End If
                                                Next
                                                If bFound = True Then Exit For
                                            Next
                                        End If
                                    Else
                                        sTemp = GetCacheObjectValue(.lDestID, .iDestTypeID)
                                    End If
                                Else : sTemp = ""
                                End If
                                If .lLocationID > 0 AndAlso .iLocationTypeID > 0 Then
                                    sTemp &= " (" & GetCacheObjectValue(.lLocationID, .iLocationTypeID) & ")"
                                End If
                                If lblDestination(lListIdx).Caption <> sTemp Then lblDestination(lListIdx).Caption = sTemp

                                If .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eMineral Then
                                    sTemp = "Load: Any/All Minerals"
                                ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eComponentCache Then
                                    sTemp = "Load: Any Components"
                                ElseIf .iLoadItemTypeID = ObjectType.eMineral Then
                                    sTemp = ""
                                    For Z As Int32 = 0 To glMineralUB
                                        If glMineralIdx(Z) = .lLoadItemID Then
                                            sTemp = "Load: " & goMinerals(Z).MineralName
                                            Exit For
                                        End If
                                    Next Z
                                ElseIf .iLoadItemTypeID = ObjectType.eColonists Then
                                    sTemp = "Load Colonists"
                                ElseIf .iLoadItemTypeID = ObjectType.eOfficers Then
                                    sTemp = "Load Officers"
                                ElseIf .iLoadItemTypeID = ObjectType.eEnlisted Then
                                    sTemp = "Load Enlisted"
                                ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = -1 Then
                                    sTemp = "Any Cargo"
                                ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eArmorTech Then
                                    sTemp = "All Armor"
                                ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eEngineTech Then
                                    sTemp = "All Engines"
                                ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eRadarTech Then
                                    sTemp = "All Radar"
                                ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eShieldTech Then
                                    sTemp = "All Shields"
                                ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eWeaponTech Then
                                    sTemp = "All Weapons"
                                ElseIf .iLoadItemTypeID = ObjectType.eUnitDef OrElse .iLoadItemTypeID = ObjectType.eFacilityDef Then
                                    sTemp = ""
                                    For Z As Int32 = 0 To glEntityDefUB
                                        If glEntityDefIdx(Z) = .lLoadItemID AndAlso goEntityDefs(Z).ObjTypeID = .iLoadItemTypeID Then
                                            sTemp = "Build: " & goEntityDefs(X).DefName
                                            Exit For
                                        End If
                                    Next Z
                                ElseIf .iLoadItemTypeID = Int16.MaxValue Then
                                    sTemp = "Wait " & .lLoadItemID & " seconds"
                                ElseIf .iLoadItemTypeID = ObjectType.eWormhole Then
                                    sTemp = "Jump"
                                ElseIf .lLoadItemID > -1 AndAlso .iLoadItemTypeID > -1 Then
                                    sTemp = GetCacheObjectValue(.lLoadItemID, .iLoadItemTypeID)
                                Else : sTemp = ""
                                End If

                                If (.yFlag And 1) <> 0 Then
                                    If sTemp = "" Then sTemp = "Unload Only" Else sTemp &= " (Unload)"
                                End If

                                If .iDestTypeID = ObjectType.eWormhole Then
                                    sTemp = "Jump"
                                End If

                                If lblMineral(lListIdx).Caption <> sTemp Then lblMineral(lListIdx).Caption = sTemp
                            End With
                        Next X
                    End If
                End If
            End If
            If mbScrollOne = True AndAlso vscScroll.Value < vscScroll.MaxValue Then
                mbScrollOne = False
                vscScroll.Value += 1
            End If
        Catch
        End Try
	End Sub

	Private Sub ConfigureItemAsUnused(ByVal lIndex As Int32)
		If lblDestination(lIndex).Visible = True Then lblDestination(lIndex).Visible = False
		If lblMineral(lIndex).Visible = True Then lblMineral(lIndex).Visible = False
		If btnRemoveItem(lIndex).Visible = True Then btnRemoveItem(lIndex).Visible = False
		If btnSetMineral(lIndex).Visible = True Then btnSetMineral(lIndex).Visible = False
		If btnSetLocation(lIndex).Visible = True Then btnSetLocation(lIndex).Visible = False
		If lblMineral(lIndex).Caption <> "Unload" Then lblMineral(lIndex).Caption = "Unload"
		If lblDestination(lIndex).Caption <> "Unassigned" Then lblDestination(lIndex).Caption = "Unassigned"
	End Sub
	Private Sub ConfigureItemAsAdd(ByVal lIndex As Int32)
		If lblDestination(lIndex).Visible = False Then lblDestination(lIndex).Visible = True
		If lblMineral(lIndex).Visible = False Then lblMineral(lIndex).Visible = True
		If btnRemoveItem(lIndex).Visible = True Then btnRemoveItem(lIndex).Visible = False
		If btnSetMineral(lIndex).Visible = True Then btnSetMineral(lIndex).Visible = False
		If btnSetLocation(lIndex).Visible = False Then btnSetLocation(lIndex).Visible = True
		If lblMineral(lIndex).Caption <> "" Then lblMineral(lIndex).Caption = ""
		If lblDestination(lIndex).Caption <> "" Then lblDestination(lIndex).Caption = ""
	End Sub
	Private Sub ConfigureItemAsUsed(ByVal lIndex As Int32)
		If lblDestination(lIndex).Visible = False Then lblDestination(lIndex).Visible = True
		If lblMineral(lIndex).Visible = False Then lblMineral(lIndex).Visible = True
		If btnRemoveItem(lIndex).Visible = False Then btnRemoveItem(lIndex).Visible = True
		If btnSetMineral(lIndex).Visible = False Then btnSetMineral(lIndex).Visible = True
		If btnSetLocation(lIndex).Visible = False Then btnSetLocation(lIndex).Visible = True
	End Sub

	Private Sub HandleRemoveItemButtonClick(ByVal sName As String)
		Try
			Dim lIdx As Int32 = sName.IndexOf("("c)
			If lIdx > -1 Then
				lIdx = CInt(Val(sName.Substring(lIdx + 1)))
				lIdx += vscScroll.Value

				Dim yMsg(9) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eRemoveRouteItem).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(lIdx).CopyTo(yMsg, lPos) : lPos += 4
				MyBase.moUILib.SendMsgToPrimary(yMsg)
			End If
		Catch
		End Try
	End Sub
	Private Sub HandleSetMineralButtonClick(ByVal sName As String)
		If mlEntityIndex = -1 Then Return

		mbInPopup = True

		SetAllForPopup()

		msCurrentSetMineralControl = sName
		If moPopup Is Nothing Then
			moPopup = New fraSelectMin(MyBase.moUILib)
			Me.AddChild(CType(moPopup, UIControl))
			AddHandler moPopup.CancelClicked, AddressOf Popup_CancelClick
            AddHandler moPopup.ItemAccepted, AddressOf Popup_ItemClick
		End If
		moPopup.Left = Me.Width \ 2 - moPopup.Width \ 2
		moPopup.Top = Me.Height \ 2 - moPopup.Height \ 2
		Try
			Dim lIdx As Int32 = sName.IndexOf("("c)
			If lIdx > -1 Then
				lIdx = CInt(Val(sName.Substring(lIdx + 1)))
				lIdx += vscScroll.Value

                Dim iTypeID As Int16 = 0
                Dim yFlag As Byte = 1
				If goCurrentEnvir Is Nothing = False AndAlso mlEntityIndex <> -1 AndAlso goCurrentEnvir.lEntityIdx(mlEntityIndex) <> -1 Then
                    Dim lRouteIdx As Int32 = lIdx
                    lIdx = goCurrentEnvir.oEntity(mlEntityIndex).uRoute(lRouteIdx).lLoadItemID
                    iTypeID = goCurrentEnvir.oEntity(mlEntityIndex).uRoute(lRouteIdx).iLoadItemTypeID
                    yFlag = goCurrentEnvir.oEntity(mlEntityIndex).uRoute(lRouteIdx).yFlag
				End If
                moPopup.SetFromSelected(lIdx, iTypeID, yFlag)

			End If
		Catch
		End Try

		moPopup.Visible = True


	End Sub
	Private Sub HandleSetLocationButtonClick(ByVal sName As String)
		msCurrentSetLocControl = sName
		Me.Visible = False
		MyBase.moUILib.lUISelectState = UILib.eSelectState.eSelectRouteLoc
		MyBase.moUILib.AddNotification("Left-click a location for this way point.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		MyBase.moUILib.AddNotification("Right-click to cancel.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eRoutePlusClick, -1, -1, -1, "")
		End If
	End Sub

    Private Sub Popup_ItemClick(ByVal lItemID As Int32, ByVal iTypeID As Int16, ByVal bUnload As Boolean)
        Try
            Dim lIdx As Int32 = msCurrentSetMineralControl.IndexOf("("c)
            If lIdx > -1 Then
                lIdx = CInt(Val(msCurrentSetMineralControl.Substring(lIdx + 1)))
                lIdx += vscScroll.Value

                If goCurrentEnvir Is Nothing = False AndAlso mlEntityIndex <> -1 AndAlso goCurrentEnvir.lEntityIdx(mlEntityIndex) <> -1 Then
                    'lIdx = goCurrentEnvir.oEntity(mlEntityIndex).uRoute(lIdx).lLoadItemID
                    Dim uRoute As RouteItem = goCurrentEnvir.oEntity(mlEntityIndex).uRoute(lIdx)

                    Dim yMsg(22) As Byte
                    Dim lPos As Int32 = 0

                    System.BitConverter.GetBytes(GlobalMessageCode.eSetRouteMineral).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(lIdx).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(uRoute.lDestID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(uRoute.iDestTypeID).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lItemID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2

                    If bUnload = True Then
                        yMsg(lPos) = 1
                    Else : yMsg(lPos) = 0
                    End If
                    lPos += 1

                    MyBase.moUILib.SendMsgToPrimary(yMsg)
                End If
            End If
        Catch
        End Try
        moPopup.Visible = False
        mbInPopup = False
        SetAllForPopup()
    End Sub
	Private Sub Popup_CancelClick()
		moPopup.Visible = False
		mbInPopup = False
		SetAllForPopup()
	End Sub
	Private Sub SetAllForPopup()
		Dim bValue As Boolean = Not mbInPopup
		For X As Int32 = 0 To ml_ROUTE_ITEMS - 1
			btnRemoveItem(X).Enabled = bValue
			btnSetMineral(X).Enabled = bValue
			btnSetLocation(X).Enabled = bValue
		Next X
		btnClose.Enabled = bValue
		btnBegin.Enabled = bValue
		btnPause.Enabled = bValue
        btnForceNext.Enabled = bValue
        btnRunOnce.Enabled = bValue
        btnClear.Enabled = bValue
        cboTemplates.Enabled = bValue
        btnEditTemplate.Enabled = bValue
	End Sub

	Public Sub SetLocResultVector(ByVal vecLoc As Vector3)
		Try
			Dim sName As String = msCurrentSetLocControl
			Dim lIdx As Int32 = sName.IndexOf("("c)
			If lIdx > -1 Then
				lIdx = CInt(Val(sName.Substring(lIdx + 1)))
				lIdx += vscScroll.Value

				Dim yMsg(23) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(lIdx).CopyTo(yMsg, lPos) : lPos += 4

                Dim lDestID As Int32 = -1
                Dim iDestTypeID As Int16 = -1
                If goGalaxy.CurrentSystemIdx <> -1 Then
                    If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx <> -1 AndAlso (glCurrentEnvirView = CurrentView.ePlanetMapView OrElse glCurrentEnvirView = CurrentView.ePlanetView) Then
                        lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).ObjectID
                        iDestTypeID = ObjectType.ePlanet
                    Else
                        lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID
                        iDestTypeID = ObjectType.eSolarSystem
                    End If
                Else
                    Return
                End If

                'goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                System.BitConverter.GetBytes(lDestID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(iDestTypeID).CopyTo(yMsg, lPos) : lPos += 2

				System.BitConverter.GetBytes(CInt(vecLoc.X)).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(CInt(vecLoc.Z)).CopyTo(yMsg, lPos) : lPos += 4

				If NewTutorialManager.TutorialOn = True Then
					NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eRouteSelectDest, -1, CInt(vecLoc.X), CInt(vecLoc.Z), "")
				End If

				MyBase.moUILib.SendMsgToPrimary(yMsg)
			End If
		Catch
		End Try

		Me.Visible = True
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        mbScrollOne = True
	End Sub
	Public Sub SetLocResultGUID(ByRef oEntity As BaseEntity)

		If NewTutorialManager.TutorialOn = True Then
            If oEntity Is Nothing = False Then
                ReDim sParms(1)
                sParms(0) = CByte(oEntity.yProductionType).ToString
                sParms(1) = oEntity.ObjTypeID.ToString
                If MyBase.moUILib.CommandAllowedWithParms(True, "RouteSetLocation", sParms, False) = False Then Return
            End If
		End If

		Try
			Dim sName As String = msCurrentSetLocControl
			Dim lIdx As Int32 = sName.IndexOf("("c)
			If lIdx > -1 Then
				lIdx = CInt(Val(sName.Substring(lIdx + 1)))
				lIdx += vscScroll.Value

				Dim yMsg(23) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(lIdx).CopyTo(yMsg, lPos) : lPos += 4

				oEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
				System.BitConverter.GetBytes(CInt(oEntity.LocX)).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(CInt(oEntity.LocZ)).CopyTo(yMsg, lPos) : lPos += 4

                If NewTutorialManager.TutorialOn = True Then
					'If sParms Is Nothing = False AndAlso sParms.Length >= 3 Then
					'	If oEntity.yProductionType.ToString <> sParms(0) OrElse oEntity.ObjTypeID.ToString <> sParms(2) Then		'oEntity.ObjectID.ToString <> sParms(1) OrElse 
					'		Return
					'	End If
					'End If
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eRouteSelectDest, oEntity.ObjectID, oEntity.ObjTypeID, oEntity.yProductionType, "")
                End If

                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
        Catch
        End Try

		Me.Visible = True
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        mbScrollOne = True
	End Sub
	Public Sub SetLocResultWormhole(ByRef oWormhole As Wormhole, ByVal bSystem1 As Boolean)
		Try
			Dim sName As String = msCurrentSetLocControl
			Dim lIdx As Int32 = sName.IndexOf("("c)
			If lIdx > -1 Then
				lIdx = CInt(Val(sName.Substring(lIdx + 1)))
				lIdx += vscScroll.Value

				Dim yMsg(23) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(lIdx).CopyTo(yMsg, lPos) : lPos += 4

				oWormhole.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
				If bSystem1 = True Then
					System.BitConverter.GetBytes(oWormhole.LocX1).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(oWormhole.LocY1).CopyTo(yMsg, lPos) : lPos += 4
				Else
					System.BitConverter.GetBytes(oWormhole.LocX2).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(oWormhole.LocY2).CopyTo(yMsg, lPos) : lPos += 4
				End If
				
				MyBase.moUILib.SendMsgToPrimary(yMsg)
			End If
		Catch
		End Try

		Me.Visible = True
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        mbScrollOne = True
	End Sub

#Region "  Select Mineral Popup  "
	'Interface created from Interface Builder
	Public Class fraSelectMin
		Inherits UIWindow

		Private lblTitle As UILabel
        Private tvwItems As UITreeView 'UIListBox
		Private WithEvents btnAccept As UIButton
		Private WithEvents btnCancel As UIButton

        Private chkUnload As UICheckBox

		Public Event CancelClicked()
        Public Event ItemAccepted(ByVal lItemID As Int32, ByVal iTypeID As Int16, ByVal bUnload As Boolean)

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

			'fraSelectMin initial props
			With Me
				.ControlName = "fraSelectMin"
                .Left = 410
                .Top = 169
				.Width = 250
                .Height = 250 ' 220
                .Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
				.BorderLineWidth = 1
			End With

			'lblTitle initial props
			lblTitle = New UILabel(oUILib)
			With lblTitle
				.ControlName = "lblTitle"
				.Left = 5
				.Top = 0
				.Width = 184
				.Height = 24
				.Enabled = True
				.Visible = True
                .Caption = "Select an item to pick up"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblTitle, UIControl))

            'tvwItems initial props
            tvwItems = New UITreeView(oUILib)
            With tvwItems
                .ControlName = "tvwItems"
                .Left = 5
                .Top = 25
                .Width = 240
                .Height = Me.Height - 85   '113
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            End With
            Me.AddChild(CType(tvwItems, UIControl))

			'btnAccept initial props
			btnAccept = New UIButton(oUILib)
			With btnAccept
				.ControlName = "btnAccept"
				.Left = 15
                .Top = Me.Height - 26
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Accept"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
            Me.AddChild(CType(btnAccept, UIControl))

            'chkUnload initial props
            chkUnload = New UICheckBox(oUILib)
            With chkUnload
                .ControlName = "chkUnload"
                .Left = 15
                .Top = tvwItems.Top + tvwItems.Height + 5
                .Width = 120
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Unload Cargo"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = True
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "If checked, any cargo will be dropped off at this waypoint."
            End With
            Me.AddChild(CType(chkUnload, UIControl))

			'btnCancel initial props
			btnCancel = New UIButton(oUILib)
			With btnCancel
				.ControlName = "btnCancel"
				.Left = 135
                .Top = btnAccept.Top
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Clear"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			Me.AddChild(CType(btnCancel, UIControl))

			FillMineralList()
		End Sub

		Private Sub FillMineralList()
            tvwItems.Clear()

            tvwItems.AddNode("Do Not Load Cargo", -1, -2, -1, Nothing, Nothing)
            tvwItems.AddNode("Any Cargo", -1, -1, -1, Nothing, Nothing)
            tvwItems.AddNode("Any/All Alloys and Minerals", -1, ObjectType.eMineral, -1, Nothing, Nothing)
            tvwItems.AddNode("Any/All Components", -1, ObjectType.eComponentCache, -1, Nothing, Nothing)
            'tvwItems.AddNode("Colonists", -1, ObjectType.eColonists, -1, Nothing, Nothing)
            'tvwItems.AddNode("Enlisted", -1, ObjectType.eEnlisted, -1, Nothing, Nothing)
            'tvwItems.AddNode("Officers", -1, ObjectType.eOfficers, -1, Nothing, Nothing)

            Dim oAlloys As UITreeView.UITreeViewItem = tvwItems.AddNode("Alloys", Int32.MinValue, Int32.MinValue, Int32.MinValue, Nothing, Nothing)
            Dim oComponents As UITreeView.UITreeViewItem = tvwItems.AddNode("Components", Int32.MinValue, Int32.MinValue, Int32.MinValue, Nothing, Nothing)
            Dim oMinerals As UITreeView.UITreeViewItem = tvwItems.AddNode("Minerals", Int32.MinValue, Int32.MinValue, Int32.MinValue, Nothing, Nothing)
            Dim oPersonnel As UITreeView.UITreeViewItem = tvwItems.AddNode("Personnel", Int32.MinValue, Int32.MaxValue, Int32.MinValue, Nothing, Nothing)

            oMinerals.bItemBold = True : oMinerals.bUseFillColor = True
            oAlloys.bItemBold = True : oAlloys.bUseFillColor = True
            oComponents.bItemBold = True : oComponents.bUseFillColor = True
            oPersonnel.bItemBold = True : oPersonnel.bUseFillColor = True

            tvwItems.AddNode("Colonists", 0, ObjectType.eColonists, -1, oPersonnel, Nothing)
            tvwItems.AddNode("Enlisted", 0, ObjectType.eEnlisted, -1, oPersonnel, Nothing)
            tvwItems.AddNode("Officers", 0, ObjectType.eOfficers, -1, oPersonnel, Nothing)


            Dim lSorted() As Int32 = GetSortedMineralIdxArray(True)
			If lSorted Is Nothing = False Then
                For X As Int32 = 0 To lSorted.GetUpperBound(0)
                    If glMineralIdx(lSorted(X)) <> -1 AndAlso (goMinerals(lSorted(X)).ObjectID <= 157 OrElse goMinerals(lSorted(X)).ObjectID = 41991) Then
                        tvwItems.AddNode(goMinerals(lSorted(X)).MineralName, goMinerals(lSorted(X)).ObjectID, goMinerals(lSorted(X)).ObjTypeID, -1, oMinerals, Nothing)
                    End If
                Next X
                For X As Int32 = 0 To lSorted.GetUpperBound(0)
                    If glMineralIdx(lSorted(X)) <> -1 AndAlso goMinerals(lSorted(X)).ObjectID > 157 AndAlso goMinerals(lSorted(X)).ObjectID <> 41991 Then
                        tvwItems.AddNode(goMinerals(lSorted(X)).MineralName, goMinerals(lSorted(X)).ObjectID, goMinerals(lSorted(X)).ObjTypeID, -1, oAlloys, Nothing)
                    End If
                Next X
            End If

            Dim lTypeID() As Int32 = {ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech}
            Dim bExp(lTypeID.GetUpperBound(0)) As Boolean
            For X As Int32 = 0 To lTypeID.GetUpperBound(0)
                bExp(X) = False
                Dim oNode As UITreeView.UITreeViewItem = tvwItems.GetNodeByItemData2(-1, lTypeID(X))
                If oNode Is Nothing = False Then bExp(X) = oNode.bExpanded
            Next X

            Dim oArmor As UITreeView.UITreeViewItem = tvwItems.AddNode("All Armor", -1, ObjectType.eArmorTech, 0, oComponents, Nothing)
            Dim oEngine As UITreeView.UITreeViewItem = tvwItems.AddNode("All Engines", -1, ObjectType.eEngineTech, 0, oComponents, Nothing)
            Dim oRadar As UITreeView.UITreeViewItem = tvwItems.AddNode("All Radar", -1, ObjectType.eRadarTech, 0, oComponents, Nothing)
            Dim oShield As UITreeView.UITreeViewItem = tvwItems.AddNode("All Shields", -1, ObjectType.eShieldTech, 0, oComponents, Nothing)
            Dim oWeapon As UITreeView.UITreeViewItem = tvwItems.AddNode("All Weapons", -1, ObjectType.eWeaponTech, 0, oComponents, Nothing)
 
            oArmor.bItemBold = True : oArmor.bUseFillColor = True
            oEngine.bItemBold = True : oEngine.bUseFillColor = True 
            oRadar.bItemBold = True : oRadar.bUseFillColor = True
            oShield.bItemBold = True : oShield.bUseFillColor = True 
            oWeapon.bItemBold = True : oWeapon.bUseFillColor = True

            Dim oOther As UITreeView.UITreeViewItem = Nothing

            If goCurrentPlayer Is Nothing = False Then

                For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                    If goCurrentPlayer.moTechs(X) Is Nothing = False Then

                        If goCurrentPlayer.moTechs(X).bArchived = False Then 'OrElse bFilterArchived = False Then
                            Dim iTemp As Int16 = goCurrentPlayer.moTechs(X).ObjTypeID

                            If iTemp = ObjectType.eAlloyTech OrElse iTemp = ObjectType.eHullTech OrElse iTemp = ObjectType.ePrototype OrElse iTemp = ObjectType.eSpecialTech Then Continue For

                            Dim sValue As String = goCurrentPlayer.moTechs(X).GetComponentName

                            Dim oRoot As UITreeView.UITreeViewItem = GetRootNode(goCurrentPlayer.moTechs(X).ObjTypeID, goCurrentPlayer.moTechs(X).GetTechHullTypeID())
                            Dim oTech As UITreeView.UITreeViewItem = tvwItems.AddNode(sValue, goCurrentPlayer.moTechs(X).ObjectID, goCurrentPlayer.moTechs(X).ObjTypeID, -1, oRoot, Nothing)
                            If oRoot.oParentNode Is Nothing = False Then oRoot.oParentNode.lItemData3 += 1
                            oRoot.lItemData3 += 1
                        End If

                    End If
                Next X
            End If

            For X As Int32 = 0 To lTypeID.GetUpperBound(0)
                Dim oNode As UITreeView.UITreeViewItem = tvwItems.GetNodeByItemData2(-1, lTypeID(X))
                If oNode Is Nothing = False Then
                    oNode.bExpanded = bExp(X)
                    oNode.sItem = oNode.sItem & " (" & oNode.lItemData3 & ")"
                    oNode.clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                    If lTypeID(X) <> ObjectType.eAlloyTech AndAlso lTypeID(X) <> ObjectType.eArmorTech AndAlso lTypeID(X) <> ObjectType.eSpecialTech Then
                        Dim oTmpNode As UITreeView.UITreeViewItem = oNode.oFirstChild
                        While oTmpNode Is Nothing = False
                            oTmpNode.sItem = oTmpNode.sItem & " (" & oTmpNode.lItemData3 & ")"
                            oTmpNode = oTmpNode.oNextSibling
                        End While
                    End If
                End If
            Next X
        End Sub

        Private Function GetRootNode(ByVal iTypeID As Int16, ByVal lHullTypeID As Int32) As UITreeView.UITreeViewItem

            Dim oNode As UITreeView.UITreeViewItem = tvwItems.GetNodeByItemData2(-(lHullTypeID + 2), iTypeID)
            If oNode Is Nothing Then
                Dim oRoot As UITreeView.UITreeViewItem = tvwItems.GetNodeByItemData2(-1, iTypeID)
                If oRoot Is Nothing Then Return Nothing
                If iTypeID = ObjectType.eArmorTech OrElse iTypeID = ObjectType.eAlloyTech OrElse iTypeID = ObjectType.eSpecialTech Then Return oRoot
                oNode = tvwItems.AddNode(Base_Tech.GetHullTypeName(CByte(lHullTypeID)), -(lHullTypeID + 2), iTypeID, 0, oRoot, Nothing)
                oNode.bItemBold = True : oNode.bUseFillColor = True : oNode.clrFillColor = System.Drawing.Color.FromArgb(255, 64, 64, 64)
            End If
            Return oNode
        End Function

        Public Sub SetFromSelected(ByVal lItemID As Int32, ByVal iTypeID As Int16, ByVal yFlag As Byte)
            tvwItems.oSelectedNode = tvwItems.GetNodeByItemData2(lItemID, iTypeID)

            chkUnload.Value = (yFlag And 1) <> 0
        End Sub

		Private Sub btnAccept_Click(ByVal sName As String) Handles btnAccept.Click
            'If lstMinerals.ListIndex > -1 Then
            '	RaiseEvent MineralAccept(lstMinerals.ItemData(lstMinerals.ListIndex))
            'End If

            Dim oNode As UITreeView.UITreeViewItem = tvwItems.oSelectedNode
            If oNode Is Nothing = False AndAlso oNode.lItemData <> Int32.MinValue Then
                RaiseEvent ItemAccepted(oNode.lItemData, CShort(oNode.lItemData2), chkUnload.Value)
            Else
                MyBase.moUILib.AddNotification("Select an item in the list first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If

		End Sub

		Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
			RaiseEvent CancelClicked()
		End Sub
	End Class
#End Region

	Private Sub btnBegin_Click(ByVal sName As String) Handles btnBegin.Click
		Try

            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    If goCurrentEnvir.oEntity(X).bSelected = True Then
                        Dim yMsg(9) As Byte
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMsg, lPos) : lPos += 2
                        System.BitConverter.GetBytes(goCurrentEnvir.oEntity(X).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(Int32.MaxValue).CopyTo(yMsg, lPos) : lPos += 4
                        MyBase.moUILib.SendMsgToPrimary(yMsg)
                    End If
                End If
            Next X

			If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBeginRouteClicked, -1, -1, -1, "")
                If NewTutorialManager.lRouteEngineerID < 1 Then
                    NewTutorialManager.lRouteEngineerID = goCurrentEnvir.oEntity(mlEntityIndex).ObjectID
                End If
            End If

            MyBase.moUILib.RemoveWindow(Me.ControlName)
            MyBase.moUILib.RemoveWindow("frmBehavior")
		Catch
		End Try
	End Sub

	Private Sub btnClear_Click(ByVal sName As String) Handles btnClear.Click
		If btnClear.Caption.ToUpper = "CONFIRM" Then

			If mlEntityIndex < 0 Then Return
			Dim oEnvir As BaseEnvironment = goCurrentEnvir
			If oEnvir Is Nothing Then Return
			If oEnvir.lEntityUB < mlEntityIndex Then Return
			If oEnvir.lEntityIdx(mlEntityIndex) < 1 Then Return

			Dim oEntity As BaseEntity = oEnvir.oEntity(mlEntityIndex)
			If oEntity Is Nothing Then Return

			Dim yMsg(9) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eRemoveRouteItem).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(oEntity.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yMsg, lPos) : lPos += 4
			MyBase.moUILib.SendMsgToPrimary(yMsg)
			btnClear.Caption = "Clear"
		Else : btnClear.Caption = "Confirm"
		End If
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnForceNext_Click(ByVal sName As String) Handles btnForceNext.Click
		Try
			Dim yMsg(9) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yMsg, lPos) : lPos += 4
			MyBase.moUILib.SendMsgToPrimary(yMsg)
		Catch
		End Try
	End Sub

	Private Sub btnPause_Click(ByVal sName As String) Handles btnPause.Click
		Try
			Dim yMsg(9) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(-2I).CopyTo(yMsg, lPos) : lPos += 4
			MyBase.moUILib.SendMsgToPrimary(yMsg)
		Catch
		End Try
	End Sub

	Private Sub btnRunOnce_Click(ByVal sName As String) Handles btnRunOnce.Click
		Try
			Dim yMsg(9) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(-4I).CopyTo(yMsg, lPos) : lPos += 4
			MyBase.moUILib.SendMsgToPrimary(yMsg)
		Catch
		End Try
	End Sub

    Private Sub cboTemplates_ItemSelected(ByVal lItemIndex As Integer) Handles cboTemplates.ItemSelected
        If cboTemplates.ListIndex > -1 Then
            Dim lTemplateID As Int32 = cboTemplates.ItemData(cboTemplates.ListIndex)

            Try
                Dim yMsg(13) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eApplyRouteToEntities).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(lTemplateID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                SetFromEntity(mlEntityIndex)
            Catch
            End Try

        End If
    End Sub

    Private Sub btnEditTemplate_Click(ByVal sName As String) Handles btnEditTemplate.Click
        Dim oFrm As frmRouteTemplate = CType(MyBase.moUILib.GetWindow("frmRouteTemplate"), frmRouteTemplate)
        If oFrm Is Nothing = True Then oFrm = New frmRouteTemplate(MyBase.moUILib)
        oFrm.Visible = True
        oFrm.Left = Me.Left
        oFrm.Top = Me.Top - oFrm.Height - 1
        If cboTemplates Is Nothing = False Then
            If cboTemplates.ListIndex > -1 Then
                oFrm.SetFromTemplateID(cboTemplates.ItemData(cboTemplates.ListIndex))
            End If
        End If
        oFrm = Nothing
    End Sub
End Class