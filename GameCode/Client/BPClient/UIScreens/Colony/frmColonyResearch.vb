Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmColonyResearch
	Inherits UIWindow

	Private Const ml_ITEM_HEIGHT As Int32 = 24

	Private txtCost As UITextBox
	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private lnDiv2 As UILine
	Private lblResearchProjects As UILabel
	Private lblItemCosts As UILabel
	Private lblResFacs As UILabel
	Private lblResQueue As UILabel
	Private lblFacsToAssign As UILabel
    Private txtAssign As UITextBox
    Private lblQuantity As UILabel
    Private txtQuantity As UITextBox

    Private WithEvents tvwItems As UITreeView
    Private WithEvents btnClose As UIButton
    Private WithEvents lstQueue As UIListBox
    Private WithEvents btnAddToQueue As UIButton
    Private WithEvents btnRemove As UIButton

    Private Shared lFormHeight As Int32 = 0

    Private vscrResLabs As UIScrollBar

    Private mbLoading As Boolean = True

    Private Structure uProducer
        Public lObjectID As Int32
        Public iObjTypeID As Int16
        Public lProdFactor As Int32
        Public lProjectID As Int32
        Public iProjectTypeID As Int16
        Public yProgress As Byte
        Public bIsPrimary As Boolean
        Public lCyclesRemaining As Int32
        Public lCycleStarted As Int32

        Public sProducerName As String
        Public bCancelled As Boolean

        Public lProdCnt As Int32
        Public yProductionType As Byte
    End Structure
    Private muFactories() As uProducer = Nothing
    Private muResearchers() As uProducer = Nothing
    Private Structure uQueueItem
        Public lObjectID As Int32
        Public iObjTypeID As Int16
        Public yAssignCnt As Byte
        Public yQty As Byte
    End Structure
    Private muQueue() As uQueueItem = Nothing

    Private mbIgnoreQueueEvents As Boolean = False

    Private mlLastRequest As Int32

    Private moSysFont As System.Drawing.Font
    Private moSysFontBold As System.Drawing.Font

    Private rcResearch As Rectangle
    Private rcProduction As Rectangle

    Private myViewState As Byte = 0     'research = 0, production = 1
    Public Property ViewState() As Byte
        Get
            Return myViewState
        End Get
        Set(ByVal value As Byte)
            If myViewState <> value Then
                vscrResLabs.Value = 0
                'tvwItems.something to tell it to scroll to the top
            End If

            myViewState = value
            If value = 0 Then
                'research
                FillTechList()
                lblTitle.Caption = "Colonial Management"
                lblResearchProjects.Caption = "Researchable Projects"
                lblItemCosts.Caption = "Project Costs"
                lblResQueue.Caption = "Colonial Research Queue"
                lblResFacs.Caption = "Research Facilities"
                If muResearchers Is Nothing = False AndAlso muResearchers.GetUpperBound(0) > 0 Then
                    lblResFacs.Caption = "Research Facilities (" & (muResearchers.GetUpperBound(0) + 1).ToString & ")"
                End If
                txtQuantity.Visible = False
                lblQuantity.Visible = False
            Else
                'production
                FillProdList()
                lblTitle.Caption = "Colonial Management"
                lblResearchProjects.Caption = "Buildable Items"
                lblItemCosts.Caption = "Item Costs"
                lblResQueue.Caption = "Colonial Production Queue"
                lblResFacs.Caption = "Production Facilities"
                If muResearchers Is Nothing = False AndAlso muFactories.GetUpperBound(0) > 0 Then
                    lblResFacs.Caption = "Production Facilities (" & (muFactories.GetUpperBound(0) + 1).ToString & ")"
                End If
                txtQuantity.Visible = True
                lblQuantity.Visible = True
            End If
            Me.IsDirty = True
        End Set
    End Property

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenColonyProdResWindow)

        'frmColonyResearch initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eColonyResearch
            .ControlName = "frmColonyResearch"
            .Width = 512
            '.Height = 512
            'If lFormHeight > 0 Then .Height = lFormHeight Else .Height = 512
            'If .Top + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then .Height = 512

            '.Left = 245
            '.Top = 76
            'If muSettings.ColonialManagementLeft < 0 OrElse muSettings.ColonialManagementTop < 0 OrElse muSettings.ColonialManagementLeft > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width OrElse muSettings.ColonialManagementTop > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height OrElse NewTutorialManager.TutorialOn = True Then
            '    .Left = 245
            '    .Top = 76
            '    muSettings.ColonialManagementLeft = .Left
            '    muSettings.ColonialManagementTop = .Top
            'Else
            '    .Left = muSettings.ColonialManagementLeft
            '    .Top = muSettings.ColonialManagementTop
            'End If

            .DoWindowInitialPosition(245, 76, 512, 512, muSettings.ColonialManagementLeft, muSettings.ColonialManagementTop, -1, muSettings.ColonialManagementHeight, False)

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .mbAcceptReprocessEvents = True
            .mlResizeType = ResizeType.Down
            .ResizeLimits = New System.Drawing.Rectangle(512, 512, 512, 1024) 'Min-Width, Min-Height, Max-Width, Max-Height
            .Resizable = True
        End With

        'lstTechs initial props
        tvwItems = New UITreeView(oUILib)
        With tvwItems
            .ControlName = "tvwItems"
            .Left = 5
            .Top = 55
            .Width = 365
            .Height = 130
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(tvwItems, UIControl))

        'txtCost initial props
        txtCost = New UITextBox(oUILib)
        With txtCost
            .ControlName = "txtCost"
            .Left = 375
            .Top = 55
            .Width = 132
            .Height = 130
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .Locked = True
            .MultiLine = True
        End With
        Me.AddChild(CType(txtCost, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 200
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Colonial Management"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = Me.BorderLineWidth \ 2
            .Top = 30
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 485
            .Top = 4
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'lblResearchProjects initial props
        lblResearchProjects = New UILabel(oUILib)
        With lblResearchProjects
            .ControlName = "lblResearchProjects"
            .Left = 5
            .Top = 35
            .Width = 160
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Researchable Projects"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblResearchProjects, UIControl))

        'lblItemCosts initial props
        lblItemCosts = New UILabel(oUILib)
        With lblItemCosts
            .ControlName = "lblItemCosts"
            .Left = 375
            .Top = 35
            .Width = 95
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Project Costs"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblItemCosts, UIControl))

        'lblResFacs initial props
        lblResFacs = New UILabel(oUILib)
        With lblResFacs
            .ControlName = "lblResFacs"
            .Left = 5
            .Top = 315
            .Width = 200
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Research Facilities"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblResFacs, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = Me.BorderLineWidth \ 2
            .Top = 335
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'New Control initial props
        lstQueue = New UIListBox(oUILib)
        With lstQueue
            .ControlName = "lstQueue"
            .Left = 5
            .Top = 210
            .Width = 365
            .Height = 101
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstQueue, UIControl))

        'lblResQueue initial props
        lblResQueue = New UILabel(oUILib)
        With lblResQueue
            .ControlName = "lblResQueue"
            .Left = 5
            .Top = 190
            .Width = 225
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Colonial Research Queue"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblResQueue, UIControl))

        'lblFacsToAssign initial props
        lblFacsToAssign = New UILabel(oUILib)
        With lblFacsToAssign
            .ControlName = "lblFacsToAssign"
            .Left = 375
            .Top = 190
            .Width = 130
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Facilities to Assign"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblFacsToAssign, UIControl))

        'txtAssign initial props
        txtAssign = New UITextBox(oUILib)
        With txtAssign
            .ControlName = "txtAssign"
            .Left = 410
            .Top = 210
            .Width = 80
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "1"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 2
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
            .ToolTipText = "The number of facilities to assign."
        End With
        Me.AddChild(CType(txtAssign, UIControl))

        'lblQuantity initial props
        lblQuantity = New UILabel(oUILib)
        With lblQuantity
            .ControlName = "lblQuantity"
            .Left = 375
            .Top = 235
            .Width = 30
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Qty:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblQuantity, UIControl))

        'txtQuantity initial props
        txtQuantity = New UITextBox(oUILib)
        With txtQuantity
            .ControlName = "txtQuantity"
            .Left = txtAssign.Left
            .Top = lblQuantity.Top
            .Width = txtAssign.Width
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "1"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 2
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
            .ToolTipText = "The quantity that each facility will produce."
        End With
        Me.AddChild(CType(txtQuantity, UIControl))

        'btnAddToQueue initial props
        btnAddToQueue = New UIButton(oUILib)
        With btnAddToQueue
            .ControlName = "btnAddToQueue"
            .Left = 390
            .Top = 260
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Queue"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAddToQueue, UIControl))

        'btnRemove initial props
        btnRemove = New UIButton(oUILib)
        With btnRemove
            .ControlName = "btnRemove"
            .Left = 390
            .Top = 290
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Remove"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRemove, UIControl))

        'vscrResLabs initial props
        vscrResLabs = New UIScrollBar(oUILib, True)
        With vscrResLabs
            .ControlName = "vscrResLabs"
            .Left = Me.Width - 26
            .Top = 336
            .Width = 24
            .Height = Me.Height - .Top - 10
            .Enabled = True
            .Visible = True
            .ReverseDirection = True
            .Value = 0
            .MaxValue = 0
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 4
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(vscrResLabs, UIControl))

        rcProduction = New Rectangle(btnClose.Left - 102, btnClose.Top, 100, btnClose.Height)
        rcResearch = New Rectangle(btnClose.Left - 202, btnClose.Top, 100, btnClose.Height)

        moSysFont = New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0)
        moSysFontBold = New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Bold, GraphicsUnit.Point, 0)

        ViewState = 0

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
    End Sub

    Private Sub FillProdList()
        Dim lOrigID As Int32 = -1
        Dim iOrigTypeID As Int16 = -1

        Dim oNode As UITreeView.UITreeViewItem = tvwItems.oSelectedNode
        If oNode Is Nothing = False Then
            lOrigID = oNode.lItemData
            iOrigTypeID = CShort(oNode.lItemData2)
        End If

        tvwItems.Clear()
        'Ok, need to fill this list with 
        '  components that are researched
        '  all researched prototypes (in Prototype Parts)
        'Dim oAlloy As UITreeView.UITreeViewItem = tvwItems.AddNode("Alloys", Int32.MinValue, ObjectType.eAlloyTech, 0, Nothing, Nothing)
        Dim oArmor As UITreeView.UITreeViewItem = tvwItems.AddNode("Armor", Int32.MinValue, ObjectType.eArmorTech, 0, Nothing, Nothing)
        Dim oEngine As UITreeView.UITreeViewItem = tvwItems.AddNode("Engines", Int32.MinValue, ObjectType.eEngineTech, 0, Nothing, Nothing)
        Dim oRadar As UITreeView.UITreeViewItem = tvwItems.AddNode("Radar", Int32.MinValue, ObjectType.eRadarTech, 0, Nothing, Nothing)
        Dim oShield As UITreeView.UITreeViewItem = tvwItems.AddNode("Shields", Int32.MinValue, ObjectType.eShieldTech, 0, Nothing, Nothing)
        Dim oWeapon As UITreeView.UITreeViewItem = tvwItems.AddNode("Weapons", Int32.MinValue, ObjectType.eWeaponTech, 0, Nothing, Nothing)
        Dim oPrototypes As UITreeView.UITreeViewItem = tvwItems.AddNode("Prototype Parts", Int32.MinValue, ObjectType.ePrototype, 0, Nothing, Nothing)

        Dim oUnits As UITreeView.UITreeViewItem = tvwItems.AddNode("Units", Int32.MinValue, ObjectType.eUnitDef, 0, Nothing, Nothing)

        Dim oPPFacilities As UITreeView.UITreeViewItem = tvwItems.AddNode("Facilities", Int32.MinValue, ObjectType.ePrototype, 0, oPrototypes, Nothing)
        Dim oPPUnits As UITreeView.UITreeViewItem = tvwItems.AddNode("Units", Int32.MinValue, ObjectType.ePrototype, 0, oPrototypes, Nothing)

        'oAlloy.bItemBold = True : oAlloy.bUseFillColor = True
        oArmor.bItemBold = True : oArmor.bUseFillColor = True
        oEngine.bItemBold = True : oEngine.bUseFillColor = True
        oPrototypes.bItemBold = True : oPrototypes.bUseFillColor = True
        oRadar.bItemBold = True : oRadar.bUseFillColor = True
        oShield.bItemBold = True : oShield.bUseFillColor = True
        oWeapon.bItemBold = True : oWeapon.bUseFillColor = True
        oUnits.bItemBold = True : oUnits.bUseFillColor = True

        'Dim lAlloys As Int32 = 0
        Dim lArmor As Int32 = 0
        Dim lEngines As Int32 = 0
        Dim lPrototypes As Int32 = 0
        Dim lRadars As Int32 = 0
        Dim lShields As Int32 = 0
        Dim lWeapons As Int32 = 0

        For X As Int32 = 0 To goCurrentPlayer.mlTechUB
            Dim oTech As Base_Tech = goCurrentPlayer.moTechs(X)
            If oTech Is Nothing = False AndAlso oTech.ObjTypeID <> ObjectType.eSpecialTech AndAlso oTech.ObjTypeID <> ObjectType.eHullTech Then
                If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched AndAlso oTech.bArchived = False Then
                    Dim sValue As String = oTech.GetComponentName()

                    Dim oParent As UITreeView.UITreeViewItem = Nothing
                    Select Case oTech.ObjTypeID
                        Case ObjectType.eAlloyTech
                            'oParent = oAlloy
                            'lAlloys += 1
                            Continue For
                        Case ObjectType.eArmorTech
                            oParent = oArmor
                            lArmor += 1
                        Case ObjectType.eEngineTech
                            oParent = oEngine
                            lEngines += 1
                        Case ObjectType.ePrototype
                            oParent = oPrototypes
                            lPrototypes += 1
                            Dim oHull As HullTech = CType(goCurrentPlayer.GetTech(CType(oTech, PrototypeTech).HullID, ObjectType.eHullTech), HullTech)
                            If oHull Is Nothing = False Then

                                Dim yType As Byte = HullTech.GetHullTypeID(oHull.yTypeID, oHull.ySubTypeID)
                                If yType = HullTech.eyHullType.Facility OrElse yType = HullTech.eyHullType.SpaceStation Then
                                    oParent = oPPFacilities
                                Else : oParent = oPPUnits
                                End If
                            End If
                        Case ObjectType.eRadarTech
                            oParent = oRadar
                            lRadars += 1
                        Case ObjectType.eShieldTech
                            oParent = oShield
                            lShields += 1
                        Case ObjectType.eWeaponTech
                            oParent = oWeapon
                            lWeapons += 1
                    End Select

                    tvwItems.AddNode(sValue, oTech.ObjectID, oTech.ObjTypeID, 0, oParent, Nothing)
                End If
            End If
        Next X

        'Remove unused parents
        'If oAlloy.oFirstChild Is Nothing Then tvwItems.RemoveNode(oAlloy) Else oAlloy.sItem &= " (" & lAlloys & ")"
        If oArmor.oFirstChild Is Nothing Then tvwItems.RemoveNode(oArmor) Else oArmor.sItem &= " (" & lArmor & ")"
        If oEngine.oFirstChild Is Nothing Then tvwItems.RemoveNode(oEngine) Else oEngine.sItem &= " (" & lEngines & ")"
        If oPrototypes.oFirstChild Is Nothing Then tvwItems.RemoveNode(oPrototypes) Else oPrototypes.sItem &= " (" & lPrototypes & ")"
        If oRadar.oFirstChild Is Nothing Then tvwItems.RemoveNode(oRadar) Else oRadar.sItem &= " (" & lRadars & ")"
        If oShield.oFirstChild Is Nothing Then tvwItems.RemoveNode(oShield) Else oShield.sItem &= " (" & lShields & ")"
        If oWeapon.oFirstChild Is Nothing Then tvwItems.RemoveNode(oWeapon) Else oWeapon.sItem &= " (" & lWeapons & ")"

        '  all units able to be built at this colony
        Dim sName() As String = Nothing
        Dim lID() As Int32 = Nothing
        Dim iTypeID() As Int16 = Nothing
        Dim lUB As Int32 = -1
        For X As Int32 = 0 To glEntityDefUB
            If glEntityDefIdx(X) <> -1 AndAlso goEntityDefs(X).ObjTypeID = ObjectType.eUnitDef Then
                If goEntityDefs(X).yChassisType <> ChassisType.eSpaceBased Then
                    If (goEntityDefs(X).lExtendedFlags And 1) <> 0 Then Continue For

                    'ok, lets check the filter
                    Dim lPrototypeID As Int32 = goEntityDefs(X).PrototypeID
                    If lPrototypeID > -1 Then
                        Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lPrototypeID, ObjectType.ePrototype)
                        If oTech Is Nothing = False Then
                            If oTech.bArchived = True Then Continue For
                        End If
                    End If

                    Dim sTemp As String = goEntityDefs(X).DefName
                    lUB += 1
                    ReDim Preserve sName(lUB)
                    ReDim Preserve iTypeID(lUB)
                    ReDim Preserve lID(lUB)
                    Dim lIdx As Int32 = lUB
                    For Y As Int32 = 0 To lUB - 1
                        If sTemp.ToUpper < sName(Y).ToUpper Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                    For Y As Int32 = lUB To lIdx + 1 Step -1
                        sName(Y) = sName(Y - 1)
                        lID(Y) = lID(Y - 1)
                        iTypeID(Y) = iTypeID(Y - 1)
                    Next Y
                    sName(lIdx) = sTemp

                    lID(lIdx) = goEntityDefs(X).ObjectID
                    iTypeID(lIdx) = goEntityDefs(X).ObjTypeID
                End If
            End If
        Next X
        For X As Int32 = 0 To lUB
            tvwItems.AddNode(sName(X), lID(X), iTypeID(X), -1, oUnits, Nothing)
        Next X

    End Sub

    Private Sub FillTechList()

        Dim lOrigID As Int32 = -1
        Dim iOrigTypeID As Int16 = -1

        Dim oNode As UITreeView.UITreeViewItem = tvwItems.oSelectedNode
        If oNode Is Nothing = False Then
            lOrigID = oNode.lItemData
            iOrigTypeID = CShort(oNode.lItemData2)
        End If

        tvwItems.Clear()
        For X As Int32 = 0 To goCurrentPlayer.mlTechUB
            Dim oTech As Base_Tech = goCurrentPlayer.moTechs(X)
            If oTech Is Nothing = False Then
                If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentResearching AndAlso oTech.bArchived = False Then
                    If oTech.ObjTypeID <> ObjectType.eSpecialTech OrElse (oTech.ObjTypeID = ObjectType.eSpecialTech And CType(goCurrentPlayer.GetTech(oTech.ObjectID, oTech.ObjTypeID), SpecialTech).bInTheTank = False) Then
                        Dim sValue As String = oTech.GetComponentName

                        If sValue.Length > 20 Then sValue = sValue.Substring(0, 20)
                        sValue = sValue.PadRight(22, " "c)
                        sValue &= (Base_Tech.GetComponentTypeName(oTech.ObjTypeID)).PadRight(12, " "c)

                        If goCurrentPlayer.moTechs(X).Researchers <> 0 Then
                            sValue &= "Assigned (" & goCurrentPlayer.moTechs(X).Researchers & ")"
                        End If

                        'tvwTechs.AddItem(sValue, False)
                        'tvwTechs.ItemData(tvwTechs.NewIndex) = oTech.ObjectID : tvwTechs.ItemData2(tvwTechs.NewIndex) = oTech.ObjTypeID
                        tvwItems.AddNode(sValue, oTech.ObjectID, oTech.ObjTypeID, -1, Nothing, Nothing)
                    End If
                End If
            End If
        Next X

        'Add Minerals
        'Dim lMinSorted() As Int32 = GetSortedMineralIdxArray(True)
        'If lMinSorted Is Nothing = False Then
        '    For X As Int32 = 0 To lMinSorted.GetUpperBound(0)
        '        If goMinerals(lMinSorted(X)).bDiscovered = False Then
        '            goUILib.AddNotification("Something : " & goMinerals(lMinSorted(X)).MineralName & " (" & ObjectType.eMineralTech.ToString & ") (" & goMinerals(lMinSorted(X)).ObjTypeID.ToString & ") (" & goMinerals(lMinSorted(X)).ObjectID.ToString & ")", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '            'tvwItems.AddNode(goMinerals(lMinSorted(X)).MineralName, goMinerals(lMinSorted(X)).ObjectID, goMinerals(lMinSorted(X)).ObjTypeID, -1, Nothing, Nothing)
        '            tvwItems.AddNode(goMinerals(lMinSorted(X)).MineralName, goMinerals(lMinSorted(X)).ObjectID, ObjectType.eMineralTech, -1, Nothing, Nothing)
        '        End If
        '    Next X
        'End If

        'goUILib.AddNotification("Something : " & oTech.GetComponentName, Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

        'For X As Int32 = 0 To tvwTechs.ListCount - 1
        '    If tvwTechs.ItemData(X) = lOrigID AndAlso tvwTechs.ItemData2(X) = iOrigTypeID Then
        '        tvwTechs.ListIndex = X
        '        Exit For
        '    End If
        'Next X
        tvwItems.oSelectedNode = tvwItems.GetNodeByItemData2(lOrigID, iOrigTypeID)
    End Sub

    Private Sub FillQueueList()
        If lstQueue Is Nothing = False Then
            Dim lID As Int32 = -1
            Dim iTypeID As Int16 = -1
            If lstQueue.ListIndex > -1 Then
                lID = lstQueue.ItemData(lstQueue.ListIndex)
                iTypeID = CShort(lstQueue.ItemData2(lstQueue.ListIndex))
            End If

            mbIgnoreQueueEvents = True

            lstQueue.Clear()
            lstQueue.ListIndex = -1
            If muQueue Is Nothing = False Then
                Try
                    For X As Int32 = 0 To muQueue.GetUpperBound(0)
                        Dim oTech As Base_Tech = goCurrentPlayer.GetTech(muQueue(X).lObjectID, muQueue(X).iObjTypeID)
                        If oTech Is Nothing = False Then
                            Dim sValue As String = oTech.GetComponentName
                            If sValue.Length > 20 Then sValue = sValue.Substring(0, 20)
                            sValue = sValue.PadRight(22, " "c)
                            sValue &= (Base_Tech.GetComponentTypeName(oTech.ObjTypeID)).PadRight(12, " "c)

                            lstQueue.AddItem(sValue, False)
                            lstQueue.ItemData(lstQueue.NewIndex) = oTech.ObjectID : lstQueue.ItemData2(lstQueue.NewIndex) = oTech.ObjTypeID
                        End If
                    Next X
                Catch
                End Try
            End If

            If lID > 0 AndAlso iTypeID > 0 Then
                For X As Int32 = 0 To lstQueue.ListCount - 1
                    If lstQueue.ItemData(X) = lID AndAlso lstQueue.ItemData2(X) = iTypeID Then
                        lstQueue.ListIndex = X
                        Exit For
                    End If
                Next X
            End If

            Dim sName As String
            If myViewState = 0 Then
                sName = "Colonial Research Queue (" & lstQueue.ListCount.ToString & ")"
            Else
                sName = "Colonial Production Queue (" & lstQueue.ListCount.ToString & ")"
            End If
            If lblResQueue.Caption <> sName Then lblResQueue.Caption = sName
            mbIgnoreQueueEvents = False

        End If
    End Sub

    Public Sub HandleGetColonyResearchList(ByRef yData() As Byte)

        Dim lPos As Int32 = 2   'for msgcode
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lColonyID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lQueueCnt As Int32 = yData(lPos) : lPos += 1
        Dim uTmpQueue(lQueueCnt - 1) As uQueueItem
        For X As Int32 = 0 To lQueueCnt - 1
            With uTmpQueue(X)
                .lObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .yAssignCnt = yData(lPos) : lPos += 1
                .yQty = yData(lPos) : lPos += 1

                Dim oTech As Base_Tech = goCurrentPlayer.GetTech(.lObjectID, .iObjTypeID)
                If oTech Is Nothing = False Then
                    oTech.Researchers = .yAssignCnt
                End If
            End With
        Next X
        muQueue = uTmpQueue

        Dim lFacCnt As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim uTempFac(-1) As uProducer
        Dim lFacUB As Int32 = -1
        Dim uTempRes(-1) As uProducer
        Dim lResUB As Int32 = -1

        For X As Int32 = 0 To lFacCnt - 1

            Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim lProdFac As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yProdType As Byte = yData(lPos) : lPos += 1

            Dim uNew As uProducer
            With uNew
                .lObjectID = lID
                .iObjTypeID = iTypeID
                .lProdFactor = lProdFac
                .yProductionType = yProdType
                .lProjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iProjectTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim yProgress As Byte = yData(lPos) : lPos += 1
                .bCancelled = False

                If (yProgress And 128) <> 0 Then
                    If .yProductionType = ProductionType.eResearch Then
                        yProgress = CByte(yProgress Xor 128)
                        .bIsPrimary = True
                    Else
                        .bIsPrimary = False
                        .bCancelled = True
                    End If
                Else : .bIsPrimary = False
                End If
                .yProgress = yProgress

                .lCyclesRemaining = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lCycleStarted = glCurrentCycle

                .lProdCnt = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                .sProducerName = GetCacheObjectValue(.lObjectID, .iObjTypeID)
            End With
            Dim bPowered As Boolean = False
            For Y As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(Y) > -1 Then
                    Dim oEnt As BaseEntity = goCurrentEnvir.oEntity(Y)
                    If oEnt.ObjTypeID = iTypeID AndAlso oEnt.ObjectID = lID Then
                        If (oEnt.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                            bPowered = True
                        End If
                        Exit For
                    End If
                End If
            Next Y
            If bPowered = True Then
                If yProdType = ProductionType.eResearch Then
                    lResUB += 1
                    ReDim Preserve uTempRes(lResUB)
                    uTempRes(lResUB) = uNew
                ElseIf (yProdType And ProductionType.eProduction) <> 0 Then
                    lFacUB += 1
                    ReDim Preserve uTempFac(lFacUB)
                    uTempFac(lFacUB) = uNew
                End If
            End If
        Next X


        Dim lProdFactor As Int32
        Dim sProducerName As String
        Dim sProducerNameCompare As String
        Dim lObjectID As Int32
        'muResearchers = uTempRes
        'Re-Sort Researchers
        Dim lSortedRes() As uProducer = Nothing
        Dim lSortedResUB As Int32 = -1
        For X As Int32 = 0 To lResUB
            Dim lIdx As Int32 = -1

            lProdFactor = uTempRes(X).lProdFactor
            sProducerName = GetAlphaNumericSortable(uTempRes(X).sProducerName.ToUpper)
            lObjectID = uTempRes(X).lObjectID
            For Y As Int32 = 0 To lSortedResUB
                sProducerNameCompare = GetAlphaNumericSortable(lSortedRes(Y).sProducerName.ToUpper)
                If lSortedRes(Y).lProdFactor < lProdFactor OrElse (lSortedRes(Y).lProdFactor = lProdFactor And sProducerNameCompare > sProducerName) OrElse (lSortedRes(Y).lProdFactor = lProdFactor And sProducerNameCompare = sProducerName And lSortedRes(Y).lObjectID > lObjectID) Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedResUB += 1
            ReDim Preserve lSortedRes(lSortedResUB)
            If lIdx = -1 Then
                lSortedRes(lSortedResUB) = uTempRes(X)
            Else
                For Y As Int32 = lSortedResUB To lIdx + 1 Step -1
                    lSortedRes(Y) = lSortedRes(Y - 1)
                Next Y
                lSortedRes(lIdx) = uTempRes(X)
            End If
        Next X
        ''
        muResearchers = lSortedRes

        'muFactories = uTempFac
        'Re-Sort Factories
        Dim lSortedFac() As uProducer = Nothing
        Dim lSortedFacUB As Int32 = -1
        For X As Int32 = 0 To lFacUB
            Dim lIdx As Int32 = -1

            lProdFactor = uTempFac(X).lProdFactor
            sProducerName = GetAlphaNumericSortable(uTempFac(X).sProducerName.ToUpper)
            lObjectID = uTempFac(X).lObjectID
            For Y As Int32 = 0 To lSortedFacUB
                sProducerNameCompare = GetAlphaNumericSortable(lSortedFac(Y).sProducerName.ToUpper)
                If lSortedFac(Y).lProdFactor < lProdFactor OrElse (lSortedFac(Y).lProdFactor = lProdFactor And sProducerNameCompare > sProducerName) OrElse (lSortedFac(Y).lProdFactor = lProdFactor And sProducerNameCompare = sProducerName And lSortedFac(Y).lObjectID > lObjectID) Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedFacUB += 1
            ReDim Preserve lSortedFac(lSortedFacUB)
            If lIdx = -1 Then
                lSortedFac(lSortedFacUB) = uTempFac(X)
            Else
                For Y As Int32 = lSortedFacUB To lIdx + 1 Step -1
                    lSortedFac(Y) = lSortedFac(Y - 1)
                Next Y
                lSortedFac(lIdx) = uTempFac(X)
            End If
        Next X
        ''
        muFactories = lSortedFac

        'Add a count to the label
        Dim sName As String = ""
        If myViewState = 0 Then
            sName = "Research Facilities (" & (lResUB + 1).ToString & ")"
            If lblResFacs.Caption <> sName Then lblResFacs.Caption = sName
        Else
            sName = "Production Facilities (" & (lFacUB + 1).ToString & ")"
            If lblResFacs.Caption <> sName Then lblResFacs.Caption = sName
        End If
        FillQueueList()
    End Sub

    Private Sub frmColonyResearch_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        'If (myViewState = 0 AndAlso muResearchers Is Nothing) OrElse (myViewState = 1 AndAlso muFactories Is Nothing) Then Return

        Dim lY As Int32 = lMouseY - Me.GetAbsolutePosition.Y
        If lY > vscrResLabs.Top Then
            'ok, determine what we are over...
            lY -= vscrResLabs.Top
            lY -= 3
            Dim lIdx As Int32 = lY \ ml_ITEM_HEIGHT

            Dim lX As Int32 = lMouseX - Me.GetAbsolutePosition.X
            lIdx += vscrResLabs.Value

            If lIdx > -1 AndAlso ((myViewState = 0 AndAlso muResearchers Is Nothing = False AndAlso lIdx <= muResearchers.GetUpperBound(0)) OrElse (myViewState = 1 AndAlso muFactories Is Nothing = False AndAlso lIdx <= muFactories.GetUpperBound(0))) Then
                Dim uSel As uProducer

                If myViewState = 0 Then uSel = muResearchers(lIdx) Else uSel = muFactories(lIdx)

                With uSel
                    'Facility Name (3,3,180,18)
                    'Project Name (185,3,180,18)
                    'Prod Factor (365,3,35,18)
                    'Progress (410,7,64,12)
                    'Cancel Button (480,5,16,16)
                    If lX > 3 AndAlso lX < 185 Then
                        If goCamera Is Nothing = False Then goCamera.TrackingIndex = -1
                        'MyBase.moUILib.ClearSelection()
                        Dim oEnvir As BaseEnvironment = goCurrentEnvir
                        oEnvir.DeselectAll()

                        Dim lID As Int32 = .lObjectID
                        Dim iTypeID As Int16 = .iObjTypeID

                        If lID < 1 OrElse iTypeID < 1 Then Return


                        If oEnvir Is Nothing = False Then
                            Dim lCurUB As Int32 = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
                            For X As Int32 = 0 To lCurUB
                                If oEnvir.lEntityIdx(X) = lID Then
                                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                                    If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iTypeID Then
                                        oEntity.bSelected = True
                                        goCamera.SimplyPlaceCamera(CInt(oEntity.LocX), CInt(oEntity.LocY), CInt(oEntity.LocZ)) 'goCamera.TrackingIndex = X
                                        Exit For
                                    End If
                                End If
                            Next X
                        End If
                    ElseIf lX > 460 AndAlso lX < 480 Then
                        If myViewState = 1 Then
                            .bCancelled = Not .bCancelled
                            Dim yData(17) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
                            System.BitConverter.GetBytes(.lObjectID).CopyTo(yData, 2)
                            System.BitConverter.GetBytes(.iObjTypeID).CopyTo(yData, 6)
                            System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yData, 8)
                            System.BitConverter.GetBytes(-1S).CopyTo(yData, 12)
                            System.BitConverter.GetBytes(-1I).CopyTo(yData, 14)

                            If goSound Is Nothing = False Then
                                goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                            End If
                            If .bCancelled = True Then
                                MyBase.moUILib.AddNotification("Request sent for facility to no longer be part of colony queue.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Else
                                MyBase.moUILib.AddNotification("Request sent for facility to be part of colony queue.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            End If

                            MyBase.moUILib.SendMsgToPrimary(yData)
                        Else
                            If .lProjectID > 0 AndAlso .iProjectTypeID > 0 AndAlso .bCancelled = False Then
                                .bCancelled = True
                                Dim yData(17) As Byte
                                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
                                System.BitConverter.GetBytes(.lObjectID).CopyTo(yData, 2)
                                System.BitConverter.GetBytes(.iObjTypeID).CopyTo(yData, 6)
                                System.BitConverter.GetBytes(-1I).CopyTo(yData, 8)
                                System.BitConverter.GetBytes(-1S).CopyTo(yData, 12)
                                System.BitConverter.GetBytes(-1I).CopyTo(yData, 14)
                                MyBase.moUILib.SendMsgToPrimary(yData)
                                Me.IsDirty = True
                            End If
                        End If
                    End If
                End With
            End If
        Else
            Dim lTmpX As Int32 = lMouseX - Me.GetAbsolutePosition.X
            If rcResearch.Contains(lTmpX, lY) = True Then
                ViewState = 0
            ElseIf rcProduction.Contains(lTmpX, lY) = True Then
                ViewState = 1
            End If
        End If
    End Sub

    Private Sub frmColonyResearch_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim lX As Int32 = lMouseX - Me.GetAbsolutePosition.X
        Dim lY As Int32 = lMouseY - Me.GetAbsolutePosition.Y

        If vscrResLabs Is Nothing Then Return

        If lY < vscrResLabs.Top Then Return
        goUILib.SetToolTip(False)

        If lY > vscrResLabs.Top Then
            'ok, determine what we are over...
            lY -= vscrResLabs.Top
            lY -= 3
            Dim lIdx As Int32 = lY \ ml_ITEM_HEIGHT

            lIdx += vscrResLabs.Value
            If myViewState = 0 Then
                If muResearchers Is Nothing OrElse lIdx < 0 OrElse lIdx > muResearchers.GetUpperBound(0) Then Return
            ElseIf myViewState = 1 Then
                If muFactories Is Nothing OrElse lIdx < 0 OrElse lIdx > muFactories.GetUpperBound(0) Then Return
            End If
        End If

        Dim sTemp As String = ""
        If lX > 3 AndAlso lX < 175 Then
            If ViewState = 0 Then
                sTemp = "The research facility name within this colony" & vbCrLf & "that this line item refers to"
            Else
                sTemp = "The production facility name within this colony" & vbCrLf & "that this line item refers to"
            End If
        ElseIf lX > 175 AndAlso lX < 345 Then
            If ViewState = 0 Then
                sTemp = "The project that the facility is currently researching." & vbCrLf & "The project name will be bolded if the facility is the primary researcher."
            Else
                sTemp = "The project that the facility is currently producing."
            End If
        ElseIf lX > 345 AndAlso lX < 390 Then
            If ViewState = 0 Then
                sTemp = "The facility's base research production factor."
            Else
                sTemp = "The facility's base production factor."
            End If
        ElseIf lX > 390 AndAlso lX < 460 Then
            If ViewState = 0 Then
                sTemp = "The progress of the current project that this facility" & vbCrLf & "is researching. If this facility is not researching, this area is empty."
            Else
                sTemp = "The progress of the current project that this facility" & vbCrLf & "is producing. If this facility is not producing, this area is empty."
            End If
        ElseIf lX > 460 AndAlso lX < 480 Then
            If myViewState = 0 Then
                sTemp = "Click to remove the current project from this facility."
            Else
                sTemp = "Toggles whether this facility accepts orders from the colony production queue."
            End If
        End If
            If sTemp <> "" Then goUILib.SetToolTip(sTemp, lMouseX, lMouseY)
    End Sub

    Private mlForceNextFramesDirty As Int32 = 0
    Private Sub frmColonyResearch_OnNewFrame() Handles Me.OnNewFrame
        If mlForceNextFramesDirty <> 0 Then
            Me.IsDirty = True
            mlForceNextFramesDirty -= 1
        End If
        If myViewState = 0 Then
            If muResearchers Is Nothing = False Then
                For X As Int32 = 0 To muResearchers.GetUpperBound(0)
                    Dim sTemp As String = GetCacheObjectValue(muResearchers(X).lObjectID, muResearchers(X).iObjTypeID)
                    If muResearchers(X).sProducerName <> sTemp Then
                        muResearchers(X).sProducerName = sTemp
                        Me.IsDirty = True
                    End If
                Next X
            End If
        Else
            If muFactories Is Nothing = False Then
                For X As Int32 = 0 To muFactories.GetUpperBound(0)
                    Dim sTemp As String = GetCacheObjectValue(muFactories(X).lObjectID, muFactories(X).iObjTypeID)
                    If muFactories(X).sProducerName <> sTemp Then
                        muFactories(X).sProducerName = sTemp
                        Me.IsDirty = True
                    End If
                Next X
            End If
        End If


        Dim lTechCnt As Int32 = 0
        If myViewState = 0 Then
            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentResearching Then
                    Dim oTech As Base_Tech = goCurrentPlayer.moTechs(X)
                    If oTech.ObjTypeID <> ObjectType.eSpecialTech OrElse (oTech.ObjTypeID = ObjectType.eSpecialTech And CType(goCurrentPlayer.GetTech(oTech.ObjectID, oTech.ObjTypeID), SpecialTech).bInTheTank = False) Then
                        lTechCnt += 1
                    End If
                End If
            Next X
            If lTechCnt <> tvwItems.TotalNodeCount Then FillTechList()
        End If


        If glCurrentCycle - mlLastRequest > 300 Then
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing Then Return

            Dim lID As Int32 = oEnvir.ObjectID
            Dim iTypeID As Int16 = oEnvir.ObjTypeID

            If oEnvir.ObjTypeID = ObjectType.eSolarSystem Then

                Dim bGood As Boolean = False
                Try
                    For X As Int32 = 0 To oEnvir.lEntityUB
                        If oEnvir.lEntityIdx(X) > -1 Then
                            Dim oEnt As BaseEntity = oEnvir.oEntity(X)
                            If oEnt.bSelected = True Then
                                If oEnt.yProductionType = ProductionType.eSpaceStationSpecial Then
                                    bGood = True
                                    lID = oEnt.ObjectID
                                    iTypeID = oEnt.ObjTypeID
                                    Exit For
                                End If
                            End If
                        End If
                    Next X
                Catch
                End Try
                If bGood = False Then
                    btnClose_Click("")
                    Return
                End If
            End If

            mlLastRequest = glCurrentCycle

            Dim yMsg(7) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eGetColonyResearchList).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(lID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, 6)
            MyBase.moUILib.SendMsgToPrimary(yMsg)
        End If

    End Sub

    Private Sub frmColonyResearch_OnRender() Handles Me.OnRender
        Dim lItemCnt As Int32 = vscrResLabs.Height \ ml_ITEM_HEIGHT

        Dim uTmp() As uProducer = Nothing
        If myViewState = 0 Then
            uTmp = muResearchers
        Else : uTmp = muFactories
        End If

        If uTmp Is Nothing = False Then
            Dim lScrlMax As Int32 = Math.Max(0, uTmp.Length - lItemCnt)


            If vscrResLabs.MaxValue <> lScrlMax Then
                vscrResLabs.MaxValue = lScrlMax
                If vscrResLabs.Value > vscrResLabs.MaxValue Then vscrResLabs.Value = vscrResLabs.MaxValue
            End If
            If lScrlMax = 0 Then
                If vscrResLabs.Enabled <> False Then vscrResLabs.Enabled = False
            ElseIf vscrResLabs.Enabled <> True Then
                vscrResLabs.Enabled = True
            End If



            Dim rcFacName As Rectangle = New Rectangle(3, 3, 170, 18)
            Dim rcProjName As Rectangle = New Rectangle(175, 3, 170, 18)
            Dim rcProdFactor As Rectangle = New Rectangle(345, 3, 35, 18)
            Dim rcProgress As Rectangle = New Rectangle(390, 7, 64, 9)
            Dim rcCancel As Rectangle = New Rectangle(460, 5, 17, 17)

            If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

            Using oFillSprite As New Sprite(MyBase.moUILib.oDevice)
                oFillSprite.Begin(SpriteFlags.AlphaBlend)
                Dim oSelColor As System.Drawing.Color
                oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)
                If ViewState = 0 Then
                    DoMultiColorFill(rcResearch, oSelColor, rcResearch.Location, oFillSprite)
                Else
                    DoMultiColorFill(rcProduction, oSelColor, rcProduction.Location, oFillSprite)
                End If
                oFillSprite.End()
                oFillSprite.Dispose()
            End Using


            Using oFont As Font = New Font(MyBase.moUILib.oDevice, moSysFont)
                Using oBoldFont As Font = New Font(MyBase.moUILib.oDevice, moSysFontBold)
                    Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                        oTextSpr.Begin(SpriteFlags.AlphaBlend)

                        rcFacName.Y += vscrResLabs.Top
                        rcProjName.Y += vscrResLabs.Top
                        rcProdFactor.Y += vscrResLabs.Top

                        'Now, render our items...
                        For X As Int32 = vscrResLabs.Value To Math.Min(vscrResLabs.Value + lItemCnt, uTmp.GetUpperBound(0))
                            With uTmp(X)
                                oFont.DrawText(oTextSpr, .sProducerName, rcFacName, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)

                                Dim sTechName As String = "Unassigned"

                                Dim clrProj As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                                If .lProjectID > 0 AndAlso .iProjectTypeID > 0 Then
                                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(.lProjectID, .iProjectTypeID)
                                    If oTech Is Nothing = False Then
                                        sTechName = oTech.GetComponentName()
                                        clrProj = muSettings.InterfaceBorderColor
                                    End If
                                End If

                                If .bIsPrimary = True Then
                                    oBoldFont.DrawText(oTextSpr, sTechName, rcProjName, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, clrProj)
                                Else
                                    oFont.DrawText(oTextSpr, sTechName, rcProjName, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, clrProj)
                                End If

                                oFont.DrawText(oTextSpr, .lProdFactor.ToString, rcProdFactor, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                            End With
                            rcFacName.Y += ml_ITEM_HEIGHT
                            rcProjName.Y += ml_ITEM_HEIGHT
                            rcProdFactor.Y += ml_ITEM_HEIGHT
                        Next X

                        MyBase.RenderRoundedBorder(rcResearch, 1, muSettings.InterfaceBorderColor)
                        MyBase.RenderRoundedBorder(rcProduction, 1, muSettings.InterfaceBorderColor)

                        Try
                            oBoldFont.DrawText(oTextSpr, "Research", rcResearch, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                            oBoldFont.DrawText(oTextSpr, "Production", rcProduction, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                        Catch
                        End Try




                        oTextSpr.End()
                        oTextSpr.Dispose()
                    End Using
                    oBoldFont.Dispose()
                End Using
                oFont.Dispose()
            End Using

            'Now, render our icons 
            Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                oSprite.Begin(SpriteFlags.AlphaBlend)

                rcCancel.Y += vscrResLabs.Top
                rcProgress.Y += vscrResLabs.Top

                Dim rcCancelButtonSrc As Rectangle = New Rectangle(127, 143, 17, 17)
                For X As Int32 = vscrResLabs.Value To Math.Min(vscrResLabs.Value + lItemCnt, uTmp.GetUpperBound(0))
                    With uTmp(X)
                        If .lProjectID > 0 AndAlso .iProjectTypeID > 0 Then
                            Dim lTemp As Int32 = CInt(.yProgress) \ 10
                            Dim rcSrc As Rectangle = grc_UI(elInterfaceRectangle.eMinBar_0 + lTemp) 'Rectangle.FromLTRB(193, 157 + (lTemp * 9), 256, 157 + ((lTemp + 1) * 9))

                            Dim clrVal As System.Drawing.Color
                            If lTemp < 4 Then
                                clrVal = Color.Red
                            ElseIf lTemp < 7 Then
                                clrVal = Color.Yellow
                            Else : clrVal = Color.Green
                            End If

                            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcProgress, Point.Empty, 0, rcProgress.Location, clrVal)

                            If .bCancelled = True Then clrVal = Color.Red Else clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcCancelButtonSrc, rcCancel, Point.Empty, 0, rcCancel.Location, clrVal)
                        ElseIf myViewState = 1 Then
                            Dim clrVal As System.Drawing.Color
                            If .bCancelled = True Then clrVal = Color.Red Else clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcCancelButtonSrc, rcCancel, Point.Empty, 0, rcCancel.Location, clrVal)
                        End If
                        rcCancel.Y += ml_ITEM_HEIGHT
                        rcProgress.Y += ml_ITEM_HEIGHT
                    End With
                Next X

                oSprite.End()
                oSprite.Dispose()
            End Using
        Else
            Using oFillSprite As New Sprite(MyBase.moUILib.oDevice)
                oFillSprite.Begin(SpriteFlags.AlphaBlend)
                Dim oSelColor As System.Drawing.Color
                oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)
                If ViewState = 0 Then
                    DoMultiColorFill(rcResearch, oSelColor, rcResearch.Location, oFillSprite)
                Else
                    DoMultiColorFill(rcProduction, oSelColor, rcProduction.Location, oFillSprite)
                End If
                oFillSprite.End()
                oFillSprite.Dispose()
            End Using

            Using oBoldFont As Font = New Font(MyBase.moUILib.oDevice, moSysFontBold)
                Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                    MyBase.RenderRoundedBorder(rcResearch, 1, muSettings.InterfaceBorderColor)
                    MyBase.RenderRoundedBorder(rcProduction, 1, muSettings.InterfaceBorderColor)

                    oTextSpr.Begin(SpriteFlags.AlphaBlend)
                    Try
                        oBoldFont.DrawText(oTextSpr, "Research", rcResearch, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                        oBoldFont.DrawText(oTextSpr, "Production", rcProduction, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                    Catch
                    End Try
                    oTextSpr.End()
                    oTextSpr.Dispose()
                End Using
            End Using

        End If

    End Sub

    Private Sub DoMultiColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point, ByRef oSpr As Sprite)
        Dim rcSrc As Rectangle

        Dim fX As Single
        Dim fY As Single

        If rcDest.Width = 0 OrElse rcDest.Height = 0 Then Exit Sub

        rcSrc.Location = New Point(192, 0)
        rcSrc.Width = 62
        rcSrc.Height = 64

        'Now, draw it...
        With oSpr
            fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
            fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))
            .Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFill)
        End With
    End Sub

    Private Sub lstQueue_ItemClick(ByVal lIndex As Integer) Handles lstQueue.ItemClick
        If mbIgnoreQueueEvents = True Then Return
        If lIndex < 0 Then Return
        If muQueue Is Nothing Then Return

        Dim lID As Int32 = lstQueue.ItemData(lIndex)
        Dim lTypeID As Int32 = lstQueue.ItemData2(lIndex)

        Try
            For X As Int32 = 0 To muQueue.GetUpperBound(0)
                If muQueue(X).lObjectID = lID AndAlso muQueue(X).iObjTypeID = lTypeID Then
                    txtAssign.Caption = CStr(muQueue(X).yAssignCnt)
                    Exit For
                End If
            Next X
        Catch
        End Try

    End Sub

    Private Sub btnAddToQueue_Click(ByVal sName As String) Handles btnAddToQueue.Click
        Dim oNode As UITreeView.UITreeViewItem = tvwItems.oSelectedNode
        If oNode Is Nothing = False Then
            Dim lTechID As Int32 = oNode.lItemData
            Dim iTechTypeID As Int16 = CShort(oNode.lItemData2)

            Dim bInList As Boolean = False
            For X As Int32 = 0 To lstQueue.ListCount - 1
                If lstQueue.ItemData(X) = lTechID AndAlso lstQueue.ItemData2(X) = iTechTypeID Then
                    bInList = True
                    Exit For
                End If
            Next X
            If bInList = False AndAlso lstQueue.ListCount > 9 Then
                MyBase.moUILib.AddNotification("There is a maximum of 10 items in the research queue.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then
                    goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                End If
                Return
            End If
            If iTechTypeID = ObjectType.eSpecialTech Then
                Dim oTech As SpecialTech = CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), SpecialTech)
                If oTech Is Nothing OrElse oTech.bInTheTank = True Then
                    MyBase.moUILib.AddNotification("Unable to place bypassed special projects into the queue.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then
                        goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    End If
                    Return
                End If
            Else
                Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lTechID, iTechTypeID)
                If oTech Is Nothing = False AndAlso oTech.MajorDesignFlaw = 31 Then
                    MyBase.moUILib.AddNotification("That component cannot be researched or produced.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then
                        goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    End If
                    Return
                End If
            End If

            Dim yMsg(16) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetColonyResearchQueue).CopyTo(yMsg, lPos) : lPos += 2

            Dim lID As Int32 = goCurrentEnvir.ObjectID
            Dim iTypeID As Int16 = goCurrentEnvir.ObjTypeID

            If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                Dim bGood As Boolean = False
                Try
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) > -1 Then
                            Dim oEnt As BaseEntity = goCurrentEnvir.oEntity(X)
                            If oEnt.bSelected = True Then
                                If oEnt.yProductionType = ProductionType.eSpaceStationSpecial Then
                                    bGood = True
                                    lID = oEnt.ObjectID
                                    iTypeID = oEnt.ObjTypeID
                                    Exit For
                                End If
                            End If
                        End If
                    Next X
                Catch
                End Try
                If bGood = False Then Return
            End If
            'goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2

            yMsg(lPos) = 0 : lPos += 1      '0 for add
            System.BitConverter.GetBytes(lTechID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iTechTypeID).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = CByte(Val(txtAssign.Caption)) : lPos += 1
            If myViewState = 0 Then yMsg(lPos) = 1 Else yMsg(lPos) = CByte(Val(txtQuantity.Caption))
            lPos += 1

            MyBase.moUILib.SendMsgToPrimary(yMsg)
        Else
            MyBase.moUILib.AddNotification("Select a tech to add to the queue first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnRemove_Click(ByVal sName As String) Handles btnRemove.Click
        If lstQueue.ListIndex > -1 Then
            Dim lTechID As Int32 = lstQueue.ItemData(lstQueue.ListIndex)
            Dim iTechTypeID As Int16 = CShort(lstQueue.ItemData2(lstQueue.ListIndex))

            Dim yMsg(15) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetColonyResearchQueue).CopyTo(yMsg, lPos) : lPos += 2

            Dim lID As Int32 = goCurrentEnvir.ObjectID
            Dim iTypeID As Int16 = goCurrentEnvir.ObjTypeID

            If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                Dim bGood As Boolean = False
                Try
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) > -1 Then
                            Dim oEnt As BaseEntity = goCurrentEnvir.oEntity(X)
                            If oEnt.bSelected = True Then
                                If oEnt.yProductionType = ProductionType.eSpaceStationSpecial Then
                                    bGood = True
                                    lID = oEnt.ObjectID
                                    iTypeID = oEnt.ObjTypeID
                                    Exit For
                                End If
                            End If
                        End If
                    Next X
                Catch
                End Try
                If bGood = False Then Return
            End If
            'goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2

            yMsg(lPos) = 1 : lPos += 1      '1 for remove
            System.BitConverter.GetBytes(lTechID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iTechTypeID).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = 0 : lPos += 1

            MyBase.moUILib.SendMsgToPrimary(yMsg)
        Else
            MyBase.moUILib.AddNotification("Select a queue item to remove first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub tvwTechs_NodeSelected(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwItems.NodeSelected
        Dim sTemp As String = ""

        If oNode Is Nothing = False Then
            Dim lID As Int32 = oNode.lItemData
            Dim iTypeID As Int16 = CShort(oNode.lItemData2)
            If myViewState = 0 Then
                Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
                If oTech Is Nothing = False Then
                    If oTech.oResearchCost Is Nothing = False Then
                        sTemp = oTech.oResearchCost.GetBuildCostText("")
                    End If
                End If
                oTech = Nothing
            Else
                If iTypeID = ObjectType.eUnitDef OrElse iTypeID = ObjectType.eFacilityDef Then

                    For X As Int32 = 0 To glEntityDefUB
                        If glEntityDefIdx(X) = lID AndAlso goEntityDefs(X).ObjTypeID = iTypeID Then
                            Dim sPower As String = ""
                            If iTypeID = ObjectType.eFacilityDef Then
                                sPower = "Power: " & goEntityDefs(X).PowerFactor & vbCrLf
                                'If goEntityDefs(X).ProductionTypeID = ProductionType.ePowerCenter Then
                                '    sPower = "Power: (" & goEntityDefs(X).PowerFactor & ")" & vbCrLf
                                'Else : sPower = "Power: " & goEntityDefs(X).PowerFactor & vbCrLf
                                'End If
                                sPower &= "Jobs: " & goEntityDefs(X).WorkerFactor & vbCrLf

                                If goEntityDefs(X).ProductionTypeID = ProductionType.eColonists Then
                                    sPower &= "Housing: " & goEntityDefs(X).ProdFactor & vbCrLf
                                ElseIf goEntityDefs(X).ProductionTypeID = ProductionType.eResearch Then
                                    sPower &= "Research Ability: " & goEntityDefs(X).ProdFactor & vbCrLf
                                ElseIf (goEntityDefs(X).ProductionTypeID And ProductionType.eProduction) <> 0 Then
                                    sPower &= "Production: " & goEntityDefs(X).ProdFactor & vbCrLf
                                End If
                            End If

                            sTemp = goEntityDefs(X).ProductionCost.GetBuildCostText(sPower)
                            Exit For
                        End If
                    Next X

                Else
                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
                    If oTech Is Nothing = False Then
                        If oTech.oProductionCost Is Nothing = False Then
                            sTemp = oTech.oProductionCost.GetBuildCostText("")
                        End If
                    End If
                    oTech = Nothing
                End If
            End If
        End If

        txtCost.Caption = sTemp
    End Sub

    Private Sub frmColonyResearch_OnResize() Handles Me.OnResize
        If mbLoading = True Then Return
        vscrResLabs.Height = Me.Height - vscrResLabs.Top - 10
        muSettings.ColonialManagementHeight = Me.Height
    End Sub

    Private Sub frmColonyResearch_WindowMoved() Handles Me.WindowMoved
        If mbLoading = True Then Return
        muSettings.ColonialManagementLeft = Me.Left
        muSettings.ColonialManagementTop = Me.Top
        muSettings.ColonialManagementHeight = Me.Height
    End Sub
End Class