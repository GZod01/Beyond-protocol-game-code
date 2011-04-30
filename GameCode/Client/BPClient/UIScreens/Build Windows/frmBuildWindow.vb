Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

#Region "Old Code"
''Build Screen
'Public Class frmBuildWindow
'    Inherits UIWindow

'    Private Const ml_REQUEST_DELAY As Int32 = 3000

'    Private WithEvents lstBuildable As UIListBox
'    Private lblBuildable As UILabel
'    Private txtItemDetails As UITextBox
'    Private lblDetails As UILabel
'    Private WithEvents lstQueue As UIListBox
'    Private lblBuildQueue As UILabel
'    Private lblAvailResources As UILabel
'    Private txtAvailResources As UITextBox
'    Private WithEvents btnAddToQueue As UIButton
'    Private WithEvents btnRemoveFromQueue As UIButton
'    Private WithEvents btnRemoveQueueItem As UIButton
'    Private WithEvents btnClearQueue As UIButton
'    Private WithEvents btnRefresh As UIButton

'    Private WithEvents txtQuantity As UITextBox
'    Private lblQty As UILabel

'    Private mlEntityIndex As Int32

'    Private mbUseChild As Boolean = False
'    Private mlChildID As Int32 = -1
'    Private miChildTypeID As Int16 = -1

'    Private mbLoading As Boolean = True

'    Private mlLastAvailResUpdate As Int32 = -1

'    Private msw_Delay As Stopwatch

'    Private mbUnitBuild As Boolean

'    Public Sub New(ByRef oUILib As UILib, ByVal bUnitBuild As Boolean)
'        MyBase.New(oUILib)

'        mbUnitBuild = bUnitBuild

'        'frmFacilityBuild initial props
'        With Me
'            .ControlName = "frmBuildWindow"
'            .Left = 364
'            .Top = 95
'            .Width = 289
'            .Height = 388
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            .BorderColor = muSettings.InterfaceBorderColor
'            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = False
'            .mbAcceptReprocessEvents = True
'        End With
'        Debug.Write(Me.ControlName & " Newed" & vbCrLf)
'        'lstBuildable initial props
'        lstBuildable = New UIListBox(oUILib)
'        With lstBuildable
'            .ControlName = "lstBuildable"
'            .Left = 3
'            .Top = 20
'            .Width = 159
'            .Height = 100
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'            '.FillColor = System.Drawing.Color.FromArgb(-16760704)
'            .FillColor = muSettings.InterfaceFillColor
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            '.HighlightColor = System.Drawing.Color.FromArgb(-16744193)
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(lstBuildable, UIControl))

'        'lblBuildable initial props
'        lblBuildable = New UILabel(oUILib)
'        With lblBuildable
'            .ControlName = "lblBuildable"
'            .Left = 3
'            .Top = 0
'            .Width = 159
'            .Height = 20
'            .Enabled = True
'            .Visible = True
'            .Caption = "Buildable Items"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblBuildable, UIControl))

'        'txtItemDetails initial props
'        txtItemDetails = New UITextBox(oUILib)
'        With txtItemDetails
'            .ControlName = "txtItemDetails"
'            .Left = 165
'            .Top = 20
'            .Width = 120 '151 
'            .Height = 100 '340   
'            If bUnitBuild Then .Height = 100 Else .Height = 220
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Arial", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = 0
'            .BackColorEnabled = muSettings.InterfaceFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
'            .MaxLength = 0
'            .BorderColor = muSettings.InterfaceBorderColor
'            .Locked = True
'            .MultiLine = True
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(txtItemDetails, UIControl))

'        'lblDetails initial props
'        lblDetails = New UILabel(oUILib)
'        With lblDetails
'            .ControlName = "lblDetails"
'            .Left = 165
'            .Top = 0
'            .Width = 159
'            .Height = 20
'            .Enabled = True
'            .Visible = True
'            .Caption = "Item Details"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblDetails, UIControl))

'        If bUnitBuild = False Then
'            'btnRefresh initial props
'            btnRefresh = New UIButton(oUILib)
'            With btnRefresh
'                .ControlName = "btnRefresh"
'                .Left = 187
'                .Top = 266
'                .Width = 100
'                .Height = 18
'                .Enabled = True
'                .Visible = True
'                .Caption = "Refresh"
'                .ForeColor = muSettings.InterfaceBorderColor
'                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'                .DrawBackImage = True
'                .FontFormat = CType(5, DrawTextFormat)
'                .ControlImageRect = New Rectangle(0, 0, 120, 32)
'            End With
'            Me.AddChild(CType(btnRefresh, UIControl))

'            'lstQueue initial props
'            lstQueue = New UIListBox(oUILib)
'            With lstQueue
'                .ControlName = "lstQueue"
'                .Left = 3
'                .Top = 140
'                .Width = 159
'                .Height = 100
'                .Enabled = True
'                .Visible = True
'                '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'                .BorderColor = muSettings.InterfaceBorderColor
'                '.FillColor = System.Drawing.Color.FromArgb(-16760704)
'                .FillColor = muSettings.InterfaceFillColor
'                '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'                .ForeColor = muSettings.InterfaceBorderColor
'                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'                '.HighlightColor = System.Drawing.Color.FromArgb(-16744193)
'                .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'                .mbAcceptReprocessEvents = True
'            End With
'            Me.AddChild(CType(lstQueue, UIControl))

'            'lblBuildQueue initial props
'            lblBuildQueue = New UILabel(oUILib)
'            With lblBuildQueue
'                .ControlName = "lblBuildQueue"
'                .Left = 3
'                .Top = 120
'                .Width = 159
'                .Height = 20
'                .Enabled = True
'                .Visible = True
'                .Caption = "Build Queue"
'                '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'                .ForeColor = muSettings.InterfaceBorderColor
'                .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
'                .DrawBackImage = False
'                .FontFormat = DrawTextFormat.VerticalCenter
'            End With
'            Me.AddChild(CType(lblBuildQueue, UIControl))

'            'lblAvailResources initial props
'            lblAvailResources = New UILabel(oUILib)
'            With lblAvailResources
'                .ControlName = "lblAvailResources"
'                .Left = 3
'                .Top = 264
'                .Width = 159
'                .Height = 20
'                .Enabled = True
'                .Visible = True
'                .Caption = "Available Resources"
'                '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'                .ForeColor = muSettings.InterfaceTextBoxForeColor
'                .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
'                .DrawBackImage = False
'                .FontFormat = DrawTextFormat.VerticalCenter
'            End With
'            Me.AddChild(CType(lblAvailResources, UIControl))

'            'txtAvailResources initial props
'            txtAvailResources = New UITextBox(oUILib)
'            With txtAvailResources
'                .ControlName = "txtAvailResources"
'                .Left = 3
'                .Top = 284
'                .Width = 282 '313 
'                .Height = 100
'                .Enabled = True
'                .Visible = True
'                .Caption = ""
'                .ForeColor = muSettings.InterfaceBorderColor
'                .SetFont(New System.Drawing.Font("Arial", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
'                .DrawBackImage = False
'                .FontFormat = 0
'                .BackColorEnabled = muSettings.InterfaceFillColor
'                .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
'                .MaxLength = 0
'                .BorderColor = muSettings.InterfaceBorderColor
'                .Locked = True
'                .MultiLine = True
'                .mbAcceptReprocessEvents = True
'            End With
'            Me.AddChild(CType(txtAvailResources, UIControl))

'            'btnMoveUp initial props
'            btnAddToQueue = New UIButton(oUILib)
'            With btnAddToQueue
'                .ControlName = "btnAddToQueue"
'                .Left = 204 ' 84
'                .Top = 241
'                .Width = 20
'                .Height = 20
'                .Enabled = True
'                .Visible = True
'                .Caption = "+"
'                .ForeColor = System.Drawing.Color.FromArgb(-1)
'                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
'                .DrawBackImage = True
'                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
'                .ToolTipText = "Adds Quantity of the current selection to the build queue."
'            End With
'            Me.AddChild(CType(btnAddToQueue, UIControl))

'            'btnMoveDown initial props
'            btnRemoveFromQueue = New UIButton(oUILib)
'            With btnRemoveFromQueue
'                .ControlName = "btnMoveDown"
'                .Left = 224 ' 104
'                .Top = 241
'                .Width = 20
'                .Height = 20
'                .Enabled = True
'                .Visible = True
'                .Caption = "-"
'                .ForeColor = System.Drawing.Color.FromArgb(-1)
'                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
'                .DrawBackImage = True
'                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
'                .ToolTipText = "Removes Quantity of the current selection from the build queue."
'            End With
'            Me.AddChild(CType(btnRemoveFromQueue, UIControl))

'            'btnRemove initial props
'            btnRemoveQueueItem = New UIButton(oUILib)
'            With btnRemoveQueueItem
'                .ControlName = "btnRemove"
'                .Left = 244 '124
'                .Top = 241
'                .Width = 20
'                .Height = 20
'                .Enabled = True
'                .Visible = True
'                .Caption = "x"
'                .ForeColor = System.Drawing.Color.FromArgb(-1)
'                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
'                .DrawBackImage = True
'                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
'                .ToolTipText = "Removes the current queue item selection from the queue."
'            End With
'            Me.AddChild(CType(btnRemoveQueueItem, UIControl))

'            'btnClear initial props
'            btnClearQueue = New UIButton(oUILib)
'            With btnClearQueue
'                .ControlName = "btnClear"
'                .Left = 264 ' 144
'                .Top = 241
'                .Width = 20
'                .Height = 20
'                .Enabled = True
'                .Visible = True
'                .Caption = "c"
'                .ForeColor = System.Drawing.Color.FromArgb(-1)
'                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
'                .DrawBackImage = True
'                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
'                .ToolTipText = "Clears the build queue."
'            End With
'            Me.AddChild(CType(btnClearQueue, UIControl))

'            'txtQuantity initial props
'            txtQuantity = New UITextBox(oUILib)
'            With txtQuantity
'                .ControlName = "txtQuantity"
'                .Left = 95 ' 160
'                .Top = 242
'                .Width = 96
'                .Height = 18
'                .Enabled = True
'                .Visible = True
'                .Caption = "1"
'                .ForeColor = muSettings.InterfaceTextBoxForeColor
'                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'                .DrawBackImage = False
'                .FontFormat = CType(4, DrawTextFormat)
'                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'                .MaxLength = 7
'                .BorderColor = muSettings.InterfaceBorderColor
'            End With
'            Me.AddChild(CType(txtQuantity, UIControl))

'            'lblQty initial props
'            lblQty = New UILabel(oUILib)
'            With lblQty
'                .ControlName = "lblQty"
'                .Left = 30 '95
'                .Top = 241
'                .Width = 61
'                .Height = 18
'                .Enabled = True
'                .Visible = True
'                .Caption = "Quantity:"
'                .ForeColor = muSettings.InterfaceBorderColor
'                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'                .DrawBackImage = False
'                .FontFormat = CType(6, DrawTextFormat)
'            End With
'            Me.AddChild(CType(lblQty, UIControl))
'        Else
'            'Me.Height = 123 '363 
'            lstBuildable.Height = Me.Height - lstBuildable.Top - 5
'            txtItemDetails.Height = lstBuildable.Height
'            'lstBuildable.Height = 340
'        End If

'        'Now, position me better
'        Dim lFinalLeft As Int32
'        Dim lFinalTop As Int32
'        If muSettings.BuildWindowLeft < 0 AndAlso muSettings.BuildWindowTop < 0 Then
'            'Ok, check if the lower right is taken...
'            Dim ofrmChat As UIWindow = oUILib.GetWindow("frmChat")
'            Dim bTopSet As Boolean = False
'            lFinalLeft = oUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
'            If ofrmChat Is Nothing = False Then
'                If ofrmChat.Left < lFinalLeft AndAlso ofrmChat.Left + ofrmChat.Width > lFinalLeft Then
'                    lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
'                    lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
'                    bTopSet = True
'                End If

'                'lFinalLeft = ofrmChat.Left + ofrmChat.Width '+ Me.Width
'                'Else : lFinalLeft = oUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
'            End If

'            If lFinalLeft + Me.Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then
'                'we exceed it... so center the form in the screen
'                lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
'                'Center the top too
'                lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
'            ElseIf bTopSet = False Then
'                lFinalTop = oUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height
'            End If
'        Else
'            If muSettings.BuildWindowTop < 0 OrElse (muSettings.BuildWindowTop + Me.Height) > oUILib.oDevice.PresentationParameters.BackBufferHeight Then
'                'Ok, center vertically
'                lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
'            Else : lFinalTop = muSettings.BuildWindowTop
'            End If
'            If muSettings.BuildWindowLeft < 0 OrElse (muSettings.BuildWindowLeft + Me.Width) > oUILib.oDevice.PresentationParameters.BackBufferWidth Then
'                'Center horizontally
'                lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
'            Else : lFinalLeft = muSettings.BuildWindowLeft
'            End If
'        End If

'        Me.Left = lFinalLeft
'        Me.Top = lFinalTop

'        muSettings.BuildWindowLeft = lFinalLeft
'        muSettings.BuildWindowTop = lFinalTop

'        'Ensure that we are unique
'        oUILib.RemoveWindow(Me.ControlName)

'        'Now, add me...
'        oUILib.AddWindow(Me)

'        msw_Delay = Stopwatch.StartNew

'        For X As Int32 = 0 To glEntityDefUB
'            If glEntityDefIdx(X) <> -1 Then
'                If goEntityDefs(X).RequiredProductionTypeID = ProductionType.eEnlisted Then
'                    goEntityDefs(X).ProductionCost.ColonistCost = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnlistedsColonistsCost)
'                ElseIf goEntityDefs(X).RequiredProductionTypeID = ProductionType.eOfficers Then
'                    goEntityDefs(X).ProductionCost.EnlistedCost = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eOfficersEnlistedCost)
'                End If
'            End If
'        Next X

'        mbLoading = False

'    End Sub

'    Private Sub FillBuildList(ByVal yProductionType As Byte)
'        'Const y_LEVEL_COMPARE As Byte = ProductionType.e2ndLevel Or ProductionType.e3rdLevel

'        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
'            If yProductionType = ProductionType.eEnlisted Then
'                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.BarracksBuildWindow)
'            End If
'        End If

'        Dim X As Int32
'        Dim yAcceptedChassis As Byte

'        lstBuildable.Clear()

'        If goCurrentEnvir Is Nothing Then Exit Sub

'        Dim bHasCC As Boolean = False
'        Dim bHasTradepost As Boolean = False
'        For X = 0 To goCurrentEnvir.lEntityUB
'            If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
'                If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eCommandCenterSpecial Then bHasCC = True
'                If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eTradePost Then bHasTradepost = True
'            End If
'        Next X

'        If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
'            yAcceptedChassis = ChassisType.eSpaceBased
'        ElseIf goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
'            yAcceptedChassis = ChassisType.eAtmospheric Or ChassisType.eGroundBased
'        End If

'        Dim sName() As String = Nothing
'        Dim lID() As Int32 = Nothing
'        Dim iTypeID() As Int16 = Nothing
'        Dim lUB As Int32 = -1

'        For X = 0 To glEntityDefUB
'            If glEntityDefIdx(X) <> -1 Then
'                If goEntityDefs(X).RequiredProductionTypeID = yProductionType Then
'                    If (bHasCC = False OrElse goEntityDefs(X).ProductionTypeID <> ProductionType.eCommandCenterSpecial) AndAlso (bHasTradepost = False OrElse goEntityDefs(X).ProductionTypeID <> ProductionType.eTradePost) Then
'                        If (goEntityDefs(X).yChassisType And yAcceptedChassis) <> 0 Then

'                            Dim sTemp As String = goEntityDefs(X).DefName
'                            lUB += 1
'                            ReDim Preserve sName(lUB)
'                            ReDim Preserve iTypeID(lUB)
'                            ReDim Preserve lID(lUB)
'                            Dim lIdx As Int32 = lUB
'                            For Y As Int32 = 0 To lUB - 1
'                                If sTemp < sName(Y) Then
'                                    lIdx = Y
'                                    Exit For
'                                End If
'                            Next Y
'                            For Y As Int32 = lUB To lIdx + 1 Step -1
'                                sName(Y) = sName(Y - 1)
'                                lID(Y) = lID(Y - 1)
'                                iTypeID(Y) = iTypeID(Y - 1)
'                            Next Y
'                            sName(lIdx) = sTemp

'                            lID(lIdx) = goEntityDefs(X).ObjectID
'                            iTypeID(lIdx) = goEntityDefs(X).ObjTypeID

'                            ''ok, we can build htis here
'                            'lstBuildable.AddItem(goEntityDefs(X).DefName)
'                            'lstBuildable.ItemData(lstBuildable.NewIndex) = goEntityDefs(X).ObjectID
'                            'lstBuildable.ItemData2(lstBuildable.NewIndex) = goEntityDefs(X).ObjTypeID
'                        End If
'                    End If
'                End If
'            End If
'        Next X

'        'Now, for refinery stuff
'        If yProductionType = ProductionType.eRefining Then
'            For X = 0 To goCurrentPlayer.mlTechUB
'                If goCurrentPlayer.moTechs(X) Is Nothing = False Then
'                    If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech AndAlso goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
'                        Dim sTemp As String = goCurrentPlayer.moTechs(X).GetComponentName
'                        lUB += 1
'                        ReDim Preserve sName(lUB)
'                        ReDim Preserve iTypeID(lUB)
'                        ReDim Preserve lID(lUB)
'                        Dim lIdx As Int32 = lUB
'                        For Y As Int32 = 0 To lUB - 1
'                            If sTemp < sName(Y) Then
'                                lIdx = Y
'                                Exit For
'                            End If
'                        Next Y
'                        For Y As Int32 = lUB To lIdx + 1 Step -1
'                            sName(Y) = sName(Y - 1)
'                            lID(Y) = lID(Y - 1)
'                            iTypeID(Y) = iTypeID(Y - 1)
'                        Next Y
'                        sName(lIdx) = sTemp

'                        lID(lIdx) = CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID
'                        iTypeID(lIdx) = ObjectType.eMineral

'                        'lstBuildable.AddItem(goCurrentPlayer.moTechs(X).GetComponentName())
'                        'lstBuildable.ItemData(lstBuildable.NewIndex) = CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID
'                        'lstBuildable.ItemData2(lstBuildable.NewIndex) = ObjectType.eMineral
'                    End If
'                End If
'            Next X
'        End If

'        'Now, for all production-capable facilities
'        If (yProductionType And ProductionType.eProduction) <> 0 Then
'            For X = 0 To goCurrentPlayer.mlTechUB
'                If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
'                    If goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eAlloyTech AndAlso goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.ePrototype AndAlso goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eSpecialTech AndAlso goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eHullTech Then
'                        Dim sTemp As String = goCurrentPlayer.moTechs(X).GetComponentName
'                        lUB += 1
'                        ReDim Preserve sName(lUB)
'                        ReDim Preserve iTypeID(lUB)
'                        ReDim Preserve lID(lUB)
'                        Dim lIdx As Int32 = lUB
'                        For Y As Int32 = 0 To lUB - 1
'                            If sTemp < sName(Y) Then
'                                lIdx = Y
'                                Exit For
'                            End If
'                        Next Y
'                        For Y As Int32 = lUB To lIdx + 1 Step -1
'                            sName(Y) = sName(Y - 1)
'                            lID(Y) = lID(Y - 1)
'                            iTypeID(Y) = iTypeID(Y - 1)
'                        Next Y
'                        sName(lIdx) = sTemp

'                        lID(lIdx) = goCurrentPlayer.moTechs(X).ObjectID
'                        iTypeID(lIdx) = goCurrentPlayer.moTechs(X).ObjTypeID

'                        'lstBuildable.AddItem(goCurrentPlayer.moTechs(X).GetComponentName)
'                        'lstBuildable.ItemData(lstBuildable.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
'                        'lstBuildable.ItemData2(lstBuildable.NewIndex) = goCurrentPlayer.moTechs(X).ObjTypeID
'                    End If
'                End If
'            Next X
'        End If

'        For X = 0 To lUB
'            lstBuildable.AddItem(sName(X))
'            lstBuildable.ItemData(lstBuildable.NewIndex) = lID(X)
'            lstBuildable.ItemData2(lstBuildable.NewIndex) = iTypeID(X)
'        Next X

'        If lstBuildable.ListCount = 0 Then
'            Me.Visible = False
'        ElseIf mbUnitBuild = False Then
'            lstBuildable.ListIndex = 0 
'        End If

'    End Sub

'    Public Function UpdateFromEntity(ByVal lEntityIndex As Int32, Optional ByVal lChildID As Int32 = -1, Optional ByVal iChildTypeID As Int16 = -1) As Boolean
'        mlEntityIndex = lEntityIndex
'        If goCurrentEnvir Is Nothing Then Return False
'        If mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return False

'        mlChildID = lChildID
'        miChildTypeID = iChildTypeID

'        mbUseChild = False
'        If mlChildID > 0 AndAlso miChildTypeID > 0 Then
'            'Ok, find the child object...
'            Dim oChild As StationChild = goCurrentEnvir.oEntity(mlEntityIndex).GetChild(lChildID, iChildTypeID)

'            If oChild Is Nothing = False AndAlso oChild.oChildDef Is Nothing = False Then
'                mbUseChild = True
'                FillBuildList(oChild.oChildDef.ProductionTypeID)
'            End If
'        End If
'        If mbUseChild = False Then FillBuildList(goCurrentEnvir.oEntity(mlEntityIndex).yProductionType)

'        'ok, send our msg for entity production...
'        Dim yMsg(7) As Byte
'        System.BitConverter.GetBytes(EpicaMessageCode.eGetEntityProduction).CopyTo(yMsg, 0)

'        If mbUseChild = True Then
'            System.BitConverter.GetBytes(lChildID).CopyTo(yMsg, 2)
'            System.BitConverter.GetBytes(iChildTypeID).CopyTo(yMsg, 6)
'        Else : goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
'        End If
'        MyBase.moUILib.SendMsgToPrimary(yMsg)

'        msw_Delay.Reset()
'        msw_Delay.Start()

'        Return True
'    End Function

'    Public Sub HandleProductionQueueMsg(ByVal yData() As Byte)
'        If goCurrentEnvir Is Nothing OrElse mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Exit Sub

'        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
'        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
'        Dim lCurrentCycle As Int32 = System.BitConverter.ToInt32(yData, 8)
'        Dim iQueueCnt As Int16 = System.BitConverter.ToInt16(yData, 12)

'        Dim X As Int32
'        Dim lPos As Int32 = 14

'        Dim lProdID As Int32
'        Dim iProdTypeID As Int16
'        Dim lQty As Int32
'        Dim yPercComplete As Byte
'        Dim lFinishCycle As Int32

'        If lstQueue Is Nothing Then Exit Sub

'        lstQueue.Clear()

'        For X = 0 To iQueueCnt - 1
'            lProdID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'            iProdTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
'            lQty = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'            yPercComplete = yData(lPos) : lPos += 1
'            lFinishCycle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

'            lstQueue.AddItem(GetProdName(lProdID, iProdTypeID) & " x" & lQty)
'            lstQueue.ItemData(lstQueue.NewIndex) = lProdID
'            lstQueue.ItemData2(lstQueue.NewIndex) = iProdTypeID

'        Next X

'    End Sub

'    Private Sub lstBuildable_ItemClick(ByVal lIndex As Integer) Handles lstBuildable.ItemClick
'        Dim sVal As String = ""
'        Dim X As Int32

'        Dim lID As Int32
'        Dim iTypeID As Int16

'        Dim oProdCost As ProductionCost = Nothing
'        Dim sPower As String = ""
'        Dim lED_Idx As Int32 = -1

'        If lIndex > -1 Then
'            lID = lstBuildable.ItemData(lIndex)
'            iTypeID = CShort(lstBuildable.ItemData2(lIndex))

'            'Ok, find our buildable item...

'            'see if it is a tech component...
'            Dim oTech As Epica_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
'            If oTech Is Nothing = False Then
'                oProdCost = oTech.oProductionCost
'            ElseIf iTypeID = ObjectType.eMineral Then
'                'Ok, a mineral... refinery...
'                For X = 0 To goCurrentPlayer.mlTechUB
'                    If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech Then
'                        If CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID = lID Then
'                            oProdCost = goCurrentPlayer.moTechs(X).oProductionCost
'                            Exit For
'                        End If
'                    End If
'                Next X
'            Else
'                For X = 0 To glEntityDefUB
'                    If glEntityDefIdx(X) = lID AndAlso goEntityDefs(X).ObjTypeID = iTypeID Then
'                        oProdCost = goEntityDefs(X).ProductionCost
'                        lED_Idx = X

'                        If iTypeID = ObjectType.eFacilityDef Then
'                            sPower = "Power: " & goEntityDefs(X).PowerFactor & vbCrLf
'                            'If goEntityDefs(X).ProductionTypeID = ProductionType.ePowerCenter Then
'                            '    sPower = "Power: (" & goEntityDefs(X).PowerFactor & ")" & vbCrLf
'                            'Else : sPower = "Power: " & goEntityDefs(X).PowerFactor & vbCrLf
'                            'End If
'                            sPower &= "Jobs: " & goEntityDefs(X).WorkerFactor & vbCrLf

'                            If goTutorial Is Nothing = False Then
'                                If goTutorial.TutorialOn = True Then
'                                    If goEntityDefs(X).ProductionTypeID = ProductionType.eCommandCenterSpecial Then
'                                        goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.CommandCenterSelectedInBuildWindow)
'                                    End If
'                                End If

'                                If TutorialManager.bFirstFactoryBuilt = False AndAlso (goEntityDefs(X).ProductionTypeID And ProductionType.eProduction) <> 0 Then
'                                    Dim ofrm As frmTwoHourAlert = CType(goUILib.GetWindow("frmTwoHourAlert"), frmTwoHourAlert)
'                                    If ofrm Is Nothing Then ofrm = New frmTwoHourAlert(goUILib)
'                                    ofrm = Nothing
'                                End If
'                            End If
'                        End If
'                        Exit For
'                    End If
'                Next X
'            End If

'            If oProdCost Is Nothing Then
'                sVal = "Costs: Acquiring Data..."
'            Else
'                sVal = oProdCost.GetBuildCostText(sPower)
'                If iTypeID = ObjectType.eFacilityDef AndAlso lED_Idx <> -1 Then
'                    'Ok, set our buildghost       
'                    MyBase.moUILib.BuildGhost = goResMgr.GetMesh(goEntityDefs(lED_Idx).ModelID)
'                    MyBase.moUILib.vecBuildGhostLoc = New Vector3(0, 0, 0)
'                    MyBase.moUILib.BuildGhostAngle = 0
'                    MyBase.moUILib.BuildGhostID = goEntityDefs(lED_Idx).ObjectID
'                    MyBase.moUILib.BuildGhostTypeID = goEntityDefs(lED_Idx).ObjTypeID
'                End If
'            End If
'        End If

'        txtItemDetails.Caption = sVal
'    End Sub

'    Private Sub lstBuildable_ItemDblClick(ByVal lIndex As Integer) Handles lstBuildable.ItemDblClick
'        'TODO: ItemDblClick does not work
'    End Sub

'    Private Sub btnAddToQueue_Click(ByVal sName As String) Handles btnAddToQueue.Click
'        Dim lIndex As Int32 = lstBuildable.ListIndex
'        Dim lID As Int32
'        Dim iTypeID As Int16

'        Dim lQty As Int32 = Math.Abs(CInt(Val(txtQuantity.Caption)))

'        If frmMain.mbShiftKeyDown = True Then lQty = 10
'        If frmMain.mbCtrlKeyDown = True Then lQty = 100

'        If lIndex > -1 AndAlso lQty <> 0 Then
'            lID = lstBuildable.ItemData(lIndex)
'            iTypeID = CShort(lstBuildable.ItemData2(lIndex))
'            CreateAndSendSetEntityProductionMsg(lID, iTypeID, lQty)
'        End If
'    End Sub

'    Protected Overrides Sub Finalize()
'        Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
'        msw_Delay.Stop()
'        msw_Delay = Nothing
'        MyBase.Finalize()
'    End Sub

'    Private Sub frmBuildWindow_OnNewFrame() Handles Me.OnNewFrame
'        'TODO: Probably want to do this a bit smarter... probably synthesize the effects of this from when the user clicks the build button..
'        '  and on a SetEntityProduction failure, we remove the item...

'        If goCurrentEnvir Is Nothing OrElse mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 OrElse goCurrentEnvir.oEntity(mlEntityIndex).bSelected = False Then
'            MyBase.moUILib.RemoveWindow(Me.ControlName)
'        End If

'        If glCurrentEnvirView <> CurrentView.eSystemView AndAlso glCurrentEnvirView <> CurrentView.ePlanetView Then
'            MyBase.moUILib.RemoveWindow(Me.ControlName)
'        End If

'        If msw_Delay.ElapsedMilliseconds > ml_REQUEST_DELAY Then
'            'ok, send our msg for entity production...
'            Dim yMsg(7) As Byte

'            System.BitConverter.GetBytes(EpicaMessageCode.eGetEntityProduction).CopyTo(yMsg, 0)
'            If mbUseChild = True Then
'                System.BitConverter.GetBytes(mlChildID).CopyTo(yMsg, 2)
'                System.BitConverter.GetBytes(miChildTypeID).CopyTo(yMsg, 6)
'            Else : goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
'            End If
'            MyBase.moUILib.SendMsgToPrimary(yMsg)

'            msw_Delay.Reset()
'            msw_Delay.Start()
'        End If

'        If txtAvailResources Is Nothing = False AndAlso goAvailableResources Is Nothing = False Then
'            'Figure out if our available resources still apply to us...
'            With goAvailableResources
'                If .LastUpdate <> -1 Then
'                    If .iColonyParentTypeID = ObjectType.eFacility Then
'                        'space station, is the station still selected?
'                        If goCurrentEnvir.oEntity(mlEntityIndex).ObjectID <> .lColonyParentID OrElse goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID <> .iColonyParentTypeID Then
'                            'apparently not... so, clear our resources
'                            .ClearResources()
'                            txtAvailResources.Caption = ""
'                        End If
'                    ElseIf goCurrentEnvir.ObjectID <> .lColonyParentID OrElse goCurrentEnvir.ObjTypeID <> .iColonyParentTypeID Then
'                        .ClearResources()
'                    End If

'                    If .LastUpdate <> mlLastAvailResUpdate Then
'                        txtAvailResources.Caption = .GetTextBoxText()
'                        mlLastAvailResUpdate = .LastUpdate
'                    End If
'                End If
'            End With
'        End If
'    End Sub

'    Private Sub frmBuildWindow_OnResize() Handles Me.OnResize
'        If mbLoading = True Then Return
'        muSettings.BuildWindowLeft = Me.Left
'        muSettings.BuildWindowTop = Me.Top
'    End Sub

'    Private Sub btnRefresh_Click(ByVal sName As String) Handles btnRefresh.Click
'        Dim yMsg(7) As Byte
'        If goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eFacility Then
'            System.BitConverter.GetBytes(EpicaMessageCode.eGetAvailableResources).CopyTo(yMsg, 0)
'            If mbUseChild = False Then
'                goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
'            Else
'                System.BitConverter.GetBytes(mlChildID).CopyTo(yMsg, 2)
'                System.BitConverter.GetBytes(miChildTypeID).CopyTo(yMsg, 6)
'            End If
'            MyBase.moUILib.SendMsgToPrimary(yMsg)
'        End If
'    End Sub

'    Private Function GetProdName(ByVal lProdID As Int32, ByVal iProdTypeID As Int16) As String
'        'Now... set our caption...
'        If iProdTypeID = ObjectType.eFacilityDef OrElse iProdTypeID = ObjectType.eUnitDef OrElse iProdTypeID = ObjectType.eEnlisted OrElse iProdTypeID = ObjectType.eOfficers Then
'            For X As Int32 = 0 To glEntityDefUB
'                If glEntityDefIdx(X) = lProdID AndAlso goEntityDefs(X).ObjTypeID = iProdTypeID Then
'                    Return goEntityDefs(X).DefName
'                    Exit For
'                End If
'            Next X
'        ElseIf iProdTypeID = ObjectType.eMineral Then
'            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
'                If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech Then
'                    If CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID = lProdID Then
'                        Return goCurrentPlayer.moTechs(X).GetComponentName
'                        Exit For
'                    End If
'                End If
'            Next X
'        ElseIf iProdTypeID = ObjectType.eMineralTech Then
'            For X As Int32 = 0 To glMineralUB
'                If glMineralIdx(X) = lProdID Then
'                    If goMinerals(X).bDiscovered = True Then
'                        Return "Study " & goMinerals(X).MineralName
'                    Else : Return "Discover " & goMinerals(X).MineralName
'                    End If
'                    Exit For
'                End If
'            Next X
'        ElseIf iProdTypeID = ObjectType.eRepairItem Then
'            Select Case lProdID
'                Case -2
'                    Return "Armor Repairs (F)"
'                Case -3
'                    Return "Armor Repairs (L)"
'                Case -4
'                    Return "Armor Repairs (B)"
'                Case -5
'                    Return "Armor Repairs (R)"
'                Case -6
'                    Return "Structural Repairs"
'                Case Else
'                    Return "Component Repairs"
'            End Select
'        Else
'            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
'                If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjectID = lProdID AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = iProdTypeID Then
'                    Return goCurrentPlayer.moTechs(X).GetComponentName
'                    Exit For
'                End If
'            Next X
'        End If

'        Return "Unknown"
'    End Function

'    Private Sub btnClearQueue_Click(ByVal sName As String) Handles btnClearQueue.Click
'        CreateAndSendSetEntityProductionMsg(-1, -1, -1)
'    End Sub

'    Private Sub btnRemoveFromQueue_Click(ByVal sName As String) Handles btnRemoveFromQueue.Click
'        Dim lIndex As Int32 = lstBuildable.ListIndex
'        Dim lID As Int32
'        Dim iTypeID As Int16
'        Dim lQty As Int32 = CInt(Val(txtQuantity.Caption))

'        lQty = -(Math.Abs(lQty))
'        If frmMain.mbShiftKeyDown = True Then lQty = -10
'        If frmMain.mbCtrlKeyDown = True Then lQty = -100

'        If lIndex > -1 AndAlso lQty <> 0 Then
'            lID = lstBuildable.ItemData(lIndex)
'            iTypeID = CShort(lstBuildable.ItemData2(lIndex))

'            CreateAndSendSetEntityProductionMsg(lID, iTypeID, lQty)
'        End If

'    End Sub

'    Private Sub btnRemoveQueueItem_Click(ByVal sName As String) Handles btnRemoveQueueItem.Click
'        Dim lIndex As Int32 = lstQueue.ListIndex
'        Dim lID As Int32
'        Dim iTypeID As Int16
'        Dim lQty As Int32

'        lQty = lIndex

'        If lIndex > -1 Then

'            lID = lstQueue.ItemData(lIndex)
'            iTypeID = CShort(lstQueue.ItemData2(lIndex))

'            iTypeID = -(Math.Abs(iTypeID))

'            CreateAndSendSetEntityProductionMsg(lID, iTypeID, lQty)
'        End If
'    End Sub

'    Private Sub CreateAndSendSetEntityProductionMsg(ByVal lProdID As Int32, ByVal iProdTypeID As Int16, ByVal lQty As Int32)
'        Dim lBuilderID As Int32 = -1
'        Dim iBuilderTypeID As Int16

'        If goCurrentEnvir Is Nothing Then Return

'        If mbUseChild = True Then
'            'Ok, simply make it...
'            lBuilderID = mlChildID : iBuilderTypeID = miChildTypeID
'        Else
'            If iProdTypeID <> ObjectType.eFacilityDef Then
'                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
'                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
'                        lBuilderID = goCurrentEnvir.oEntity(X).ObjectID
'                        iBuilderTypeID = goCurrentEnvir.oEntity(X).ObjTypeID
'                        Exit For
'                    End If
'                Next X
'            ElseIf goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eSpaceStationSpecial Then
'                lBuilderID = goCurrentEnvir.oEntity(mlEntityIndex).ObjectID
'                iBuilderTypeID = goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID
'            Else
'                If MyBase.moUILib.BuildGhost Is Nothing = False Then MyBase.moUILib.BuildGhost = Nothing
'            End If
'        End If

'        If lBuilderID = -1 Then Return

'        Dim yData(17) As Byte
'        System.BitConverter.GetBytes(EpicaMessageCode.eSetEntityProduction).CopyTo(yData, 0)
'        System.BitConverter.GetBytes(lBuilderID).CopyTo(yData, 2)
'        System.BitConverter.GetBytes(iBuilderTypeID).CopyTo(yData, 6)
'        System.BitConverter.GetBytes(lProdID).CopyTo(yData, 8)
'        System.BitConverter.GetBytes(iProdTypeID).CopyTo(yData, 12)
'        System.BitConverter.GetBytes(lQty).CopyTo(yData, 14)

'        MyBase.moUILib.SendMsgToPrimary(yData)
'    End Sub

'    Private Sub txtQuantity_TextChanged() Handles txtQuantity.TextChanged
'        If txtQuantity.Caption <> "" Then
'            Dim lTemp As Int32 = CInt(Val(txtQuantity.Caption))
'            If lTemp.ToString <> txtQuantity.Caption Then
'                txtQuantity.Caption = lTemp.ToString
'            End If
'        End If
'    End Sub
'End Class

'Build Screen
#End Region

Public Class frmBuildWindow
    Inherits UIWindow

    Private Const ml_REQUEST_DELAY As Int32 = 3000

    Private WithEvents tvwBuildable As UITreeView
    Private lblBuildable As UILabel
    Private txtItemDetails As UITextBox
    Private lblDetails As UILabel
    Private WithEvents lstQueue As UIListBox
    Private lblBuildQueue As UILabel

    Private WithEvents btnAddToQueue As UIButton
    Private WithEvents btnRemoveFromQueue As UIButton
    Private WithEvents btnRemoveQueueItem As UIButton
    Private WithEvents btnClearQueue As UIButton
    'Private WithEvents btnAvailResources As UIButton

    Private WithEvents cboBuildFilter As UIComboBox
    Private WithEvents btnEditFilter As UIButton

    Private WithEvents chkFilter As UICheckBox

    Private WithEvents txtQuantity As UITextBox
    Private lblQty As UILabel

    Private mlEntityIndex As Int32

    Private mbUseChild As Boolean = False
    Private mlChildID As Int32 = -1
    Private miChildTypeID As Int16 = -1

    Private mbLoading As Boolean = True

    'Private mlLastAvailResUpdate As Int32 = -1

    Private msw_Delay As Stopwatch

    Private mbUnitBuild As Boolean

    Private Shared mbArmorExpanded As Boolean = False
    Private Shared mbEngineExpanded As Boolean = False
    Private Shared mbRadarExpanded As Boolean = False
    Private Shared mbShieldExpanded As Boolean = False
    Private Shared mbWeaponExpanded As Boolean = False
    Private Shared mbFacilityExpanded As Boolean = False
    Private Shared mbMineralExpanded As Boolean = False
    Private Shared mbUnitExpanded As Boolean = False

    Public Sub New(ByRef oUILib As UILib, ByVal bUnitBuild As Boolean)
        MyBase.New(oUILib)

        mbUnitBuild = bUnitBuild

        'frmFacilityBuild initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eBuildWindow
            .ControlName = "frmBuildWindow"
            .Left = 118
            .Top = 76
            .Width = 511
            .Height = 511
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .mbAcceptReprocessEvents = True
            .bRoundedBorder = True
            .BorderLineWidth = 2
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'lstBuildable initial props
        tvwBuildable = New UITreeView(oUILib)
        With tvwBuildable
            .ControlName = "lstBuildable"       'NOTE: Left the name the same to not break any tutorial scripts
            .Left = 4
            .Top = 28
            .Width = 280
            .Height = 95
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(tvwBuildable, UIControl))

        'lblBuildable initial props
        lblBuildable = New UILabel(oUILib)
        With lblBuildable
            .ControlName = "lblBuildable"
            .Left = 4
            .Top = 4
            .Width = 90
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Buildable Items"
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblBuildable, UIControl))

        'txtItemDetails initial props
        txtItemDetails = New UITextBox(oUILib)
        With txtItemDetails
            .ControlName = "txtItemDetails"
            .Left = 290
            .Top = 40 '20
            .Width = 215
            .Height = 210
            'If bUnitBuild Then .Height = 100 Else .Height = 220
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Arial", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = 0
            .BackColorEnabled = muSettings.InterfaceFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .Locked = True
            .MultiLine = True
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(txtItemDetails, UIControl))

        'btnEditFilter initial props
        btnEditFilter = New UIButton(oUILib)
        With btnEditFilter
            .ControlName = "btnEditFilter"
            .Width = 80
            .Height = 18
            .Left = 290
            .Top = 5
            .Enabled = True
            .Visible = False
            .Caption = "Edit Filter"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnEditFilter, UIControl))

        'cboBuildFilter initial props
        cboBuildFilter = New UIComboBox(oUILib)
        With cboBuildFilter
            .ControlName = "cboBuildFilter"
            .Left = btnEditFilter.Left + btnEditFilter.Width + 1
            .Top = btnEditFilter.Top
            .Width = Me.Width - .Left - 5
            .Height = 18
            .Enabled = True
            .Visible = False
            .l_ListBoxHeight = 200
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboBuildFilter, UIControl))

        'lblDetails initial props
        lblDetails = New UILabel(oUILib)
        With lblDetails
            .ControlName = "lblDetails"
            .Left = 290
            .Top = btnEditFilter.Top + btnEditFilter.Height + 1
            .Width = 90
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Item Details"
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblDetails, UIControl))

        'chkFilter initial props
        chkFilter = New UICheckBox(oUILib)
        With chkFilter
            .ControlName = "chkFilter"
            .Left = tvwBuildable.Left + tvwBuildable.Width - 110
            .Top = lblBuildable.Top
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Filter Archived"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkFilter, UIControl))

        tvwBuildable.Height = Me.Height - tvwBuildable.Top - 5
        If bUnitBuild = False Then
            ''btnAvailResources initial props
            'btnAvailResources = New UIButton(oUILib)
            'With btnAvailResources
            '	.ControlName = "btnAvailResources"
            '	.Left = 4
            '	.Top = 233
            '	.Width = 150
            '	.Height = 18
            '	.Enabled = True
            '	.Visible = True
            '	.Caption = "Available Resources"
            '	.ForeColor = muSettings.InterfaceBorderColor
            '	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '	.DrawBackImage = True
            '	.FontFormat = CType(5, DrawTextFormat)
            '	.ControlImageRect = New Rectangle(0, 0, 120, 32)
            'End With
            'Me.AddChild(CType(btnAvailResources, UIControl))

            'lstQueue initial props
            lstQueue = New UIListBox(oUILib)
            With lstQueue
                .ControlName = "lstQueue"
                .Left = 290
                .Top = Me.Height - 200
                .Width = Me.Width - .Left - 5
                .Height = Me.Height - .Top - 5 '95
                .Enabled = True
                .Visible = True
                '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
                .BorderColor = muSettings.InterfaceBorderColor
                '.FillColor = System.Drawing.Color.FromArgb(-16760704)
                .FillColor = muSettings.InterfaceFillColor
                '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
                '.HighlightColor = System.Drawing.Color.FromArgb(-16744193)
                .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(lstQueue, UIControl))

            'lblBuildQueue initial props
            lblBuildQueue = New UILabel(oUILib)
            With lblBuildQueue
                .ControlName = "lblBuildQueue"
                .Left = 290
                .Top = lstQueue.Top - 20
                .Width = Me.Width - .Left - 5 ' 90
                .Height = 20
                .Enabled = True
                .Visible = True
                .Caption = "Build Queue"
                '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblBuildQueue, UIControl))

            'btnClear initial props
            btnClearQueue = New UIButton(oUILib)
            With btnClearQueue
                .ControlName = "btnClear"
                .Left = Me.Width - 25
                .Top = lblBuildQueue.Top - 25
                .Width = 20
                .Height = 20
                .Enabled = True
                .Visible = True
                .Caption = "c"
                .ForeColor = System.Drawing.Color.FromArgb(-1)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
                .ToolTipText = "Clears the build queue."
            End With
            Me.AddChild(CType(btnClearQueue, UIControl))

            'btnRemove initial props
            btnRemoveQueueItem = New UIButton(oUILib)
            With btnRemoveQueueItem
                .ControlName = "btnRemove"
                .Left = btnClearQueue.Left - 20
                .Top = btnClearQueue.Top
                .Width = 20
                .Height = 20
                .Enabled = True
                .Visible = True
                .Caption = "x"
                .ForeColor = System.Drawing.Color.FromArgb(-1)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
                .ToolTipText = "Removes the current queue item selection from the queue."
            End With
            Me.AddChild(CType(btnRemoveQueueItem, UIControl))

            'btnMoveDown initial props
            btnRemoveFromQueue = New UIButton(oUILib)
            With btnRemoveFromQueue
                .ControlName = "btnRemoveFromQueue"
                .Left = btnRemoveQueueItem.Left - 20
                .Top = btnClearQueue.Top
                .Width = 20
                .Height = 20
                .Enabled = True
                .Visible = True
                .Caption = "-"
                .ForeColor = System.Drawing.Color.FromArgb(-1)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
                .ToolTipText = "Removes Quantity of the current selection from the build queue."
            End With
            Me.AddChild(CType(btnRemoveFromQueue, UIControl))

            'btnMoveUp initial props
            btnAddToQueue = New UIButton(oUILib)
            With btnAddToQueue
                .ControlName = "btnAddToQueue"
                .Left = btnRemoveFromQueue.Left - 20
                .Top = btnClearQueue.Top
                .Width = 20
                .Height = 20
                .Enabled = True
                .Visible = True
                .Caption = "+"
                .ForeColor = System.Drawing.Color.FromArgb(-1)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
                .ToolTipText = "Adds Quantity of the current selection to the build queue."
            End With
            Me.AddChild(CType(btnAddToQueue, UIControl))

            'txtQuantity initial props
            txtQuantity = New UITextBox(oUILib)
            With txtQuantity
                .ControlName = "txtQuantity"
                .Left = btnAddToQueue.Left - 106
                .Top = btnClearQueue.Top
                .Width = 96
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
                .MaxLength = 7
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(txtQuantity, UIControl))

            'lblQty initial props
            lblQty = New UILabel(oUILib)
            With lblQty
                .ControlName = "lblQty"
                .Left = txtQuantity.Left - 65
                .Top = btnClearQueue.Top
                .Width = 61
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Qty:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
            End With
            Me.AddChild(CType(lblQty, UIControl))

            txtItemDetails.Height = lblQty.Top - 5 - txtItemDetails.Top
        Else
            txtItemDetails.Height = Me.Height - txtItemDetails.Top - 5
        End If

        'Now, position me better
        Dim lFinalLeft As Int32
        Dim lFinalTop As Int32
        If (muSettings.BuildWindowLeft < 0 AndAlso muSettings.BuildWindowTop < 0) OrElse NewTutorialManager.TutorialOn = True Then
            'Ok, check if the lower right is taken...
            Dim ofrmChat As UIWindow = oUILib.GetWindow("frmChat")
            Dim bTopSet As Boolean = False
            lFinalLeft = oUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
            If ofrmChat Is Nothing = False Then
                If ofrmChat.Left < lFinalLeft AndAlso ofrmChat.Left + ofrmChat.Width > lFinalLeft Then
                    lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
                    lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
                    bTopSet = True
                End If

                'lFinalLeft = ofrmChat.Left + ofrmChat.Width '+ Me.Width
                'Else : lFinalLeft = oUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
            End If

            If lFinalLeft + Me.Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then
                'we exceed it... so center the form in the screen
                lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
                'Center the top too
                lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
            ElseIf bTopSet = False Then
                lFinalTop = oUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height
            End If
        Else
            If muSettings.BuildWindowTop < 0 OrElse (muSettings.BuildWindowTop + Me.Height) > oUILib.oDevice.PresentationParameters.BackBufferHeight Then
                'Ok, center vertically
                lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
            Else : lFinalTop = muSettings.BuildWindowTop
            End If
            If muSettings.BuildWindowLeft < 0 OrElse (muSettings.BuildWindowLeft + Me.Width) > oUILib.oDevice.PresentationParameters.BackBufferWidth Then
                'Center horizontally
                lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
            Else : lFinalLeft = muSettings.BuildWindowLeft
            End If
        End If

        Me.Left = lFinalLeft
        Me.Top = lFinalTop

        muSettings.BuildWindowLeft = lFinalLeft
        muSettings.BuildWindowTop = lFinalTop

        'Ensure that we are unique
        oUILib.RemoveWindow(Me.ControlName)

        'Now, add me...
        If HasAliasedRights(AliasingRights.eAddProduction Or AliasingRights.eCancelProduction) = True Then
            oUILib.AddWindow(Me)
        Else : Return
        End If

        msw_Delay = Stopwatch.StartNew

        For X As Int32 = 0 To glEntityDefUB
            If glEntityDefIdx(X) <> -1 Then
                If goEntityDefs(X).RequiredProductionTypeID = ProductionType.eEnlisted Then
                    goEntityDefs(X).ProductionCost.ColonistCost = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnlistedsColonistsCost)
                ElseIf goEntityDefs(X).RequiredProductionTypeID = ProductionType.eOfficers Then
                    goEntityDefs(X).ProductionCost.EnlistedCost = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eOfficersEnlistedCost)
                End If
            End If
        Next X

        mbLoading = False

    End Sub

    Private Sub FillBuildList(ByVal yProductionType As Byte)
        'Const y_LEVEL_COMPARE As Byte = ProductionType.e2ndLevel Or ProductionType.e3rdLevel

        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
            If yProductionType = ProductionType.eEnlisted Then
                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.BarracksBuildWindow)
            End If
        End If

        Dim X As Int32
        Dim yAcceptedChassis As Byte

        tvwBuildable.Clear()

        If mbUnitBuild = False Then
            If tvwBuildable.oIconTexture Is Nothing OrElse tvwBuildable.oIconTexture.Disposed = True Then
                tvwBuildable.oIconTexture = goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
            End If
            'lstBuildable.RenderIcons = True
        End If

        If goCurrentEnvir Is Nothing Then Return

        Dim bHasCC As Boolean = False
        Dim bHasTradepost As Boolean = False
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eCommandCenterSpecial Then bHasCC = True
                If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eTradePost Then
                    Dim bSetTradepost As Boolean = True
                    If goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.lGuildHallID = goCurrentEnvir.oEntity(X).ObjectID Then bSetTradepost = False
                    If bSetTradepost = True Then
                        bHasTradepost = True
                    End If
                End If
            End If
        Next X

        If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
            yAcceptedChassis = ChassisType.eSpaceBased
        ElseIf goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
            Dim bNavalPossible As Boolean = False
            If CType(goCurrentEnvir.oGeoObject, Planet).MapTypeID = PlanetType.eAcidic Then
                'check for Naval 3
            ElseIf CType(goCurrentEnvir.oGeoObject, Planet).MapTypeID = PlanetType.eGeoPlastic Then
                'check for naval 4
            ElseIf CType(goCurrentEnvir.oGeoObject, Planet).MapTypeID <> PlanetType.eBarren Then
                bNavalPossible = True
            End If
            yAcceptedChassis = ChassisType.eAtmospheric Or ChassisType.eGroundBased 'Or ChassisType.eNavalBased
            If bNavalPossible = True Then yAcceptedChassis = yAcceptedChassis Or ChassisType.eNavalBased
        End If

        Dim sName() As String = Nothing
        Dim lID() As Int32 = Nothing
        Dim iTypeID() As Int16 = Nothing
        Dim lUB As Int32 = -1

        Dim bFilter As Boolean = chkFilter.Value

        For X = 0 To glEntityDefUB
            If glEntityDefIdx(X) <> -1 Then
                If goEntityDefs(X).RequiredProductionTypeID = yProductionType Then
                    If (bHasCC = False OrElse goEntityDefs(X).ProductionTypeID <> ProductionType.eCommandCenterSpecial) AndAlso (bHasTradepost = False OrElse goEntityDefs(X).ProductionTypeID <> ProductionType.eTradePost) Then
                        If (goEntityDefs(X).yChassisType And yAcceptedChassis) <> 0 Then

                            'ok, lets check the filter
                            If bFilter = True Then
                                Dim lPrototypeID As Int32 = goEntityDefs(X).PrototypeID
                                If lPrototypeID > -1 Then
                                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lPrototypeID, ObjectType.ePrototype)
                                    If oTech Is Nothing = False Then
                                        If oTech.bArchived = True Then Continue For
                                    End If
                                End If
                            End If

                            If (goEntityDefs(X).lExtendedFlags And 1) <> 0 Then Continue For
                            If (goEntityDefs(X).lExtendedFlags And 2) <> 0 Then Continue For

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

                            ''ok, we can build htis here
                            'lstBuildable.AddItem(goEntityDefs(X).DefName)
                            'lstBuildable.ItemData(lstBuildable.NewIndex) = goEntityDefs(X).ObjectID
                            'lstBuildable.ItemData2(lstBuildable.NewIndex) = goEntityDefs(X).ObjTypeID
                        End If
                    End If
                End If
            End If
        Next X

        'Now, for refinery stuff
        If yProductionType = ProductionType.eRefining Then
            For X = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X) Is Nothing = False Then
                    If (goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech OrElse goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eArmorTech) AndAlso goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then

                        If goCurrentPlayer.moTechs(X).bArchived = False OrElse bFilter = False Then
                            Dim sTemp As String = goCurrentPlayer.moTechs(X).GetComponentName
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

                            If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech Then
                                lID(lIdx) = CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID
                                iTypeID(lIdx) = ObjectType.eMineral
                            Else
                                'armor
                                lID(lIdx) = goCurrentPlayer.moTechs(X).ObjectID
                                iTypeID(lIdx) = goCurrentPlayer.moTechs(X).ObjTypeID
                            End If


                            'lstBuildable.AddItem(goCurrentPlayer.moTechs(X).GetComponentName())
                            'lstBuildable.ItemData(lstBuildable.NewIndex) = CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID
                            'lstBuildable.ItemData2(lstBuildable.NewIndex) = ObjectType.eMineral
                        End If

                    End If
                End If
            Next X
        End If

        'Now, for all production-capable facilities
        If (yProductionType And ProductionType.eProduction) <> 0 Then
            For X = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                    If goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eAlloyTech AndAlso goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.ePrototype AndAlso goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eSpecialTech AndAlso goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eHullTech Then
                        If bFilter = False OrElse goCurrentPlayer.moTechs(X).bArchived = False Then

                            If goCurrentPlayer.moTechs(X).MajorDesignFlaw = 31 Then Continue For

                            Dim sTemp As String = goCurrentPlayer.moTechs(X).GetComponentName
                            lUB += 1
                            ReDim Preserve sName(lUB)
                            ReDim Preserve iTypeID(lUB)
                            ReDim Preserve lID(lUB)
                            Dim lIdx As Int32 = lUB
                            For Y As Int32 = 0 To lUB - 1
                                If goCurrentPlayer.moTechs(X).ObjTypeID < iTypeID(Y) OrElse (goCurrentPlayer.moTechs(X).ObjTypeID = iTypeID(Y) And sTemp.ToUpper < sName(Y).ToUpper) Then
                                    'If sTemp.ToUpper < sName(Y).ToUpper Then
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

                            lID(lIdx) = goCurrentPlayer.moTechs(X).ObjectID
                            iTypeID(lIdx) = goCurrentPlayer.moTechs(X).ObjTypeID

                            'lstBuildable.AddItem(goCurrentPlayer.moTechs(X).GetComponentName)
                            'lstBuildable.ItemData(lstBuildable.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
                            'lstBuildable.ItemData2(lstBuildable.NewIndex) = goCurrentPlayer.moTechs(X).ObjTypeID
                        End If

                    End If
                End If
            Next X
        End If

        'Add our root nodes
        Dim oArmorRoot As UITreeView.UITreeViewItem = tvwBuildable.AddNode("Armors", -1, ObjectType.eArmorTech, -1, Nothing, Nothing)
        Dim oEngineRoot As UITreeView.UITreeViewItem = tvwBuildable.AddNode("Engines", -1, ObjectType.eEngineTech, -1, Nothing, Nothing)
        Dim oRadarRoot As UITreeView.UITreeViewItem = tvwBuildable.AddNode("Radars", -1, ObjectType.eRadarTech, -1, Nothing, Nothing)
        Dim oShieldRoot As UITreeView.UITreeViewItem = tvwBuildable.AddNode("Shields", -1, ObjectType.eShieldTech, -1, Nothing, Nothing)
        Dim oWeaponRoot As UITreeView.UITreeViewItem = tvwBuildable.AddNode("Weapons", -1, ObjectType.eWeaponTech, -1, Nothing, Nothing)
        Dim oFacilityRoot As UITreeView.UITreeViewItem = tvwBuildable.AddNode("Facilities", -1, ObjectType.eFacilityDef, -1, Nothing, Nothing)
        Dim oMineralRoot As UITreeView.UITreeViewItem = tvwBuildable.AddNode("Minerals", -1, ObjectType.eMineral, -1, Nothing, Nothing)
        Dim oUnitRoot As UITreeView.UITreeViewItem = tvwBuildable.AddNode("Units", -1, ObjectType.eUnitDef, -1, Nothing, Nothing)

        With oArmorRoot
            .bRenderIcon = True : .IconColor = AvailResources.GetTypeIDIconColor(CShort(.lItemData2)) : .IconRect = AvailResources.GetTypeIDIconRect(CShort(.lItemData2))
            .clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : .bUseFillColor = True : .bItemBold = True : .bItemLocked = True
            .bExpanded = mbArmorExpanded
        End With
        With oEngineRoot
            .bRenderIcon = True : .IconColor = AvailResources.GetTypeIDIconColor(CShort(.lItemData2)) : .IconRect = AvailResources.GetTypeIDIconRect(CShort(.lItemData2))
            .clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : .bUseFillColor = True : .bItemBold = True : .bItemLocked = True
            .bExpanded = mbEngineExpanded
        End With
        With oRadarRoot
            .bRenderIcon = True : .IconColor = AvailResources.GetTypeIDIconColor(CShort(.lItemData2)) : .IconRect = AvailResources.GetTypeIDIconRect(CShort(.lItemData2))
            .clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : .bUseFillColor = True : .bItemBold = True : .bItemLocked = True
            .bExpanded = mbRadarExpanded
        End With
        With oShieldRoot
            .bRenderIcon = True : .IconColor = AvailResources.GetTypeIDIconColor(CShort(.lItemData2)) : .IconRect = AvailResources.GetTypeIDIconRect(CShort(.lItemData2))
            .clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : .bUseFillColor = True : .bItemBold = True : .bItemLocked = True
            .bExpanded = mbShieldExpanded
        End With
        With oWeaponRoot
            .bRenderIcon = True : .IconColor = AvailResources.GetTypeIDIconColor(CShort(.lItemData2)) : .IconRect = AvailResources.GetTypeIDIconRect(CShort(.lItemData2))
            .clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : .bUseFillColor = True : .bItemBold = True : .bItemLocked = True
            .bExpanded = mbWeaponExpanded
        End With
        With oFacilityRoot
            .bRenderIcon = True : .IconColor = AvailResources.GetTypeIDIconColor(CShort(.lItemData2)) : .IconRect = AvailResources.GetTypeIDIconRect(CShort(.lItemData2))
            .clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : .bUseFillColor = True : .bItemBold = True : .bItemLocked = True
            .bExpanded = mbFacilityExpanded
        End With
        With oMineralRoot
            .bRenderIcon = True : .IconColor = AvailResources.GetTypeIDIconColor(CShort(.lItemData2)) : .IconRect = AvailResources.GetTypeIDIconRect(CShort(.lItemData2))
            .clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : .bUseFillColor = True : .bItemBold = True : .bItemLocked = True
            .bExpanded = mbMineralExpanded
        End With
        With oUnitRoot
            .bRenderIcon = True : .IconColor = AvailResources.GetTypeIDIconColor(CShort(.lItemData2)) : .IconRect = AvailResources.GetTypeIDIconRect(CShort(.lItemData2))
            .clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : .bUseFillColor = True : .bItemBold = True : .bItemLocked = True
            .bExpanded = mbUnitExpanded
        End With

        Dim bHasItems As Boolean = False

        For X = 0 To lUB
            'Get the root node...
            bHasItems = True
            Dim oRoot As UITreeView.UITreeViewItem = tvwBuildable.GetNodeByItemData2(-1, iTypeID(X))
            If oRoot Is Nothing Then
                oRoot = tvwBuildable.GetNodeByItemData2(-1, -1)
                If oRoot Is Nothing Then
                    oRoot = tvwBuildable.AddNode("Other", -1, -1, -1, Nothing, Nothing)
                End If
            End If
            Dim oNode As UITreeView.UITreeViewItem = tvwBuildable.AddNode(sName(X), lID(X), iTypeID(X), -1, oRoot, Nothing)
            'tvwBuildable.AddItem(sName(X))
            'tvwBuildable.ItemData(tvwBuildable.NewIndex) = lID(X)
            'tvwBuildable.ItemData2(tvwBuildable.NewIndex) = iTypeID(X)
            'tvwBuildable.ApplyIconOffset(tvwBuildable.NewIndex) = True
            'tvwBuildable.IconForeColor(tvwBuildable.NewIndex) = AvailResources.GetTypeIDIconColor(iTypeID(X))
            'tvwBuildable.rcIconRectangle(tvwBuildable.NewIndex) = AvailResources.GetTypeIDIconRect(iTypeID(X))
            oNode.bRenderIcon = True
            oNode.IconColor = AvailResources.GetTypeIDIconColor(iTypeID(X))
            oNode.IconRect = AvailResources.GetTypeIDIconRect(iTypeID(X))
            If (iTypeID(X) = ObjectType.eUnitDef AndAlso sName(X) = "Engineer") OrElse (iTypeID(X) = ObjectType.eOfficers AndAlso sName(X).StartsWith("Officers")) OrElse (iTypeID(X) = ObjectType.eEnlisted AndAlso sName(X).StartsWith("Enlisted")) Then
                tvwBuildable.oSelectedNode = oNode
            End If
        Next X

        If oArmorRoot Is Nothing = False AndAlso oArmorRoot.oFirstChild Is Nothing Then tvwBuildable.RemoveNode(oArmorRoot)
        If oEngineRoot Is Nothing = False AndAlso oEngineRoot.oFirstChild Is Nothing Then tvwBuildable.RemoveNode(oEngineRoot)
        If oRadarRoot Is Nothing = False AndAlso oRadarRoot.oFirstChild Is Nothing Then tvwBuildable.RemoveNode(oRadarRoot)
        If oShieldRoot Is Nothing = False AndAlso oShieldRoot.oFirstChild Is Nothing Then tvwBuildable.RemoveNode(oShieldRoot)
        If oWeaponRoot Is Nothing = False AndAlso oWeaponRoot.oFirstChild Is Nothing Then tvwBuildable.RemoveNode(oWeaponRoot)
        If oFacilityRoot Is Nothing = False AndAlso oFacilityRoot.oFirstChild Is Nothing Then tvwBuildable.RemoveNode(oFacilityRoot)
        If oMineralRoot Is Nothing = False AndAlso oMineralRoot.oFirstChild Is Nothing Then tvwBuildable.RemoveNode(oMineralRoot)
        If oUnitRoot Is Nothing = False AndAlso oUnitRoot.oFirstChild Is Nothing Then tvwBuildable.RemoveNode(oUnitRoot)

        If tvwBuildable.oRootNode Is Nothing = False Then
            If tvwBuildable.oRootNode.oNextSibling Is Nothing Then tvwBuildable.oRootNode.bExpanded = True
        End If

        If bHasItems = False Then
            Me.Visible = False
            If yProductionType = ProductionType.eRefining Then
                MyBase.moUILib.AddNotification("No produceable alloys available. You must research the alloys to build them.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            ElseIf yProductionType = ProductionType.eFacility Then '(yProductionType And ProductionType.eFacility) <> 0 Then
                MyBase.moUILib.AddNotification("No produceable buildings available. Research a prototype in order to build them.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            ElseIf yProductionType <> 0 Then
                MyBase.moUILib.AddNotification("No produceable items in the build list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If

        ElseIf mbUnitBuild = False Then
            'tvwBuildable.ListIndex = 0
        End If

    End Sub

    Public Function UpdateFromEntity(ByVal lEntityIndex As Int32, Optional ByVal lChildID As Int32 = -1, Optional ByVal iChildTypeID As Int16 = -1) As Boolean
        mlEntityIndex = lEntityIndex

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return False
        If mlEntityIndex = -1 OrElse oEnvir.lEntityIdx Is Nothing OrElse oEnvir.lEntityUB < mlEntityIndex OrElse oEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return False

        Dim oEntity As BaseEntity = oEnvir.oEntity(mlEntityIndex)
        If oEntity Is Nothing = True Then Return False

        mlChildID = lChildID
        miChildTypeID = iChildTypeID

        mbUseChild = False
        If mlChildID > 0 AndAlso miChildTypeID > 0 Then
            'Ok, find the child object...
            Dim oChild As StationChild = oEntity.GetChild(lChildID, iChildTypeID)

            If oChild Is Nothing = False AndAlso oChild.oChildDef Is Nothing = False Then
                mbUseChild = True
                FillBuildList(oChild.oChildDef.ProductionTypeID)
            End If
        End If
        If mbUseChild = False Then FillBuildList(oEntity.yProductionType)

        'ok, send our msg for entity production...
        Dim yMsg(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityProduction).CopyTo(yMsg, 0)

        If mbUseChild = True Then
            System.BitConverter.GetBytes(lChildID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(iChildTypeID).CopyTo(yMsg, 6)
        Else : oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
        End If
        MyBase.moUILib.SendMsgToPrimary(yMsg)

        If msw_Delay Is Nothing Then msw_Delay = New Stopwatch
        msw_Delay.Reset()
        msw_Delay.Start()

        Return True
    End Function

    Public Sub HandleProductionQueueMsg(ByVal yData() As Byte)
        If goCurrentEnvir Is Nothing OrElse mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Exit Sub

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lCurrentCycle As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iQueueCnt As Int16 = System.BitConverter.ToInt16(yData, 12)

        Dim X As Int32
        Dim lPos As Int32 = 14

        Dim lProdID As Int32
        Dim iProdTypeID As Int16
        Dim lQty As Int32
        Dim yPercComplete As Byte
        Dim lFinishCycle As Int32

        Dim lRallyX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lRallyZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(mlEntityIndex)
        If oEntity Is Nothing = False Then
            oEntity.RallyX = lRallyX
            oEntity.RallyZ = lRallyZ
        End If

        If lstQueue Is Nothing Then Exit Sub
        lstQueue.Clear()

        For X = 0 To iQueueCnt - 1
            lProdID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iProdTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lQty = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            yPercComplete = yData(lPos) : lPos += 1
            lFinishCycle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            lstQueue.AddItem(GetProdName(lProdID, iProdTypeID) & " x" & lQty)
            lstQueue.ItemData(lstQueue.NewIndex) = lProdID
            lstQueue.ItemData2(lstQueue.NewIndex) = iProdTypeID
            lstQueue.ItemData3(lstQueue.NewIndex) = lQty

        Next X

        Dim sName As String = "Build Queue"
        If iQueueCnt > 0 Then
            sName &= " (" & iQueueCnt.ToString("#,###") & ")"
        Else
            sName &= " (Empty)"
        End If
        If lblBuildQueue.Caption <> sName Then lblBuildQueue.Caption = sName
        lstQueue.RestorePosition()
    End Sub

    Private Sub btnAddToQueue_Click(ByVal sName As String) Handles btnAddToQueue.Click
        Dim oNode As UITreeView.UITreeViewItem = tvwBuildable.oSelectedNode

        Dim lID As Int32
        Dim iTypeID As Int16

        Dim lQty As Int32 = Math.Abs(CInt(Val(txtQuantity.Caption)))

        If frmMain.mbShiftKeyDown = True Then
            If NewTutorialManager.TutorialOn = False OrElse goUILib.CommandAllowed(True, "ShiftClickBuild") = True Then
                lQty = 10
            End If
        End If
        If frmMain.mbCtrlKeyDown = True Then
            If NewTutorialManager.TutorialOn = False OrElse goUILib.CommandAllowed(True, "ControlClickBuild") = True Then
                lQty = 100
            End If
        End If

        If oNode Is Nothing = False AndAlso lQty <> 0 Then
            lID = oNode.lItemData
            iTypeID = CShort(oNode.lItemData2)

            If iTypeID < 1 OrElse lID < 0 Then Return

            CreateAndSendSetEntityProductionMsg(lID, iTypeID, lQty)
        End If

        'If NewTutorialManager.TutorialOn = True Then
        txtQuantity.Caption = "1"
        'End If
    End Sub

    Protected Overrides Sub Finalize()
        'MyBase.moUILib.RemoveWindow("frmAvailRes")
        Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
        msw_Delay.Stop()
        msw_Delay = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub frmBuildWindow_OnNewFrame() Handles Me.OnNewFrame
        Dim oEnvir As BaseEnvironment = goCurrentEnvir

        If oEnvir Is Nothing OrElse mlEntityIndex = -1 OrElse oEnvir.lEntityIdx(mlEntityIndex) = -1 Then
            MyBase.moUILib.RemoveWindow(Me.ControlName)
            Return
        End If

        Dim oEntity As BaseEntity = oEnvir.oEntity(mlEntityIndex)
        If oEntity Is Nothing = True OrElse oEntity.bSelected = False Then
            MyBase.moUILib.RemoveWindow(Me.ControlName)
            Return
        End If

        If glCurrentEnvirView <> CurrentView.eSystemView AndAlso glCurrentEnvirView <> CurrentView.ePlanetView Then
            MyBase.moUILib.RemoveWindow(Me.ControlName)
        End If

        If msw_Delay.ElapsedMilliseconds > ml_REQUEST_DELAY Then
            'ok, send our msg for entity production...
            Dim yMsg(7) As Byte

            System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityProduction).CopyTo(yMsg, 0)
            If mbUseChild = True Then
                System.BitConverter.GetBytes(mlChildID).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(miChildTypeID).CopyTo(yMsg, 6)
            Else : oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
            End If
            MyBase.moUILib.SendMsgToPrimary(yMsg)

            msw_Delay.Reset()
            msw_Delay.Start()
        End If

    End Sub

    Private Sub frmBuildWindow_OnResize() Handles Me.OnResize
        If mbLoading = True Then Return
        muSettings.BuildWindowLeft = Me.Left
        muSettings.BuildWindowTop = Me.Top
    End Sub

    'Private Sub btnRefresh_Click(ByVal sName As String) Handles btnRefresh.Click
    '    Dim yMsg(7) As Byte
    '    If goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eFacility Then
    '        System.BitConverter.GetBytes(EpicaMessageCode.eGetAvailableResources).CopyTo(yMsg, 0)
    '        If mbUseChild = False Then
    '            goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
    '        Else
    '            System.BitConverter.GetBytes(mlChildID).CopyTo(yMsg, 2)
    '            System.BitConverter.GetBytes(miChildTypeID).CopyTo(yMsg, 6)
    '        End If
    '        MyBase.moUILib.SendMsgToPrimary(yMsg)
    '    End If
    'End Sub

    Private Function GetProdName(ByVal lProdID As Int32, ByVal iProdTypeID As Int16) As String
        'Now... set our caption...
        If iProdTypeID = ObjectType.eFacilityDef OrElse iProdTypeID = ObjectType.eUnitDef OrElse iProdTypeID = ObjectType.eEnlisted OrElse iProdTypeID = ObjectType.eOfficers Then
            For X As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(X) = lProdID AndAlso goEntityDefs(X).ObjTypeID = iProdTypeID Then
                    Return goEntityDefs(X).DefName
                    Exit For
                End If
            Next X
        ElseIf iProdTypeID = ObjectType.eMineral Then
            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech Then
                    If CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID = lProdID Then
                        Return goCurrentPlayer.moTechs(X).GetComponentName
                        Exit For
                    End If
                End If
            Next X
        ElseIf iProdTypeID = ObjectType.eMineralTech Then
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = lProdID Then
                    If goMinerals(X).bDiscovered = True Then
                        Return "Study " & goMinerals(X).MineralName
                    Else : Return "Discover " & goMinerals(X).MineralName
                    End If
                    Exit For
                End If
            Next X
        ElseIf iProdTypeID = ObjectType.eRepairItem Then
            Select Case lProdID
                Case -2
                    Return "Armor Repairs (F)"
                Case -3
                    Return "Armor Repairs (L)"
                Case -4
                    Return "Armor Repairs (B)"
                Case -5
                    Return "Armor Repairs (R)"
                Case -6
                    Return "Structural Repairs"
                Case Else
                    Return "Component Repairs"
            End Select
        Else
            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjectID = lProdID AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = iProdTypeID Then
                    Return goCurrentPlayer.moTechs(X).GetComponentName
                    Exit For
                End If
            Next X
        End If

        Return "Unknown"
    End Function

    Private Sub btnClearQueue_Click(ByVal sName As String) Handles btnClearQueue.Click
        CreateAndSendSetEntityProductionMsg(-1, -1, -1)
    End Sub

    Private Sub btnRemoveFromQueue_Click(ByVal sName As String) Handles btnRemoveFromQueue.Click
        Dim oNode As UITreeView.UITreeViewItem = tvwBuildable.oSelectedNode
        Dim lID As Int32
        Dim iTypeID As Int16
        Dim lQty As Int32 = CInt(Val(txtQuantity.Caption))

        lQty = -(Math.Abs(lQty))
        If frmMain.mbShiftKeyDown = True Then lQty = -10
        If frmMain.mbCtrlKeyDown = True Then lQty = -100

        If oNode Is Nothing = False AndAlso lQty <> 0 Then
            lID = oNode.lItemData
            iTypeID = CShort(oNode.lItemData2)

            If iTypeID < 1 OrElse lID < 0 Then Return

            CreateAndSendSetEntityProductionMsg(lID, iTypeID, lQty)
        End If

    End Sub

    Private Sub btnRemoveQueueItem_Click(ByVal sName As String) Handles btnRemoveQueueItem.Click
        Dim lIndex As Int32 = lstQueue.ListIndex
        Dim lID As Int32
        Dim iTypeID As Int16
        Dim lQty As Int32

        lQty = lIndex

        If lIndex > -1 Then

            lID = lstQueue.ItemData(lIndex)
            iTypeID = CShort(lstQueue.ItemData2(lIndex))

            'iTypeID = -(Math.Abs(iTypeID))
            lQty = -lstQueue.ItemData3(lIndex)

            CreateAndSendSetEntityProductionMsg(lID, iTypeID, lQty)
        End If
    End Sub

    Private Sub CreateAndSendSetEntityProductionMsg(ByVal lProdID As Int32, ByVal iProdTypeID As Int16, ByVal lQty As Int32)
        Dim lBuilderID As Int32 = -1
        Dim iBuilderTypeID As Int16

        If goCurrentEnvir Is Nothing Then Return

        If mbUseChild = True Then
            'Ok, simply make it...
            lBuilderID = mlChildID : iBuilderTypeID = miChildTypeID
        Else
            If iProdTypeID <> ObjectType.eFacilityDef Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                        lBuilderID = goCurrentEnvir.oEntity(X).ObjectID
                        iBuilderTypeID = goCurrentEnvir.oEntity(X).ObjTypeID
                        Exit For
                    End If
                Next X
            ElseIf goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eSpaceStationSpecial Then
                lBuilderID = goCurrentEnvir.oEntity(mlEntityIndex).ObjectID
                iBuilderTypeID = goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID
            Else
                If MyBase.moUILib.BuildGhost Is Nothing = False Then MyBase.moUILib.BuildGhost = Nothing
            End If
        End If

        If lBuilderID = -1 Then Return

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBuildOrderSent, lProdID, iProdTypeID, lQty, "")
        End If

        Dim yData(17) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
        System.BitConverter.GetBytes(lBuilderID).CopyTo(yData, 2)
        System.BitConverter.GetBytes(iBuilderTypeID).CopyTo(yData, 6)
        System.BitConverter.GetBytes(lProdID).CopyTo(yData, 8)
        System.BitConverter.GetBytes(iProdTypeID).CopyTo(yData, 12)
        System.BitConverter.GetBytes(lQty).CopyTo(yData, 14)

        MyBase.moUILib.SendMsgToPrimary(yData)
    End Sub

    Private Sub txtQuantity_TextChanged() Handles txtQuantity.TextChanged
        If txtQuantity.Caption <> "" Then
            Dim lTemp As Int32 = CInt(Val(txtQuantity.Caption))
            If lTemp.ToString <> txtQuantity.Caption Then
                txtQuantity.Caption = lTemp.ToString
            End If
        End If
    End Sub

    'Private Sub btnAvailResources_Click(ByVal sName As String) Handles btnAvailResources.Click
    '	If mlEntityIndex = -1 OrElse goCurrentEnvir Is Nothing OrElse goCurrentEnvir.lEntityUB < mlEntityIndex OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return
    '	Dim ofrm As frmAvailRes = CType(MyBase.moUILib.GetWindow("frmAvailRes"), frmAvailRes)
    '	If ofrm Is Nothing Then
    '		ofrm = New frmAvailRes(MyBase.moUILib, False, CType(goCurrentEnvir.oEntity(mlEntityIndex), Base_GUID))
    '	Else : ofrm.SetGUID(CType(goCurrentEnvir.oEntity(mlEntityIndex), Base_GUID))
    '	End If
    '	ofrm.Visible = True
    'End Sub

    Private Sub chkFilter_Click() Handles chkFilter.Click
        MyBase.moUILib.GetMsgSys.LoadArchived()
        If mlChildID > 0 AndAlso miChildTypeID > 0 Then
            'Ok, find the child object...
            Dim oChild As StationChild = goCurrentEnvir.oEntity(mlEntityIndex).GetChild(mlChildID, miChildTypeID)

            If oChild Is Nothing = False AndAlso oChild.oChildDef Is Nothing = False Then
                mbUseChild = True
                FillBuildList(oChild.oChildDef.ProductionTypeID)
            End If
        End If
        If mbUseChild = False Then FillBuildList(goCurrentEnvir.oEntity(mlEntityIndex).yProductionType)
    End Sub


    Private Sub frmBuildWindow_WindowMoved() Handles Me.WindowMoved
        muSettings.BuildWindowLeft = Me.Left
        muSettings.BuildWindowTop = Me.Top
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmSelectFac")
        If ofrm Is Nothing = False Then
            ofrm.Left = Me.Left - ofrm.Width
            ofrm.Top = Me.Top
        End If
    End Sub

    Private Sub cboBuildFilter_ItemSelected(ByVal lItemIndex As Integer) Handles cboBuildFilter.ItemSelected
        'xxx
    End Sub

    Private Sub btnEditFilter_Click(ByVal sName As String) Handles btnEditFilter.Click
        'xxx
    End Sub

    Private Sub tvwBuildable_NodeExpanded(ByVal oNode As UITreeView.UITreeViewItem) Handles tvwBuildable.NodeExpanded
        If oNode Is Nothing = False Then
            If oNode.oFirstChild Is Nothing = False Then
                Select Case oNode.lItemData2
                    Case ObjectType.eArmorTech
                        mbArmorExpanded = True
                    Case ObjectType.eEngineTech
                        mbEngineExpanded = True
                    Case ObjectType.eRadarTech
                        mbRadarExpanded = True
                    Case ObjectType.eShieldTech
                        mbShieldExpanded = True
                    Case ObjectType.eWeaponTech
                        mbWeaponExpanded = True
                    Case ObjectType.eFacilityDef
                        mbFacilityExpanded = True
                    Case ObjectType.eMineral
                        mbMineralExpanded = True
                    Case ObjectType.eUnitDef
                        mbUnitExpanded = True
                End Select
            End If
        End If
    End Sub

    Private Sub tvwBuildable_NodeSelected(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwBuildable.NodeSelected
        Dim sVal As String = ""
        Dim X As Int32

        Dim lID As Int32
        Dim iTypeID As Int16

        Dim oProdCost As ProductionCost = Nothing
        Dim sPower As String = ""
        Dim lED_Idx As Int32 = -1

        If oNode Is Nothing = False Then
            lID = oNode.lItemData
            iTypeID = CShort(oNode.lItemData2)

            'Ok, find our buildable item...
            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSelectItemInBuildWindow, lID, iTypeID, -1, "")
            End If

            'see if it is a tech component...
            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
            If oTech Is Nothing = False Then
                oProdCost = oTech.oProductionCost
            ElseIf iTypeID = ObjectType.eMineral Then
                'Ok, a mineral... refinery...
                For X = 0 To goCurrentPlayer.mlTechUB
                    If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech Then
                        If CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID = lID Then
                            oProdCost = goCurrentPlayer.moTechs(X).oProductionCost
                            Exit For
                        End If
                    End If
                Next X
            Else
                For X = 0 To glEntityDefUB
                    If glEntityDefIdx(X) = lID AndAlso goEntityDefs(X).ObjTypeID = iTypeID Then
                        oProdCost = goEntityDefs(X).ProductionCost
                        lED_Idx = X

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

                            If goTutorial Is Nothing = False Then
                                If goTutorial.TutorialOn = True Then
                                    If goEntityDefs(X).ProductionTypeID = ProductionType.eCommandCenterSpecial Then
                                        goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.CommandCenterSelectedInBuildWindow)
                                    End If
                                End If

                                'If TutorialManager.bFirstFactoryBuilt = False AndAlso (goEntityDefs(X).ProductionTypeID And ProductionType.eProduction) <> 0 Then
                                '	Dim ofrm As frmTwoHourAlert = CType(goUILib.GetWindow("frmTwoHourAlert"), frmTwoHourAlert)
                                '	If ofrm Is Nothing Then ofrm = New frmTwoHourAlert(goUILib)
                                '	ofrm = Nothing
                                'End If
                            End If
                        End If
                        Exit For
                    End If
                Next X
            End If

            If oProdCost Is Nothing Then
                sVal = "Costs: Acquiring Data..."
            Else
                sVal = oProdCost.GetBuildCostText(sPower)

                If goCurrentEnvir.oEntity(mlEntityIndex).yProductionType <> ProductionType.eSpaceStationSpecial AndAlso iTypeID = ObjectType.eFacilityDef AndAlso lED_Idx <> -1 Then
                    'Hide the window for building placement
                    Me.tvwBuildable.oSelectedNode = Nothing     '?
                    Me.Visible = False


                    'Ok, set our buildghost       
                    MyBase.moUILib.BuildGhost = goResMgr.GetMesh(goEntityDefs(lED_Idx).ModelID)
                    MyBase.moUILib.vecBuildGhostLoc = New Vector3(0, 0, 0)
                    MyBase.moUILib.BuildGhostAngle = 0
                    MyBase.moUILib.BuildGhostID = goEntityDefs(lED_Idx).ObjectID
                    MyBase.moUILib.BuildGhostTypeID = goEntityDefs(lED_Idx).ObjTypeID
                    MyBase.moUILib.bBuildGhostNaval = (goEntityDefs(lED_Idx).yChassisType And ChassisType.eNavalBased) <> 0
                End If
            End If
        End If

        txtItemDetails.Caption = sVal
    End Sub

    Private Sub tvwBuildable_NodeUnexpanded(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwBuildable.NodeUnexpanded
        If oNode Is Nothing = False Then
            If oNode.oFirstChild Is Nothing = False Then
                Select Case oNode.lItemData2
                    Case ObjectType.eArmorTech
                        mbArmorExpanded = False
                    Case ObjectType.eEngineTech
                        mbEngineExpanded = False
                    Case ObjectType.eRadarTech
                        mbRadarExpanded = False
                    Case ObjectType.eShieldTech
                        mbShieldExpanded = False
                    Case ObjectType.eWeaponTech
                        mbWeaponExpanded = False
                    Case ObjectType.eFacilityDef
                        mbFacilityExpanded = False
                    Case ObjectType.eMineral
                        mbMineralExpanded = False
                    Case ObjectType.eUnitDef
                        mbUnitExpanded = False
                End Select
            End If
        End If
    End Sub

    Private Sub tvwBuildable_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles tvwBuildable.OnMouseMove
        Dim sVal As String = ""
        Dim X As Int32


        Dim lID As Int32
        Dim iTypeID As Int16


        Dim oProdCost As ProductionCost = Nothing
        Dim sPower As String = ""
        Dim lED_Idx As Int32 = -1

        Dim oNode As UITreeView.UITreeViewItem = tvwBuildable.GetNodeUnderMouse(lMouseX, lMouseY)

        If tvwBuildable.oSelectedNode Is Nothing = False Then
            If tvwBuildable.oSelectedNode.lItemData > 0 Then
                Return
            End If
        End If

        'if we are hovering over a build list item and don't have a build list item selected display hovered over item.
        If oNode Is Nothing = False Then
            lID = oNode.lItemData
            iTypeID = CShort(oNode.lItemData2)

            If iTypeID < 1 OrElse lID < 0 Then Return

            'see if it is a tech component...
            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
            If oTech Is Nothing = False Then
                oProdCost = oTech.oProductionCost
            ElseIf iTypeID = ObjectType.eMineral Then
                'Ok, a mineral... refinery...
                For X = 0 To goCurrentPlayer.mlTechUB
                    If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech Then
                        If CType(goCurrentPlayer.moTechs(X), AlloyTech).AlloyResultID = lID Then
                            oProdCost = goCurrentPlayer.moTechs(X).oProductionCost
                            Exit For
                        End If
                    End If
                Next X
            Else
                For X = 0 To glEntityDefUB
                    If glEntityDefIdx(X) = lID AndAlso goEntityDefs(X).ObjTypeID = iTypeID Then
                        oProdCost = goEntityDefs(X).ProductionCost
                        lED_Idx = X


                        If iTypeID = ObjectType.eFacilityDef Then
                            sPower = "Power: " & goEntityDefs(X).PowerFactor.ToString("#,##0") & vbCrLf
                            'If goEntityDefs(X).ProductionTypeID = ProductionType.ePowerCenter Then
                            '    sPower = "Power: (" & goEntityDefs(X).PowerFactor & ")" & vbCrLf
                            'Else : sPower = "Power: " & goEntityDefs(X).PowerFactor & vbCrLf
                            'End If
                            sPower &= "Jobs: " & goEntityDefs(X).WorkerFactor.ToString("#,##0") & vbCrLf


                            If goEntityDefs(X).ProductionTypeID = ProductionType.eColonists Then
                                sPower &= "Housing: " & goEntityDefs(X).ProdFactor.ToString("#,##0") & vbCrLf
                            ElseIf goEntityDefs(X).ProductionTypeID = ProductionType.eResearch Then
                                sPower &= "Research Ability: " & goEntityDefs(X).ProdFactor.ToString("#,##0") & vbCrLf
                            ElseIf (goEntityDefs(X).ProductionTypeID And ProductionType.eProduction) <> 0 Then
                                sPower &= "Production: " & goEntityDefs(X).ProdFactor.ToString("#,##0") & vbCrLf
                            End If
                        End If
                        Exit For
                    End If
                Next X
            End If


            If oProdCost Is Nothing Then
                sVal = "Costs: Acquiring Data..."
            Else
                sVal = oProdCost.GetBuildCostText(sPower)
            End If
            txtItemDetails.Caption = sVal
        End If
    End Sub
End Class
