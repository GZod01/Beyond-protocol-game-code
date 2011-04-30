Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmFleet
    Inherits UIWindow

    Private lnDiv1 As UILine
    Private lblTitle As UILabel
    Private lblSelect As UILabel
    Private lblElements As UILabel
    Private lnDiv2 As UILine
    Private lblDetails As UILabel
    Private lnDiv3 As UILine
    Private lblReinforcers As UILabel
    Private lblDefaultFormation As UILabel

    Private WithEvents btnCreate As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnRename As UIButton
    Private WithEvents btnAdd As UIButton
    Private WithEvents btnRemove As UIButton
    Private WithEvents btnClose As UIButton
    Private WithEvents btnOrders As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents btnRemoveReinforcer As UIButton
    Private WithEvents btnFormations As UIButton

    Private WithEvents lstElements As UIListBox
    Private WithEvents lstFleet As UIListBox
    Private WithEvents lstReinforcers As UIListBox

    Private WithEvents cboDefaultFormation As UIComboBox
    
    Private WithEvents txtName As UITextBox

    Private mbHasUnknown As Boolean = False
    Private mlMoverCnt As Int32 = 0
    Private mlLastMoverUpdate As Int32

    Private mbIgnoreFormation As Boolean = True

    'If this is negative 1 then normal... otherwise, special select battlegroup look and feel
    Private mlReinforceFacilityID As Int32 = -1

    Private mlLastMsgUpdate As Int32 = -1
    Private moLaunchUnit As UnitGroup.UnitGroupElement

    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib, ByVal lFacilityID As Int32)
        MyBase.New(oUILib)

        mlReinforceFacilityID = lFacilityID
        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenBattlegroupWindow)

        'frmFleet initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eFleet
            .ControlName = "frmFleet"
            '.Left = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 245
            '.Top = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 220
            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.FleetX
                lTop = muSettings.FleetY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 245
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 220
            If lLeft + 490 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 490
            If lTop + 511 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 511

            .Left = lLeft
            .Top = lTop

            .Width = 490
            .Height = 511
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 1
        End With

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = Me.BorderLineWidth
            .Top = 25
            .Width = Me.Width - Me.BorderLineWidth
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
            .Top = 5
            .Width = 281
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Battlegroup Management"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lblSelect initial props
        lblSelect = New UILabel(oUILib)
        With lblSelect
            .ControlName = "lblSelect"
            .Left = 5
            .Top = 30
            .Width = 341
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select a Battlegroup:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSelect, UIControl))

        'lstFleet initial props
        lstFleet = New UIListBox(oUILib)
        With lstFleet
            .ControlName = "lstFleet"
            .Left = 5
            .Top = 50
            .Width = 480
            .Height = 100
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstFleet, UIControl))

        'btnCreate initial props
        btnCreate = New UIButton(oUILib)
        With btnCreate
            .ControlName = "btnCreate"
            .Left = 385
            .Top = 30
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Create New"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Select to Create a new battlegroup with the currently selected units."
        End With
        Me.AddChild(CType(btnCreate, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = Me.BorderLineWidth
            .Top = 155
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'lblDetails initial props
        lblDetails = New UILabel(oUILib)
        With lblDetails
            .ControlName = "lblDetails"
            .Left = 5
            .Top = 155
            .Width = 400 '150
            .Height = 18
            .Enabled = True
            .Visible = True
            '.Caption = "Battlegroup Name     Total Warpoint Upkeep: "
            .Caption = "Battlegroup Name"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDetails, UIControl))

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Left = 385
            .Top = 175
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Disband"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Disbands the entire battlegroup" & vbCrLf & "All elements will be detached from" & vbCrLf & "the fleet/army and will be individual."
        End With
        Me.AddChild(CType(btnDelete, UIControl))

        'txtName initial props
        txtName = New UITextBox(oUILib)
        With txtName
            .ControlName = "txtName"
            .Left = 5
            .Top = 175
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtName, UIControl))

        'btnRename initial props
        btnRename = New UIButton(oUILib)
        With btnRename
            .ControlName = "btnRename"
            .Left = 210
            .Top = 175
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Rename"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to change the name of the battlegroup."
        End With
        Me.AddChild(CType(btnRename, UIControl))

        'lblElements initial props
        lblElements = New UILabel(oUILib)
        With lblElements
            .ControlName = "lblElements"
            .Left = 5
            .Top = 200
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Elements:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblElements, UIControl))

        'lstElements initial props
        lstElements = New UIListBox(oUILib)
        With lstElements
            .ControlName = "lstElements"
            .Left = 5
            .Top = 220
            .Width = 480
            .Height = 142 '179
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstElements, UIControl))

        'btnAdd initial props
        btnAdd = New UIButton(oUILib)
        With btnAdd
            .ControlName = "btnAdd"
            .Left = 5
            .Top = lstElements.Top + lstElements.Height + 5
            .Width = 180
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Add Current Selection"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Adds units currently selected in the" & vbCrLf & "environment to the selected battlegroup."
        End With
        Me.AddChild(CType(btnAdd, UIControl))

        'btnRemove initial props
        btnRemove = New UIButton(oUILib)
        With btnRemove
            .ControlName = "btnRemove"
            .Left = 310
            .Top = btnAdd.Top
            .Width = 180
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Remove Selected"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Detaches the selected elements from the battlegroup."
        End With
        Me.AddChild(CType(btnRemove, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 467
            .Top = 2
            .Width = 22
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to Close"
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = btnClose.Left - 23
            .Top = btnClose.Top
            .Width = 22
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to begin the tutorial for this window"
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'btnOrders initial props
        btnOrders = New UIButton(oUILib)
        With btnOrders
            .ControlName = "btnOrders"
            .Left = 185 + 10
            .Top = btnAdd.Top
            .Width = 105
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Orders"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to give the battlegroup orders"
        End With
        Me.AddChild(CType(btnOrders, UIControl))

        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = Me.BorderLineWidth
            .Top = btnAdd.Top + btnAdd.Height + 3
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'lblReinforcers initial props
        lblReinforcers = New UILabel(oUILib)
        With lblReinforcers
            .ControlName = "lblReinforcers"
            .Left = 5
            .Top = lnDiv3.Top + 2
            .Width = 368
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Reinforcing Facilities"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblReinforcers, UIControl))

        'lstReinforcers initial props
        lstReinforcers = New UIListBox(oUILib)
        With lstReinforcers
            .ControlName = "lstReinforcers"
            .Left = 5
            .Top = lblReinforcers.Top + lblReinforcers.Height + 2
            .Width = 480
            .Height = Me.Height - 7 - .Top
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstReinforcers, UIControl))

        'btnRemoveReinforcer initial props
        btnRemoveReinforcer = New UIButton(oUILib)
        With btnRemoveReinforcer
            .ControlName = "btnRemoveReinforcer"
            .Left = 310 'lstReinforcers.Left + lstReinforcers.Width - 180
            .Top = lnDiv3.Top + 1
            .Width = 180
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Remove Reinforcer"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Removes the selected reinforcing facility from this battlegroup."
        End With
        Me.AddChild(CType(btnRemoveReinforcer, UIControl))

        'lblDefaultFormation initial props
        lblDefaultFormation = New UILabel(oUILib)
        With lblDefaultFormation
            .ControlName = "lblDefaultFormation"
            .Left = 0
            .Top = 200
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Formation:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Sets the default formation for the battlegroup of units within the same environment."
        End With
        Me.AddChild(CType(lblDefaultFormation, UIControl))

        'btnFormations initial props
        btnFormations = New UIButton(oUILib)
        With btnFormations
            .ControlName = "btnFormations"
            .Left = lstElements.Left + lstElements.Width - 30
            .Top = lblDefaultFormation.Top
            .Width = 30
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "..."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to go to the Formation Management window."
        End With
        Me.AddChild(CType(btnFormations, UIControl))

        cboDefaultFormation = New UIComboBox(oUILib)
        With cboDefaultFormation
            .ControlName = "cboDefaultFormation"
            .Left = lstElements.Left + lstElements.Width - 190
            .Top = lblDefaultFormation.Top
            .Width = 155
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
            .ToolTipText = "Sets the default formation for the battlegroup of units within the same environment."
        End With
        Me.AddChild(CType(cboDefaultFormation, UIControl))
        lblDefaultFormation.Left = cboDefaultFormation.Left - lblDefaultFormation.Width

        FillFormationList()

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewBattleGroups) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack rights to view the battlegroups interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        If mlReinforceFacilityID <> -1 Then
            Me.Height = 155
            btnCreate.Caption = "Reinforce"
            lblTitle.Caption = "Battlegroup Reinforcement Configuration"
        End If

        mbIgnoreFormation = False

        mbLoading = False
    End Sub

    Private Sub FillFormationList()
        Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.oFormations, goCurrentPlayer.lFormationIdx, GetSortedIndexArrayType.eFormationName)
        If lSorted Is Nothing = False Then
            cboDefaultFormation.Clear()
            cboDefaultFormation.AddItem("No Formation")
            cboDefaultFormation.ItemData(cboDefaultFormation.NewIndex) = -1
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                cboDefaultFormation.AddItem(goCurrentPlayer.oFormations(lSorted(X)).sName)
                cboDefaultFormation.ItemData(cboDefaultFormation.NewIndex) = goCurrentPlayer.oFormations(lSorted(X)).FormationID
            Next X
            cboDefaultFormation.ListIndex = 0
        End If
    End Sub

    Public Sub RefreshFleetList()
        Dim sValue As String
        Try
            lstFleet.Clear()
            lstElements.Clear()

            mlMoverCnt = 0
            Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.moUnitGroups, goCurrentPlayer.mlUnitGroupIdx, GetSortedIndexArrayType.eFleetName)
            If lSorted Is Nothing = False Then
                For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                    If goCurrentPlayer.mlUnitGroupIdx(lSorted(X)) <> -1 Then
                        With goCurrentPlayer.moUnitGroups(lSorted(X))
                            sValue = .sName.PadRight(23, " "c)

                            If .lInterSystemOriginID <> -1 AndAlso .lInterSystemTargetID <> -1 AndAlso .iParentTypeID = ObjectType.eGalaxy Then
                                'Moving
                                mlMoverCnt += 1
                                sValue &= "Moving to " & goGalaxy.GetSystemName(.lInterSystemTargetID) & " (" & .GetInterSystemMovementETA & ")"
                            ElseIf .lInterSystemOriginID <> -1 AndAlso .lInterSystemTargetID <> -1 Then
                                sValue &= "Regrouping for new destination"
                                Dim sTemp As String = ""
                                For y As Int32 = 0 To goGalaxy.mlSystemUB
                                    If goGalaxy.moSystems(y).ObjectID = .lInterSystemTargetID Then
                                        sTemp = goGalaxy.moSystems(y).SystemName
                                        Exit For
                                    End If
                                Next
                                If sTemp = "" Then sTemp = GetCacheObjectValue(.lInterSystemTargetID, ObjectType.eSolarSystem)
                                If sTemp <> "Unknown" AndAlso sTemp <> "" Then sValue &= " (" & sTemp & ")"
                            Else
                                'TODO: Put more here as more becomes determined
                                'Deployed
                                If .iParentTypeID = ObjectType.eSolarSystem Then
                                    sValue &= "Deployed In " & goGalaxy.GetSystemName(.lParentID)
                                ElseIf .iParentTypeID = ObjectType.ePlanet Then
                                    Dim sTemp As String = goGalaxy.GetPlanetName(.lParentID)
                                    If sTemp = "" Then sTemp = GetCacheObjectValue(.lParentID, .iParentTypeID)
                                    If sTemp = "Unknown" Then sTemp = "A Planet"
                                    sValue &= "Deployed On " & sTemp

                                Else
                                    Dim sTemp As String = goGalaxy.GetPlanetName(.lParentID)
                                    If sTemp = "" Then sTemp = GetCacheObjectValue(.lParentID, .iParentTypeID)
                                    If sTemp = "Unknown" Then sTemp = "A Planet"
                                    sValue &= "In a Hangar (" & sTemp & ")"
                                End If
                            End If

                            lstFleet.AddItem(sValue)
                            lstFleet.ItemData(lstFleet.NewIndex) = .ObjectID
                        End With
                    End If
                Next X
            End If
            If lstFleet.ListIndex = -1 AndAlso lstFleet.ListCount >= 0 Then
                'If btnAdd.Enabled = False Then btnAdd.Enabled = True
                'If btnOrders.Enabled = False Then btnOrders.Enabled = True
                'If btnRemove.Enabled = False Then btnRemove.Enabled = True
                'If btnRemoveReinforcer.Enabled = False Then btnRemoveReinforcer.Enabled = True
                'If btnDelete.Enabled = False Then btnDelete.Enabled = True
                'If txtName.Enabled = False Then txtName.Enabled = True
                'lstFleet.ListIndex = 0
            ElseIf lstFleet.ListIndex >= lstFleet.ListCount Then
                'lstFleet.ListIndex = lstFleet.ListCount - 1
            ElseIf lstFleet.ListCount > 0 Then
                FillElementList()
            End If
            If btnRename.Enabled = True Then btnRename.Enabled = False

        Catch
        End Try

    End Sub

	Private Sub btnAdd_Click(ByVal sName As String) Handles btnAdd.Click
		If lstFleet.ListIndex = -1 Then Return
		If lstFleet.ItemData(lstFleet.ListIndex) < 1 Then Return

		If HasAliasedRights(AliasingRights.eModifyBattleGroups) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter battlegroups.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		'Add the currently selected units in the environment to the fleet/army
		Dim yMsg() As Byte
		Dim lPos As Int32 = 0
		Dim lCnt As Int32 = 0

		Dim oFleet As UnitGroup = Nothing
		Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)

		For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
			If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
				oFleet = goCurrentPlayer.moUnitGroups(X)
				Exit For
			End If
		Next X

		For X As Int32 = 0 To goCurrentEnvir.lEntityUB
			If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
				Dim bFound As Boolean = False
				For Y As Int32 = 0 To oFleet.lUnitUB
					If oFleet.uUnitIDs(Y).lUnitID = goCurrentEnvir.oEntity(X).ObjectID Then
						bFound = True
						Exit For
					End If
				Next Y
				If bFound = False Then lCnt += 1
			End If
		Next X

		If oFleet.iParentTypeID = ObjectType.eGalaxy Then
			goUILib.AddNotification("This battlegroup is in system-to-system movement and cannot be altered.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim lVal As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eMaxBattlegroupUnits)
		If lVal < 255 AndAlso lCnt + lstElements.ListCount > lVal Then
			goUILib.AddNotification("The maximum number of units able to be assigned to a battlegroup is " & lVal & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
			Return
		End If

		ReDim yMsg(9 + (lCnt * 4))

		System.BitConverter.GetBytes(GlobalMessageCode.eAddToFleet).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(lstFleet.ItemData(lstFleet.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

		For X As Int32 = 0 To goCurrentEnvir.lEntityUB
			If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
				Dim bFound As Boolean = False
				For Y As Int32 = 0 To oFleet.lUnitUB
					If oFleet.uUnitIDs(Y).lUnitID = goCurrentEnvir.oEntity(X).ObjectID Then
						bFound = True
						Exit For
					End If
				Next Y
				If bFound = False Then
					System.BitConverter.GetBytes(goCurrentEnvir.lEntityIdx(X)).CopyTo(yMsg, lPos) : lPos += 4

					'Now, this typically doesn't fail, so add them manually
					If oFleet Is Nothing = False Then oFleet.AddUnit(goCurrentEnvir.lEntityIdx(X), goCurrentEnvir.ObjectID, goCurrentEnvir.ObjTypeID)
				End If
			End If
		Next X

		MyBase.moUILib.SendMsgToPrimary(yMsg)
		FillElementList()

	End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

	Private Sub btnCreate_Click(ByVal sName As String) Handles btnCreate.Click
		'Create a new fleet/army with the currently selected units in the environment
		Dim yMsg() As Byte
		Dim lPos As Int32 = 0
		Dim lCnt As Int32 = 0

		If mlReinforceFacilityID <> -1 Then
			'Ok, selecting a battlegroup to reinforce
			If lstFleet.ListIndex > -1 Then
				Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)

				lPos = 0
				ReDim yMsg(9)
				System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetReinforcer).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(lFleetID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(mlReinforceFacilityID).CopyTo(yMsg, lPos) : lPos += 4
				MyBase.moUILib.SendMsgToPrimary(yMsg)
				MyBase.moUILib.RemoveWindow(Me.ControlName)
			Else : MyBase.moUILib.AddNotification("Select a battlegroup to reinforce.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			End If
			Return
		End If

		If HasAliasedRights(AliasingRights.eCreateBattleGroups) = False Then
			MyBase.moUILib.AddNotification("You lack rights to create battlegroups.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		If goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eMaxBattlegroups) < lstFleet.ListCount + 1 Then
			goUILib.AddNotification("Maximum number of battlegroups have been created!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
			Return
		End If

		For X As Int32 = 0 To goCurrentEnvir.lEntityUB
			If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
				lCnt += 1
			End If
		Next X

		If lCnt = 0 Then
			goUILib.AddNotification("Units must be selected in the current environment to create a new battlegroup.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim lVal As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eMaxBattlegroupUnits)
		If lVal < 255 AndAlso lCnt > goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eMaxBattlegroupUnits) Then
			goUILib.AddNotification("The maximum number of units able to assigned to a battlegroup is " & lVal, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
			Return
		End If

		ReDim yMsg(11 + (lCnt * 4))

		System.BitConverter.GetBytes(GlobalMessageCode.eCreateFleet).CopyTo(yMsg, lPos) : lPos += 2
		goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
		System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

		For X As Int32 = 0 To goCurrentEnvir.lEntityUB
			If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
				System.BitConverter.GetBytes(goCurrentEnvir.oEntity(X).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			End If
		Next X

        MyBase.moUILib.SendMsgToPrimary(yMsg)

	End Sub

	Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
		If lstFleet.ListIndex = -1 Then Return
		If lstFleet.ItemData(lstFleet.ListIndex) < 1 Then Return

		Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)

		If HasAliasedRights(AliasingRights.eModifyBattleGroups) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter battlegroups.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
        For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
            If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                If goCurrentPlayer.moUnitGroups(X).iParentTypeID = ObjectType.eGalaxy Then
                    goUILib.AddNotification("This battlegroup is in system-to-system movement and cannot be altered.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
            End If
        Next

        If btnDelete.Caption.ToUpper <> "CONFIRM" Then
            MyBase.moUILib.AddNotification("Press Confirm again to confirm disbanding the fleet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            btnDelete.Caption = "Confirm"
        Else
            'Delete the currently selected fleet/army (and set elements' unitgroupid = -1)
            Dim yMsg(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eDeleteFleet).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(lFleetID).CopyTo(yMsg, 2)
            MyBase.moUILib.SendMsgToPrimary(yMsg)

            For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                    goCurrentPlayer.moUnitGroups(X) = Nothing
                    goCurrentPlayer.mlUnitGroupIdx(X) = -1
                End If
            Next
            btnDelete.Caption = "Disband"
        End If
    End Sub

	Private Sub btnRemove_Click(ByVal sName As String) Handles btnRemove.Click
		If lstFleet.ListIndex = -1 Then Return
		If lstFleet.ItemData(lstFleet.ListIndex) < 1 Then Return
		If lstElements.ListIndex = -1 Then Return
		If lstElements.ItemData(lstElements.ListIndex) < 1 Then Return

		If HasAliasedRights(AliasingRights.eModifyBattleGroups) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter battlegroups.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		'Remove lstElements selected item from the selected fleet
		Dim yMsg() As Byte
		Dim lPos As Int32 = 0
		Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)
		Dim lUnitID As Int32 = lstElements.ItemData(lstElements.ListIndex)

		For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
			If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
				If goCurrentPlayer.moUnitGroups(X).iParentTypeID = ObjectType.eGalaxy Then
					goUILib.AddNotification("This battlegroup is in system-to-system movement and cannot be altered.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
			End If
		Next

		ReDim yMsg(9)

		System.BitConverter.GetBytes(GlobalMessageCode.eRemoveFromFleet).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(lFleetID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lUnitID).CopyTo(yMsg, lPos) : lPos += 4

		MyBase.moUILib.SendMsgToPrimary(yMsg)

		For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
			If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
				goCurrentPlayer.moUnitGroups(X).RemoveUnit(lUnitID)
				Exit For
			End If
		Next X
		FillElementList()
	End Sub

	Private Sub btnRename_Click(ByVal sName As String) Handles btnRename.Click
		If lstFleet.ListIndex = -1 Then Return
		If lstFleet.ItemData(lstFleet.ListIndex) < 1 Then Return

		If HasAliasedRights(AliasingRights.eModifyBattleGroups) = False Then
			MyBase.moUILib.AddNotification("You lack rights to modify battlegroups.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim sNewVal As String = Mid$(txtName.Caption, 1, 20)
        'For X As Int32 = 0 To lstFleet.ListCount - 1
        '    If lstFleet.List(X).ToUpper.StartsWith(sNewVal.ToUpper) = True Then
        '        goUILib.AddNotification("A battlegroup with that name already exists!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '        Return
        '    End If
        'Next X
        If sNewVal.trim = "" Then Return
        For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
            If goCurrentPlayer.mlUnitGroupIdx(X) <> -1 AndAlso goCurrentPlayer.moUnitGroups(X).sName = sNewVal Then
                goUILib.AddNotification("A battlegroup with that name already exists!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
        Next

		'Set UnitGroup's Name to txtName.Caption
		Dim yMsg(27) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityName).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(lstFleet.ItemData(lstFleet.ListIndex)).CopyTo(yMsg, 2)
		System.BitConverter.GetBytes(ObjectType.eUnitGroup).CopyTo(yMsg, 6)
		System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(yMsg, 8)
        BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eEntityChangeName, lstFleet.ItemData(lstFleet.ListIndex), ObjectType.eUnitGroup)
		MyBase.moUILib.SendMsgToPrimary(yMsg)

        Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)
        'lstFleet.List(lstFleet.ListIndex) = sNewVal

		For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
			If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
				goCurrentPlayer.moUnitGroups(X).sName = sNewVal
				Exit For
			End If
        Next
        RefreshFleetList()
	End Sub

	Private Sub frmFleet_OnNewFrame() Handles Me.OnNewFrame
		Dim lCnt As Int32 = 0

		If goCurrentPlayer Is Nothing Then Return

		For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
			If goCurrentPlayer.mlUnitGroupIdx(X) <> -1 Then lCnt += 1
		Next X

        'If lCnt = 0 OrElse lstFleet.ListIndex = -1 Then
        '    'Fleets
        '    If btnOrders.Enabled = True Then btnOrders.Enabled = False
        '    If btnDelete.Enabled = True Then btnDelete.Enabled = False
        '    If txtName.Enabled = True Then txtName.Enabled = False
        '    If txtName.Caption <> "" Then txtName.Caption = ""
        '    'Elements
        '    If btnRemove.Enabled = True Then btnRemove.Enabled = False
        '    If btnAdd.Enabled = True Then btnAdd.Enabled = False
        '    If btnRemoveReinforcer.Enabled = True Then btnRemoveReinforcer.Enabled = False
        '    If lstElements.ListCount > 0 Then lstElements.Clear()
        'Else
        '    If btnOrders.Enabled = False Then btnOrders.Enabled = True
        '    If btnDelete.Enabled = False Then btnDelete.Enabled = True
        'End If
        'If lstElements.ListCount = 0 OrElse lstElements.ListIndex = -1 Then
        '    If btnRemove.Enabled = True Then btnRemove.Enabled = False
        'Else
        '    If btnRemove.Enabled = False Then btnRemove.Enabled = True
        'End If
        'If lstReinforcers.ListCount = 0 OrElse lstReinforcers.ListIndex = -1 Then
        '    If btnRemoveReinforcer.Enabled = True Then btnRemoveReinforcer.Enabled = False
        'Else
        '    If btnRemoveReinforcer.Enabled = False Then btnRemoveReinforcer.Enabled = True
        'End If


        If lCnt <> lstFleet.ListCount Then
            RefreshFleetList()
        ElseIf mbHasUnknown = True Then
            FillElementList()
        Else
            If lstFleet.ListIndex <> -1 Then
                Dim lFleet As Int32 = lstFleet.ItemData(lstFleet.ListIndex)
                If lFleet <> -1 Then
                    For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                        If goCurrentPlayer.mlUnitGroupIdx(X) = lFleet Then
                            If goCurrentPlayer.moUnitGroups(X).LastMessageUpdate <> mlLastMsgUpdate Then
                                FillElementList()
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
        End If

        If mlMoverCnt <> 0 AndAlso (glCurrentCycle - mlLastMoverUpdate > 3) Then
            mlLastMoverUpdate = glCurrentCycle
            Dim lPrevMvrCnt As Int32 = mlMoverCnt

            mlMoverCnt = 0

            For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                If goCurrentPlayer.mlUnitGroupIdx(X) <> -1 Then
                    With goCurrentPlayer.moUnitGroups(X)
                        If .lInterSystemTargetID <> -1 AndAlso .lInterSystemOriginID <> -1 AndAlso .iParentTypeID = ObjectType.eGalaxy Then
                            'Ok, update this one only...
                            For lLstIdx As Int32 = 0 To lstFleet.ListCount - 1
                                If lstFleet.ItemData(lLstIdx) = .ObjectID Then

                                    Dim sValue As String = .sName.PadRight(23, " "c)
                                    mlMoverCnt += 1
                                    sValue &= "Moving to " & goGalaxy.GetSystemName(.lInterSystemTargetID)
                                    Dim iLen As Int32 = sValue.Length

                                    sValue = sValue & (Space(50 - iLen))
                                    sValue &= " ("
                                    iLen = (11 - .GetInterSystemMovementETA.Length)
                                    sValue = sValue & (Space(iLen))

                                    sValue &= .GetInterSystemMovementETA & ")"

                                    lstFleet.List(lLstIdx) = sValue
                                    Exit For
                                End If
                            Next lLstIdx
                        End If
                    End With
                End If
            Next X

            If mlMoverCnt <> lPrevMvrCnt Then RefreshFleetList()
        End If

        If lstFleet.ListIndex > -1 AndAlso goCurrentPlayer Is Nothing = False Then
            Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)
            Dim oUnitGroup As UnitGroup = Nothing
            For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                    oUnitGroup = goCurrentPlayer.moUnitGroups(X)
                    Exit For
                End If
            Next X
            If oUnitGroup Is Nothing = False Then
                For X As Int32 = 0 To oUnitGroup.mlReinforcerUB
                    If oUnitGroup.muReinforcers(X).lUnitID <> -1 Then
                        Dim bFound As Boolean = False
                        Dim lUnitID As Int32 = oUnitGroup.muReinforcers(X).lUnitID
                        Dim sEntry As String = GetCacheObjectValue(lUnitID, ObjectType.eFacility).PadRight(21, " "c) & GetCacheObjectValue(oUnitGroup.muReinforcers(X).lParentID, oUnitGroup.muReinforcers(X).iParentTypeID)
                        For Y As Int32 = 0 To lstReinforcers.ListCount - 1
                            If lstReinforcers.ItemData(Y) = lUnitID Then
                                If lstReinforcers.List(Y) <> sEntry Then lstReinforcers.List(Y) = sEntry
                                bFound = True
                                Exit For
                            End If
                        Next Y
                        If bFound = False Then
                            lstReinforcers.AddItem(sEntry, False)
                            lstReinforcers.ItemData(lstReinforcers.NewIndex) = lUnitID
                        End If
                    End If
                Next X
                'Now, do a reverse lookup
                Dim bDone As Boolean = False
                While bDone = False
                    bDone = True
                    For X As Int32 = 0 To lstReinforcers.ListCount - 1
                        Dim bFound As Boolean = False
                        Dim lUnitID As Int32 = lstReinforcers.ItemData(X)
                        For Y As Int32 = 0 To oUnitGroup.mlReinforcerUB
                            If oUnitGroup.muReinforcers(Y).lUnitID = lUnitID Then
                                bFound = True
                                Exit For
                            End If
                        Next Y
                        If bFound = False Then
                            bDone = False
                            lstReinforcers.RemoveItem(X)
                            Exit For
                        End If
                    Next X
                End While
            End If
        End If
        Dim iGroupCount As Int32 = lstFleet.ListCount
        Dim iMaxGroups As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eMaxBattlegroups)
        Dim sTemp As String = "Battlegroups: " & iGroupCount.ToString & " (" & iMaxGroups.ToString & " max)"
        'If iGroupCount >= iMaxGroups Then
        '    If btnCreate.Enabled = True Then
        '        btnCreate.Enabled = False
        '    End If
        'ElseIf btnCreate.Enabled = False Then
        '    btnCreate.Enabled = True
        'End If

        If lblSelect.Caption <> sTemp Then lblSelect.Caption = sTemp
        Dim iElementCount As Int32 = lstElements.ListCount
        Dim iMaxElements As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eMaxBattlegroupUnits)
        sTemp = "Elements: " & iElementCount & " (" & iMaxElements.ToString & " max)"
        'If iElementCount >= iMaxElements Then
        '    If btnAdd.Enabled = True Then
        '        btnAdd.Enabled = False
        '    End If
        'ElseIf btnAdd.Enabled = False AndAlso lstFleet.ListIndex > -1 Then
        '    btnAdd.Enabled = True
        'End If
        If lblElements.Caption <> sTemp Then lblElements.Caption = sTemp
        'If lstReinforcers.ListIndex >= lstReinforcers.ListCount Then
        '    lstReinforcers.ListIndex = lstReinforcers.ListCount - 1
        'End If
        
    End Sub

	Private Sub lstFleet_ItemClick(ByVal lIndex As Integer) Handles lstFleet.ItemClick
		FillElementList()
        lstReinforcers.Clear()
        If btnDelete.Caption <> "Disband" Then btnDelete.Caption = "Disband"
        If btnRename.Enabled = True Then btnRename.Enabled = False
    End Sub

	Private Sub lstFleet_ItemDblClick(ByVal lIndex As Integer) Handles lstFleet.ItemDblClick
		If goCurrentEnvir Is Nothing OrElse goUILib Is Nothing Then Return
		If lIndex = -1 Then Return

		'Get our fleet id
		Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)

		'Clear our selection
		goCurrentEnvir.DeselectAll()

		'Now, select what we can
		Dim oFleet As UnitGroup = Nothing
		For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
			If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
				oFleet = goCurrentPlayer.moUnitGroups(X)
				Exit For
			End If
		Next X

		If oFleet Is Nothing = False Then
			For X As Int32 = 0 To oFleet.lUnitUB
				Dim lTmpID As Int32 = oFleet.uUnitIDs(X).lUnitID
				If lTmpID <> -1 Then
					For Y As Int32 = 0 To goCurrentEnvir.lEntityUB
						If goCurrentEnvir.lEntityIdx(Y) = lTmpID AndAlso goCurrentEnvir.oEntity(Y).ObjTypeID = ObjectType.eUnit Then
							'Ok, found it
							goCurrentEnvir.oEntity(Y).bSelected = True
							goUILib.AddSelection(Y)
							Exit For
						End If
					Next Y
				End If
            Next X

            Dim oMultiSelect As frmMultiDisplay = CType(goUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
            If oMultiSelect Is Nothing = False Then oMultiSelect.SetFormationID(oFleet.lDefaultFormationID)
		End If
	End Sub

	Private Sub FillElementList()
		If lstFleet.ListIndex = -1 Then Return

		Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)
		Dim sValue As String
		Dim sParent As String

		mbHasUnknown = False

		lstElements.Clear()

		For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
			If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                With goCurrentPlayer.moUnitGroups(X)
                    mbIgnoreFormation = True
                    cboDefaultFormation.FindComboItemData(.lDefaultFormationID)
                    mbIgnoreFormation = False
                    For Y As Int32 = 0 To .lUnitUB
                        If .uUnitIDs(Y).lUnitID <> -1 Then
                            sValue = GetCacheObjectValue(.uUnitIDs(Y).lUnitID, ObjectType.eUnit)
                            mbHasUnknown = mbHasUnknown OrElse (sValue = "Unknown")
                            sValue = sValue.PadRight(23, " "c)

                            If .uUnitIDs(Y).iParentTypeID = ObjectType.eUnit OrElse .uUnitIDs(Y).iParentTypeID = ObjectType.eFacility Then
                                sParent = "In a Hangar"
                            ElseIf .uUnitIDs(Y).iParentTypeID = ObjectType.eSolarSystem Then
                                sParent = goGalaxy.GetSystemName(.uUnitIDs(Y).lParentID)
                                If sParent = "" Then sParent = "Unknown"
                            ElseIf .uUnitIDs(Y).iParentTypeID = ObjectType.ePlanet Then
                                sParent = goGalaxy.GetPlanetName(.uUnitIDs(Y).lParentID)
                                If sParent = "" Then sParent = GetCacheObjectValue(.uUnitIDs(Y).lParentID, .uUnitIDs(Y).iParentTypeID)
                                If sParent = "" Then sParent = "Unknown"
                            Else : sParent = GetCacheObjectValue(.uUnitIDs(Y).lParentID, .uUnitIDs(Y).iParentTypeID)
                            End If
                            mbHasUnknown = mbHasUnknown OrElse (sParent = "Unknown")
                            sValue &= sParent

                            lstElements.AddItem(sValue)
                            lstElements.ItemData(lstElements.NewIndex) = .uUnitIDs(Y).lUnitID
                        End If
                    Next Y

                    mlLastMsgUpdate = .LastMessageUpdate
                    'lblDetails.Caption = "Battlegroup Name     Total Warpoint Upkeep: "
                    'If .lWarpointUpkeep > -1 Then
                    '    lblDetails.Caption = "Battlegroup Name     Total Warpoint Upkeep: " & .lWarpointUpkeep.ToString("#,##0")
                    'End If

                    txtName.Caption = .sName
                End With
			End If
        Next X

        If lstElements.ListIndex = -1 AndAlso lstElements.ListIndex > -1 Then
            'If btnAdd.Enabled = False Then btnAdd.Enabled = True
            'If btnRemove.Enabled = False Then btnRemove.Enabled = True
            'If btnRemoveReinforcer.Enabled = False Then btnRemoveReinforcer.Enabled = True
            'lstElements.ListIndex = 0
        ElseIf lstElements.ListIndex >= lstElements.ListCount Then
            'lstElements.ListIndex = lstElements.ListCount - 1
        ElseIf lstElements.ListCount = 0 Then
            'If btnAdd.Enabled = True Then btnAdd.Enabled = False
            'If btnRemove.Enabled = True Then btnRemove.Enabled = False
            'If btnRemoveReinforcer.Enabled = True Then btnRemoveReinforcer.Enabled = False
        End If
    End Sub

	Private Sub btnOrders_Click(ByVal sName As String) Handles btnOrders.Click
		If lstFleet.ListIndex = -1 Then Return
		If lstFleet.ItemData(lstFleet.ListIndex) < 1 Then Return

		Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)

		Dim ofrmOrders As frmFleetOrders = CType(MyBase.moUILib.GetWindow("frmFleetOrders"), frmFleetOrders)
		If ofrmOrders Is Nothing Then ofrmOrders = New frmFleetOrders(goUILib)

        ofrmOrders.Visible = True
        ofrmOrders.Left = Me.Left - ofrmOrders.Width - 2
        ofrmOrders.Top = Me.Top
        ofrmOrders.SetFromFleet(lFleetID)
	End Sub

	Private Sub lstElements_ItemClick(ByVal lIndex As Integer) Handles lstElements.ItemClick
		If lstElements.ListIndex = -1 Then Return
		If goCurrentEnvir Is Nothing Then Return

		Dim oEnvir As BaseEnvironment = goCurrentEnvir
		Dim lUnitID As Int32 = lstElements.ItemData(lIndex)

		oEnvir.DeselectAll()

		For X As Int32 = 0 To oEnvir.lEntityUB
			If oEnvir.lEntityIdx(X) = lUnitID AndAlso oEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
				oEnvir.oEntity(X).bSelected = True
				goUILib.AddSelection(X)
				Exit For
			End If
        Next X
    End Sub

	Private Sub lstElements_ItemDblClick(ByVal lIndex As Integer) Handles lstElements.ItemDblClick
        moLaunchUnit = Nothing
        If lstElements.ListIndex = -1 Then Return
		If goCurrentEnvir Is Nothing Then Return
		If lstFleet.ListIndex = -1 Then Return

		Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)
		Dim lUnitID As Int32 = lstElements.ItemData(lstElements.ListIndex)
		Dim oUnitGroup As UnitGroup = Nothing

		For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
			If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
				oUnitGroup = goCurrentPlayer.moUnitGroups(X)
				Exit For
			End If
		Next X

		If oUnitGroup Is Nothing Then Return

		For X As Int32 = 0 To oUnitGroup.lUnitUB
			If oUnitGroup.uUnitIDs(X).lUnitID = lUnitID Then

				mlEntityID = lUnitID
				miEntityTypeID = ObjectType.eUnit

				If oUnitGroup.uUnitIDs(X).lParentID = goCurrentEnvir.ObjectID AndAlso oUnitGroup.uUnitIDs(X).iParentTypeID = goCurrentEnvir.ObjTypeID Then
					'Yes, we are... so no need to wait for switch
					FinishClickToEvent()
                Else
                    If oUnitGroup.uUnitIDs(X).iParentTypeID = ObjectType.eUnit OrElse oUnitGroup.uUnitIDs(X).iParentTypeID = ObjectType.eFacility Then

                        moLaunchUnit = oUnitGroup.uUnitIDs(X)
                        Dim oFrm As New frmMsgBox(goUILib, "Unable to go to that unit because it is in a hangar.  Remotely undock the unit now?", MsgBoxStyle.YesNo, "Confirm undock")
                        oFrm.Visible = True
                        AddHandler oFrm.DialogClosed, AddressOf LaunchUnitResult
                        'MyBase.moUILib.AddNotification("Unable to go to that unit because it is in a hangar.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    ElseIf oUnitGroup.uUnitIDs(X).iParentTypeID <> ObjectType.eSolarSystem AndAlso oUnitGroup.uUnitIDs(X).iParentTypeID <> ObjectType.ePlanet Then
                        MyBase.moUILib.AddNotification("Unable to go to that unit because it is not in a system or planet environment.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                    'no, change to that environment... but first set our select state
                    MyBase.moUILib.lUISelectState = UILib.eSelectState.eBattlegroupJumpTo
                    frmMain.ForceChangeEnvironment(oUnitGroup.uUnitIDs(X).lParentID, oUnitGroup.uUnitIDs(X).iParentTypeID)
                End If
                Exit For
            End If
		Next X
	End Sub

    Private Sub LaunchUnitResult(ByVal lResult As MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            If HasAliasedRights(AliasingRights.eDockUndockUnits) = False Then
                goUILib.AddNotification("You lack rights to dock/undock units.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                moLaunchUnit = Nothing
                Return
            End If
            Dim yData(14) As Byte

            System.BitConverter.GetBytes(GlobalMessageCode.eRequestUndock).CopyTo(yData, 0)
            'Send our object's GUID
            System.BitConverter.GetBytes(moLaunchUnit.lUnitID).CopyTo(yData, 2)
            System.BitConverter.GetBytes(ObjectType.eUnit).CopyTo(yData, 6)
            'Send our parent object's GUID
            System.BitConverter.GetBytes(moLaunchUnit.lParentID).CopyTo(yData, 8)
            System.BitConverter.GetBytes(moLaunchUnit.iParentTypeID).CopyTo(yData, 12)

            yData(14) = 0
            If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, SoundMgr.SpeechType.eUndockRequest Or SoundMgr.SpeechType.eSingleSelect)
            MyBase.moUILib.SendMsgToRegion(yData)
        End If
        moLaunchUnit = Nothing
    End Sub

    Private mlEntityID As Int32 = -1
    Private miEntityTypeID As Int16 = -1
    Public Sub FinishClickToEvent()
        goCamera.TrackingIndex = -1


        If goCurrentEnvir Is Nothing = False Then
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                glCurrentEnvirView = CurrentView.ePlanetMapView
            Else : glCurrentEnvirView = CurrentView.eSystemMapView1
            End If

            With goCamera
                .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                .mlCameraX = 0 : .mlCameraY = 1000 : .mlCameraZ = -1000
            End With

            If miEntityTypeID > -1 AndAlso mlEntityID > 0 Then
                'Ok, now find that item...

                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) = mlEntityID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = miEntityTypeID Then

                        If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                            glCurrentEnvirView = CurrentView.ePlanetView
                            goCamera.mlCameraY = 1700
                        Else
                            glCurrentEnvirView = CurrentView.eSystemView
                            goCamera.mlCameraY = 3000
                        End If

                        With goCamera
                            .mlCameraAtX = CInt(goCurrentEnvir.oEntity(X).LocX) : .mlCameraAtY = 0 : .mlCameraAtZ = CInt(goCurrentEnvir.oEntity(X).LocZ)
                            .mlCameraX = .mlCameraAtX : .mlCameraZ = .mlCameraAtZ - 500

                            Try
                                If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                    goCamera.mlCameraY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraAtX, goCamera.mlCameraAtZ, True)) + muSettings.lPlanetViewCameraY

                                    goCamera.mlCameraX += muSettings.lPlanetViewCameraX
                                    goCamera.mlCameraZ += muSettings.lPlanetViewCameraZ
                                End If
                            Catch
                            End Try
                        End With

                        'goCamera.TrackingIndex = X
                        Exit For
                    End If
                Next X
            End If
        End If

    End Sub

    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
        If goTutorial Is Nothing Then goTutorial = New TutorialManager()
        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eBattlegroup)
    End Sub

    Private Sub btnRemoveReinforcer_Click(ByVal sName As String) Handles btnRemoveReinforcer.Click
        If lstReinforcers.ListIndex > -1 AndAlso lstFleet.ListIndex > -1 Then
            Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)
            Dim yMsg(9) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetReinforcer).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lFleetID).CopyTo(yMsg, lPos) : lPos += 4
            Dim lFacID As Int32 = lstReinforcers.ItemData(lstReinforcers.ListIndex)
            lFacID = -lFacID
            System.BitConverter.GetBytes(lFacID).CopyTo(yMsg, lPos) : lPos += 4
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            Dim oUnitGroup As UnitGroup = Nothing
            For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                    oUnitGroup = goCurrentPlayer.moUnitGroups(X)
                    Exit For
                End If
            Next X
            If oUnitGroup Is Nothing = False Then oUnitGroup.RemoveReinforcer(Math.Abs(lFacID))
        Else : MyBase.moUILib.AddNotification("Select a reinforcing facility first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub cboDefaultFormation_ItemSelected(ByVal lItemIndex As Integer) Handles cboDefaultFormation.ItemSelected
        If mbIgnoreFormation = True Then Return

        Dim lFormationID As Int32 = -1
        If cboDefaultFormation.ListIndex > -1 Then lFormationID = cboDefaultFormation.ItemData(cboDefaultFormation.ListIndex)

        If lstFleet.ListIndex > -1 Then
            Dim lFleetID As Int32 = lstFleet.ItemData(lstFleet.ListIndex)
            Dim oFleet As UnitGroup = Nothing
            For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                    oFleet = goCurrentPlayer.moUnitGroups(X)
                    Exit For
                End If
            Next X

            If oFleet Is Nothing = False Then
                If oFleet.lDefaultFormationID <> lFormationID Then
                    oFleet.lDefaultFormationID = lFormationID

                    Dim yMsg(9) As Byte
                    Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetFormation).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(oFleet.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oFleet.lDefaultFormationID).CopyTo(yMsg, lPos) : lPos += 4
                    MyBase.moUILib.SendMsgToPrimary(yMsg)
                End If
            End If
        End If

    End Sub

    Private Sub FormationWindowClosed(ByVal lID As Int32)
        FillFormationList()
    End Sub

    Private Sub btnFormations_Click(ByVal sName As String) Handles btnFormations.Click

        If goCurrentPlayer Is Nothing Then Return

        Dim ofrm As frmFormations = CType(goUILib.GetWindow("frmFormations"), frmFormations)
        If ofrm Is Nothing = False Then
            If ofrm.Visible = True Then
                goUILib.RemoveWindow(ofrm.ControlName)
            Else
                ofrm.Visible = True
                AddHandler ofrm.FormationWindowClosed, AddressOf FormationWindowClosed
            End If
        Else
            ofrm = New frmFormations(goUILib)
            ofrm.Visible = True
            AddHandler ofrm.FormationWindowClosed, AddressOf FormationWindowClosed
        End If
        ofrm = Nothing

    End Sub

    Private Sub frmFleet_WindowClosed() Handles Me.WindowClosed
        Dim ofrmOrders As frmFleetOrders = CType(MyBase.moUILib.GetWindow("frmFleetOrders"), frmFleetOrders)
        If ofrmOrders Is Nothing = False AndAlso ofrmOrders.Visible = True Then
            ofrmOrders.RemoveMe()
        End If
        ofrmOrders = Nothing
    End Sub

    Private Sub frmFleet_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.FleetX = Me.Left
            muSettings.FleetY = Me.Top
        End If
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmFleetOrders")
        If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
            ofrm.Left = Me.Left - ofrm.Width - 2
            ofrm.Top = Me.Top
        End If
        ofrm = Nothing
    End Sub

    Private Sub txtName_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtName.OnKeyPress
        If btnRename.Enabled = False Then btnRename.Enabled = True
    End Sub

    Private Sub txtName_TextChanged() Handles txtName.TextChanged
        'Stop
    End Sub
End Class