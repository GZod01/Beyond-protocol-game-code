Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmFormations
	Inherits UIWindow

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private lblTotal As UILabel
	Private lblNext As UILabel
	Private lblFormations As UILabel
	Private lblName As UILabel
	Private lblBasedOn As UILabel
	Private txtNextSlot As UITextBox
	Private txtName As UITextBox
	Private chkDefault As UICheckBox

	Private lblCellSize As UILabel
	Private txtCellSize As UITextBox

	Private WithEvents ctlSlots As UIFormation
	Private WithEvents btnClose As UIButton
	Private WithEvents optInsert As UIOption
	Private WithEvents optReplace As UIOption
	Private WithEvents optRemove As UIOption
	Private WithEvents btnClear As UIButton
	Private WithEvents btnSave As UIButton
	Private WithEvents cboFormations As UIComboBox
	Private WithEvents btnCreate As UIButton
	Private WithEvents btnCopy As UIButton
	Private WithEvents btnDelete As UIButton
	Private WithEvents cboSortBy As UIComboBox

    Private lblDefaultFormation As UILabel
    Private WithEvents optStar As UIOption
    Private WithEvents optHorizontal As UIOption
    Private WithEvents optVertical As UIOption
    Private WithEvents optMass As UIOption

	Private mbIgnoreNextSlotClick As Boolean = False

    Private mbWaitingForSave As Boolean = False
    Private mbLoading As Boolean = True

    Public Event FormationWindowClosed(ByVal lFormationID As Int32)

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenFormationsWindow)

		'frmFormations initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eFormations
            .ControlName = "frmFormations"
            '.Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            '.Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.FormationsX
                lTop = muSettings.FormationsY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256
            If lLeft + 512 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 512
            If lTop + 512 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 512

            .Left = lLeft
            .Top = lTop

            .Width = 512
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
        End With

		'ctlSlots initial props
		ctlSlots = New UIFormation(oUILib)
		With ctlSlots
			.ControlName = "ctlSlots"
			.Left = 5
			.Top = 80
			.Width = 400
			.Height = 400
			.Enabled = True
			.Visible = True
		End With
		Me.AddChild(CType(ctlSlots, UIControl))

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 5
			.Width = 149
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Formation Manager"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 487
			.Top = 2
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

		'lblTotal initial props
		lblTotal = New UILabel(oUILib)
		With lblTotal
			.ControlName = "lblTotal"
			.Left = 410
			.Top = 180
			.Width = 95
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Total Slots: 625"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(0, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTotal, UIControl))

		'lblNext initial props
		lblNext = New UILabel(oUILib)
		With lblNext
			.ControlName = "lblNext"
			.Left = 410
			.Top = 200
			.Width = 66
			.Height = 18
            .Enabled = True
			.Visible = True
			.Caption = "Next Slot:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(0, DrawTextFormat)
		End With
		Me.AddChild(CType(lblNext, UIControl))

		'txtNextSlot initial props
		txtNextSlot = New UITextBox(oUILib)
		With txtNextSlot
			.ControlName = "txtNextSlot"
			.Left = 425
			.Top = 220
			.Width = 65
			.Height = 18
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "1"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 3
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(txtNextSlot, UIControl))

		'chkDefault initial props
		chkDefault = New UICheckBox(oUILib)
		With chkDefault
			.ControlName = "chkDefault"
			.Left = ctlSlots.Left + ctlSlots.Width + 10
			.Top = 55
			.Width = 60
			.Height = 24
            .Enabled = True
            .Visible = False
			.Caption = "Default"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		Me.AddChild(CType(chkDefault, UIControl))

		'optInsert initial props
		optInsert = New UIOption(oUILib)
		With optInsert
			.ControlName = "optInsert"
			.Left = 425
			.Top = 240
			.Width = 45
			.Height = 24
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Insert"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		Me.AddChild(CType(optInsert, UIControl))

		'optReplace initial props
		optReplace = New UIOption(oUILib)
		With optReplace
			.ControlName = "optReplace"
			.Left = 425
			.Top = 260
			.Width = 65
			.Height = 24
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Replace"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = True
		End With
		Me.AddChild(CType(optReplace, UIControl))

		'optRemove initial props
		optRemove = New UIOption(oUILib)
		With optRemove
			.ControlName = "optRemove"
			.Left = 425
			.Top = 280
			.Width = 65
			.Height = 24
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		Me.AddChild(CType(optRemove, UIControl))

		'btnClear initial props
		btnClear = New UIButton(oUILib)
		With btnClear
			.ControlName = "btnClear"
			.Left = 5
			.Top = 485
			.Width = 100
			.Height = 24
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Clear"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnClear, UIControl))

		'btnSave initial props
		btnSave = New UIButton(oUILib)
		With btnSave
			.ControlName = "btnSave"
			.Left = 310
			.Top = 485
			.Width = 100
			.Height = 24
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Save"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnSave, UIControl))

		'lblFormations initial props
		lblFormations = New UILabel(oUILib)
		With lblFormations
			.ControlName = "lblFormations"
			.Left = 5
			.Top = 30
			.Width = 70
			.Height = 18
            .Enabled = True
			.Visible = True
			.Caption = "Formations:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblFormations, UIControl))

		'btnCreate initial props
		btnCreate = New UIButton(oUILib)
		With btnCreate
			.ControlName = "btnCreate"
			.Left = 235
			.Top = 28
			.Width = 85
			.Height = 24
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Add New"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnCreate, UIControl))

		'btnCopy initial props
		btnCopy = New UIButton(oUILib)
		With btnCopy
			.ControlName = "btnCopy"
			.Left = 325
			.Top = 28
			.Width = 85
			.Height = 24
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Copy"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnCopy, UIControl))

		'btnDelete initial props
		btnDelete = New UIButton(oUILib)
		With btnDelete
			.ControlName = "btnDelete"
			.Left = 415
			.Top = 28
			.Width = 85
			.Height = 24
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Delete"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnDelete, UIControl))

		'lblName initial props
		lblName = New UILabel(oUILib)
		With lblName
			.ControlName = "lblName"
			.Left = 5
			.Top = 55
			.Width = 42
			.Height = 18
            .Enabled = True
			.Visible = True
			.Caption = "Name:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblName, UIControl))

		'txtName initial props
		txtName = New UITextBox(oUILib)
		With txtName
			.ControlName = "txtName"
			.Left = 50
			.Top = 55
			.Width = 110
			.Height = 18
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "New Formation"
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

		'lblCellSize initial props
		lblCellSize = New UILabel(oUILib)
		With lblCellSize
			.ControlName = "lblCellSize"
			.Left = 410
			.Top = 320
			.Width = 90
			.Height = 18
            .Enabled = True
			.Visible = True
			.Caption = "Cell Size:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblCellSize, UIControl))

		'txtCellSize initial props
		txtCellSize = New UITextBox(oUILib)
		With txtCellSize
			.ControlName = "txtCellSize"
			.Left = 425
			.Top = 340
			.Width = 65
			.Height = 18
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "1"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 5
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(txtCellSize, UIControl))

		'lblBasedOn initial props
		lblBasedOn = New UILabel(oUILib)
		With lblBasedOn
			.ControlName = "lblBasedOn"
			.Left = 185
			.Top = 55
			.Width = 65
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Based On:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
        Me.AddChild(CType(lblBasedOn, UIControl))


        'lblDefaultFormation initial props
        lblDefaultFormation = New UILabel(oUILib)
        With lblDefaultFormation
            .ControlName = "lblDefaultFormation"
            .Left = lblCellSize.Left
            .Top = txtCellSize.Top + txtCellSize.Height + 10
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Default Placing:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDefaultFormation, UIControl))

        'optStar initial props
        optStar = New UIOption(oUILib)
        With optStar
            .ControlName = "optStar"
            .Left = txtCellSize.Left
            .Top = lblDefaultFormation.Top + lblDefaultFormation.Height + 2
            .Width = 38 '65
            .Height = 18
            .Enabled = Not gbAliased
            .Visible = True
            .Caption = "Star"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
        End With
        Me.AddChild(CType(optStar, UIControl))

        'optHorizontal initial props
        optHorizontal = New UIOption(oUILib)
        With optHorizontal
            .ControlName = "optHorizontal"
            .Left = optStar.Left
            .Top = optStar.Top + optStar.Height + 2
            .Width = 75 '65
            .Height = 18
            .Enabled = Not gbAliased
            .Visible = True
            .Caption = "Horizontal"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optHorizontal, UIControl))

        'optVertical initial props
        optVertical = New UIOption(oUILib)
        With optVertical
            .ControlName = "optVertical"
            .Left = optHorizontal.Left
            .Top = optHorizontal.Top + optHorizontal.Height + 2
            .Width = 61 ' 65
            .Height = 18
            .Enabled = Not gbAliased
            .Visible = True
            .Caption = "Vertical"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optVertical, UIControl))

        'optMass initial props
        optMass = New UIOption(oUILib)
        With optMass
            .ControlName = "optMass"
            .Left = optVertical.Left
            .Top = optVertical.Top + optVertical.Height + 2
            .Width = 48 '65
            .Height = 18
            .Enabled = Not gbAliased
            .Visible = True
            .Caption = "Mass"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optMass, UIControl))

		'cboSortBy initial props
		cboSortBy = New UIComboBox(oUILib)
		With cboSortBy
			.ControlName = "cboSortBy"
			.Left = 255
			.Top = 55
			.Width = 150
			.Height = 18
            .Enabled = Not gbAliased
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .l_ListBoxHeight = 185
        End With
        Me.AddChild(CType(cboSortBy, UIControl))

        'cboFormations initial props
        cboFormations = New UIComboBox(oUILib)
        With cboFormations
            .ControlName = "cboFormations"
            .Left = 80
            .Top = 30
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .l_ListBoxHeight = 430
        End With
        Me.AddChild(CType(cboFormations, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eMoveUnits) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack rights to move units.  Therefore you can't view formation settings.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        FillLists()
        mbLoading = False
	End Sub

	Private Sub FillFormations()
		Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.oFormations, goCurrentPlayer.lFormationIdx, GetSortedIndexArrayType.eFormationName)
		If lSorted Is Nothing = False Then
			cboFormations.Clear()
			For X As Int32 = 0 To lSorted.GetUpperBound(0)
				cboFormations.AddItem(goCurrentPlayer.oFormations(lSorted(X)).sName)
				cboFormations.ItemData(cboFormations.NewIndex) = goCurrentPlayer.oFormations(lSorted(X)).FormationID
			Next X
		End If
	End Sub

	Private Sub FillLists()
		cboFormations.Clear()
		cboSortBy.Clear()

		If goCurrentPlayer Is Nothing = False Then
			FillFormations()
		End If

		cboSortBy.AddItem("Cargo Bay Size")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eCargoBayCap
		cboSortBy.AddItem("Combat Rating")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eCombatRating
		cboSortBy.AddItem("Entity Name")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eEntityName
		cboSortBy.AddItem("Hangar Capacity")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eHangarCap
		cboSortBy.AddItem("Hull Size")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eHullSize
		cboSortBy.AddItem("Radar Range")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eMaxRadarRange
		cboSortBy.AddItem("Front Armor HP")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eMostFrontArmorHP
		cboSortBy.AddItem("Shield HP")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eMostShieldHP
		cboSortBy.AddItem("Maneuver")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eManeuver
		cboSortBy.AddItem("Speed")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eSpeed
		cboSortBy.AddItem("Weapons")
		cboSortBy.ItemData(cboSortBy.NewIndex) = CriteriaType.eWeaponSlots


	End Sub

	Private Sub btnClear_Click(ByVal sName As String) Handles btnClear.Click
		ctlSlots.Clear()
		txtNextSlot.Caption = "1"
		lblTotal.Caption = "Total Slots: 0"
	End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        Dim lID As Int32 = -1
        If cboFormations.ListIndex > -1 Then
            lID = cboFormations.ItemData(cboFormations.ListIndex)
        End If
        RaiseEvent FormationWindowClosed(lID)
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

	Private Sub btnCopy_Click(ByVal sName As String) Handles btnCopy.Click
		If cboFormations.ListIndex = -1 OrElse cboFormations.ItemData(cboFormations.ListIndex) < 1 Then
			MyBase.moUILib.AddNotification("Select a formation to copy.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		cboFormations.ListIndex = -1
		txtName.Caption = "New Formation"
	End Sub

	Private Sub btnCreate_Click(ByVal sName As String) Handles btnCreate.Click
		cboFormations.ListIndex = -1
		ctlSlots.Clear()
		txtName.Caption = "New Formation"
        cboSortBy.ListIndex = -1
        txtNextSlot.Caption = "1"
	End Sub

	Private Sub btnSave_Click(ByVal sName As String) Handles btnSave.Click
		If cboSortBy.ListIndex = -1 Then
			MyBase.moUILib.AddNotification("Select a property to base this formation on.", Color.Red, -1, 1 - 1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		If ctlSlots.GetSlotCount() < 2 Then
			MyBase.moUILib.AddNotification("This formation must consist of at least 2 elements.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim lResult As Int32 = ctlSlots.Validate()
		If lResult <> 0 Then
			MyBase.moUILib.AddNotification("There are multiple slots designated as " & lResult & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		If txtName.Caption.Trim = "" Then
			MyBase.moUILib.AddNotification("This formation must have a name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		If Val(txtCellSize.Caption) < 1 Then
			MyBase.moUILib.AddNotification("Cell Size cannot be less than 1.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		'Ok, let's build our msg
		Dim yGridData() As Byte = ctlSlots.GetMsgData()
		If yGridData Is Nothing Then Return

		Dim yMsg(31 + yGridData.Length) As Byte
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(GlobalMessageCode.eAddFormation).CopyTo(yMsg, lPos) : lPos += 2
		If cboFormations.ListIndex <> -1 Then
			System.BitConverter.GetBytes(cboFormations.ItemData(cboFormations.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
		Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
		End If

        'DEFAULT OR NOT!?!?
        If optStar.Value = True Then
            yMsg(lPos) = 0
        ElseIf optHorizontal.Value = True Then
            yMsg(lPos) = 1
        ElseIf optVertical.Value = True Then
            yMsg(lPos) = 2
        Else
            yMsg(lPos) = 3
        End If
        lPos += 1

		yMsg(lPos) = CByte(cboSortBy.ItemData(cboSortBy.ListIndex)) : lPos += 1
		System.Text.ASCIIEncoding.ASCII.GetBytes(txtName.Caption).CopyTo(yMsg, lPos) : lPos += 20

		Dim lCellSize As Int32 = CInt(Val(txtCellSize.Caption))
		System.BitConverter.GetBytes(lCellSize).CopyTo(yMsg, lPos) : lPos += 4

		yGridData.CopyTo(yMsg, lPos) : lPos += yGridData.Length

		mbWaitingForSave = True
		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Private Sub cboFormations_ItemSelected(ByVal lItemIndex As Integer) Handles cboFormations.ItemSelected
		If cboFormations.ListIndex <> -1 Then
			btnDelete.Caption = "Delete"

			Dim oFormation As FormationDef = Nothing
			Dim lFormationID As Int32 = cboFormations.ItemData(cboFormations.ListIndex)
			If goCurrentPlayer Is Nothing = False Then
				For X As Int32 = 0 To goCurrentPlayer.lFormationUB
					If goCurrentPlayer.lFormationIdx(X) = lFormationID Then
						oFormation = goCurrentPlayer.oFormations(X)
						Exit For
					End If
				Next X
			End If

			If oFormation Is Nothing = False Then
				ctlSlots.SetFromFormation(oFormation)
				Me.txtName.Caption = oFormation.sName
				cboSortBy.FindComboItemData(oFormation.yCriteria)
                'chkDefault.Value = oFormation.yDefault <> 0
                optStar.Value = (oFormation.yDefault = 0)
                optHorizontal.Value = (oFormation.yDefault = 1)
                optVertical.Value = (oFormation.yDefault = 2)
                optMass.Value = (oFormation.yDefault = 3)

				txtCellSize.Caption = oFormation.lCellSize.ToString
				txtNextSlot.Caption = (ctlSlots.GetMaxSlotValue + 1).ToString
				lblTotal.Caption = "Total Slots: " & ctlSlots.GetSlotCount
			End If
		End If
	End Sub

	Private Sub ctlSlots_SlotClick(ByVal lIndexX As Integer, ByVal lIndexY As Integer) Handles ctlSlots.SlotClick
		If mbIgnoreNextSlotClick = True Then
			mbIgnoreNextSlotClick = False
			Return
		End If
		If optInsert.Value = True Then
			'ok, insert... means we find this slot and shift all others...
			Dim lSlotVal As Int32 = CInt(Val(txtNextSlot.Caption))
			If lSlotVal < 1 Then Return
			ctlSlots.InsertSlot(lIndexX, lIndexY, lSlotVal)
		ElseIf optReplace.Value = True Then
			Dim lSlotVal As Int32 = CInt(Val(txtNextSlot.Caption))
			If lSlotVal < 1 Then Return
			ctlSlots.RemoveSlot(lSlotVal)
			ctlSlots.SetSlot(lIndexX, lIndexY, lSlotVal)
		Else
			'should be remove
			ctlSlots.RemoveAndShiftSlot(lIndexX, lIndexY)
		End If
		txtNextSlot.Caption = (ctlSlots.GetMaxSlotValue() + 1).ToString
		lblTotal.Caption = "Total Slots: " & ctlSlots.GetSlotCount
	End Sub

	Private Sub optInsert_Click() Handles optInsert.Click
		optRemove.Value = False
		optReplace.Value = False
	End Sub

	Private Sub optRemove_Click() Handles optRemove.Click
		optInsert.Value = False
		optReplace.Value = False
	End Sub

	Private Sub optReplace_Click() Handles optReplace.Click
		optInsert.Value = False
		optRemove.Value = False
	End Sub

	Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
		If cboFormations.ListIndex = -1 OrElse cboFormations.ItemData(cboFormations.ListIndex) < 1 Then
			MyBase.moUILib.AddNotification("Select a formation to delete.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		
		If btnDelete.Caption.ToUpper = "CONFIRM" Then
			Dim yMsg(5) As Byte
			Dim lPos As Int32 = 0
			Dim lFormationID As Int32 = cboFormations.ItemData(cboFormations.ListIndex)

			System.BitConverter.GetBytes(GlobalMessageCode.eRemoveFormation).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lFormationID).CopyTo(yMsg, lPos) : lPos += 4
			MyBase.moUILib.SendMsgToPrimary(yMsg)

			If goCurrentPlayer Is Nothing = False Then
				For X As Int32 = 0 To goCurrentPlayer.lFormationUB
					If goCurrentPlayer.lFormationIdx(X) = lFormationID Then
						goCurrentPlayer.lFormationIdx(X) = -1
					End If
				Next X
			End If

			btnDelete.Caption = "Delete"
		Else : btnDelete.Caption = "Confirm"
		End If
	End Sub

	Private Sub frmFormations_OnNewFrame() Handles Me.OnNewFrame
		'check our counts
		Dim lCnt As Int32 = 0
		If goCurrentPlayer Is Nothing = False Then
			For X As Int32 = 0 To goCurrentPlayer.lFormationUB
				If goCurrentPlayer.lFormationIdx(X) <> -1 Then lCnt += 1
			Next X
		End If
		If lCnt <> cboFormations.ListCount Then
			FillLists()
			If mbWaitingForSave = True Then
				mbWaitingForSave = False
				Dim sName As String = txtName.Caption
				For X As Int32 = 0 To goCurrentPlayer.lFormationUB
					If goCurrentPlayer.lFormationIdx(X) <> -1 Then
						If goCurrentPlayer.oFormations(X).sName = sName Then
							cboFormations.FindComboItemData(goCurrentPlayer.oFormations(X).FormationID)
							Exit For
						End If
					End If
				Next X
			End If
		End If
	End Sub

	Private Sub cboSortBy_ItemSelected(ByVal lItemIndex As Integer) Handles cboSortBy.ItemSelected
		mbIgnoreNextSlotClick = True
	End Sub

    Private Sub frmFormations_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.FormationsX = Me.Left
            muSettings.FormationsY = Me.Top
        End If
    End Sub

    Private Sub optStar_Click() Handles optStar.Click
        optStar.Value = True
        optHorizontal.Value = False
        optVertical.Value = False
        optMass.Value = False
    End Sub

    Private Sub optHorizontal_Click() Handles optHorizontal.Click
        optStar.Value = False
        optHorizontal.Value = True
        optVertical.Value = False
        optMass.Value = False
    End Sub

    Private Sub optVertical_Click() Handles optVertical.Click
        optStar.Value = False
        optHorizontal.Value = False
        optVertical.Value = True
        optMass.Value = False
    End Sub

    Private Sub optMass_Click() Handles optMass.Click
        optStar.Value = False
        optHorizontal.Value = False
        optVertical.Value = False
        optMass.Value = True
    End Sub
End Class