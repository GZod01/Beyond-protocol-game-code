Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmBugDetails
    Inherits UIWindow

    Private lblCategory As UILabel
    Private WithEvents cboCategory As UIComboBox
    Private lblTitle As UILabel
    Private lblSubCat As UILabel
    Private WithEvents cboSubCat As UIComboBox
    Private lblPriority As UILabel
    Private WithEvents cboPriority As UIComboBox
    Private Label9 As UILabel
    Private WithEvents txtUserName As UITextBox
    Private Label5 As UILabel
    Private WithEvents txtDescription As UITextBox
    Private lblSteps As UILabel
    Private WithEvents txtStepsToProduce As UITextBox
    Private lblDevNotes As UILabel
    Private WithEvents txtDevNotes As UITextBox
    Private lblOccurence As UILabel
    Private WithEvents cboOccurence As UIComboBox
    Private lblStatus As UILabel
    Private WithEvents cboStatus As UIComboBox
    Private WithEvents btnSubmit As UIButton
    Private WithEvents btnCancel As UIButton
    Private WithEvents btnHelp As UIButton

    Private lblAssignTo As UILabel
    Private WithEvents cboAssignTo As UIComboBox

    Private mlID As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmBugDetails initial props
        With Me
            .ControlName = "frmBugDetails"
            .Left = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 350
            .Top = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 267
            .Width = 700
            .Height = 440
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True  'show regardless...
        End With

        'lblCategory initial props
        lblCategory = New UILabel(oUILib)
        With lblCategory
            .ControlName = "lblCategory"
            .Left = 10
            .Top = 50
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Category:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCategory, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 10
            .Top = 10
            .Width = 146
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Bug Details"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lblSubCat initial props
        lblSubCat = New UILabel(oUILib)
        With lblSubCat
            .ControlName = "lblSubCat"
            .Left = 240
            .Top = 50
            .Width = 82
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Sub-Category:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSubCat, UIControl))

        'lblPriority initial props
        lblPriority = New UILabel(oUILib)
        With lblPriority
            .ControlName = "lblPriority"
            .Left = 500
            .Top = 50
            .Width = 82
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Priority:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPriority, UIControl))

        'lblAssignTo initial props
        lblAssignTo = New UILabel(oUILib)
        With lblAssignTo
            .ControlName = "lblAssignTo"
            .Left = 480
            .Top = 75
            .Width = 86
            .Height = 18
            .Enabled = True
			.Visible = False
            .Caption = "Assign To:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAssignTo, UIControl))

        'Label9 initial props
        Label9 = New UILabel(oUILib)
        With Label9
            .ControlName = "Label9"
            .Left = 440
            .Top = 20
            .Width = 116
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "Submitted By:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(Label9, UIControl))

        'txtUserName initial props
        txtUserName = New UITextBox(oUILib)
        With txtUserName
            .ControlName = "txtUserName"
            .Left = Label9.Left + 84
            .Top = 20
            .Width = 140
            .Height = 18
            .Enabled = False
            .Visible = False
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtUserName, UIControl))

        'Label5 initial props
        Label5 = New UILabel(oUILib)
        With Label5
            .ControlName = "Label5"
            .Left = 10
            .Top = 95
            .Width = 180
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Describe the Problem in Detail:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(Label5, UIControl))

        'txtDescription initial props
        txtDescription = New UITextBox(oUILib)
        With txtDescription
            .ControlName = "txtDescription"
            .Left = 10
            .Top = 115
            .Width = 679
            .Height = 131
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
            .MultiLine = True
        End With
        Me.AddChild(CType(txtDescription, UIControl))

        'lblSteps initial props
        lblSteps = New UILabel(oUILib)
        With lblSteps
            .ControlName = "lblSteps"
            .Left = 10
            .Top = 250
            .Width = 201
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Steps to Reproduce (if applicable):"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSteps, UIControl))

        'txtStepsToProduce initial props
        txtStepsToProduce = New UITextBox(oUILib)
        With txtStepsToProduce
            .ControlName = "txtStepsToProduce"
            .Left = 10
            .Top = 270
            .Width = 679
            .Height = 131
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
            .MultiLine = True
        End With
        Me.AddChild(CType(txtStepsToProduce, UIControl))

        'lblDevNotes initial props
        lblDevNotes = New UILabel(oUILib)
        With lblDevNotes
            .ControlName = "lblDevNotes"
            .Left = 10
            .Top = 405
            .Width = 201
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "Developer's Notes:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDevNotes, UIControl))

        'txtDevNotes initial props
        txtDevNotes = New UITextBox(oUILib)
        With txtDevNotes
            .ControlName = "txtDevNotes"
            .Left = 10
            .Top = 425
            .Width = 679
            .Height = 60
            .Enabled = True
            .Visible = False
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
        End With
        Me.AddChild(CType(txtDevNotes, UIControl))

        'lblOccurence initial props
        lblOccurence = New UILabel(oUILib)
        With lblOccurence
            .ControlName = "lblOccurence"
            .Left = 9
            .Top = 75
            .Width = 186
            .Height = 18
            .Enabled = True
			.Visible = False
            .Caption = "How often does this bug occur?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblOccurence, UIControl))

        'lblStatus initial props
        lblStatus = New UILabel(oUILib)
        With lblStatus
            .ControlName = "lblStatus"
            .Left = 240
            .Top = 20
            .Width = 40
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "Status:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblStatus, UIControl))

        'btnSubmit initial props
        btnSubmit = New UIButton(oUILib)
        With btnSubmit
            .ControlName = "btnSubmit"
            .Left = 460
            .Top = txtStepsToProduce.Top + txtStepsToProduce.Height + 10
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Submit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSubmit, UIControl))

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = 589
            .Top = btnSubmit.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Cancel"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCancel, UIControl))

        '================================ COMBO BOXES LAST ===============================
        'cboAssignTo initial props
        cboAssignTo = New UIComboBox(oUILib)
        With cboAssignTo
            .ControlName = "cboAssignTo"
            .Left = 550
            .Top = 75
            .Width = 140
            .Height = 18
            .Enabled = True
			.Visible = False
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboAssignTo, UIControl))

        'cboOccurence initial props
        cboOccurence = New UIComboBox(oUILib)
        With cboOccurence
            .ControlName = "cboOccurence"
            .Left = 203
            .Top = 75
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = False
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboOccurence, UIControl))

        'cboPriority initial props
        cboPriority = New UIComboBox(oUILib)
        With cboPriority
            .ControlName = "cboPriority"
            .Left = 550
            .Top = 50
            .Width = 140
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboPriority, UIControl))

        'cboSubCat initial props
        cboSubCat = New UIComboBox(oUILib)
        With cboSubCat
            .ControlName = "cboSubCat"
            .Left = 330
            .Top = 50
            .Width = 140
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboSubCat, UIControl))

        'cboCategory initial props
        cboCategory = New UIComboBox(oUILib)
        With cboCategory
            .ControlName = "cboCategory"
            .Left = 70
            .Top = 50
            .Width = 140
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboCategory, UIControl))

        'cboStatus initial props
        cboStatus = New UIComboBox(oUILib)
        With cboStatus
            .ControlName = "cboStatus"
            .Left = 290
            .Top = 20
            .Width = 140
            .Height = 18
            .Enabled = True
            .Visible = False
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboStatus, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        FillValues()
    End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)

		Dim oTmpWin As UIWindow = MyBase.moUILib.GetWindow("frmBugMain")
		If oTmpWin Is Nothing = False Then oTmpWin.Visible = True
		oTmpWin = Nothing
	End Sub

    Private Sub SetControlEnable(ByVal bAdmin As Boolean, ByVal bCreator As Boolean)
        cboCategory.bReadOnly = Not bAdmin
        cboSubCat.bReadOnly = Not bAdmin
        cboPriority.bReadOnly = Not bAdmin
        cboOccurence.bReadOnly = Not bAdmin
        cboStatus.bReadOnly = Not bAdmin
        txtDevNotes.Locked = Not bAdmin

        cboAssignTo.bReadOnly = Not bAdmin

        txtDescription.Locked = Not (bAdmin = True OrElse bCreator = True)
        txtStepsToProduce.Locked = Not (bAdmin = True OrElse bCreator = True)
		btnSubmit.Visible = False '(bAdmin = True OrElse bCreator = True)
    End Sub

    Public Sub LoadFromBug(ByRef oBug As BugEntry)
        'Now, fill in the blanks
        With oBug
            mlID = .lID
            If cboCategory.FindComboItemData(CInt(.yCategory)) = True Then
                cboSubCat.FindComboItemData(CInt(.ySubCat))
            End If
            cboPriority.FindComboItemData(CInt(.yPriority))
            cboOccurence.FindComboItemData(CInt(.yOccurs))
            txtDescription.Caption = .sProblemDesc
            txtStepsToProduce.Caption = .sStepsToProduce
            txtDevNotes.Caption = .sDevNotes
            cboStatus.FindComboItemData(CInt(.yStatus))
            txtUserName.Caption = .sUser
            cboAssignTo.FindComboItemData(.lAssignedToID)
        End With

        'set our controls for viewing the details... not entry...
		'Me.Height = 534
        Label9.Visible = True
        txtUserName.Visible = True
		'lblDevNotes.Visible = True
		'txtDevNotes.Visible = True
        lblStatus.Visible = True
        cboStatus.Visible = True
		'btnSubmit.Top = 500
		'btnCancel.Top = 500

		btnSubmit.Visible = False

		SetControlEnable(False, False) 'IsAdmin() OrElse oBug.lAssignedToID = glPlayerID, glPlayerID = oBug.lCreatedByUserID)
    End Sub

	Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click
		Dim oBug As BugEntry = New BugEntry()

		With oBug
			.lID = mlID
			.sDevNotes = txtDevNotes.Caption
			.sProblemDesc = txtDescription.Caption
			.sStepsToProduce = txtStepsToProduce.Caption
			.sUser = txtUserName.Caption
			If cboCategory.ListIndex > -1 Then .yCategory = CByte(cboCategory.ItemData(cboCategory.ListIndex))
			If cboOccurence.ListIndex > -1 Then .yOccurs = CByte(cboOccurence.ItemData(cboOccurence.ListIndex))
			If cboPriority.ListIndex > -1 Then .yPriority = CByte(cboPriority.ItemData(cboPriority.ListIndex))
			If cboStatus.ListIndex > -1 Then .yStatus = CByte(cboStatus.ItemData(cboStatus.ListIndex))
			If cboSubCat.ListIndex > -1 Then .ySubCat = CByte(cboSubCat.ItemData(cboSubCat.ListIndex))
			If cboAssignTo.ListIndex > -1 Then .lAssignedToID = cboAssignTo.ItemData(cboAssignTo.ListIndex)
		End With

		Dim yMsg() As Byte = oBug.GenerateSaveMsg()
		MyBase.moUILib.SendMsgToPrimary(yMsg)

		Dim oFrm As frmBugMain = CType(MyBase.moUILib.GetWindow("frmBugMain"), frmBugMain)
		If oFrm Is Nothing = False Then oFrm.UpdateBugData(oBug)
		oFrm = Nothing

		btnCancel_Click(btnCancel.ControlName)
	End Sub

    Private Sub cboCategory_ItemSelected(ByVal lItemIndex As Integer) Handles cboCategory.ItemSelected
        cboSubCat.Clear()
        cboSubCat.ListIndex = -1

        If lItemIndex < 0 Then Return

        Select Case cboCategory.ItemData(lItemIndex)
            Case 0
                cboSubCat.AddItem("Client-Side Crash") : cboSubCat.ItemData(cboSubCat.NewIndex) = 0
                cboSubCat.AddItem("Server-Side Crash") : cboSubCat.ItemData(cboSubCat.NewIndex) = 1
                cboSubCat.AddItem("Other") : cboSubCat.ItemData(cboSubCat.NewIndex) = 2
            Case 1
                cboSubCat.AddItem("Exploit") : cboSubCat.ItemData(cboSubCat.NewIndex) = 0
                cboSubCat.AddItem("Game Logic Bug") : cboSubCat.ItemData(cboSubCat.NewIndex) = 1
                cboSubCat.AddItem("Geography Bug") : cboSubCat.ItemData(cboSubCat.NewIndex) = 2
                cboSubCat.AddItem("Unexpected Result") : cboSubCat.ItemData(cboSubCat.NewIndex) = 3
                cboSubCat.AddItem("User Interface") : cboSubCat.ItemData(cboSubCat.NewIndex) = 4
                cboSubCat.AddItem("Other") : cboSubCat.ItemData(cboSubCat.NewIndex) = 5
            Case 2
                cboSubCat.AddItem("Geography") : cboSubCat.ItemData(cboSubCat.NewIndex) = 0
                cboSubCat.AddItem("Models (units and facilities)") : cboSubCat.ItemData(cboSubCat.NewIndex) = 1
                cboSubCat.AddItem("Particle FX") : cboSubCat.ItemData(cboSubCat.NewIndex) = 2
                cboSubCat.AddItem("Textures") : cboSubCat.ItemData(cboSubCat.NewIndex) = 3
                cboSubCat.AddItem("User Interface") : cboSubCat.ItemData(cboSubCat.NewIndex) = 4
                cboSubCat.AddItem("Other") : cboSubCat.ItemData(cboSubCat.NewIndex) = 5
            Case 3
                cboSubCat.AddItem("Low Frame Rate") : cboSubCat.ItemData(cboSubCat.NewIndex) = 0
                cboSubCat.AddItem("Memory Leak") : cboSubCat.ItemData(cboSubCat.NewIndex) = 1
                cboSubCat.AddItem("Perceived Lag") : cboSubCat.ItemData(cboSubCat.NewIndex) = 2
                cboSubCat.AddItem("Game Stutter") : cboSubCat.ItemData(cboSubCat.NewIndex) = 3
                cboSubCat.AddItem("Other") : cboSubCat.ItemData(cboSubCat.NewIndex) = 4
            Case 4
                cboSubCat.AddItem("Miscellaneous") : cboSubCat.ItemData(cboSubCat.NewIndex) = 0
            Case 5
                cboSubCat.AddItem("Login Credentials") : cboSubCat.ItemData(cboSubCat.NewIndex) = 0
                cboSubCat.AddItem("Connection Issue") : cboSubCat.ItemData(cboSubCat.NewIndex) = 1
                cboSubCat.AddItem("Other") : cboSubCat.ItemData(cboSubCat.NewIndex) = 2
            Case 6
                cboSubCat.AddItem("Gameplay-Related") : cboSubCat.ItemData(cboSubCat.NewIndex) = 0
                cboSubCat.AddItem("User Interface") : cboSubCat.ItemData(cboSubCat.NewIndex) = 1
                cboSubCat.AddItem("Hotkeys") : cboSubCat.ItemData(cboSubCat.NewIndex) = 2
                cboSubCat.AddItem("Other Neat Ideas") : cboSubCat.ItemData(cboSubCat.NewIndex) = 3
            Case 7
                cboSubCat.AddItem("Battlegroups") : cboSubCat.ItemData(cboSubCat.NewIndex) = 9
                cboSubCat.AddItem("Chat") : cboSubCat.ItemData(cboSubCat.NewIndex) = 13
                cboSubCat.AddItem("Colony/Budget") : cboSubCat.ItemData(cboSubCat.NewIndex) = 12
                cboSubCat.AddItem("Combat") : cboSubCat.ItemData(cboSubCat.NewIndex) = 0
                cboSubCat.AddItem("Diplomacy") : cboSubCat.ItemData(cboSubCat.NewIndex) = 11
                cboSubCat.AddItem("Espionage") : cboSubCat.ItemData(cboSubCat.NewIndex) = 17
                cboSubCat.AddItem("Guilds/Alliances") : cboSubCat.ItemData(cboSubCat.NewIndex) = 16
                cboSubCat.AddItem("In-Game Email") : cboSubCat.ItemData(cboSubCat.NewIndex) = 10
                cboSubCat.AddItem("Mining") : cboSubCat.ItemData(cboSubCat.NewIndex) = 7
                cboSubCat.AddItem("Movement") : cboSubCat.ItemData(cboSubCat.NewIndex) = 1
                cboSubCat.AddItem("Performance") : cboSubCat.ItemData(cboSubCat.NewIndex) = 3
                cboSubCat.AddItem("Production") : cboSubCat.ItemData(cboSubCat.NewIndex) = 4
                cboSubCat.AddItem("Repair") : cboSubCat.ItemData(cboSubCat.NewIndex) = 14
                cboSubCat.AddItem("Sound") : cboSubCat.ItemData(cboSubCat.NewIndex) = 18
                cboSubCat.AddItem("Space Stations") : cboSubCat.ItemData(cboSubCat.NewIndex) = 15
                cboSubCat.AddItem("Special Techs") : cboSubCat.ItemData(cboSubCat.NewIndex) = 6
                cboSubCat.AddItem("Tech Builder") : cboSubCat.ItemData(cboSubCat.NewIndex) = 5
                cboSubCat.AddItem("Trading") : cboSubCat.ItemData(cboSubCat.NewIndex) = 8
                cboSubCat.AddItem("User Interface") : cboSubCat.ItemData(cboSubCat.NewIndex) = 2
                cboSubCat.AddItem("Other - be specific") : cboSubCat.ItemData(cboSubCat.NewIndex) = 19
        End Select
    End Sub

    Private Sub FillValues()
        'Clear all of our combo boxes
        cboCategory.Clear() : cboOccurence.Clear() : cboPriority.Clear() : cboStatus.Clear() : cboSubCat.Clear()

        'Category combo box
        cboCategory.AddItem("Crash") : cboCategory.ItemData(cboCategory.NewIndex) = 0
        cboCategory.AddItem("Gameplay") : cboCategory.ItemData(cboCategory.NewIndex) = 1
        cboCategory.AddItem("Graphical") : cboCategory.ItemData(cboCategory.NewIndex) = 2
        cboCategory.AddItem("Performance") : cboCategory.ItemData(cboCategory.NewIndex) = 3
        cboCategory.AddItem("Other") : cboCategory.ItemData(cboCategory.NewIndex) = 4
        cboCategory.AddItem("Login") : cboCategory.ItemData(cboCategory.NewIndex) = 5
        cboCategory.AddItem("Suggestions") : cboCategory.ItemData(cboCategory.NewIndex) = 6
        cboCategory.AddItem("Test Case") : cboCategory.ItemData(cboCategory.NewIndex) = 7

        cboOccurence.AddItem("Easily Reproduceable") : cboOccurence.ItemData(cboOccurence.NewIndex) = 0
        cboOccurence.AddItem("Intermittent") : cboOccurence.ItemData(cboOccurence.NewIndex) = 1
        cboOccurence.AddItem("Rarely Occurs") : cboOccurence.ItemData(cboOccurence.NewIndex) = 2

        cboPriority.AddItem("Extremely High") : cboPriority.ItemData(cboPriority.NewIndex) = 0
        cboPriority.AddItem("High") : cboPriority.ItemData(cboPriority.NewIndex) = 1
        cboPriority.AddItem("Normal") : cboPriority.ItemData(cboPriority.NewIndex) = 2
        cboPriority.AddItem("Low") : cboPriority.ItemData(cboPriority.NewIndex) = 3
        cboPriority.AddItem("Extremely Low") : cboPriority.ItemData(cboPriority.NewIndex) = 4

        cboStatus.AddItem("New") : cboStatus.ItemData(cboStatus.NewIndex) = 0
        cboStatus.AddItem("Open") : cboStatus.ItemData(cboStatus.NewIndex) = 1
        cboStatus.AddItem("In Progress") : cboStatus.ItemData(cboStatus.NewIndex) = 2
        cboStatus.AddItem("Unreleased") : cboStatus.ItemData(cboStatus.NewIndex) = 3
        cboStatus.AddItem("Released") : cboStatus.ItemData(cboStatus.NewIndex) = 4
        cboStatus.AddItem("On Hold") : cboStatus.ItemData(cboStatus.NewIndex) = 5
        cboStatus.AddItem("Waiting") : cboStatus.ItemData(cboStatus.NewIndex) = 6
        cboStatus.AddItem("Closed") : cboStatus.ItemData(cboStatus.NewIndex) = 7

        cboAssignTo.AddItem("Unassigned") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 0
        cboAssignTo.AddItem("Aurelius") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 6
        cboAssignTo.AddItem("Csaj") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 2
		'cboAssignTo.AddItem("Digi") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 9
		'cboAssignTo.AddItem("Norma") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 7
		'cboAssignTo.AddItem("Nimnode") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 22
		'cboAssignTo.AddItem("Inian") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 61
		cboAssignTo.AddItem("Havoc") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 80
		cboAssignTo.AddItem("Mikey") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 191
		cboAssignTo.AddItem("Wraith") : cboAssignTo.ItemData(cboAssignTo.NewIndex) = 200
		cboAssignTo.FindComboItemData(0)

        cboAssignTo.bReadOnly = Not IsAdmin()
    End Sub

    Private Function IsAdmin() As Boolean
		Return glPlayerID = 1 OrElse glPlayerID = 6 OrElse glPlayerID = 9 OrElse glPlayerID = 2 OrElse glPlayerID = 80 OrElse glPlayerID = 191 OrElse glPlayerID = 200
    End Function

    Private Sub txtDescription_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.OnKeyDown, txtStepsToProduce.OnKeyDown
        If e.KeyCode = Keys.V AndAlso e.Control = True Then
            'ok, get our clipboard
            If Clipboard.ContainsText = True Then
                txtDescription.Caption &= Clipboard.GetText()
            End If
            e.Handled = True
        End If
    End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eBugEntryWindow)
	End Sub
End Class