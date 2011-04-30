Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmBugMain_old
	Inherits UIWindow

	Private WithEvents lstBugs As UIListBox
	Private WithEvents btnRefresh As UIButton
	Private WithEvents btnNew As UIButton
	Private WithEvents btnOpen As UIButton
	Private WithEvents btnClose As UIButton
	Private WithEvents btnHelp As UIButton
	Private lblCols As UILabel

	'Form's properties....
	Private moBugs() As BugEntry
	Private mlBugUB As Int32 = -1

	Private mlForcefulRedrawDelay As Int32 = 0

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmBugMain initial props
		With Me
			.ControlName = "frmBugMain"
			.Left = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 475
			.Top = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 225
			.Width = 950
			.Height = 450
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = True	'always show this form...
			.mbAcceptReprocessEvents = True
		End With

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

		'lstBugs initial props
		lstBugs = New UIListBox(oUILib)
		With lstBugs
			.ControlName = "lstBugs"
			.Left = 10
			.Top = 25
			.Width = 930
			.Height = 384
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
			.mbAcceptReprocessEvents = True
		End With
		Me.AddChild(CType(lstBugs, UIControl))

		'btnRefresh initial props
		btnRefresh = New UIButton(oUILib)
		With btnRefresh
			.ControlName = "btnRefresh"
			.Left = 10
			.Top = 420
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Refresh List"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRefresh, UIControl))

		'btnNew initial props
		btnNew = New UIButton(oUILib)
		With btnNew
			.ControlName = "btnNew"
			.Left = 600
			.Top = 420
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "New Bug"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnNew, UIControl))

		'btnOpen initial props
		btnOpen = New UIButton(oUILib)
		With btnOpen
			.ControlName = "btnOpen"
			.Left = 720
			.Top = 419
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Open/View"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnOpen, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 840
			.Top = 418
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Close"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnClose, UIControl))

		'lblCols initial props
		lblCols = New UILabel(oUILib)
		With lblCols
			.ControlName = "lblCols"
			.Left = 13
			.Top = 5
			.Width = 930
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "ID  Category      Priority       Description                                                    Status        User"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
		End With
		Me.AddChild(CType(lblCols, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.RemoveWindow("frmBugDetails")
	End Sub

	Private Sub RefreshList()
		Dim X As Int32
		Dim sLine As String

		lstBugs.Clear()

		For X = 0 To mlBugUB
			With moBugs(X)
				'ID  Category      Priority       Description                                                    Status        User
				sLine = .lID.ToString.PadRight(4, " "c)
				sLine &= BugEntry.GetCategory(.yCategory).PadRight(14, " "c)
				sLine &= BugEntry.GetPriority(.yPriority).PadRight(15, " "c)
				sLine &= Mid$(.sProblemDesc.Replace(vbCrLf, " "), 1, 60).PadRight(63, " "c)
				sLine &= BugEntry.GetStatus(.yStatus).PadRight(14, " "c)
				sLine &= .sUser
				lstBugs.AddItem(sLine)

				If .lAssignedToID = glPlayerID OrElse .lCreatedByUserID = glPlayerID Then
					lstBugs.ItemBold(lstBugs.NewIndex) = True
				Else : lstBugs.ItemBold(lstBugs.NewIndex) = False
				End If

				lstBugs.ItemData(lstBugs.NewIndex) = .lID
			End With
		Next X
	End Sub

	Private Sub btnNew_Click(ByVal sName As String) Handles btnNew.Click
		'Show frmBugDetails in BugEntry mode
		Dim ofrmBugDetails As frmBugDetails = New frmBugDetails(goUILib)
		ofrmBugDetails.Visible = True
		ofrmBugDetails = Nothing
		Me.Visible = False
	End Sub

	Private Sub btnOpen_Click(ByVal sName As String) Handles btnOpen.Click
		'Show frmBugDetails in BugDetail mode passing the selected bug in
		If lstBugs.ListIndex > -1 Then
			Dim ofrmBugDetails As New frmBugDetails(goUILib)
			ofrmBugDetails.Visible = True
			For X As Int32 = 0 To mlBugUB
				If moBugs(X).lID = lstBugs.ItemData(lstBugs.ListIndex) Then
					ofrmBugDetails.LoadFromBug(moBugs(X))
					Exit For
				End If
			Next X
			ofrmBugDetails = Nothing
			Me.Visible = False
		Else
			MyBase.moUILib.AddNotification("Please select a bug first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End If
	End Sub

	Private Sub btnRefresh_Click(ByVal sName As String) Handles btnRefresh.Click
		'Refresh the bug list (send a new request to the server)
		mlBugUB = -1
		Erase moBugs
		RefreshList()

		Dim yMsg(1) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eBugList).CopyTo(yMsg, 0)
		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Public Sub HandleAddBugToList(ByVal yData() As Byte)
		Dim X As Int32
		Dim lIdx As Int32 = -1
		Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim lLen As Int32
		Dim lPos As Int32

		For X = 0 To mlBugUB
			If moBugs(X).lID = lID Then
				'Ok, just work with this one
				lIdx = X
				Exit For
			End If
		Next X

		If lIdx = -1 Then
			'ok, create one...
			mlBugUB += 1
			ReDim Preserve moBugs(mlBugUB)
			lIdx = mlBugUB
		End If

		If moBugs(lIdx) Is Nothing Then moBugs(lIdx) = New BugEntry

		With moBugs(lIdx)
			.lID = lID
			.yCategory = yData(6)
			.ySubCat = yData(7)
			.yOccurs = yData(8)
			.yPriority = yData(9)
			.yStatus = yData(10)

			lPos = 11
			.lCreatedByUserID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lAssignedToID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			'Username always takes up 20 chars
			.sUser = GetStringFromBytes(yData, lPos, 20) : lPos += 20

			'Now, next 2 bytes is length
			lLen = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.sProblemDesc = GetStringFromBytes(yData, lPos, lLen)
			lPos += lLen

			lLen = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.sStepsToProduce = GetStringFromBytes(yData, lPos, lLen)
			lPos += lLen

			lLen = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.sDevNotes = GetStringFromBytes(yData, lPos, lLen)
			lPos += lLen
		End With

		RefreshList()

	End Sub

	Private Sub frmBugMain_OnNewFrame() Handles Me.OnNewFrame
		If glCurrentCycle - mlForcefulRedrawDelay > 30 Then
			mlForcefulRedrawDelay = glCurrentCycle
			Me.IsDirty = True
		End If
	End Sub

	Public Sub UpdateBugData(ByRef oBug As BugEntry)

		If oBug.lID < 1 Then
			btnRefresh_Click(btnRefresh.ControlName)
			Return
		End If

		For X As Int32 = 0 To mlBugUB
			If moBugs(X).lID = oBug.lID Then
				moBugs(X) = oBug
				With moBugs(X)
					Dim sLine As String
					sLine = .lID.ToString.PadRight(4, " "c)
					sLine &= BugEntry.GetCategory(.yCategory).PadRight(14, " "c)
					sLine &= BugEntry.GetPriority(.yPriority).PadRight(15, " "c)
					sLine &= Mid$(.sProblemDesc.Replace(vbCrLf, " "), 1, 60).PadRight(63, " "c)
					sLine &= BugEntry.GetStatus(.yStatus).PadRight(14, " "c)
					sLine &= .sUser

					For Y As Int32 = 0 To lstBugs.ListCount - 1
						If lstBugs.ItemData(Y) = .lID Then
							lstBugs.List(Y) = sLine
							Exit For
						End If
					Next Y
				End With
				Exit For
			End If
		Next X
	End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eBugMainWindow)
	End Sub
End Class

'Interface created from Interface Builder
Public Class frmBugMain
	Inherits UIWindow

	Private txtText As UITextBox
	Private WithEvents btnClose As UIButton
	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmBugMain initial props
		With Me
			.ControlName = "frmBugMain"
			.Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
			.Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 64
			.Width = 512
			.Height = 128
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = True
			.BorderLineWidth = 2
		End With

		'lblText initial props
		txtText = New UITextBox(oUILib)
		With txtText
			.ControlName = "txtText"
			.Left = 10
			.Top = 10
			.Width = 496
			.Height = 80
			.Enabled = True
			.Visible = True
			.Caption = "The F12 Bug Entry Interface is presently offline while it is converted to interface with the website. To enter new bugs or manage existing bugs, please visit the issue tracker at http://www.beyondprotocol.com/issuetracker and use the same login you use for the website."
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.Center
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.MultiLine = True
			.Locked = True
		End With
		Me.AddChild(CType(txtText, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = Me.Width \ 2 - 50
			.Top = 100
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Close"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnClose, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub


	Public Sub HandleAddBugToList(ByVal yData() As Byte)


	End Sub

	Public Sub UpdateBugData(ByRef oBug As BugEntry)

	End Sub

End Class