Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAttachments
	Inherits UIWindow

	Private lstAttachments As UIListBox
	Private WithEvents btnView As UIButton
	Private WithEvents btnClose As UIButton
	Private WithEvents btnAdd As UIButton
	Private WithEvents btnDelete As UIButton

	Private moEvent As GuildEvent

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmAttachments initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAttachments
            .ControlName = "frmAttachments"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            .Width = 256
            .Height = 256
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 2
        End With

		'lstAttachments initial props
		lstAttachments = New UIListBox(oUILib)
		With lstAttachments
			.ControlName = "lstAttachments"
			.Left = 5
			.Top = 5
			.Width = 246
			.Height = 190
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		Me.AddChild(CType(lstAttachments, UIControl))

		'btnView initial props
		btnView = New UIButton(oUILib)
		With btnView
			.ControlName = "btnView"
			.Left = 5
			.Top = 227
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "View"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnView, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 151
			.Top = 227
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

		'btnAdd initial props
		btnAdd = New UIButton(oUILib)
		With btnAdd
			.ControlName = "btnAdd"
			.Left = 5
			.Top = 200
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Add"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnAdd, UIControl))

		'btnDelete initial props
		btnDelete = New UIButton(oUILib)
		With btnDelete
			.ControlName = "btnDelete"
			.Left = 151
			.Top = 200
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Delete"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnDelete, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

	Private Sub btnAdd_Click(ByVal sName As String) Handles btnAdd.Click
		MyBase.moUILib.yRenderUI = 2
		moEvent.AttachmentUB += 1
		ReDim Preserve moEvent.Attachments(moEvent.AttachmentUB)
		moEvent.Attachments(moEvent.AttachmentUB) = New SketchPad
		Dim ofrm As New frmSketchPad(MyBase.moUILib)
		Dim yPos As Byte = CByte(moEvent.AttachmentUB)
		ofrm.SetExternalData(moEvent.EventID, moEvent.Attachments(moEvent.AttachmentUB))
		ofrm.Visible = True
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
		If lstAttachments Is Nothing = False AndAlso lstAttachments.ListIndex > -1 Then
			If btnDelete.Caption.ToUpper = "DELETE" Then
				btnDelete.Caption = "Confirm"
			Else
				Dim yMsg(9) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eRemoveEventAttachment).CopyTo(yMsg, 0)
				System.BitConverter.GetBytes(moEvent.EventID).CopyTo(yMsg, 2)
				System.BitConverter.GetBytes(lstAttachments.ItemData(lstAttachments.ListIndex)).CopyTo(yMsg, 6)
				MyBase.moUILib.SendMsgToPrimary(yMsg)
				lstAttachments.RemoveItem(lstAttachments.ListIndex)

				btnDelete.Caption = "Delete"
			End If
		End If
	End Sub

	Private Sub btnView_Click(ByVal sName As String) Handles btnView.Click
		If lstAttachments.ListIndex > -1 Then
			Dim lID As Int32 = lstAttachments.ItemData(lstAttachments.ListIndex)
			Dim oSketchPad As SketchPad = moEvent.GetAttachment(lID)
			If oSketchPad Is Nothing = False Then
				'ok, view the sketchpad...
				Dim ofrm As New frmSketchPad(MyBase.moUILib)
				ofrm.SetExternalData(moEvent.EventID, oSketchPad)
			End If
		End If
	End Sub

	Public Sub SetFromEvent(ByRef oEvent As GuildEvent)
		lstAttachments.Clear()
		moEvent = oEvent
		If oEvent Is Nothing = False Then
			For X As Int32 = 0 To oEvent.AttachmentUB
				If oEvent.Attachments(X) Is Nothing = False Then
					With oEvent.Attachments(X)
						lstAttachments.AddItem(.sName, False)
						lstAttachments.ItemData(lstAttachments.NewIndex) = .lID
					End With
				End If
			Next X
		End If
	End Sub

	Private Sub frmAttachments_OnNewFrame() Handles Me.OnNewFrame
		If moEvent Is Nothing = False Then
			Dim lCnt As Int32 = 0
			For X As Int32 = 0 To moEvent.AttachmentUB
				If moEvent.Attachments(X) Is Nothing = False AndAlso moEvent.Attachments(X).lID > 0 Then
					lCnt += 1
				End If
			Next X
			If lCnt <> lstAttachments.ListCount Then
				lstAttachments.Clear()
				For X As Int32 = 0 To moEvent.AttachmentUB
					If moEvent.Attachments(X) Is Nothing = False AndAlso moEvent.Attachments(X).lID > 0 Then
						lstAttachments.AddItem(moEvent.Attachments(X).sName)
						lstAttachments.ItemData(lstAttachments.NewIndex) = moEvent.Attachments(X).lID
					End If
				Next X
			End If
		End If
	End Sub
End Class