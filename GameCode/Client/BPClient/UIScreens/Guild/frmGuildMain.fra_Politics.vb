'Option Strict On

'Imports Microsoft.DirectX
'Imports Microsoft.DirectX.Direct3D

'Partial Class frmGuildMain

'	'Interface created from Interface Builder
'	Private Class fra_Politics
'		Inherits guildframe


'		Private txtDetails As UITextBox
'		Private fraSpecifics As UIWindow
'		Private fraIcon As UIWindow
'		Private lblEntityName As UILabel
'		Private lblMemberCount As UILabel
'		Private lblGuildLeader As UILabel
'		Private lblNotes As UILabel
'		Private lblFirstContact As UILabel
'		Private txtFirstContactBy As UITextBox
'		Private lblFirstContactOn As UILabel
'		Private txtFirstContactOn As UITextBox
'		Private lblFirstContactAt As UILabel
'		Private txtFirstContactPlace As UITextBox
'		Private lblFirstContactWith As UILabel
'		Private txtFirstContactWith As UITextBox
'		Private lblRelStandings As UILabel
'		Private lnDiv1 As UILine
'		Private lblTowardsUs As UILabel
'		Private lblTowardsThem As UILabel

'		Private WithEvents lstRels As UIListBox
'		Private WithEvents btnGotoFCLoc As UIButton
'		Private WithEvents btnUpdateNotes As UIButton
'		Private WithEvents hscrRel As UIScrollBar
'		Private WithEvents btnSetRel As UIButton
'		Private WithEvents btnResetRel As UIButton

'		Private moSprite As Sprite = Nothing
'		Private moIconBack As Texture = Nothing
'		Private moIconFore As Texture = Nothing

'		'for rendering the icon
'		Private rcBack As Rectangle = Rectangle.Empty
'		Private rcFore1 As Rectangle
'		Private rcFore2 As Rectangle
'		Private clrBack As System.Drawing.Color
'		Private clrFore1 As System.Drawing.Color
'		Private clrFore2 As System.Drawing.Color
'		Private mlCurrIcon As Int32 = -1

'		Private mbLoading As Boolean = True

'		Public Sub New(ByRef oUILib As UILib)
'			MyBase.New(oUILib)

'			' initial props
'			With Me
'				.ControlName = "fra_Politics"
'				.Left = 15
'				.Top = ml_CONTENTS_TOP
'				.Width = 800
'				.Height = 600
'				.Enabled = True
'				.Visible = True
'				.BorderColor = muSettings.InterfaceBorderColor
'				.FillColor = muSettings.InterfaceFillColor
'				.FullScreen = False
'				.Moveable = False
'				.BorderLineWidth = 2
'				.mbAcceptReprocessEvents = True
'			End With

'			'NewControl1 initial props
'			lstRels = New UIListBox(oUILib)
'			With lstRels
'				.ControlName = "lstRels"
'				.Left = 15
'				.Top = 5
'				.Width = 225
'				.Height = 520
'				.Enabled = True
'				.Visible = True
'				.BorderColor = muSettings.InterfaceBorderColor
'				.FillColor = muSettings.InterfaceFillColor
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'				.mbAcceptReprocessEvents = True
'			End With
'			Me.AddChild(CType(lstRels, UIControl))

'			'fraSpecifics initial props
'			fraSpecifics = New UIWindow(oUILib)
'			With fraSpecifics
'				.ControlName = "fraSpecifics"
'				.Left = 250
'				.Top = 5
'				.Width = 540
'				.Height = 71
'				.Enabled = True
'				.Visible = True
'				.BorderColor = muSettings.InterfaceBorderColor
'				.FillColor = muSettings.InterfaceFillColor
'				.FullScreen = False
'				.BorderLineWidth = 2
'				.Moveable = False
'			End With
'			Me.AddChild(CType(fraSpecifics, UIControl))

'			'fraIcon initial props
'			fraIcon = New UIWindow(oUILib)
'			With fraIcon
'				.ControlName = "fraIcon"
'				.Left = 2
'				.Top = 2
'				.Width = 67
'				.Height = 67
'				.Enabled = True
'				.Visible = True
'				.BorderColor = muSettings.InterfaceBorderColor
'				.FillColor = muSettings.InterfaceFillColor
'				.FullScreen = False
'			End With
'			fraSpecifics.AddChild(CType(fraIcon, UIControl))

'			'lblEntityName initial props
'			lblEntityName = New UILabel(oUILib)
'			With lblEntityName
'				.ControlName = "lblEntityName"
'				.Left = 75
'				.Top = 5
'				.Width = 200
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Entity Name"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			fraSpecifics.AddChild(CType(lblEntityName, UIControl))

'			'lblMemberCount initial props
'			lblMemberCount = New UILabel(oUILib)
'			With lblMemberCount
'				.ControlName = "lblMemberCount"
'				.Left = 75
'				.Top = 25
'				.Width = 200
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = ""
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			fraSpecifics.AddChild(CType(lblMemberCount, UIControl))

'			'lblGuildLeader initial props
'			lblGuildLeader = New UILabel(oUILib)
'			With lblGuildLeader
'				.ControlName = "lblGuildLeader"
'				.Left = 75
'				.Top = 45
'				.Width = 200
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = ""
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			fraSpecifics.AddChild(CType(lblGuildLeader, UIControl))

'			'txtDetails initial props
'			txtDetails = New UITextBox(oUILib)
'			With txtDetails
'				.ControlName = "txtDetails"
'				.Left = 250
'				.Top = 100
'				.Width = 215
'				.Height = 390
'				.Enabled = True
'				.Visible = True
'				.Caption = ""
'				.ForeColor = muSettings.InterfaceTextBoxForeColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(0, DrawTextFormat)
'				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'				.MaxLength = 0
'				.BorderColor = muSettings.InterfaceBorderColor
'				.MultiLine = True
'			End With
'			Me.AddChild(CType(txtDetails, UIControl))

'			'lblNotes initial props
'			lblNotes = New UILabel(oUILib)
'			With lblNotes
'				.ControlName = "lblNotes"
'				.Left = 250
'				.Top = 80
'				.Width = 126
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Notes and Details"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			Me.AddChild(CType(lblNotes, UIControl))

'			'lblFirstContact initial props
'			lblFirstContact = New UILabel(oUILib)
'			With lblFirstContact
'				.ControlName = "lblFirstContact"
'				.Left = 480
'				.Top = 100
'				.Width = 126
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "First Contact By:"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			Me.AddChild(CType(lblFirstContact, UIControl))

'			'txtFirstContactBy initial props
'			txtFirstContactBy = New UITextBox(oUILib)
'			With txtFirstContactBy
'				.ControlName = "txtFirstContactBy"
'				.Left = 615
'				.Top = 100
'				.Width = 175
'				.Height = 22
'				.Enabled = True
'				.Visible = True
'				.Caption = "Enoch Dagor"
'				.ForeColor = muSettings.InterfaceTextBoxForeColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'				.MaxLength = 0
'				.BorderColor = muSettings.InterfaceBorderColor
'				.Locked = True
'			End With
'			Me.AddChild(CType(txtFirstContactBy, UIControl))

'			'lblFirstContactOn initial props
'			lblFirstContactOn = New UILabel(oUILib)
'			With lblFirstContactOn
'				.ControlName = "lblFirstContactOn"
'				.Left = 480
'				.Top = 170
'				.Width = 126
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "First Contact On:"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			Me.AddChild(CType(lblFirstContactOn, UIControl))

'			'txtFirstContactOn initial props
'			txtFirstContactOn = New UITextBox(oUILib)
'			With txtFirstContactOn
'				.ControlName = "txtFirstContactOn"
'				.Left = 615
'				.Top = 170
'				.Width = 175
'				.Height = 22
'				.Enabled = True
'				.Visible = True
'				.Caption = ""
'				.ForeColor = muSettings.InterfaceTextBoxForeColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'				.MaxLength = 0
'				.BorderColor = muSettings.InterfaceBorderColor
'				.Locked = True
'			End With
'			Me.AddChild(CType(txtFirstContactOn, UIControl))

'			'lblFirstContactAt initial props
'			lblFirstContactAt = New UILabel(oUILib)
'			With lblFirstContactAt
'				.ControlName = "lblFirstContactAt"
'				.Left = 480
'				.Top = 205
'				.Width = 126
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Location:"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			Me.AddChild(CType(lblFirstContactAt, UIControl))

'			'txtFirstContactPlace initial props
'			txtFirstContactPlace = New UITextBox(oUILib)
'			With txtFirstContactPlace
'				.ControlName = "txtFirstContactPlace"
'				.Left = 615
'				.Top = 205
'				.Width = 175
'				.Height = 22
'				.Enabled = True
'				.Visible = True
'				.Caption = "Fadlar V (Fadlar System)"
'				.ForeColor = muSettings.InterfaceTextBoxForeColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'				.MaxLength = 0
'				.BorderColor = muSettings.InterfaceBorderColor
'				.Locked = True
'			End With
'			Me.AddChild(CType(txtFirstContactPlace, UIControl))

'			'btnGotoFCLoc initial props
'			btnGotoFCLoc = New UIButton(oUILib)
'			With btnGotoFCLoc
'				.ControlName = "btnGotoFCLoc"
'				.Left = 655
'				.Top = 235
'				.Width = 100
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Goto"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = True
'				.FontFormat = CType(5, DrawTextFormat)
'				.ControlImageRect = New Rectangle(0, 0, 120, 32)
'			End With
'			Me.AddChild(CType(btnGotoFCLoc, UIControl))

'			'lblFirstContactWith initial props
'			lblFirstContactWith = New UILabel(oUILib)
'			With lblFirstContactWith
'				.ControlName = "lblFirstContactWith"
'				.Left = 480
'				.Top = 135
'				.Width = 126
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "First Contact With:"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			Me.AddChild(CType(lblFirstContactWith, UIControl))

'			'txtFirstContactWith initial props
'			txtFirstContactWith = New UITextBox(oUILib)
'			With txtFirstContactWith
'				.ControlName = "txtFirstContactWith"
'				.Left = 615
'				.Top = 135
'				.Width = 175
'				.Height = 22
'				.Enabled = True
'				.Visible = True
'				.Caption = "Rakura"
'				.ForeColor = muSettings.InterfaceTextBoxForeColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'				.MaxLength = 0
'				.BorderColor = muSettings.InterfaceBorderColor
'				.Locked = True
'			End With
'			Me.AddChild(CType(txtFirstContactWith, UIControl))

'			'btnUpdateNotes initial props
'			btnUpdateNotes = New UIButton(oUILib)
'			With btnUpdateNotes
'				.ControlName = "btnUpdateNotes"
'				.Left = 305
'				.Top = 500
'				.Width = 100
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Update"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = True
'				.FontFormat = CType(5, DrawTextFormat)
'				.ControlImageRect = New Rectangle(0, 0, 120, 32)
'			End With
'			Me.AddChild(CType(btnUpdateNotes, UIControl))

'			'lblRelStandings initial props
'			lblRelStandings = New UILabel(oUILib)
'			With lblRelStandings
'				.ControlName = "lblRelStandings"
'				.Left = 480
'				.Top = 280
'				.Width = 163
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Relationship Standings"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			Me.AddChild(CType(lblRelStandings, UIControl))

'			'lnDiv1 initial props
'			lnDiv1 = New UILine(oUILib)
'			With lnDiv1
'				.ControlName = "lnDiv1"
'				.Left = 470
'				.Top = 270
'				.Width = 325
'				.Height = 0
'				.Enabled = True
'				.Visible = True
'				.BorderColor = muSettings.InterfaceBorderColor
'			End With
'			Me.AddChild(CType(lnDiv1, UIControl))

'			'lblTowardsUs initial props
'			lblTowardsUs = New UILabel(oUILib)
'			With lblTowardsUs
'				.ControlName = "lblTowardsUs"
'				.Left = 495
'				.Top = 300
'				.Width = 280
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Towards <guildname>"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			Me.AddChild(CType(lblTowardsUs, UIControl))

'			'lblTowardsThem initial props
'			lblTowardsThem = New UILabel(oUILib)
'			With lblTowardsThem
'				.ControlName = "lblTowardsThem"
'				.Left = 495
'				.Top = 325
'				.Width = 280
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Towards Them:"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
'				.DrawBackImage = False
'				.FontFormat = CType(4, DrawTextFormat)
'			End With
'			Me.AddChild(CType(lblTowardsThem, UIControl))

'			'vscrRel initial props
'			hscrRel = New UIScrollBar(oUILib, False)
'			With hscrRel
'				.ControlName = "hscrRel"
'				.Left = 495
'				.Top = 365
'				.Width = 280
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Value = 60
'				.MaxValue = 255
'				.MinValue = 1
'				.SmallChange = 1
'				.LargeChange = 4
'				.ReverseDirection = False
'			End With
'			Me.AddChild(CType(hscrRel, UIControl))

'			'btnSetRel initial props
'			btnSetRel = New UIButton(oUILib)
'			With btnSetRel
'				.ControlName = "btnSetRel"
'				.Left = 525
'				.Top = 405
'				.Width = 100
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Set"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = True
'				.FontFormat = CType(5, DrawTextFormat)
'				.ControlImageRect = New Rectangle(0, 0, 120, 32)
'			End With
'			Me.AddChild(CType(btnSetRel, UIControl))

'			'btnResetRel initial props
'			btnResetRel = New UIButton(oUILib)
'			With btnResetRel
'				.ControlName = "btnResetRel"
'				.Left = 645
'				.Top = 405
'				.Width = 100
'				.Height = 24
'				.Enabled = True
'				.Visible = True
'				.Caption = "Reset"
'				.ForeColor = muSettings.InterfaceBorderColor
'				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'				.DrawBackImage = True
'				.FontFormat = CType(5, DrawTextFormat)
'				.ControlImageRect = New Rectangle(0, 0, 120, 32)
'			End With
'			Me.AddChild(CType(btnResetRel, UIControl))

'			FillValues()
'			mbLoading = False

'		End Sub

'		Public Overrides Sub NewFrame()
'			If lstRels Is Nothing = False Then
'				For X As Int32 = 0 To lstRels.ListCount - 1
'					Dim lID As Int32 = lstRels.ItemData(X)
'					Dim iTypeID As Int16 = CShort(lstRels.ItemData2(X))
'					Dim sName As String = GetCacheObjectValue(lID, iTypeID)
'					If lstRels.List(X) <> sName Then lstRels.List(X) = sName
'				Next
'			End If
'		End Sub

'		Public Overrides Sub RenderEnd()
'			If mlCurrIcon <> -1 Then
'				If rcBack = Rectangle.Empty Then SetIcon(mlCurrIcon)

'				If moSprite Is Nothing OrElse moSprite.Disposed = True Then
'					Device.IsUsingEventHandlers = False
'					moSprite = New Sprite(MyBase.moUILib.oDevice)
'					Device.IsUsingEventHandlers = True
'				End If
'				If moIconFore Is Nothing OrElse moIconFore.Disposed = True Then
'					Device.IsUsingEventHandlers = False
'					moIconFore = goResMgr.GetTexture("DipPlayerFore.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
'					Device.IsUsingEventHandlers = True
'				End If
'				If moIconBack Is Nothing OrElse moIconBack.Disposed = True Then
'					Device.IsUsingEventHandlers = False
'					moIconBack = goResMgr.GetTexture("DipPlayerBack.bmp", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
'					Device.IsUsingEventHandlers = True
'				End If

'				moSprite.Begin(SpriteFlags.AlphaBlend)
'				Try
'					Dim rcDest As Rectangle = New Rectangle(0, 0, 64, 64)
'					Dim ptDest As Point = fraIcon.GetAbsolutePosition()
'					ptDest.X += 2
'					ptDest.Y += 2

'					moSprite.Draw2D(moIconBack, rcBack, rcDest, ptDest, clrBack)
'					moSprite.Draw2D(moIconFore, rcFore1, rcDest, ptDest, clrFore1)
'					moSprite.Draw2D(moIconFore, rcFore2, rcDest, ptDest, clrFore2)
'				Catch
'					'do nothing?
'				End Try
'				moSprite.End()
'			End If
'		End Sub

'		Protected Overrides Sub Finalize()
'			If moSprite Is Nothing = False Then moSprite.Dispose()
'			moSprite = Nothing
'			moIconFore = Nothing
'			moIconBack = Nothing
'			MyBase.Finalize()
'		End Sub

'		Private Sub SetIcon(ByVal lIcon As Int32)
'			Dim yBackImg As Byte
'			Dim yBackClr As Byte
'			Dim yFore1Img As Byte
'			Dim yFore1Clr As Byte
'			Dim yFore2Img As Byte
'			Dim yFore2Clr As Byte

'			PlayerIconManager.FillIconValues(lIcon, yBackImg, yBackClr, yFore1Img, yFore1Clr, yFore2Img, yFore2Clr)

'			rcBack = PlayerIconManager.ReturnImageRectangle(yBackImg, True)
'			rcFore1 = PlayerIconManager.ReturnImageRectangle(yFore1Img, False)
'			rcFore2 = PlayerIconManager.ReturnImageRectangle(yFore2Img, False)

'			clrBack = PlayerIconManager.GetColorValue(yBackClr)
'			clrFore1 = PlayerIconManager.GetColorValue(yFore1Clr)
'			clrFore2 = PlayerIconManager.GetColorValue(yFore2Clr)

'			mlCurrIcon = lIcon
'			Me.IsDirty = True
'		End Sub

'		Private Sub FillValues()
'			lstRels.Clear()
'			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
'				If goCurrentPlayer.oGuild.moRels Is Nothing = False Then

'					Dim lSorted() As Int32 = GetSortedIndexArrayNoIdxArray(goCurrentPlayer.oGuild.moRels, GetSortedIndexArrayType.eGuildRelName, False)
'					If lSorted Is Nothing = False Then
'						With goCurrentPlayer.oGuild
'							'Ok, first, go thrugh Guilds...
'							For X As Int32 = 0 To lSorted.GetUpperBound(0)
'								If lSorted(X) > -1 Then
'									Dim oRel As GuildRel = .moRels(lSorted(X))
'									If oRel Is Nothing = False Then
'										If oRel.iEntityTypeID = ObjectType.eGuild Then
'											lstRels.AddItem(oRel.sName, False)
'											lstRels.ItemData(lstRels.NewIndex) = oRel.lEntityID
'											lstRels.ItemData2(lstRels.NewIndex) = oRel.iEntityTypeID
'										End If
'									End If
'								End If
'							Next X
'							'then, go thru players
'							For X As Int32 = 0 To lSorted.GetUpperBound(0)
'								If lSorted(X) > -1 Then
'									Dim oRel As GuildRel = .moRels(lSorted(X))
'									If oRel Is Nothing = False Then
'										If oRel.iEntityTypeID = ObjectType.ePlayer Then
'											lstRels.AddItem(oRel.sName, False)
'											lstRels.ItemData(lstRels.NewIndex) = oRel.lEntityID
'											lstRels.ItemData2(lstRels.NewIndex) = oRel.iEntityTypeID
'										End If
'									End If
'								End If
'							Next X
'						End With
'					End If
'				End If
'			End If
'		End Sub

'		Private Sub btnGotoFCLoc_Click(ByVal sName As String) Handles btnGotoFCLoc.Click
'			If lstRels Is Nothing = False AndAlso lstRels.ListIndex > -1 Then
'				Dim lID As Int32 = lstRels.ItemData(lstRels.ListIndex)
'				Dim iTypeID As Int16 = CShort(lstRels.ItemData2(lstRels.ListIndex))

'				'ok, get our rel
'				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
'					Dim oRel As GuildRel = goCurrentPlayer.oGuild.GetRel(lID, iTypeID)
'					If oRel Is Nothing = False Then
'						Dim uWP As PlayerComm.WPAttachment
'						With uWP
'							.AttachNumber = 1
'							.EnvirID = oRel.lLocationIDOfFC
'							.EnvirTypeID = oRel.iLocationTypeIDOfFC
'							.LocX = oRel.lLocXOfFC
'							.LocZ = oRel.lLocZOfFC
'							.sWPName = "First Contact"
'						End With
'						uWP.JumpToAttachment()
'					End If
'				End If
'			End If

'		End Sub

'		Private Sub btnResetRel_Click(ByVal sName As String) Handles btnResetRel.Click
'			If lstRels Is Nothing = False AndAlso lstRels.ListIndex > -1 Then
'				Dim lID As Int32 = lstRels.ItemData(lstRels.ListIndex)
'				Dim iTypeID As Int16 = CShort(lstRels.ItemData2(lstRels.ListIndex))

'				'ok, get our rel
'				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
'					Dim oRel As GuildRel = goCurrentPlayer.oGuild.GetRel(lID, iTypeID)
'					If oRel Is Nothing = False Then
'						hscrRel.Value = oRel.yRelTowardsThem
'						hscrRel_ValueChange()
'					End If
'				End If
'			End If
'		End Sub

'		Private Sub btnSetRel_Click(ByVal sName As String) Handles btnSetRel.Click
'			If lstRels Is Nothing = False AndAlso lstRels.ListIndex > -1 Then
'				'Send the Set Rel Msg
'				Dim yData(10) As Byte

'				Dim lID As Int32 = lstRels.ItemData(lstRels.ListIndex)
'				Dim iTypeID As Int16 = CShort(lstRels.ItemData2(lstRels.ListIndex))
'				Dim yRel As Byte = CByte(hscrRel.Value)

'				System.BitConverter.GetBytes(GlobalMessageCode.eSetGuildRel).CopyTo(yData, 0)
'				System.BitConverter.GetBytes(lID).CopyTo(yData, 2)
'				System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 6)
'				yData(8) = yRel

'				'Now, send to primary
'				MyBase.moUILib.SendMsgToPrimary(yData)
'			End If
'		End Sub

'		Private Sub btnUpdateNotes_Click(ByVal sName As String) Handles btnUpdateNotes.Click
'			'Update the notes related to this rel
'			If lstRels.ListIndex > -1 Then
'				Dim sText As String = txtDetails.Caption.Trim

'				Dim yMsg(11 + sText.Length) As Byte
'				Dim lPos As Int32 = 0
'				Dim lID As Int32 = lstRels.ItemData(lstRels.ListIndex)
'				Dim iTypeID As Int16 = CShort(lstRels.ItemData2(lstRels.ListIndex))

'				System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildRelNotes).CopyTo(yMsg, lPos) : lPos += 2
'				System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
'				System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
'				System.BitConverter.GetBytes(sText.Length).CopyTo(yMsg, lPos) : lPos += 4
'				System.Text.ASCIIEncoding.ASCII.GetBytes(sText).CopyTo(yMsg, lPos) : lPos += sText.Length

'				MyBase.moUILib.SendMsgToPrimary(yMsg)
'				MyBase.moUILib.AddNotification("Notes Update Sent...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'			End If
'		End Sub

'		Private Sub lstRels_ItemClick(ByVal lIndex As Integer) Handles lstRels.ItemClick
'			If lIndex > -1 Then
'				Dim lID As Int32 = lstRels.ItemData(lstRels.ListIndex)
'				Dim iTypeID As Int16 = CShort(lstRels.ItemData2(lstRels.ListIndex))

'				'ok, get our rel
'				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
'					Dim oRel As GuildRel = goCurrentPlayer.oGuild.GetRel(lID, iTypeID)
'					If oRel Is Nothing = False Then
'						With oRel
'							txtDetails.Caption = .sNotes
'							txtFirstContactBy.Caption = GetCacheObjectValue(.lWhoMadeFirstContact, ObjectType.ePlayer)
'							txtFirstContactOn.Caption = .dtWhenFirstContactMade.ToLocalTime.ToString
'							txtFirstContactPlace.Caption = GetCacheObjectValue(.lLocationIDOfFC, .iLocationTypeIDOfFC)
'							txtFirstContactWith.Caption = GetCacheObjectValue(.lWhoFirstContactWasMadeWith, ObjectType.ePlayer)

'							hscrRel.Value = .yRelTowardsThem
'							hscrRel_ValueChange()
'							lblTowardsUs.Caption = "Towards Us: " & GetRelValText(.yRelTowardsUs)
'							lblTowardsThem.Caption = "Towards Them: " & GetRelValText(.yRelTowardsThem)

'							lblTowardsUs.ForeColor = GetRelValColor(.yRelTowardsUs)
'							lblTowardsThem.ForeColor = GetRelValColor(.yRelTowardsThem)

'							If .iEntityTypeID = ObjectType.eGuild Then
'								lblMemberCount.Caption = "Known Member Count: " & .lKnownMemberCount
'								lblEntityName.Caption = .sName
'								lblGuildLeader.Caption = "Leader: " & GetCacheObjectValue(.lLeaderID, ObjectType.ePlayer)
'							Else
'								lblEntityName.Caption = Player.GetPlayerNameWithTitle(.yPlayerTitle, .sName, .bIsMale)

'								If .lPlayerGuildID <> -1 Then
'									lblMemberCount.Caption = "Member Of " & GetCacheObjectValue(.lPlayerGuildID, ObjectType.eGuild)
'								Else
'									Dim sCap As String = "Not Known To Be In A Guild"
'									If .iEntityTypeID = ObjectType.ePlayer Then
'										Dim oMember As GuildMember = goCurrentPlayer.oGuild.GetMember(.lEntityID)
'										If oMember Is Nothing = False Then
'											If (oMember.yMemberState And GuildMemberState.Approved) <> 0 Then
'												sCap = "Member Of " & goCurrentPlayer.oGuild.sName
'											End If
'										End If
'									End If
'									lblMemberCount.Caption = sCap
'								End If
'								lblGuildLeader.Caption = ""
'							End If

'							SetIcon(.lIcon)
'						End With
'					End If
'				End If
'			End If
'		End Sub

'		Private Sub hscrRel_ValueChange() Handles hscrRel.ValueChange
'			If mbLoading = True Then Return
'			lblTowardsThem.Caption = "Towards Them: " & GetRelValText(CByte(hscrRel.Value))
'			lblTowardsThem.ForeColor = GetRelValColor(CByte(hscrRel.Value))
'		End Sub

'		Private Function GetRelValText(ByVal yVal As Byte) As String
'			If yVal <= elRelTypes.eBloodWar Then
'				Return "BLOOD WAR (" & yVal & ")"
'			ElseIf yVal <= elRelTypes.eWar Then
'				Return "WAR (" & yVal & ")"
'			ElseIf yVal <= elRelTypes.eNeutral Then
'				Return "NEUTRAL (" & yVal & ")"
'			ElseIf yVal <= elRelTypes.ePeace Then
'				Return "PEACE (" & yVal & ")"
'			ElseIf yVal <= elRelTypes.eAlly Then
'				Return "ALLY (" & yVal & ")"
'			Else
'				Return "BLOOD ALLY (" & yVal & ")"
'			End If
'		End Function

'		Private Function GetRelValColor(ByVal yVal As Byte) As System.Drawing.Color
'			If yVal <= elRelTypes.eWar Then
'				Return System.Drawing.Color.FromArgb(255, 255, 0, 0)
'			ElseIf yVal <= elRelTypes.eNeutral Then
'				Return System.Drawing.Color.FromArgb(255, 255, 255, 255)
'			ElseIf yVal <= elRelTypes.ePeace Then
'				Return System.Drawing.Color.FromArgb(255, 0, 128, 0)
'			Else
'				Return System.Drawing.Color.FromArgb(255, 0, 255, 255)
'			End If
'		End Function

'	End Class
'End Class
