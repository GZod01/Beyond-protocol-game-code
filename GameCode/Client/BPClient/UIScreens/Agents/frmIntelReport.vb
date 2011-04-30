Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmIntelReport
	Inherits UIWindow

	Private lblName As UILabel
	Private lnDiv1 As UILine
	Private txtNotes As UITextBox
	Private WithEvents btnClose As UIButton
    Private WithEvents btnGoto As UIButton


	Private WithEvents chkSetFilter As UICheckBox
	Private mbIgnoreSetFilter As Boolean = False

	Private moItem As PlayerItemIntel = Nothing

	Private mlObjID As Int32 = -1
	Private miObjTypeID As Int16 = -1
	Private miExtID As Int16 = -1

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmIntelReport initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eIntelReport
            .ControlName = "frmIntelReport"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            .Width = 256
            .Height = 256
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

		'lblName initial props
		lblName = New UILabel(oUILib)
		With lblName
			.ControlName = "lblName"
			.Left = 5
			.Top = 5
			.Width = 218
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Item Name Stuff"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblName, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 230
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
            .Left = Me.BorderLineWidth \ 2
			.Top = 28
            .Width = Me.Width - Me.BorderLineWidth
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'txtNotes initial props
		txtNotes = New UITextBox(oUILib)
		With txtNotes
			.ControlName = "txtNotes"
			.Left = 5
			.Top = 35
			.Width = 246
			.Height = 190
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
		Me.AddChild(CType(txtNotes, UIControl))

		'btnGoto initial props
		btnGoto = New UIButton(oUILib)
		With btnGoto
			.ControlName = "btnGoto"
			.Left = txtNotes.Left + txtNotes.Width - 110
			.Top = 230
			.Width = 110
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Goto Location"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnGoto, UIControl))

		'chkSetFilter initial props
		chkSetFilter = New UICheckBox(oUILib)
		With chkSetFilter
			.ControlName = "chkSetFilter"
			.Left = txtNotes.Left + 5
			.Top = btnGoto.Top + 2
			.Width = 62
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Archive"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
			.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
		End With
		Me.AddChild(CType(chkSetFilter, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

	Public Sub SetFromTechKnowledge(ByRef oItem As PlayerTechKnowledge)
		moItem = Nothing

		txtNotes.Caption = ""
		btnGoto.Visible = False
		lblName.Caption = ""
		If oItem Is Nothing Then Return
		If oItem.oTech Is Nothing Then Return

		lblName.Caption = oItem.oTech.GetComponentName()
		txtNotes.Caption = oItem.oTech.GetIntelReport(oItem.yKnowledgeType)

		mlObjID = oItem.oTech.ObjectID
		miObjTypeID = ObjectType.ePlayerTechKnowledge
		miExtID = oItem.oTech.ObjTypeID

		mbIgnoreSetFilter = True
		chkSetFilter.Value = oItem.bArchived
		mbIgnoreSetFilter = False
	End Sub

	Public Sub SetFromItemIntel(ByRef oItem As PlayerItemIntel)
		moItem = oItem

		moItem.RequestDetails()
		lblName.Caption = GetCacheObjectValue(moItem.lItemID, moItem.iItemTypeID)
		txtNotes.Caption = moItem.GetIntelReport()
		btnGoto.Visible = False
		If moItem.yIntelType >= PlayerItemIntel.PlayerItemIntelType.eLocation Then
            If moItem.EnvirTypeID = ObjectType.ePlanet OrElse moItem.EnvirTypeID = ObjectType.eSolarSystem Then
                btnGoto.Visible = True
            End If
        End If

		mlObjID = oItem.lItemID
		miObjTypeID = ObjectType.ePlayerItemIntel
		miExtID = oItem.iItemTypeID

		mbIgnoreSetFilter = True
		chkSetFilter.Value = oItem.bArchived
		mbIgnoreSetFilter = False
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnGoto_Click(ByVal sName As String) Handles btnGoto.Click
		'need to add loc to playeritemintel
		If moItem Is Nothing = False Then

			If moItem.yIntelType >= PlayerItemIntel.PlayerItemIntelType.eLocation Then
				Dim uWP As PlayerComm.WPAttachment
				With uWP
					.AttachNumber = 0
					.EnvirID = moItem.EnvirID
					.EnvirTypeID = moItem.EnvirTypeID
					.LocX = moItem.LocX
					.LocZ = moItem.LocZ
					.sWPName = "Temp"
					.JumpToAttachment()
				End With
			End If 
		End If
	End Sub

	Private Sub frmIntelReport_OnNewFrame() Handles Me.OnNewFrame
		If moItem Is Nothing = False Then
			Dim sName As String = GetCacheObjectValue(moItem.lItemID, moItem.iItemTypeID)
			If lblName.Caption <> sName Then lblName.Caption = sName
			sName = moItem.GetIntelReport()
			If txtNotes.Caption <> sName Then txtNotes.Caption = sName
		End If
	End Sub

	Private Sub chkSetFilter_Click() Handles chkSetFilter.Click
		If mbIgnoreSetFilter = True Then Return

		If miObjTypeID = ObjectType.ePlayerItemIntel Then
			For X As Int32 = 0 To glItemIntelUB
				If glItemIntelIdx(X) <> -1 Then
					If goItemIntel(X).lItemID = mlObjID AndAlso goItemIntel(X).iItemTypeID = miExtID Then
						goItemIntel(X).bArchived = chkSetFilter.Value
						Exit For
					End If
				End If
			Next X
		Else
			For X As Int32 = 0 To glPlayerTechKnowledgeUB
				If glPlayerTechKnowledgeIdx(X) <> -1 Then
					If goPlayerTechKnowledge(X).oTech Is Nothing = False Then
						If goPlayerTechKnowledge(X).oTech.ObjectID = mlObjID AndAlso goPlayerTechKnowledge(X).oTech.ObjTypeID = miExtID Then
							goPlayerTechKnowledge(X).bArchived = chkSetFilter.Value
							Exit For
						End If
					End If
				End If
			Next X
		End If

		Dim yMsg(10) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eSetArchiveState).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(mlObjID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(miObjTypeID).CopyTo(yMsg, lPos) : lPos += 2
		If chkSetFilter.Value = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
		lPos += 1
		System.BitConverter.GetBytes(miExtID).CopyTo(yMsg, lPos) : lPos += 2

        MyBase.moUILib.SendMsgToPrimary(yMsg)

        'Update agent intel window
        Dim ofrm As frmAgentMain = CType(goUILib.GetWindow("frmAgentMain"), frmAgentMain)
        If ofrm Is Nothing = False Then
            ofrm.RefreshIntelList()
        End If
        ofrm = Nothing

    End Sub
End Class