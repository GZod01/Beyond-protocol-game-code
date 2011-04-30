Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmNewEmail
    Inherits UIWindow

    Public Enum eCreateFromMsgType As Integer
        eEditDraft = 0
        eReply = 1
        eReplyToAll = 2
        eForward = 3
    End Enum

    Private fraMsgHdr As UIWindow
    Private WithEvents btnTo As UIButton
    Private WithEvents btnBCC As UIButton
    Private lblSubject As UILabel
    Private txtTo As UITextBox
    Private txtBCC As UITextBox
    Private txtSubject As UITextBox
    Private txtBody As UITextBox
    Private WithEvents btnCancel As UIButton
    Private WithEvents btnSend As UIButton
    Private WithEvents btnSave As UIButton

    Private mfraAttachments As fraWPAttachments

    Private myAddrBookLookupButton As Byte = 0      '0 - none, 1 - TO, 2 - bcc

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmNewEmail initial props
        With Me
            .ControlName = "frmNewEmail"
            .lWindowMetricID = BPMetrics.eWindow.eNewEmail
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 250
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 170
            .Width = 500
            .Height = 390
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .BorderLineWidth = 1
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
        End With

        'fraMsgHdr initial props
        fraMsgHdr = New UIWindow(oUILib)
        With fraMsgHdr
            .ControlName = "fraMsgHdr"
            .Left = 10
            .Top = 10
            .Width = 480
            .Height = 90
            .Enabled = False
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With
        Me.AddChild(CType(fraMsgHdr, UIControl))

        'btnTo initial props
        btnTo = New UIButton(oUILib)
        With btnTo
            .ControlName = "btnTo"
            .Left = 15
            .Top = 15
            .Width = 45
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "TO:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnTo, UIControl))

        'btnBCC initial props
        btnBCC = New UIButton(oUILib)
        With btnBCC
            .ControlName = "btnBCC"
            .Left = 15
            .Top = 45
            .Width = 45
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "BCC:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnBCC, UIControl))

        'lblSubject initial props
        lblSubject = New UILabel(oUILib)
        With lblSubject
            .ControlName = "lblSubject"
            .Left = 15
            .Top = 75
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Subject:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSubject, UIControl))

        'txtTo initial props
        txtTo = New UITextBox(oUILib)
        With txtTo
            .ControlName = "txtTo"
            .Left = 75
            .Top = 17
            .Width = 410
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
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtTo, UIControl))

        'txtBCC initial props
        txtBCC = New UITextBox(oUILib)
        With txtBCC
            .ControlName = "txtBCC"
            .Left = 75
            .Top = 47
            .Width = 410
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
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtBCC, UIControl))

        'txtSubject initial props
        txtSubject = New UITextBox(oUILib)
        With txtSubject
            .ControlName = "txtSubject"
            .Left = 75
            .Top = 75
            .Width = 410
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
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtSubject, UIControl))

        'txtBody initial props
        txtBody = New UITextBox(oUILib)
        With txtBody
            .ControlName = "txtBody"
            .Left = 170 '10
            .Top = 110
            .Width = 320 '480
            .Height = 242
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
        Me.AddChild(CType(txtBody, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = 390
            .Top = 360
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

        'btnSend initial props
        btnSend = New UIButton(oUILib)
        With btnSend
            .ControlName = "btnSend"
            .Left = 285
            .Top = 360
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Send"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSend, UIControl))

        'btnSave initial props
        btnSave = New UIButton(oUILib)
        With btnSave
            .ControlName = "btnSave"
            .Left = 10
            .Top = 360
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Save Draft"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSave, UIControl))

        'mfraAttachments
        mfraAttachments = New fraWPAttachments(oUILib)
        With mfraAttachments
            .Left = 10
            .Top = txtBody.Top
            .Enabled = True
            .Visible = True
            .Width = 155
            .Height = 240
        End With
        Me.AddChild(CType(mfraAttachments, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
        goUILib.FocusedControl = txtTo
        txtTo.HasFocus = True

    End Sub

	Private Sub btnBCC_Click(ByVal sName As String) Handles btnBCC.Click
		myAddrBookLookupButton = 2
		Dim ofrmAddr As New frmAddressBook(goUILib)
		AddHandler ofrmAddr.ContactSelected, AddressOf AddressBookResult
		ofrmAddr = Nothing
	End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnSend_Click(ByVal sName As String) Handles btnSend.Click
		If txtTo.Caption.Trim = "" OrElse Replace(Replace(txtTo.Caption, ";", ""), ",", "").Trim = "" Then
			goUILib.AddNotification("You must enter a TO address.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		SendEmailMsg(GlobalMessageCode.eSendEmail)

		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnTo_Click(ByVal sName As String) Handles btnTo.Click
		myAddrBookLookupButton = 1
		Dim ofrmAddr As New frmAddressBook(goUILib)
		AddHandler ofrmAddr.ContactSelected, AddressOf AddressBookResult
		ofrmAddr = Nothing
	End Sub

	Private Sub AddressBookResult(ByVal sValue As String)
		Select Case myAddrBookLookupButton
			Case 1
				'add it to the TO
				Dim sText As String = txtTo.Caption.Trim
				If sText.Length > 0 Then
					If sText.EndsWith(";") = False Then
						sText &= "; "
					End If
                End If
                If txtTo.Caption.Contains(sValue) = False Then
                    sText &= sValue
                    txtTo.Caption = sText
                End If
            Case 2
                'add it to the BCC
                Dim sText As String = txtBCC.Caption.Trim
                If sText.Length > 0 Then
                    If sText.EndsWith(";") = False Then
                        sText &= "; "
                    End If
                End If
                If txtTo.Caption.Contains(sValue) = False Then
                    sText &= sValue
                    txtBCC.Caption = sText
                End If
        End Select
	End Sub

    Private Sub btnSave_Click(ByVal sName As String) Handles btnSave.Click
        BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eDraftEmail, 0, 0)

        SendEmailMsg(GlobalMessageCode.eSaveEmailDraft)
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

	Private Sub SendEmailMsg(ByVal iMsgCode As Int16)
		Dim yMsg() As Byte
		Dim lPos As Int32 = 0

		'TO, BCC, Subject, Body
		Dim sTo As String = txtTo.Caption.Trim
		Dim sBCC As String = txtBCC.Caption.Trim
		Dim sSubject As String = txtSubject.Caption
        Dim sBody As String = txtBody.Caption

        If sSubject.Length = 0 Then sSubject = "(No Subject)"
        If sBody.Length = 0 Then sBody = "(No Body)"

		ReDim yMsg(25 + sTo.Length + sBCC.Length + sSubject.Length + sBody.Length + ((mfraAttachments.mlWPUB + 1) * 35))

		System.BitConverter.GetBytes(iMsgCode).CopyTo(yMsg, lPos) : lPos += 2

		'This player is sending it... althought this will need to be verified on the server
		System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(sTo.Length).CopyTo(yMsg, lPos) : lPos += 4
		If sTo.Length <> 0 Then
			System.Text.ASCIIEncoding.ASCII.GetBytes(sTo).CopyTo(yMsg, lPos) : lPos += sTo.Length
		End If
		System.BitConverter.GetBytes(sBCC.Length).CopyTo(yMsg, lPos) : lPos += 4
		If sBCC.Length <> 0 Then
			System.Text.ASCIIEncoding.ASCII.GetBytes(sBCC).CopyTo(yMsg, lPos) : lPos += sBCC.Length
		End If
		System.BitConverter.GetBytes(sSubject.Length).CopyTo(yMsg, lPos) : lPos += 4
		If sSubject.Length <> 0 Then
			System.Text.ASCIIEncoding.ASCII.GetBytes(sSubject).CopyTo(yMsg, lPos) : lPos += sSubject.Length
		End If
		System.BitConverter.GetBytes(sBody.Length).CopyTo(yMsg, lPos) : lPos += 4
		If sSubject.Length <> 0 Then
			System.Text.ASCIIEncoding.ASCII.GetBytes(sBody).CopyTo(yMsg, lPos) : lPos += sBody.Length
		End If

		System.BitConverter.GetBytes(mfraAttachments.mlWPUB + 1).CopyTo(yMsg, lPos) : lPos += 4
		For X As Int32 = 0 To mfraAttachments.mlWPUB
			With mfraAttachments.muWP(X)
				yMsg(lPos) = .AttachNumber : lPos += 1
				System.BitConverter.GetBytes(.EnvirID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.EnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(.LocX).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.LocZ).CopyTo(yMsg, lPos) : lPos += 4

				If .sWPName = "" Then .sWPName = "WP " & .AttachNumber
				System.Text.ASCIIEncoding.ASCII.GetBytes(.sWPName).CopyTo(yMsg, lPos) : lPos += 20
			End With
		Next X

		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Public Sub SetFromCurrentMsg(ByRef oMsg As PlayerComm, ByVal lType As eCreateFromMsgType)
		Select Case lType
			Case eCreateFromMsgType.eEditDraft
				'ok, basically msg is being edited
				If oMsg.SentToList Is Nothing = False Then txtTo.Caption = oMsg.SentToList
				If oMsg.BCCList Is Nothing = False Then txtBCC.Caption = oMsg.BCCList
				If oMsg.MsgTitle Is Nothing = False Then txtSubject.Caption = oMsg.MsgTitle
				If oMsg.MsgBody Is Nothing = False Then txtBody.Caption = oMsg.MsgBody

				mfraAttachments.SetFromMessage(oMsg)
			Case eCreateFromMsgType.eForward
				txtTo.Caption = ""
				txtBCC.Caption = ""
				txtSubject.Caption = "FW: " & oMsg.MsgTitle
				Dim oSB As New System.Text.StringBuilder
				oSB.AppendLine()
				oSB.AppendLine()
				oSB.AppendLine("---- Message Forwarded ----")
				oSB.AppendLine("  From: " & oMsg.sSender)
				oSB.AppendLine("  To: " & oMsg.SentToList)
				oSB.AppendLine("  Sent On: " & GetDateTimeFromNumeric(oMsg.lSendOn))
				oSB.AppendLine("  Subject: " & oMsg.MsgTitle)
				oSB.AppendLine("---------------------------")
				oSB.AppendLine()

				txtBody.Caption = oSB.ToString & oMsg.MsgBody
				txtBody.SelStart = 0

				mfraAttachments.SetFromMessage(oMsg)
            Case eCreateFromMsgType.eReply, eCreateFromMsgType.eReplyToAll
                If oMsg.MsgTitle.StartsWith("RE: ") = True Then
                    txtSubject.Caption = oMsg.MsgTitle
                Else
                    txtSubject.Caption = "RE: " & oMsg.MsgTitle
                End If

                If lType = eCreateFromMsgType.eReply Then
                    txtTo.Caption = oMsg.sSender
                Else : txtTo.Caption = oMsg.sSender & "; " & oMsg.SentToList
                End If

                Dim oSB As New System.Text.StringBuilder
                oSB.AppendLine()
                oSB.AppendLine()
                oSB.AppendLine("---- Original Message ----")
                oSB.AppendLine("  From: " & oMsg.sSender)
                oSB.AppendLine("  To: " & oMsg.SentToList)
                oSB.AppendLine("  Sent On: " & GetDateTimeFromNumeric(oMsg.lSendOn))
                oSB.AppendLine("  Subject: " & oMsg.MsgTitle)
                oSB.AppendLine("---------------------------")
                oSB.AppendLine()
                Dim sTmp As String = oMsg.MsgBody
                If sTmp.Contains("---- Original Message ----") = True Then sTmp = sTmp.Substring(0, sTmp.IndexOf("---- Original Message ----") - 2)
                txtBody.Caption = oSB.ToString & sTmp
                txtBody.SelStart = 0


                If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
                goUILib.FocusedControl = txtBody
                txtBody.HasFocus = True
        End Select
	End Sub

	Public Sub SetToText(ByVal sText As String)
		If Me.txtTo.Caption <> "" AndAlso Me.txtTo.Caption.Trim.EndsWith(";") = False Then Me.txtTo.Caption &= "; "
		Me.txtTo.Caption &= sText
		Me.txtTo.SelStart = Me.txtTo.Caption.Length
	End Sub


	'Interface created from Interface Builder
	Private Class fraWPAttachments
		Inherits UIWindow

		Private Const ml_MAX_WAYPOINTS As Int32 = 10

		Private WithEvents lstWP As UIListBox
		Private WithEvents btnAdd As UIButton
		Private WithEvents btnDelete As UIButton
		Private txtName As UITextBox

		Public muWP() As PlayerComm.WPAttachment
		Public mlWPUB As Int32 = -1

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

			'fraWPAttachments initial props
			With Me
				.ControlName = "fraWPAttachments"
				.Left = 14
				.Top = 103
				.Width = 155
				.Height = 240
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
				.BorderLineWidth = 1
				.Caption = "Attached Waypoints"
			End With

			'lstWP initial props
			lstWP = New UIListBox(oUILib)
			With lstWP
				.ControlName = "lstWP"
				.Left = 5
				.Top = 15
				.Width = 145
				.Height = 133
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
			End With
			Me.AddChild(CType(lstWP, UIControl))

			'btnAdd initial props
			btnAdd = New UIButton(oUILib)
			With btnAdd
				.ControlName = "btnAdd"
				.Left = 5
				.Top = 182
				.Width = 145
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Add Current Location"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			Me.AddChild(CType(btnAdd, UIControl))

			'btnDelete initial props
			btnDelete = New UIButton(oUILib)
			With btnDelete
				.ControlName = "btnDelete"
				.Left = 5
				.Top = 210
				.Width = 145
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Delete Selected Item"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			Me.AddChild(CType(btnDelete, UIControl))

			'txtName initial props
			txtName = New UITextBox(oUILib)
			With txtName
				.ControlName = "txtName"
				.Left = 5
				.Top = 155
				.Width = 145
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
		End Sub

		Public Sub SetFromMessage(ByRef oMsg As PlayerComm)
			mlWPUB = oMsg.lAttachmentUB
			muWP = oMsg.uAttachments

			lstWP.Clear()
			For X As Int32 = 0 To oMsg.lAttachmentUB
				lstWP.AddItem(oMsg.uAttachments(X).sWPName)
				lstWP.ItemData(lstWP.NewIndex) = oMsg.uAttachments(X).AttachNumber
            Next X
            Me.Caption = oMsg.lAttachmentUB + 1 & " Attached Waypoints"
        End Sub

		Private Sub btnAdd_Click(ByVal sName As String) Handles btnAdd.Click
			Dim lEnvirID As Int32 = -1
			Dim iEnvirTypeID As Int16 = -1

			If glCurrentEnvirView = CurrentView.ePlanetView Then
				If goGalaxy.CurrentSystemIdx <> -1 Then
					Dim oSystem As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
					If oSystem.CurrentPlanetIdx <> -1 Then
						With oSystem.moPlanets(oSystem.CurrentPlanetIdx)
							lEnvirID = .ObjectID
							iEnvirTypeID = .ObjTypeID
						End With
					End If
				End If
			ElseIf glCurrentEnvirView = CurrentView.eSystemView Then
				If goGalaxy.CurrentSystemIdx <> -1 Then
					With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
						lEnvirID = .ObjectID
						iEnvirTypeID = .ObjTypeID
					End With
				End If
			End If

			If lEnvirID = -1 OrElse iEnvirTypeID = -1 Then
				MyBase.moUILib.AddNotification("You must be in Planet view or System view to add a waypoint attachment.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If

			If lstWP.ListCount >= ml_MAX_WAYPOINTS Then
				MyBase.moUILib.AddNotification("There is a maximum of 10 waypoint attachments per email.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If

			Dim yNumber As Byte = 0
			If lstWP.ListCount <> 0 Then
				yNumber = CByte(lstWP.ItemData(lstWP.ListCount - 1) + 1)
			End If

			mlWPUB += 1
			ReDim Preserve muWP(mlWPUB)
			With muWP(mlWPUB)
				.AttachNumber = yNumber
				.EnvirID = lEnvirID
				.EnvirTypeID = iEnvirTypeID
				.LocX = goCamera.mlCameraAtX
				.LocZ = goCamera.mlCameraAtZ
				.sWPName = txtName.Caption
				If .sWPName = "" Then .sWPName = "WP " & .AttachNumber

				lstWP.AddItem(.sWPName)
				lstWP.ItemData(lstWP.NewIndex) = .AttachNumber
            End With
            Me.Caption = mlWPUB + 1 & " Attached Waypoints"

        End Sub

		Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
			If lstWP.ListIndex <> -1 Then
				Dim lNumber As Int32 = lstWP.ItemData(lstWP.ListIndex)

				'Now, find that item...
				Dim lIdx As Int32 = -1
				For X As Int32 = 0 To mlWPUB
					If muWP(X).AttachNumber = lNumber Then
						lIdx = X
						Exit For
					End If
				Next X
				If lIdx <> -1 Then
					For X As Int32 = lIdx To mlWPUB - 1
						muWP(X) = muWP(X + 1)
					Next X
					mlWPUB -= 1
					ReDim Preserve muWP(mlWPUB)

					For X As Int32 = 0 To mlWPUB
						If muWP(X).AttachNumber > lNumber Then
							muWP(X).AttachNumber = CByte(muWP(X).AttachNumber - 1)
						End If
					Next X
				End If

				lstWP.Clear()
				For X As Int32 = 0 To mlWPUB
					lstWP.AddItem(muWP(X).sWPName)
					lstWP.ItemData(lstWP.NewIndex) = muWP(X).AttachNumber
				Next X
            End If
            Me.Caption = mlWPUB + 1 & " Attached Waypoints"
        End Sub

		Private Sub lstWP_ItemClick(ByVal lIndex As Integer) Handles lstWP.ItemClick
			If lIndex = -1 Then Return

			For X As Int32 = 0 To mlWPUB
				If muWP(X).AttachNumber = lstWP.ItemData(lstWP.ListIndex) Then
					txtName.Caption = muWP(X).sWPName
					Exit For
				End If
			Next X
		End Sub

		Private Sub lstWP_ItemDblClick(ByVal lIndex As Integer) Handles lstWP.ItemDblClick
			If lIndex = -1 Then Return
			For X As Int32 = 0 To mlWPUB
				If muWP(X).AttachNumber = lstWP.ItemData(lstWP.ListIndex) Then
					muWP(X).JumpToAttachment()
					Exit For
				End If
			Next X
		End Sub
	End Class
End Class