Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmEmailMain

    Private Class fraEmailDetail
        Inherits UIWindow

        Private mfraHeader As fraHeader

        Private WithEvents txtBody As UITextBox
        Private WithEvents btnMoveToFolder As UIButton
        Public WithEvents cboMoveToFolder As UIComboBox
        Private lblMoveToFolder As UILabel
        Private WithEvents btnReply As UIButton
        Private WithEvents btnReplyAll As UIButton
        Private WithEvents btnForward As UIButton
        Private WithEvents btnDelete As UIButton

        Private lblWaypoints As UILabel
        Private WithEvents lstWaypoints As UIListBox

        Public Event DeleteMsg()
        Public Event ForwardMsg()
        Public Event MoveToFolder(ByVal lPCF_ID As Int32)
        Public Event Reply()
        Public Event ReplyToAll()

		Private moMsg As PlayerComm = Nothing
		Private mlLastMsgUpdate As Int32 = 0

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraEmailDetail initial props
            With Me
                .ControlName = "fraEmailDetail"
                .Left = 110
                .Top = 124
                .Width = 680
                .Height = 330
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 1
                .Moveable = False
            End With

            'txtBody initial props
            txtBody = New UITextBox(oUILib)
            With txtBody
                .ControlName = "txtBody"
                .Left = 155 '5
                .Top = 125
                .Width = 520
                .Height = 200
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(0, DrawTextFormat)
                .BackColorEnabled = Color.FromArgb(0, 0, 0, 0)
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 0
                .BorderColor = muSettings.InterfaceBorderColor
                .MultiLine = True
                .Locked = True
            End With
            Me.AddChild(CType(txtBody, UIControl))

            'lblWaypoints initial props
            lblWaypoints = New UILabel(oUILib)
            With lblWaypoints
                .ControlName = "lblWaypoints"
                .Left = 5
                .Top = 125
                .Width = 100
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Waypoints"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
            End With
            Me.AddChild(CType(lblWaypoints, UIControl))

            'lstWaypoints initial props
            lstWaypoints = New UIListBox(oUILib)
            With lstWaypoints
                .ControlName = "lstWaypoints"
                .Left = 5
                .Top = 145
                .Width = 145
                .Height = 180
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            End With
            Me.AddChild(CType(lstWaypoints, UIControl))

            'btnMoveToFolder initial props
            btnMoveToFolder = New UIButton(oUILib)
            With btnMoveToFolder
                .ControlName = "btnMoveToFolder"
                .Left = 160
                .Top = 95
                .Width = 70
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Move"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnMoveToFolder, UIControl))

            'cboMoveToFolder initial props
            cboMoveToFolder = New UIComboBox(oUILib)
            With cboMoveToFolder
                .ControlName = "cboMoveToFolder"
                .Left = 5
                .Top = 95
                .Width = 150
                .Height = 20
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(cboMoveToFolder, UIControl))

            'lblMoveToFolder initial props
            lblMoveToFolder = New UILabel(oUILib)
            With lblMoveToFolder
                .ControlName = "lblMoveToFolder"
                .Left = 5
                .Top = 75
                .Width = 100
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Move To Folder:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMoveToFolder, UIControl))

            'btnReply initial props
            btnReply = New UIButton(oUILib)
            With btnReply
                .ControlName = "btnReply"
                .Left = 260
                .Top = 95
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Reply"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnReply, UIControl))

            'btnReplyAll initial props
            btnReplyAll = New UIButton(oUILib)
            With btnReplyAll
                .ControlName = "btnReplyAll"
                .Left = 365
                .Top = 95
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Reply To All"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnReplyAll, UIControl))

            'btnForward initial props
            btnForward = New UIButton(oUILib)
            With btnForward
                .ControlName = "btnForward"
                .Left = 470
                .Top = 95
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Forward"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnForward, UIControl))

            'btnDelete initial props
            btnDelete = New UIButton(oUILib)
            With btnDelete
                .ControlName = "btnDelete"
                .Left = 575
                .Top = 95
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

            'mfraHeader Initial Props
            mfraHeader = New fraHeader(oUILib)
            With mfraHeader
                .Left = 0
                .Top = 0
                .Visible = True
                .Enabled = True
            End With
            Me.AddChild(CType(mfraHeader, UIControl))
        End Sub

        'Interface created from Interface Builder
        Private Class fraHeader
            Inherits UIWindow

            Private lblFrom As UILabel
            Private lblTo As UILabel
            Private lblSentOn As UILabel
            Private lblSubject As UILabel

            Public Sub New(ByRef oUILib As UILib)
                MyBase.New(oUILib)

                'fraHeader initial props
                With Me
                    .ControlName = "fraHeader"
                    .Left = 119
                    .Top = 213
                    .Width = 680
                    .Height = 70
                    .Enabled = True
                    .Visible = True
                    .BorderColor = muSettings.InterfaceBorderColor
                    .FillColor = muSettings.InterfaceFillColor
                    .FullScreen = False
                    .BorderLineWidth = 1
                    .Moveable = False
                End With

                'lblFrom initial props
                lblFrom = New UILabel(oUILib)
                With lblFrom
                    .ControlName = "lblFrom"
                    .Left = 5
                    .Top = 5
                    .Width = 330
                    .Height = 18
                    .Enabled = True
                    .Visible = True
                    .Caption = "From:"
                    .ForeColor = muSettings.InterfaceBorderColor
                    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                    .DrawBackImage = False
                    .FontFormat = CType(4, DrawTextFormat)
                End With
                Me.AddChild(CType(lblFrom, UIControl))

                'lblTo initial props
                lblTo = New UILabel(oUILib)
                With lblTo
                    .ControlName = "lblTo"
                    .Left = 345
                    .Top = 5
                    .Width = 330
                    .Height = 18
                    .Enabled = True
                    .Visible = True
                    .Caption = "To:"
                    .ForeColor = muSettings.InterfaceBorderColor
                    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                    .DrawBackImage = False
                    .FontFormat = CType(4, DrawTextFormat)
                End With
                Me.AddChild(CType(lblTo, UIControl))

                'lblSentOn initial props
                lblSentOn = New UILabel(oUILib)
                With lblSentOn
                    .ControlName = "lblSentOn"
                    .Left = 5
                    .Top = 25
                    .Width = 670
                    .Height = 18
                    .Enabled = True
                    .Visible = True
                    .Caption = "Sent On:"
                    .ForeColor = muSettings.InterfaceBorderColor
                    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                    .DrawBackImage = False
                    .FontFormat = CType(4, DrawTextFormat)
                End With
                Me.AddChild(CType(lblSentOn, UIControl))

                'lblSubject initial props
                lblSubject = New UILabel(oUILib)
                With lblSubject
                    .ControlName = "lblSubject"
                    .Left = 5
                    .Top = 45
                    .Width = 670
                    .Height = 18
                    .Enabled = True
                    .Visible = True
                    .Caption = "Subject:"
                    .ForeColor = muSettings.InterfaceBorderColor
                    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                    .DrawBackImage = False
                    .FontFormat = CType(4, DrawTextFormat)
                End With
                Me.AddChild(CType(lblSubject, UIControl))
            End Sub

            Public Sub SetValues(ByVal sFrom As String, ByVal sTo As String, ByVal sSentOn As String, ByVal sSubject As String)
                lblFrom.Caption = sFrom
                lblTo.Caption = sTo
                'GTS 23/41/5000 at 05:51 (10/05/2006 at 5:51 PM GMT Earth Time)
                lblSentOn.Caption = sSentOn
                lblSubject.Caption = sSubject
            End Sub
        End Class

		Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
			RaiseEvent DeleteMsg()
		End Sub

		Private Sub btnForward_Click(ByVal sName As String) Handles btnForward.Click
			RaiseEvent ForwardMsg()
		End Sub

		Private Sub btnMoveToFolder_Click(ByVal sName As String) Handles btnMoveToFolder.Click
			If cboMoveToFolder.ListIndex = -1 Then
				goUILib.AddNotification("Please select a folder to move the message to.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
			Dim lPCF_ID As Int32 = cboMoveToFolder.ItemData(cboMoveToFolder.ListIndex)
			RaiseEvent MoveToFolder(lPCF_ID)
		End Sub

		Private Sub btnReply_Click(ByVal sName As String) Handles btnReply.Click
			RaiseEvent Reply()
		End Sub

		Private Sub btnReplyAll_Click(ByVal sName As String) Handles btnReplyAll.Click
			RaiseEvent ReplyToAll()
		End Sub

        Public Sub ClearValues()
            mfraHeader.SetValues("From:", "To:", "Sent:", "Subject:")
            txtBody.Caption = ""
        End Sub

        Public Sub SetFromMsg(ByRef oMsg As PlayerComm)
            moMsg = Nothing

            With oMsg
                mfraHeader.SetValues("From: " & .sSender, "To: " & .SentToList, "Sent: " & GetDateTimeFromNumeric(.lSendOn), "Subject: " & .MsgTitle)
				mlLastMsgUpdate = .lLastMsgUpdate
                txtBody.Caption = .MsgBody

                lstWaypoints.Clear()
                If oMsg.lAttachmentUB > 1 andalso oMsg.uAttachments(0).sWPName.Contains(" ") AndAlso oMsg.uAttachments(0).sWPName.Contains("/") Then
                    Dim lSorted() As Int32 = Nothing
                    Dim lSortedUB As Int32 = -1
                    For X As Int32 = 0 To oMsg.lAttachmentUB
                        Dim lIdx As Int32 = -1

                        Dim sName As String = GetRomanNumeralSortStr(oMsg.uAttachments(X).sWPName.Split(" "c)(0))

                        For Y As Int32 = 0 To lSortedUB
                            Dim sOtherName As String = GetRomanNumeralSortStr(oMsg.uAttachments(lSorted(Y)).sWPName.Split(" "c)(0))
                            If sOtherName > sName Then
                                lIdx = Y
                                Exit For
                            End If
                        Next Y
                        lSortedUB += 1
                        ReDim Preserve lSorted(lSortedUB)
                        If lIdx = -1 Then
                            lSorted(lSortedUB) = X
                        Else
                            For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                                lSorted(Y) = lSorted(Y - 1)
                            Next Y
                            lSorted(lIdx) = X
                        End If
                    Next X

                    For X As Int32 = 0 To oMsg.lAttachmentUB
                        lstWaypoints.AddItem(oMsg.uAttachments(lSorted(X)).sWPName)
                        lstWaypoints.ItemData(lstWaypoints.NewIndex) = oMsg.uAttachments(lSorted(X)).AttachNumber
                    Next X
                Else
                    For X As Int32 = 0 To oMsg.lAttachmentUB
                        lstWaypoints.AddItem(oMsg.uAttachments(X).sWPName)
                        lstWaypoints.ItemData(lstWaypoints.NewIndex) = oMsg.uAttachments(X).AttachNumber
                    Next X
                End If
                moMsg = oMsg
            End With
            lblWaypoints.Caption = "Waypoints (" & oMsg.lAttachmentUB + 1 & ")"
        End Sub

        Private Sub lstWaypoints_ItemDblClick(ByVal lIndex As Integer) Handles lstWaypoints.ItemDblClick
            If lIndex = -1 Then Return
            If moMsg Is Nothing Then Return

            Dim yNumber As Byte = CByte(lstWaypoints.ItemData(lIndex))
            For X As Int32 = 0 To moMsg.lAttachmentUB
                If moMsg.uAttachments(X).AttachNumber = yNumber Then
                    moMsg.uAttachments(X).JumpToAttachment()
                    Exit For
                End If
            Next X
        End Sub 
	End Class

End Class