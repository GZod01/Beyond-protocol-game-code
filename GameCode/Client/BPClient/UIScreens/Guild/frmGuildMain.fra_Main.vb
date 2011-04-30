Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmGuildMain

	Private Class fra_Main
		Inherits guildframe

		Private moIconFrame As UIWindow
		Private moFrame As UIWindow
		Private lblFormedOn As UILabel
		Private lblJoinedOn As UILabel
		Private lblCurrentRank As UILabel
        Private lblUpcoming As UILabel

        Private lblGuildTreasury As UILabel
        Private WithEvents btnWithdraw As UIButton
        Private WithEvents btnDeposit As UIButton
        Private WithEvents txtTransactionAmt As UITextBox
 
		Private WithEvents btnUpdateMOTD As UIButton
		Private WithEvents lstUpcoming As UIListBox
		Private WithEvents txtMOTD As UITextBox

		'for rendering the icon
		Private rcBack As Rectangle = Rectangle.Empty
		Private rcFore1 As Rectangle
		Private rcFore2 As Rectangle
		Private clrBack As System.Drawing.Color
		Private clrFore1 As System.Drawing.Color
		Private clrFore2 As System.Drawing.Color
		Private mlCurrIcon As Int32 = -1
        Private moSprite As Sprite
        Private moIconFore As Texture
		Private moIconBack As Texture

		Private mbEditMOTD As Boolean = False
        Private mbFirst As Boolean = True

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

            With Me
                .lWindowMetricID = BPMetrics.eWindow.eGuildMain_Main
                .ControlName = "fra_Main"
                .Left = 15
                .Top = 5
                .Width = 128
                .Height = 128
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 2
                .Moveable = False
            End With

			moIconFrame = New UIWindow(oUILib)
			With moIconFrame
				.ControlName = "moIconFrame"
				.Left = 15
				.Top = 5
				.Width = 128
				.Height = 128
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Moveable = False
			End With
			Me.AddChild(CType(moIconFrame, UIControl))

			moFrame = New UIWindow(MyBase.moUILib)
			With moFrame
				.ControlName = "fraMOTD"
				.Left = moIconFrame.Left + moIconFrame.Width + 5
				.Top = 5
				.Width = 640 '770
				.Height = 150
				.Enabled = True
				.Visible = True
				.Caption = "Message of the Day"
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Moveable = False
			End With
			Me.AddChild(CType(moFrame, UIControl))

			'txtMOTD initial props
			txtMOTD = New UITextBox(MyBase.moUILib)
			With txtMOTD
				.ControlName = "txtMOTD"
				.Left = 5
				.Top = 15
				.Width = 630 '750
				.Height = 100
				.Enabled = True
				.Visible = True
				.Caption = ""
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(0, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
				.MaxLength = 1000
				.BorderColor = muSettings.InterfaceBorderColor
				.MultiLine = True
				.Locked = True
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Locked = Not goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeMOTD)
				End If
			End With
			moFrame.AddChild(CType(txtMOTD, UIControl))

			'btnUpdateMOTD initial props
			btnUpdateMOTD = New UIButton(MyBase.moUILib)
			With btnUpdateMOTD
				.ControlName = "btnUpdateMOTD"
				.Left = moFrame.Width - 120
				.Top = 120
				.Width = 115
				.Height = 24
				.Enabled = True
				.Visible = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Visible = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeMOTD)
				End If
				.Caption = "Update MOTD"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			moFrame.AddChild(CType(btnUpdateMOTD, UIControl))

			'lblFormedOn initial props
			lblFormedOn = New UILabel(MyBase.moUILib)
			With lblFormedOn
				.ControlName = "lblFormedOn"
				.Left = 28
				.Top = 160
				.Width = 325
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Formed On: Requesting..." '04/18/2008 at 1:09 AM GMT"
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Caption = "Formed On: " & goCurrentPlayer.oGuild.dtFormed.ToString("MM/dd/yyyy") & " at " & goCurrentPlayer.oGuild.dtFormed.ToString("HH:mm") & " GMT"
				End If
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblFormedOn, UIControl))

			'lblJoinedOn initial props
			lblJoinedOn = New UILabel(MyBase.moUILib)
			With lblJoinedOn
				.ControlName = "lblJoinedOn"
				.Left = 34
				.Top = 185
				.Width = 325
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Joined On: Requesting..."
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Caption = "Joined On: " & goCurrentPlayer.oGuild.dtJoined.ToString("MM/dd/yyyy") & " at " & goCurrentPlayer.oGuild.dtJoined.ToString("HH:mm") & " GMT"
				End If
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblJoinedOn, UIControl))

			'lblCurrentRank initial props
			lblCurrentRank = New UILabel(MyBase.moUILib)
			With lblCurrentRank
				.ControlName = "lblCurrentRank"
				.Left = 15
				.Top = 210
				.Width = 325
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Current Rank: Requesting..."
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Caption = "Current Rank: " & goCurrentPlayer.oGuild.GetRankName(goCurrentPlayer.oGuild.lCurrentRankID)
				End If
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
            Me.AddChild(CType(lblCurrentRank, UIControl))

            'lblGuildTreasury initial props
            lblGuildTreasury = New UILabel(MyBase.moUILib)
            With lblGuildTreasury
                .ControlName = "lblGuildTreasury"
                .Left = moFrame.Left + moFrame.Width - 200
                .Top = lblFormedOn.Top
                .Width = 200
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Treasury: Requesting..." '04/18/2008 at 1:09 AM GMT"
                If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
                    .Caption = "Treasury: " & goCurrentPlayer.oGuild.blTreasury.ToString("#,###")
                End If
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblGuildTreasury, UIControl))

            'txtTransactionAmt initial props
            txtTransactionAmt = New UITextBox(MyBase.moUILib)
            With txtTransactionAmt
                .ControlName = "txtTransactionAmt"
                .Left = lblGuildTreasury.Left + 50
                .Top = lblGuildTreasury.Top + lblGuildTreasury.Height + 5
                .Width = 100
                .Height = 18
                .Enabled = Not gbAliased
                .Visible = True
                .Caption = "0"
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(0, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 12
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
                .Locked = False
                .MultiLine = False
                'If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
                '.Locked = Not goCurrentPlayer.oGuild.HasPermission(RankPermissions)
                'End If
            End With
            Me.AddChild(CType(txtTransactionAmt, UIControl))

            'btnDeposit initial props
            btnDeposit = New UIButton(MyBase.moUILib)
            With btnDeposit
                .ControlName = "btnDeposit"
                .Left = lblGuildTreasury.Left
                .Top = txtTransactionAmt.Top + txtTransactionAmt.Height + 5
                .Width = 95
                .Height = 24
                .Enabled = Not gbAliased
                .Visible = True
                .Caption = "Deposit"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnDeposit, UIControl))

            'btnWithdraw initial props
            btnWithdraw = New UIButton(MyBase.moUILib)
            With btnWithdraw
                .ControlName = "btnWithdraw"
                .Left = btnDeposit.Left + btnDeposit.Width + 10
                .Top = btnDeposit.Top
                .Width = 95
                .Height = 24
                .Enabled = False
                .Visible = True
                If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso Not gbAliased Then
                    .Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.WithdrawLowSec)
                End If
                .Caption = "Withdraw"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnWithdraw, UIControl))

			'lblUpcoming initial props
			lblUpcoming = New UILabel(MyBase.moUILib)
			With lblUpcoming
				.ControlName = "lblUpcoming"
				.Left = 15
				.Top = 250
				.Width = 325
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Upcoming Events"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblUpcoming, UIControl))

			'lstUpcoming initial props
			lstUpcoming = New UIListBox(MyBase.moUILib)
			With lstUpcoming
				.ControlName = "lstUpcoming"
				.Left = 15
				.Top = 275
				.Width = 770
				.Height = 240
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
			End With
			Me.AddChild(CType(lstUpcoming, UIControl))

			lstUpcoming.oIconTexture = goResMgr.GetTexture("GuildIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "gi.pak")
			lstUpcoming.RenderIcons = True
            mbEditMOTD = False

            'Me.IsDirty = False
            'Me.IsDirty = True
		End Sub

		Private Sub btnUpdateMOTD_Click(ByVal sName As String) Handles btnUpdateMOTD.Click
			Dim sMOTD As String = txtMOTD.Caption

			If sMOTD.Length > 200 Then sMOTD = sMOTD.Substring(0, 200)
			Dim yMsg(201) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eUpdateMOTD).CopyTo(yMsg, lPos) : lPos += 2
			System.Text.ASCIIEncoding.ASCII.GetBytes(sMOTD).CopyTo(yMsg, lPos) : lPos += 200

			mbEditMOTD = False

			MyBase.moUILib.SendMsgToPrimary(yMsg)
		End Sub

		Private Sub lstUpcoming_ItemDblClick(ByVal lIndex As Integer) Handles lstUpcoming.ItemDblClick
			If lIndex < 0 Then Return

			'eventid
			Dim lEventID As Int32 = lstUpcoming.ItemData(lIndex)
			If lEventID > 0 Then
				'ok, now, find our event
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.moEvents Is Nothing = False Then
					Dim oEvent As GuildEvent = Nothing
					With goCurrentPlayer.oGuild

						If .HasPermission(RankPermissions.ViewEvents) = False Then Return

						For X As Int32 = 0 To .moEvents.GetUpperBound(0)
							If .moEvents(X) Is Nothing = False Then
								If .moEvents(X).EventID = lEventID Then
									oEvent = .moEvents(X)
									Exit For
								End If
							End If
						Next X
						If oEvent Is Nothing Then Return
					End With

					'ok, found the event...
					CType(Me.ParentControl, frmGuildMain).mlCurrentSelectedTab = eTabProgCtrl.EventTab

					'find the calendar
					Dim oCalendar As ctlCalendar = CType(CType(Me.ParentControl, frmGuildMain).GetFullReferencedControl("fra_Events.moCalendar"), ctlCalendar)
					If oCalendar Is Nothing Then Return

					'ok, now, set our calendar to the appropriate
					Dim dtTemp As Date = oEvent.dtStartsAt.ToLocalTime
					oCalendar.SetSelectDate(dtTemp.Day, dtTemp.Month, dtTemp.Year)

				End If
			End If
		End Sub

        Public Overrides Sub NewFrame()
            If mbFirst = True Then
                If Me.IsDirty = True Then
                    Me.IsDirty = False
                Else
                    Me.IsDirty = True
                    mbFirst = False
                End If
            End If
            'ok, fill in the upcoming events...
            If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
                With goCurrentPlayer.oGuild

                    Dim bHasViewEventPermission As Boolean = .HasPermission(RankPermissions.ViewEvents)
                    If .moEvents Is Nothing = False AndAlso bHasViewEventPermission = True Then
                        'fill the upcoming events listbox
                        Dim lSorted(19) As Int32
                        Dim dtSortDt(19) As Date
                        Dim lSortedUB As Int32 = 19
                        Dim dtNow As Date = Now
                        Dim dt1Mo As Date = Now.AddMonths(1)
                        Dim lCnt As Int32 = 0

                        For X As Int32 = 0 To lSortedUB
                            lSorted(X) = -1
                            dtSortDt(X) = Date.MinValue
                        Next X

                        If bHasViewEventPermission = True Then
                            For X As Int32 = 0 To .moEvents.GetUpperBound(0)
                                If .moEvents(X) Is Nothing = False Then
                                    Dim dt As Date = .moEvents(X).dtStartsAt.ToLocalTime.Add(New TimeSpan(0, .moEvents(X).lDuration, 0)) '  .moEvents(X).dtEndsAt.ToLocalTime
                                    Dim dtStart As Date = .moEvents(X).dtStartsAt.ToLocalTime
                                    If dt > dtNow AndAlso dt < dt1Mo Then
                                        Dim lIdx As Int32 = -1
                                        For Y As Int32 = 0 To lSortedUB
                                            If lSorted(Y) = -1 Then
                                                lIdx = Y
                                                Exit For
                                            ElseIf dtSortDt(Y) > dtStart Then
                                                lIdx = Y
                                                Exit For
                                            End If
                                        Next Y
                                        If lIdx <> -1 Then
                                            lCnt += 1
                                            For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                                                lSorted(Y) = lSorted(Y - 1)
                                                dtSortDt(Y) = dtSortDt(Y - 1)
                                            Next Y
                                            lSorted(lIdx) = X
                                            dtSortDt(lIdx) = dtStart
                                        End If
                                    End If
                                End If
                            Next X
                        End If


                        If lstUpcoming.ListCount <> lCnt Then
                            lstUpcoming.Clear()
                            For X As Int32 = 0 To lSortedUB
                                If lSorted(X) <> -1 Then
                                    Dim oEvent As GuildEvent = .moEvents(lSorted(X))
                                    If oEvent Is Nothing = False Then
                                        lstUpcoming.AddItem(oEvent.ListboxString, False)
                                        lstUpcoming.ItemData(lstUpcoming.NewIndex) = oEvent.EventID
                                        lstUpcoming.ApplyIconOffset(lstUpcoming.NewIndex) = True
                                    End If
                                Else : Exit For
                                End If
                            Next X
                        Else
                            For X As Int32 = 0 To lSortedUB
                                If lSorted(X) <> -1 Then
                                    Dim oEvent As GuildEvent = .moEvents(lSorted(X))
                                    If oEvent Is Nothing = False Then
                                        If lstUpcoming.ItemData(X) <> oEvent.EventID Then
                                            lstUpcoming.Clear()
                                            Exit For
                                        Else
                                            If lstUpcoming.List(X) <> oEvent.ListboxString Then lstUpcoming.List(X) = oEvent.ListboxString
                                            Dim rcTemp As Rectangle = oEvent.rcIconRect
                                            If lstUpcoming.rcIconRectangle(X) <> rcTemp Then lstUpcoming.rcIconRectangle(X) = rcTemp

                                            If lstUpcoming.ItemCustomColor(X) <> oEvent.ListBoxColor Then
                                                lstUpcoming.ItemCustomColor(X) = oEvent.ListBoxColor
                                                Me.IsDirty = True
                                            End If
                                        End If
                                    End If
                                End If
                            Next X
                        End If

                    End If

                    If mbEditMOTD = False AndAlso txtMOTD.Caption <> .sMOTD Then
                        txtMOTD.Caption = .sMOTD
                        mbEditMOTD = False
                    End If

                    Dim sText As String = "Formed On: Requesting..."
                    If .dtFormed <> Date.MinValue Then sText = "Formed On: " & .dtFormed.ToLocalTime.ToShortDateString & " " & .dtFormed.ToLocalTime.ToShortTimeString
                    If lblFormedOn.Caption <> sText Then lblFormedOn.Caption = sText

                    sText = "Joined On: Requesting..."
                    If .dtJoined <> Date.MinValue Then sText = "Joined On: " & .dtJoined.ToLocalTime.ToShortDateString & " " & .dtJoined.ToLocalTime.ToShortTimeString
                    If lblJoinedOn.Caption <> sText Then lblJoinedOn.Caption = sText

                    sText = "Current Rank: Requesting..."
                    Dim oRank As GuildRank = .GetRankByID(.lCurrentRankID)
                    If oRank Is Nothing = False Then sText = "Current Rank: " & oRank.sRankName
                    If lblCurrentRank.Caption <> sText Then lblCurrentRank.Caption = sText
                End With
            End If

        End Sub

		Public Overrides Sub RenderEnd()
			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
				If rcBack = Rectangle.Empty OrElse mlCurrIcon <> goCurrentPlayer.oGuild.lIcon Then SetIcon(goCurrentPlayer.oGuild.lIcon)

                If moSprite Is Nothing OrElse moSprite.Disposed = True Then
                    Device.IsUsingEventHandlers = False
                    moSprite = New Sprite(MyBase.moUILib.oDevice)
                    Device.IsUsingEventHandlers = True
                End If
				If moIconFore Is Nothing OrElse moIconFore.Disposed = True Then
					Device.IsUsingEventHandlers = False
					moIconFore = goResMgr.GetTexture("DipPlayerFore.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
					Device.IsUsingEventHandlers = True
				End If
				If moIconBack Is Nothing OrElse moIconBack.Disposed = True Then
					Device.IsUsingEventHandlers = False
					moIconBack = goResMgr.GetTexture("DipPlayerBack.bmp", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
					Device.IsUsingEventHandlers = True
				End If

				moSprite.Begin(SpriteFlags.AlphaBlend)
				Try
					Dim rcDest As Rectangle = New Rectangle(0, 0, 128, 128)
					Dim ptDest As Point = moIconFrame.GetAbsolutePosition()
					ptDest.X += 2
					ptDest.Y += 2

					ptDest.X \= 2
					ptDest.Y \= 2

					moSprite.Draw2D(moIconBack, rcBack, rcDest, ptDest, clrBack)
					moSprite.Draw2D(moIconFore, rcFore1, rcDest, ptDest, clrFore1)
					moSprite.Draw2D(moIconFore, rcFore2, rcDest, ptDest, clrFore2)
				Catch
					'do nothing?
				End Try
				moSprite.End()
			End If
		End Sub

		Private Sub SetIcon(ByVal lIcon As Int32)
			Dim yBackImg As Byte
			Dim yBackClr As Byte
			Dim yFore1Img As Byte
			Dim yFore1Clr As Byte
			Dim yFore2Img As Byte
			Dim yFore2Clr As Byte

			PlayerIconManager.FillIconValues(lIcon, yBackImg, yBackClr, yFore1Img, yFore1Clr, yFore2Img, yFore2Clr)

			rcBack = PlayerIconManager.ReturnImageRectangle(yBackImg, True)
			rcFore1 = PlayerIconManager.ReturnImageRectangle(yFore1Img, False)
			rcFore2 = PlayerIconManager.ReturnImageRectangle(yFore2Img, False)

			clrBack = PlayerIconManager.GetColorValue(yBackClr)
			clrFore1 = PlayerIconManager.GetColorValue(yFore1Clr)
			clrFore2 = PlayerIconManager.GetColorValue(yFore2Clr)

			mlCurrIcon = lIcon
			Me.IsDirty = True
		End Sub

		Protected Overrides Sub Finalize()
			If moSprite Is Nothing = False Then moSprite.Dispose()
			moSprite = Nothing
			moIconFore = Nothing
			moIconBack = Nothing
			MyBase.Finalize()
		End Sub

		Private Sub txtMOTD_TextChanged() Handles txtMOTD.TextChanged
			mbEditMOTD = True
		End Sub

        Private Sub btnWithdraw_Click(ByVal sName As String) Handles btnWithdraw.Click
            If IsNumeric(txtTransactionAmt.Caption) = False OrElse Val(txtTransactionAmt.Caption) < 0 Then
                goUILib.AddNotification("Enter a numeric value greater than 0 for the amount.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            txtTransactionAmt.Caption = Math.Min(100000000000, CLng(txtTransactionAmt.Caption)).ToString
            Dim blQty As Int64 = CLng(txtTransactionAmt.Caption)
            GuildBankTrans(1, 1, ObjectType.eCredits, blQty)
        End Sub

        Private Sub btnDeposit_Click(ByVal sName As String) Handles btnDeposit.Click
            If IsNumeric(txtTransactionAmt.Caption) = False OrElse Val(txtTransactionAmt.Caption) < 0 Then
                goUILib.AddNotification("Enter a numeric value greater than 0 for the amount.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            txtTransactionAmt.Caption = Math.Min(100000000000, CLng(txtTransactionAmt.Caption)).ToString
            Dim blQty As Int64 = CLng(txtTransactionAmt.Caption)
            GuildBankTrans(0, 1, ObjectType.eCredits, blQty)
        End Sub

        Private Sub GuildBankTrans(ByVal yType As Byte, ByVal lID As Int32, ByVal iTypeID As Int16, ByVal blQty As Int64)
            Dim yMsg(20) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eGuildBankTransaction).CopyTo(yMsg, lPos) : lPos += 2

            yMsg(lPos) = yType : lPos += 1          '1 indicates deposit
            System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(blQty).CopyTo(yMsg, lPos) : lPos += 8
            Dim lTPID As Int32 = -1 'cboTradepost.ItemData(cboTradepost.ListIndex)
            System.BitConverter.GetBytes(lTPID).CopyTo(yMsg, lPos) : lPos += 4

            If yType = 0 Then
                MyBase.moUILib.AddNotification("Request sent to deposit credits into guild bank account.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Else
                MyBase.moUILib.AddNotification("Request sent to withdraw credits into guild bank account.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If

            MyBase.moUILib.SendMsgToPrimary(yMsg)
        End Sub

        Private Sub txtTransactionAmt_OnGotFocus() Handles txtTransactionAmt.OnGotFocus
        End Sub

        Private Sub txtTransactionAmt_TextChanged() Handles txtTransactionAmt.TextChanged
            Me.IsDirty = True
        End Sub

        Private Sub fra_Main_OnNewFrame() Handles Me.OnNewFrame
            If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
                Dim sTemp As String = "Treasury: " & goCurrentPlayer.oGuild.blTreasury.ToString("#,##0")
                If lblGuildTreasury.Caption <> sTemp Then lblGuildTreasury.Caption = sTemp
            End If
        End Sub
    End Class

End Class
