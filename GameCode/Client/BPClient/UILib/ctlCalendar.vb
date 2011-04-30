Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class ctlCalendar
	Inherits UIWindow

	Private txtSunday As UITextBox
	Private txtMonday As UITextBox
	Private WithEvents lblMonthYear As UILabel
	Private txtTuesday As UITextBox
	Private txtWednesday As UITextBox
	Private txtThursday As UITextBox
	Private txtFriday As UITextBox
	Private txtSaturday As UITextBox
	Private WithEvents vscrMonth As UIScrollBar

	Private mlFirstDayOfMonth As Int32
	Private mlDaysInMonth As Int32

	Private mlSelectedMonth As Int32
	Private mlSelectedDay As Int32
	Private mlSelectedYear As Int32

	Private mlDisplayedMonth As Int32
	Private mlDisplayedYear As Int32

    'Private Shared moLine As Line
    'Private Shared moFont As Font
    Private Shared moSprite As Sprite
    Public Shared Sub ReleaseDefaultPool()
        'If moLine Is Nothing = False Then moLine.Dispose()
        'If moFont Is Nothing = False Then moFont.Dispose()
        If moSprite Is Nothing = False Then moSprite.Dispose()
        'moLine = Nothing
        'moFont = Nothing
        moSprite = Nothing
    End Sub
    Private moTex As Texture

	Public Event DayClick(ByVal lDay As Int32, ByVal lMonth As Int32, ByVal lYear As Int32)

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'ctlCalendar initial props
		With Me
			.ControlName = "ctlCalendar"
			.Left = 395
			.Top = 187
			.Width = 256
			.Height = 256
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.BorderLineWidth = 2
		End With

		'txtSunday initial props
		txtSunday = New UITextBox(oUILib)
		With txtSunday
			.ControlName = "txtSunday"
			.Left = 2
			.Top = 21
			.Width = 32
			.Height = 18
			.Enabled = False
			.Visible = True
			.Caption = "S"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 128, 128, 128)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bCanGetFocus = False
		End With
		Me.AddChild(CType(txtSunday, UIControl))

		'txtMonday initial props
		txtMonday = New UITextBox(oUILib)
		With txtMonday
			.ControlName = "txtMonday"
			.Left = 34
			.Top = 21
			.Width = 32
			.Height = 18
			.Enabled = False
			.Visible = True
			.Caption = "M"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 128, 128, 128)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bCanGetFocus = False
		End With
		Me.AddChild(CType(txtMonday, UIControl))

		'lblMonthYear initial props
		lblMonthYear = New UILabel(oUILib)
		With lblMonthYear
			.ControlName = "lblMonthYear"
			.Left = 1
			.Top = 1
			.Width = 226
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "April 2008"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMonthYear, UIControl))

		'txtTuesday initial props
		txtTuesday = New UITextBox(oUILib)
		With txtTuesday
			.ControlName = "txtTuesday"
			.Left = 66
			.Top = 21
			.Width = 32
			.Height = 18
			.Enabled = False
			.Visible = True
			.Caption = "T"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 128, 128, 128)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bCanGetFocus = False
		End With
		Me.AddChild(CType(txtTuesday, UIControl))

		'txtWednesday initial props
		txtWednesday = New UITextBox(oUILib)
		With txtWednesday
			.ControlName = "txtWednesday"
			.Left = 98
			.Top = 21
			.Width = 32
			.Height = 18
			.Enabled = False
			.Visible = True
			.Caption = "W"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 128, 128, 128)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bCanGetFocus = False
		End With
		Me.AddChild(CType(txtWednesday, UIControl))

		'txtThursday initial props
		txtThursday = New UITextBox(oUILib)
		With txtThursday
			.ControlName = "txtThursday"
			.Left = 130
			.Top = 21
			.Width = 32
			.Height = 18
			.Enabled = False
			.Visible = True
			.Caption = "T"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 128, 128, 128)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bCanGetFocus = False
		End With
		Me.AddChild(CType(txtThursday, UIControl))

		'txtFriday initial props
		txtFriday = New UITextBox(oUILib)
		With txtFriday
			.ControlName = "txtFriday"
			.Left = 162
			.Top = 21
			.Width = 32
			.Height = 18
			.Enabled = False
			.Visible = True
			.Caption = "F"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 128, 128, 128)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bCanGetFocus = False
		End With
		Me.AddChild(CType(txtFriday, UIControl))

		'txtSaturday initial props
		txtSaturday = New UITextBox(oUILib)
		With txtSaturday
			.ControlName = "txtSaturday"
			.Left = 194
			.Top = 21
			.Width = 32
			.Height = 18
			.Enabled = False
			.Visible = True
			.Caption = "S"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 128, 128, 128)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bCanGetFocus = False
		End With
		Me.AddChild(CType(txtSaturday, UIControl))

		'vscrMonth initial props
		vscrMonth = New UIScrollBar(oUILib, True)
		With vscrMonth
			.ControlName = "vscrMonth"
			.Left = 229
			.Top = 21
			.Width = 24
			.Height = 231
			.Enabled = True
			.Visible = True
			.MaxValue = 10
			.MinValue = 0
			.Value = 5
			.SmallChange = 1
			.LargeChange = 1
			.ReverseDirection = False
		End With
		Me.AddChild(CType(vscrMonth, UIControl))

		Me.SetMonthAndYear(Now.Month, Now.Year)
    End Sub

	Private Sub SetMonthAndYear(ByVal lMonth As Int32, ByVal lYear As Int32)
		mlDisplayedMonth = lMonth
		mlDisplayedYear = lYear

		Dim sVal As String = GetMonthName(lMonth)
		If sVal <> "" Then sVal &= ", "

		If lYear < 2000 Then lYear += 2000
		sVal &= lYear.ToString("####")

		lblMonthYear.Caption = sVal

		'Now, configure our days...
		mlDaysInMonth = Date.DaysInMonth(lYear, lMonth)
		Dim dtval As Date = GetLocaleSpecificDT(lMonth.ToString & "/01/" & lYear)
		mlFirstDayOfMonth = dtval.DayOfWeek

		Me.IsDirty = True
	End Sub

	Public Sub SetSelectDate(ByVal lDay As Int32, ByVal lMonth As Int32, ByVal lYear As Int32)
		mlSelectedDay = lDay
		mlSelectedMonth = lMonth
		mlSelectedYear = lYear
		Me.IsDirty = True
	End Sub

	Private Sub ctlCalendar_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
		'ok, determine where we clicked...
		Dim ptLoc As Point = Me.GetAbsolutePosition()
		lMouseX -= ptLoc.X
		lMouseY -= ptLoc.Y

		Dim lTop As Int32 = txtSunday.Top + txtSunday.Height

		'ok, reduce mouse y by the top
		lMouseY -= lTop
		If lMouseY > 0 Then
			'ok, now, divide by 32
			lMouseY \= 32
			'mousey tells us our row...

			'mouse x...
			lMouseX -= 2
			If lMouseX < 0 Then Return

			'divide by 32
			lMouseX \= 32

			'mousex tells us our column
			Dim lDayOfCalendar As Int32 = lMouseY * 7 + lMouseX + 1

			'Ok, day of calender is the day on the calendar if each day on the calendar is visible and represents a valid day...
			' this is not always the case as months may start on any day... not just sunday
			lDayOfCalendar -= mlFirstDayOfMonth

			If lDayOfCalendar < 1 Then Return
			If lDayOfCalendar > mlDaysInMonth Then Return

			mlSelectedMonth = mlDisplayedMonth
			mlSelectedYear = mlDisplayedYear
			mlSelectedDay = lDayOfCalendar

			RaiseEvent DayClick(mlSelectedDay, mlSelectedMonth, mlSelectedYear)
			Me.IsDirty = True
		End If
		
	End Sub

	Private Sub ctlCalendar_OnRender() Handles Me.OnRender
		Dim lDay As Int32 = mlFirstDayOfMonth

		Dim lXOffset As Int32
		Dim lYOffset As Int32

		Dim ptTemp As Point = Me.GetAbsolutePosition()
		lXOffset = ptTemp.X + 2
		lYOffset = ptTemp.Y

		Dim lTop As Int32 = txtSunday.Top + txtSunday.Height + lYOffset

        'If moLine Is Nothing OrElse moLine.Disposed = True Then
        '	moLine = Nothing
        '	Device.IsUsingEventHandlers = False
        '	moLine = New Line(MyBase.moUILib.oDevice)
        '	Device.IsUsingEventHandlers = True
        'End If
        'If moFont Is Nothing OrElse moFont.Disposed = True Then
        '	moFont = Nothing
        '	Device.IsUsingEventHandlers = False
        '          moFont = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '	Device.IsUsingEventHandlers = True
        '      End If
        If moSprite Is Nothing OrElse moSprite.Disposed = True Then
            moSprite = Nothing
            Device.IsUsingEventHandlers = False
            moSprite = New Sprite(MyBase.moUILib.oDevice)
            Device.IsUsingEventHandlers = True
        End If

        'With moLine
        '	.Antialias = False
        '	.Width = 1.5F
        '	.Begin()
        'End With
        BPLine.PrepareMultiDraw(1.5F, False)

		For X As Int32 = 0 To mlDaysInMonth - 1

			If mlDisplayedMonth = mlSelectedMonth AndAlso mlDisplayedYear = mlSelectedYear AndAlso (X + 1) = mlSelectedDay Then
				Dim rcTemp As Rectangle = New Rectangle(lDay * 32 + lXOffset, lTop, 32, 32)
				MyBase.moUILib.DoAlphaBlendColorFill(rcTemp, System.Drawing.Color.FromArgb(255, 128, 128, 160), rcTemp.Location)
			End If

			'ok, draw our box
			Dim vecs(4) As Vector2
			vecs(0).X = (lDay * 32) + lXOffset : vecs(0).Y = lTop
			vecs(1).X = (lDay * 32) + 32 + lXOffset : vecs(1).Y = lTop
			vecs(2).X = (lDay * 32) + 32 + lXOffset : vecs(2).Y = lTop + 32
			vecs(3).X = (lDay * 32) + lXOffset : vecs(3).Y = lTop + 32
			vecs(4).X = (lDay * 32) + lXOffset : vecs(4).Y = lTop
            'moLine.Draw(vecs, muSettings.InterfaceBorderColor)
            BPLine.MultiDrawLine(vecs, muSettings.InterfaceBorderColor)

			lDay += 1
			If lDay > 6 Then
				lDay = 0
				lTop += 32
			End If
        Next X

        'moLine.End()
        BPLine.EndMultiDraw()

		'Now, go back through and render our day numbers
		lDay = mlFirstDayOfMonth
        lTop = txtSunday.Top + txtSunday.Height + lYOffset

        Device.IsUsingEventHandlers = False
        Try
            Using moFont As New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                For X As Int32 = 0 To mlDaysInMonth - 1

                    Dim rcDest As Rectangle = New Rectangle((lDay * 32) + lXOffset, lTop, 32, 32)
                    moFont.DrawText(Nothing, (X + 1).ToString, rcDest, DrawTextFormat.Top Or DrawTextFormat.Right, muSettings.InterfaceBorderColor)
                    lDay += 1
                    If lDay > 6 Then
                        lDay = 0
                        lTop += 32
                    End If
                Next X
            End Using
        Catch
        End Try
        Device.IsUsingEventHandlers = True

		'Finally, do our render of event icons
		If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.moEvents Is Nothing = False Then
			With goCurrentPlayer.oGuild

				If moSprite Is Nothing OrElse moSprite.Disposed = True Then
					moSprite = Nothing
					Device.IsUsingEventHandlers = False
					moSprite = New Sprite(MyBase.moUILib.oDevice)
					Device.IsUsingEventHandlers = True
				End If
				If moTex Is Nothing OrElse moTex.Disposed = True Then moTex = goResMgr.GetTexture("GuildIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "gi.pak")

				Dim bSpriteBegun As Boolean = False
				Try
					Dim yDayIconLvl(30, 2) As Byte
					For X As Int32 = 0 To 30
						For Y As Int32 = 0 To 2
							yDayIconLvl(X, Y) = 255
						Next
					Next X

					Dim lBaseTop As Int32 = txtSunday.Top + txtSunday.Height + lYOffset + 11

					For X As Int32 = 0 To .moEvents.GetUpperBound(0)
						Dim oEvent As GuildEvent = .moEvents(X)
						If oEvent Is Nothing = False Then
							Dim dt As Date = oEvent.dtStartsAt.ToLocalTime
							'event on this date?
							If dt.Month = mlDisplayedMonth AndAlso dt.Year = mlDisplayedYear Then
								'ok, determine where to render this...
								Dim lAdjX As Int32 = Int32.MaxValue
								Dim lAdjY As Int32 = Int32.MaxValue
								For Y As Int32 = 0 To 2
									If yDayIconLvl(dt.Day, Y) = oEvent.yEventIcon Then
										lAdjX = Int32.MinValue : lAdjY = Int32.MinValue
										Exit For
									ElseIf yDayIconLvl(dt.Day, Y) = 255 Then
										Select Case Y
											Case 0
												lAdjX = 0 : lAdjY = 0
											Case 1
												lAdjX = 0 : lAdjY = 16
											Case 2
												lAdjX = 16 : lAdjY = 16
										End Select
										yDayIconLvl(dt.Day, Y) = oEvent.yEventIcon
										Exit For
									End If
								Next Y

								If lAdjX <> Int32.MaxValue AndAlso lAdjX <> Int32.MinValue AndAlso lAdjY <> Int32.MaxValue AndAlso lAdjY <> Int32.MinValue Then
									lAdjX += Me.Left
									lAdjY += Me.Top

									lDay = mlFirstDayOfMonth
									lTop = lBaseTop
									For Y As Int32 = 0 To mlDaysInMonth - 1

										If Y + 1 = dt.Day Then

											Dim rcDest As Rectangle = New Rectangle((lDay * 32) + lXOffset + lAdjX, lTop + lAdjY, 16, 16)
											If bSpriteBegun = False Then
												moSprite.Begin(SpriteFlags.AlphaBlend)
												bSpriteBegun = True
											End If 
											moSprite.Draw2D(moTex, oEvent.rcIconRect, rcDest, New Point(16, 16), 0, rcDest.Location, Color.White)	 ' rcDest.Location, Color.White)

										Else
											lDay += 1
											If lDay > 6 Then
												lDay = 0
												lTop += 32
											End If
										End If
 
									Next Y


								End If

							End If
						End If
					Next X
				Catch
				Finally
					If bSpriteBegun = True Then
						Try
							moSprite.End()
						Catch
							moSprite.Dispose()
							moSprite = Nothing
						End Try
					End If
				End Try

			End With
		End If
		
	End Sub

	Private Sub vscrMonth_ValueChange() Handles vscrMonth.ValueChange
		Dim lMonthChange As Int32 = 5 - vscrMonth.Value
		Dim lCurMonth As Int32 = Now.Month
		Dim lYear As Int32 = Now.Year

		Dim lNewMonth As Int32 = lCurMonth + lMonthChange
		If lNewMonth < 1 Then
			lNewMonth += 12
			lYear -= 1
		ElseIf lNewMonth > 12 Then
			lNewMonth -= 12
			lYear += 1
		End If
		Me.SetMonthAndYear(lNewMonth, lYear)
	End Sub

	Public Sub GetSelectedDate(ByRef lDay As Int32, ByRef lMonth As Int32, ByRef lYear As Int32)
		lDay = mlSelectedDay
		lMonth = mlSelectedMonth
		lYear = mlSelectedYear
	End Sub

	Protected Overrides Sub Finalize()
		If moSprite Is Nothing = False Then moSprite.Dispose()
		moSprite = Nothing
		MyBase.Finalize()
	End Sub
End Class

Public Class CalenderTest
	Inherits UIWindow
	Private moCalender As ctlCalendar
	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)
		With Me
			.ControlName = "CalendarTest"
			.Left = 395
			.Top = 187
			.Width = 512
			.Height = 512
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = True
			.Moveable = True
		End With
		moCalender = New ctlCalendar(oUILib)
		With moCalender
			.Left = 5
			.Top = 5
			.Visible = True
			.Enabled = True
		End With
		Me.AddChild(CType(moCalender, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub
End Class