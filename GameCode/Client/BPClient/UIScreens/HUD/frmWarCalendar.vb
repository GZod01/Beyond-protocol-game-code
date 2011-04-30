Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmWarCalendar
#Region "Version 1, Hours are horizontal (too big)"
    'Inherits UIWindow

    'Private WithEvents btnClose As UIButton
    'Private WithEvents btnHelp As UIButton

    'Private mbLoading As Boolean = True

    'Private Structure CalendarEntry
    '    Public sDate As Date
    '    Public iWeek As Int32
    '    Public iDay As Int32
    '    Public iHour As Int32
    '    Public iWarable As Boolean
    '    Public iColumn As Int32
    '    Public iRow As Int32

    '    'X and Y for Click detection
    '    Public iStartX As Int32
    '    Public iStartY As Int32
    '    Public iEndX As Int32
    '    Public iEndY As Int32
    'End Structure

    'Private iCalendarEntries(336) As CalendarEntry

    'Public Sub New(ByRef oUILib As UILib)
    '    MyBase.New(oUILib)
    '    With Me
    '        .ControlName = "frmWarCalendar"

    '        Dim lLeft As Int32 = -1
    '        Dim lTop As Int32 = -1

    '        lLeft = muSettings.WarCalendarX
    '        lTop = muSettings.WarCalendarY

    '        If lLeft < 0 Then lLeft = 0
    '        If lTop < 0 Then lTop = 0
    '        .Width = (140 * 7)
    '        .Height = 861 + 24 + 5
    '        If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
    '        If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
    '        .Top = lTop
    '        .Left = lLeft

    '        .Enabled = True
    '        .Visible = True
    '        .BorderColor = muSettings.InterfaceBorderColor
    '        .FillColor = muSettings.InterfaceFillColor
    '        .FullScreen = False
    '        .BorderLineWidth = 2
    '    End With

    '    'btnClose initial props
    '    btnClose = New UIButton(oUILib)
    '    With btnClose
    '        .ControlName = "btnClose"
    '        .Left = Me.Width - 24 - Me.BorderLineWidth
    '        .Top = Me.BorderLineWidth
    '        .Width = 24
    '        .Height = 24
    '        .Enabled = True
    '        .Visible = True
    '        .Caption = "X"
    '        .ForeColor = muSettings.InterfaceBorderColor
    '        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '        .DrawBackImage = True
    '        .FontFormat = CType(5, DrawTextFormat)
    '        .ControlImageRect = New Rectangle(0, 0, 120, 32)
    '    End With
    '    Me.AddChild(CType(btnClose, UIControl))

    '    'btnHelp initial props
    '    btnHelp = New UIButton(oUILib)
    '    With btnHelp
    '        .ControlName = "btnHelp"
    '        .Left = btnClose.Left - 25
    '        .Top = btnClose.Top
    '        .Width = 24
    '        .Height = 24
    '        .Enabled = True
    '        .Visible = True
    '        .Caption = "?"
    '        .ForeColor = muSettings.InterfaceBorderColor
    '        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '        .DrawBackImage = True
    '        .FontFormat = CType(5, DrawTextFormat)
    '        .ControlImageRect = New Rectangle(0, 0, 120, 32)
    '    End With
    '    Me.AddChild(CType(btnHelp, UIControl))

    '    Dim lnDiv1 As UILine
    '    Dim lblTitle As UILabel

    '    Dim iDayWidth As Int32 = 140
    '    Dim iDayHeight As Int32 = (22 * 12) + 22 + 1 'Hours + div + title
    '    Dim DP As Int32 = DatePart(DateInterval.Weekday, Now())
    '    Dim HH As Int32 = DatePart(DateInterval.Hour, Now())
    '    For iWeek As Int32 = 0 To 2
    '        'Draw Week Seperator
    '        lnDiv1 = New UILine(oUILib)
    '        With lnDiv1
    '            .ControlName = "lnDiv1"
    '            .Left = 3
    '            .Top = iDayHeight * iWeek + 26
    '            .Width = Me.Width - 4
    '            .Height = 0
    '            .Enabled = True
    '            .Visible = True
    '            .BorderColor = muSettings.InterfaceBorderColor
    '        End With
    '        Me.AddChild(CType(lnDiv1, UIControl))

    '        lnDiv1 = New UILine(oUILib)
    '        With lnDiv1
    '            .ControlName = "lnDiv1"
    '            .Left = 3
    '            .Top = iDayHeight * iWeek + 22 + 22
    '            .Width = Me.Width - 4
    '            .Height = 0
    '            .Enabled = True
    '            .Visible = True
    '            .BorderColor = muSettings.InterfaceBorderColor
    '        End With
    '        Me.AddChild(CType(lnDiv1, UIControl))

    '        For IDay As Int32 = 0 To 6
    '            'Add Day Div
    '            If IDay > 0 Then
    '                lnDiv1 = New UILine(oUILib)
    '                With lnDiv1
    '                    .ControlName = "lnDiv1"
    '                    .Left = IDay * iDayWidth
    '                    If IDay = 1 Then
    '                        Debug.Print(iWeek & ":" & .Left)
    '                    End If
    '                    .Top = iDayHeight * iWeek + 22 + 4
    '                    .Width = 1
    '                    .Height = iDayHeight
    '                    .Enabled = True
    '                    .Visible = True
    '                    .BorderColor = muSettings.InterfaceBorderColor
    '                End With
    '                Me.AddChild(CType(lnDiv1, UIControl))
    '            End If
    '            If (iWeek = 0 AndAlso IDay + 1 >= DP) OrElse iWeek = 1 OrElse (iWeek = 2 And IDay + 1 <= DP) Then
    '                'Label Day
    '                lblTitle = New UILabel(oUILib)
    '                With lblTitle
    '                    .ControlName = "lblTitle"
    '                    .Left = 3 + (IDay * iDayWidth)
    '                    .Top = iDayHeight * iWeek + 23
    '                    .Width = iDayWidth
    '                    .Height = 22
    '                    .Enabled = True
    '                    .Visible = True
    '                    .Caption = GetDayOfWeekName(IDay)
    '                    .ForeColor = muSettings.InterfaceBorderColor
    '                    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '                    .DrawBackImage = False
    '                    .FontFormat = CType(4, DrawTextFormat)
    '                End With
    '                Me.AddChild(CType(lblTitle, UIControl))
    '            End If

    '            For iHour As Int32 = 0 To 23
    '                If (iWeek = 0 And IDay + 1 >= DP And ((IDay + 1 = DP And iHour > HH) Or IDay + 1 > DP)) Or iWeek = 1 Or (iWeek = 2 And IDay + 1 <= DP And ((IDay + 1 = DP And iHour < HH) Or IDay + 1 < DP)) Then
    '                    'Draw Hour Label
    '                    lblTitle = New UILabel(oUILib)
    '                    With lblTitle
    '                        .ControlName = "lblTitle"
    '                        .Width = iDayWidth
    '                        .Height = 22
    '                        .Enabled = True
    '                        .Visible = True
    '                        If iHour < 12 Then
    '                            If iHour = 0 Then
    '                                .Caption = "12 AM"
    '                            Else
    '                                .Caption = iHour & " AM"
    '                            End If
    '                            .Left = 3 + (IDay * iDayWidth)
    '                            .Top = iDayHeight * iWeek + 23 + (iHour * 22)
    '                        Else
    '                            If iHour = 12 Then
    '                                .Caption = "12 PM"
    '                            Else
    '                                .Caption = iHour - 12 & " PM"
    '                            End If
    '                            .Left = 3 + (IDay * iDayWidth) + iDayWidth \ 2
    '                            .Top = iDayHeight * iWeek + 23 + ((iHour - 12) * 22)
    '                        End If
    '                        .Top += 22
    '                        .Left += 25
    '                        .ForeColor = muSettings.InterfaceBorderColor
    '                        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 7.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '                        .DrawBackImage = False
    '                        .FontFormat = CType(4, DrawTextFormat)
    '                    End With
    '                    Me.AddChild(CType(lblTitle, UIControl))
    '                End If
    '            Next iHour
    '        Next IDay
    '    Next iWeek

    '    Dim iEntry As Int32 = 1
    '    For iWeek As Int32 = 0 To 2
    '        For IDay As Int32 = 0 To 6
    '            For iHour As Int32 = 0 To 23
    '                If (iWeek = 0 And IDay + 1 >= DP And ((IDay + 1 = DP And iHour > HH) Or IDay + 1 > DP)) Or iWeek = 1 Or (iWeek = 2 And IDay + 1 <= DP And ((IDay + 1 = DP And iHour < HH) Or IDay + 1 < DP)) Then
    '                    With iCalendarEntries(iEntry)
    '                        .iColumn = IDay
    '                        .iRow = iHour
    '                        .sDate = DateAdd(DateInterval.Day, IDay, Now())
    '                        .iWarable = CBool(CInt(1 * Rnd()))
    '                        .iWeek = iWeek
    '                        .iDay = IDay
    '                        .iHour = iHour
    '                        Debug.Print("W=" & iWeek & " D=" & IDay & " H=" & iHour & " W=" & .iWarable)
    '                    End With
    '                    iEntry += 1
    '                End If
    '            Next
    '        Next
    '    Next

    '    MyBase.moUILib.RemoveWindow(Me.ControlName)
    '    MyBase.moUILib.AddWindow(Me)
    '    mbLoading = False
    'End Sub

    'Private Sub frmWarCalendar_OnRender() Handles Me.OnRender
    '    Dim iDayWidth As Int32 = 140
    '    Dim iDayHeight As Int32 = (22 * 12) + 22 + 1 'Hours + div + title

    '    Dim rcCancel As Rectangle = New Rectangle(460, 5, 17, 17)
    '    Dim clrVal As System.Drawing.Color
    '    Using oSprite As New Sprite(MyBase.moUILib.oDevice)
    '        oSprite.Begin(SpriteFlags.AlphaBlend)

    '        For iEntry As Int32 = 1 To 335
    '            With iCalendarEntries(iEntry)
    '                clrVal = Color.Red
    '                Dim rcCancelButtonSrc As Rectangle = New Rectangle(127, 143, 17, 17)
    '                If .iHour < 12 Then
    '                    rcCancel.X = 3 + (.iDay * iDayWidth)
    '                    rcCancel.Y = iDayHeight * .iWeek + 23 + (.iHour * 22)
    '                Else

    '                    rcCancel.X = 3 + (.iDay * iDayWidth) + iDayWidth \ 2
    '                    rcCancel.Y = iDayHeight * .iWeek + 23 + ((.iHour - 12) * 22)
    '                End If
    '                rcCancel.Y += 25
    '                If .iWarable = True Then
    '                    clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
    '                Else
    '                    clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
    '                End If
    '                oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcCancelButtonSrc, rcCancel, Point.Empty, 0, rcCancel.Location, clrVal)

    '            End With
    '        Next

    '        oSprite.End()
    '        oSprite.Dispose()
    '    End Using
    'End Sub
#End Region
#Region "Version 2, Hours are horizontal"
    'Inherits UIWindow

    'Private WithEvents btnClose As UIButton
    'Private WithEvents btnHelp As UIButton

    'Private mbLoading As Boolean = True

    'Private Structure CalendarEntry
    '    Public sDate As Date
    '    Public iWeek As Int32
    '    Public iDay As Int32
    '    Public iHour As Int32
    '    Public iWarable As Boolean
    '    Public iColumn As Int32
    '    Public iRow As Int32

    '    'X and Y for Click detection
    '    Public iStartX As Int32
    '    Public iStartY As Int32

    '    Public bHasEntry As Boolean

    '    Public bEditable As Boolean
    'End Structure

    'Private iCalendarEntries(336) As CalendarEntry
    'Private miSpacerY As Int32 = 30
    'Private miSpacerX As Int32 = 30

    'Public Sub New(ByRef oUILib As UILib)
    '    MyBase.New(oUILib)
    '    miSpacerY = 25
    '    With Me
    '        .ControlName = "frmWarCalendar"

    '        Dim lLeft As Int32 = -1
    '        Dim lTop As Int32 = -1

    '        lLeft = muSettings.WarCalendarX
    '        lTop = muSettings.WarCalendarY

    '        If lLeft < 0 Then lLeft = 0
    '        If lTop < 0 Then lTop = 0
    '        .Width = (24 * miSpacerY) + 80 + 10
    '        .Height = 330
    '        If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
    '        If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
    '        .Top = lTop
    '        .Left = lLeft

    '        .Enabled = True
    '        .Visible = True
    '        .BorderColor = muSettings.InterfaceBorderColor
    '        .FillColor = muSettings.InterfaceFillColor
    '        .FullScreen = False
    '        .BorderLineWidth = 2
    '    End With

    '    'btnClose initial props
    '    btnClose = New UIButton(oUILib)
    '    With btnClose
    '        .ControlName = "btnClose"
    '        .Left = Me.Width - 24 - Me.BorderLineWidth
    '        .Top = Me.BorderLineWidth
    '        .Width = 24
    '        .Height = 24
    '        .Enabled = True
    '        .Visible = True
    '        .Caption = "X"
    '        .ForeColor = muSettings.InterfaceBorderColor
    '        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '        .DrawBackImage = True
    '        .FontFormat = CType(5, DrawTextFormat)
    '        .ControlImageRect = New Rectangle(0, 0, 120, 32)
    '    End With
    '    Me.AddChild(CType(btnClose, UIControl))

    '    'btnHelp initial props
    '    btnHelp = New UIButton(oUILib)
    '    With btnHelp
    '        .ControlName = "btnHelp"
    '        .Left = btnClose.Left - 25
    '        .Top = btnClose.Top
    '        .Width = 24
    '        .Height = 24
    '        .Enabled = True
    '        .Visible = True
    '        .Caption = "?"
    '        .ForeColor = muSettings.InterfaceBorderColor
    '        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '        .DrawBackImage = True
    '        .FontFormat = CType(5, DrawTextFormat)
    '        .ControlImageRect = New Rectangle(0, 0, 120, 32)
    '    End With
    '    Me.AddChild(CType(btnHelp, UIControl))

    '    Dim lnDiv1 As UILine
    '    Dim lblTitle As UILabel

    '    Dim iDayWidth As Int32 = 140
    '    Dim iDayHeight As Int32 = 22
    '    Dim DP As Int32 = DatePart(DateInterval.Weekday, Now()) - 1
    '    Dim HH As Int32 = DatePart(DateInterval.Hour, Now())

    '    Dim iDayi As Int32 = 0
    '    Dim iTop As Int32 = 20
    '    Dim iWeekStart As Int32 = 22 + iTop
    '    Dim iLeft As Int32 = 85

    '    'Dim tLocalTime As DateTimeOffset = New DateTimeOffset(Now())
    '    'Dim tOtherTime As DateTimeOffset = tLocalTime.ToOffset(TimeSpan.Zero)
    '    'Dim iOffset As Int32 = Math.Abs(tOtherTime.Hour - tLocalTime.Hour)

    '    Dim iOffset As Int32 = 4
    '    For iDay As Int32 = 0 + DP To 14 + DP
    '        If iDay > 13 Then
    '            iDayi = iDay - 14
    '        ElseIf iDay > 6 Then
    '            iDayi = iDay - 7
    '        Else
    '            iDayi = iDay
    '        End If

    '        'Day Label
    '        lblTitle = New UILabel(oUILib)
    '        With lblTitle
    '            .ControlName = "lblTitle"
    '            .Left = 5
    '            .Top = iTop + 22
    '            .Width = 125
    '            .Height = 22
    '            .Enabled = True
    '            .Visible = True
    '            .Caption = GetDayOfWeekName(iDayi)
    '            .ForeColor = muSettings.InterfaceBorderColor
    '            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '            .DrawBackImage = False
    '            .FontFormat = CType(4, DrawTextFormat)
    '        End With
    '        Me.AddChild(CType(lblTitle, UIControl))

    '        'Draw Day Seperator
    '        lnDiv1 = New UILine(oUILib)
    '        With lnDiv1
    '            .ControlName = "lnDiv1"
    '            .Left = 3
    '            .Top = iTop + 22
    '            .Width = 2 + 80 + (24 * miSpacerY)
    '            .Height = 0
    '            .Enabled = True
    '            .Visible = True
    '            .BorderColor = muSettings.InterfaceBorderColor
    '        End With
    '        Me.AddChild(CType(lnDiv1, UIControl))

    '        If iDayi = 6 Then
    '            lnDiv1 = New UILine(oUILib)
    '            With lnDiv1
    '                .ControlName = "lnDiv1"
    '                .Left = 3
    '                .Top = iTop + 44
    '                .Width = 3 + 80 + (24 * miSpacerY)
    '                .Height = 0
    '                .Enabled = True
    '                .Visible = True
    '                .BorderColor = muSettings.InterfaceBorderColor
    '            End With
    '            Me.AddChild(CType(lnDiv1, UIControl))
    '        End If



    '        iTop += iDayHeight
    '        If iDayi = 6 Then
    '            iLeft = 85
    '            For iHour As Int32 = 0 To 24
    '                lnDiv1 = New UILine(oUILib)
    '                With lnDiv1
    '                    .ControlName = "lnDiv1"
    '                    .Left = iLeft
    '                    .Top = iWeekStart
    '                    .Width = 1
    '                    .Height = iTop - iWeekStart + 22
    '                    .Enabled = True
    '                    .Visible = True
    '                    .BorderColor = muSettings.InterfaceBorderColor
    '                End With
    '                Me.AddChild(CType(lnDiv1, UIControl))
    '                'iLeft += 35
    '                iLeft += miSpacerY
    '            Next iHour
    '            iTop += iDayHeight
    '            iWeekStart = iTop + 22
    '        End If
    '    Next
    '    'Closing Entry
    '    If iDayi < 6 Then
    '        iLeft = 85
    '        For iHour As Int32 = 0 To 24
    '            lnDiv1 = New UILine(oUILib)
    '            With lnDiv1
    '                .ControlName = "lnDiv1"
    '                .Left = iLeft
    '                .Top = iWeekStart
    '                .Width = 1
    '                .Height = iTop - iWeekStart + 22
    '                .Enabled = True
    '                .Visible = True
    '                .BorderColor = muSettings.InterfaceBorderColor
    '            End With
    '            Me.AddChild(CType(lnDiv1, UIControl))
    '            iLeft += miSpacerY
    '        Next iHour

    '        lnDiv1 = New UILine(oUILib)
    '        With lnDiv1
    '            .ControlName = "lnDiv1"
    '            .Left = 3
    '            .Top = iTop + 22
    '            .Width = 2 + 80 + (24 * miSpacerY)
    '            .Height = 0
    '            .Enabled = True
    '            .Visible = True
    '            .BorderColor = muSettings.InterfaceBorderColor
    '        End With
    '        Me.AddChild(CType(lnDiv1, UIControl))
    '        iTop += iDayHeight
    '    End If

    '    Me.Height = iTop + 5

    '    iLeft = 88
    '    For iHour As Int32 = 0 To 23
    '        lblTitle = New UILabel(oUILib)
    '        With lblTitle
    '            .ControlName = "lblTitle"
    '            .Left = iLeft
    '            .Top = 24
    '            .Width = miSpacerY
    '            .Height = 22
    '            .Enabled = True
    '            .Visible = True
    '            If iHour = 0 Then
    '                .Caption = "12a"
    '            ElseIf iHour > 11 Then
    '                If (iHour - 12) = 0 Then
    '                    .Caption = "12p"
    '                Else
    '                    .Caption = iHour - 12 & "p"
    '                End If
    '            Else
    '                .Caption = iHour.ToString & "a"
    '            End If
    '            .ForeColor = muSettings.InterfaceBorderColor
    '            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '            .DrawBackImage = False
    '            .FontFormat = CType(4, DrawTextFormat)
    '            'iLeft += 35
    '            iLeft += miSpacerY
    '        End With
    '        Me.AddChild(CType(lblTitle, UIControl))
    '    Next

    '    Dim iEntry As Int32 = 1
    '    Dim iRow As Int32 = 0
    '    For iWeek As Int32 = 0 To 2
    '        For IDay As Int32 = 0 To 6
    '            For iHour As Int32 = 0 To 23
    '                If (iWeek = 0 And IDay >= DP And ((IDay = DP And iHour > HH) Or IDay > DP)) Or iWeek = 1 Or (iWeek = 2 And IDay <= DP And ((IDay = DP And iHour < HH) Or IDay < DP)) Then
    '                    With iCalendarEntries(iEntry)
    '                        .bHasEntry = True
    '                        .iRow = iRow - DP
    '                        .sDate = DateAdd(DateInterval.Day, IDay, Now())
    '                        '.iWarable = CBool(CInt(1 * Rnd()))
    '                        .iWarable = True
    '                        If iEntry > 168 Then
    '                            .bEditable = True
    '                        End If
    '                        If .bEditable = True Then .iWarable = False

    '                        'Downtime Protection
    '                        If IDay > 0 And iHour - iOffset = 6 Then
    '                            .iWarable = False
    '                            .bEditable = False
    '                            .bHasEntry = False
    '                        End If
    '                        .iWeek = iWeek
    '                        .iDay = IDay
    '                        .iHour = iHour

    '                        Debug.Print("R=" & .iRow.ToString & " C=" & .iHour.ToString)
    '                    End With
    '                    iEntry += 1
    '                End If
    '            Next
    '            iRow += 1
    '            If IDay = 6 Then iRow += 1
    '        Next
    '    Next
    '    Debug.Print("E=" & iEntry.ToString)
    '    MyBase.moUILib.RemoveWindow(Me.ControlName)
    '    MyBase.moUILib.AddWindow(Me)
    '    mbLoading = False
    'End Sub

    'Private Sub frmWarCalendar_OnRender() Handles Me.OnRender

    '    Dim rcCancel As Rectangle = New Rectangle(460, 5, 17, 17)
    '    Dim clrVal As System.Drawing.Color
    '    Using oSprite As New Sprite(MyBase.moUILib.oDevice)
    '        oSprite.Begin(SpriteFlags.AlphaBlend)

    '        For iEntry As Int32 = 1 To 335
    '            With iCalendarEntries(iEntry)
    '                If .bHasEntry = True Then
    '                    clrVal = Color.Red
    '                    Dim rcCancelButtonSrc As Rectangle = New Rectangle(127, 143, 17, 17)
    '                    rcCancel.X = 90 + (miSpacerY * .iHour)
    '                    rcCancel.Y = 25 + (22 * .iRow) + 20

    '                    'rcCancel.Y += 25
    '                    If .iWarable = True Then
    '                        clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
    '                    Else
    '                        clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
    '                    End If
    '                    oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcCancelButtonSrc, rcCancel, Point.Empty, 0, rcCancel.Location, clrVal)
    '                End If
    '            End With
    '        Next

    '        oSprite.End()
    '        oSprite.Dispose()
    '    End Using
    'End Sub
#End Region
#Region "Version 3, Enoch Edition"
#End Region
    'Private Class fraCurrentSettings
    Inherits UIWindow

    Private WithEvents btnClose As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents lnLine As UILine

    Private mfraEdit As fraEdit

    Private mbLoading As Boolean = True
    Private moLines() As UILine

    Shared mbDataHasChanged As Boolean = False
    Private moSysFont As System.Drawing.Font
    Private moSysFontBold As System.Drawing.Font

    Private rcDays(7) As Rectangle
    Private miDayLabels(7) As Int32
    Private miDayEditing As Int32 = -1

    Private Structure CalendarEntry
        Public sDate As Date
        Public iWeek As Int32
        Public iDay As Int32
        Public iHour As Int32
        Public iWarable As Boolean
        Public iWarableOrig As Boolean
        Public iColumn As Int32
        Public iRow As Int32

        'X and Y for Click detection
        Public iStartX As Int32
        Public iStartY As Int32

        Public bHasEntry As Boolean

        Public bEditable As Boolean
    End Structure

    Private Shared moCalendarEntries(336) As CalendarEntry
    Private miSpacerY As Int32 = 30
    Private miSpacerX As Int32 = 30

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
        miSpacerY = 25
        With Me
            .ControlName = "frmWarCalendar"

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1
            .bRoundedBorder = False
            lLeft = muSettings.WarCalendarX
            lTop = muSettings.WarCalendarY

            If lLeft < 0 Then lLeft = 0
            If lTop < 0 Then lTop = 0
            .Width = 865
            .Height = 390
            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Top = lTop
            .Left = lLeft

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
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

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = btnClose.Left - 25
            .Top = btnClose.Top
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

        Dim lnDiv1 As UILine
        'Dim lblTitle As UILabel

        'Dim iDayWidth As Int32 = 140
        'Dim iDayHeight As Int32 = 22
        Dim DP As Int32 = DatePart(DateInterval.Weekday, Now()) - 1
        Dim HH As Int32 = DatePart(DateInterval.Hour, Now())

        'Dim iDayi As Int32 = 0
        'Dim iTop As Int32 = 20
        'Dim iWeekStart As Int32 = 22 + iTop
        'Dim iLeft As Int32 = 85

        'Dim tLocalTime As DateTimeOffset = New DateTimeOffset(Now())
        'Dim tOtherTime As DateTimeOffset = tLocalTime.ToOffset(TimeSpan.Zero)
        'Dim iOffset As Int32 = Math.Abs(tOtherTime.Hour - tLocalTime.Hour)



        ''Time Lables
        'For iDay As Int32 = 0 + DP To 14 + DP

        'Next

        ''Hour Slots
        'For iDay As Int32 = 0 + DP To 14 + DP

        'Next


        Dim iOffset As Int32 = 4
        'For iDay As Int32 = 0 + DP To 14 + DP
        '    If iDay > 13 Then
        '        iDayi = iDay - 14
        '    ElseIf iDay > 6 Then
        '        iDayi = iDay - 7
        '    Else
        '        iDayi = iDay
        '    End If

        '    'Day Label
        '    lblTitle = New UILabel(oUILib)
        '    With lblTitle
        '        .ControlName = "lblTitle"
        '        .Left = 5
        '        .Top = iTop + 22
        '        .Width = 125
        '        .Height = 22
        '        .Enabled = True
        '        .Visible = True
        '        .Caption = GetDayOfWeekName(iDayi)
        '        .ForeColor = muSettings.InterfaceBorderColor
        '        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '        .DrawBackImage = False
        '        .FontFormat = CType(4, DrawTextFormat)
        '    End With
        '    Me.AddChild(CType(lblTitle, UIControl))

        '    'Draw Day Seperator
        '    lnDiv1 = New UILine(oUILib)
        '    With lnDiv1
        '        .ControlName = "lnDiv1"
        '        .Left = 3
        '        .Top = iTop + 22
        '        .Width = 2 + 80 + (24 * miSpacerY)
        '        .Height = 0
        '        .Enabled = True
        '        .Visible = True
        '        .BorderColor = muSettings.InterfaceBorderColor
        '    End With
        '    Me.AddChild(CType(lnDiv1, UIControl))

        '    If iDayi = 6 Then
        '        lnDiv1 = New UILine(oUILib)
        '        With lnDiv1
        '            .ControlName = "lnDiv1"
        '            .Left = 3
        '            .Top = iTop + 44
        '            .Width = 3 + 80 + (24 * miSpacerY)
        '            .Height = 0
        '            .Enabled = True
        '            .Visible = True
        '            .BorderColor = muSettings.InterfaceBorderColor
        '        End With
        '        Me.AddChild(CType(lnDiv1, UIControl))
        '    End If

        '    iTop += iDayHeight
        '    If iDayi = 6 Then
        '        iLeft = 85
        '        For iHour As Int32 = 0 To 24
        '            'Add the verticle line per hour block, starting at week start/stop for snazzyness
        '            lnDiv1 = New UILine(oUILib)
        '            With lnDiv1
        '                .ControlName = "lnDiv1"
        '                .Left = iLeft
        '                .Top = iWeekStart
        '                .Width = 1
        '                .Height = iTop - iWeekStart + 22
        '                .Enabled = True
        '                .Visible = True
        '                .BorderColor = muSettings.InterfaceBorderColor
        '            End With
        '            Me.AddChild(CType(lnDiv1, UIControl))
        '            iLeft += miSpacerY
        '        Next iHour
        '        iTop += iDayHeight
        '        iWeekStart = iTop + 22
        '    End If
        'Next

        ''Closing Entry
        'If iDayi < 6 Then
        '    iLeft = 85
        '    For iHour As Int32 = 0 To 24
        '        lnDiv1 = New UILine(oUILib)
        '        With lnDiv1
        '            .ControlName = "lnDiv1"
        '            .Left = iLeft
        '            .Top = iWeekStart
        '            .Width = 1
        '            .Height = iTop - iWeekStart + 22
        '            .Enabled = True
        '            .Visible = True
        '            .BorderColor = muSettings.InterfaceBorderColor
        '        End With
        '        Me.AddChild(CType(lnDiv1, UIControl))
        '        iLeft += miSpacerY
        '    Next iHour

        '    lnDiv1 = New UILine(oUILib)
        '    With lnDiv1
        '        .ControlName = "lnDiv1"
        '        .Left = 3
        '        .Top = iTop + 22
        '        .Width = 2 + 80 + (24 * miSpacerY)
        '        .Height = 0
        '        .Enabled = True
        '        .Visible = True
        '        .BorderColor = muSettings.InterfaceBorderColor
        '    End With
        '    Me.AddChild(CType(lnDiv1, UIControl))
        '    iTop += iDayHeight
        'End If

        'Me.Height = iTop + 5

        'iLeft = 88
        'For iHour As Int32 = 0 To 23
        '    lblTitle = New UILabel(oUILib)
        '    With lblTitle
        '        .ControlName = "lblTitle"
        '        .Left = iLeft
        '        .Top = 24
        '        .Width = miSpacerY
        '        .Height = 22
        '        .Enabled = True
        '        .Visible = True
        '        If iHour = 0 Then
        '            .Caption = "12a"
        '        ElseIf iHour > 11 Then
        '            If (iHour - 12) = 0 Then
        '                .Caption = "12p"
        '            Else
        '                .Caption = iHour - 12 & "p"
        '            End If
        '        Else
        '            .Caption = iHour.ToString & "a"
        '        End If
        '        .ForeColor = muSettings.InterfaceBorderColor
        '        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '        .DrawBackImage = False
        '        .FontFormat = CType(4, DrawTextFormat)
        '        iLeft += miSpacerY
        '    End With
        '    Me.AddChild(CType(lblTitle, UIControl))
        'Next

        Dim iEntry As Int32 = 1
        ReDim moLines(335)
        Dim iRow As Int32 = 0

        Dim iDays As Int32 = 0
        Dim iDayBtns As Int32 = 0
        Dim iLeft As Int32 = 14
        Dim iTop As Int32 = 40
        Dim iLabelLeft As Int32 = 10
        Dim bFirstWeek As Boolean = True
        Dim bNewDay As Boolean = False
        Dim bStarted As Boolean = False 'Tells us if we have reached an approperiate start day/hour to begin
        moSysFont = New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0)
        moSysFontBold = New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Bold, GraphicsUnit.Point, 0)
        'HH = 23
        HH = 0
        DP = 1

        Dim iLineWidth As Int32 = 4
        For iWeek As Int32 = 0 To 2
            For iDay As Int32 = 0 To 6
                bNewDay = True
                For iHour As Int32 = 0 To 23
                    If (iWeek = 0 And iDay >= DP And ((iDay = DP And iHour > HH) Or iDay > DP)) Or iWeek = 1 Or (iWeek = 2 And iDay <= DP And ((iDay = DP And iHour < HH) Or iDay < DP)) Then
                        'If (iWeek = 0 And iDay >= DP And ((iDay = DP And iHour >= HH) Or iDay > DP)) Or (iDays > 7 And iDay <= DP And ((iDay = DP And iHour < HH) Or iDay < DP)) Then
                        'If (iDays < 7 AndAlso iDay >= DP And iHour >= HH) Or (iDays >= 7 And iDay <= DP Or (iDay = DP And iHour < HH)) Then
                        If bNewDay = True And iDayBtns < 8 AndAlso iEntry < 168 Then
                            Debug.Print("Doing Day Button - " & iDay & " @ " & iLeft)
                            bNewDay = False
                            rcDays(iDayBtns) = New Rectangle(10 + iDayBtns * 100, iTop + 25 + 25, 90, 20)
                            miDayLabels(iDayBtns) = iDay
                            iDayBtns += 1
                        End If
                        If iHour = 0 Then
                            iLeft += iLineWidth
                            Debug.Print("Day Spacer @ " & iWeek & "-" & iDay & "-" & iHour)
                            iDays += 1
                        End If
                        If iDays = 7 Then
                            If iHour = 0 Then ' HH Then
                                Debug.Print("Next Line @ " & iWeek & "-" & iDay & "-" & iHour)
                                bFirstWeek = False
                                iTop += 25
                                iLeft = 10
                                'bStarted = False
                            End If
                        End If
                        If bStarted = False Then
                            iLeft += (HH * iLineWidth)
                            bStarted = True
                        End If
                        With moCalendarEntries(iEntry)
                            .bHasEntry = True
                            .sDate = DateAdd(DateInterval.Day, iDay, Now())
                            .iWarable = True
                            If iEntry > 168 Then
                                .bEditable = True
                            End If


                            .iWeek = iWeek
                            .iDay = iDay
                            .iHour = iHour
                            If iDays > 8 And iDays < 10 And iHour > 8 And iHour < 16 Then
                                .iWarable = False
                            End If
                            If iDays > 13 And iDays < 24 And iHour > 8 And iHour < 16 Then
                                .iWarable = False
                            End If
                            'Downtime Protection
                            If iDay > 0 And iHour - iOffset = 6 Then
                                .iWarable = False
                                .bEditable = False
                                .bHasEntry = False
                                If .iHour <> 10 Then
                                    Dim xxx As String = ""
                                End If
                            End If
                            .iWarableOrig = .iWarable
                            If bFirstWeek = True AndAlso iHour Mod 4 = 0 Then
                                'lblTitle = New UILabel(oUILib)
                                'With lblTitle
                                '    .ControlName = "lblTitle"
                                '    .Left = iLabelLeft
                                '    .Top = 20
                                '    .Width = 5
                                '    .Height = 15
                                '    .Enabled = True
                                '    .Visible = True
                                '    .Caption = Mid(GetDayOfWeekName(iDay), 1, 3)
                                '    .ForeColor = muSettings.InterfaceBorderColor
                                '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
                                '    .DrawBackImage = False
                                '    .FontFormat = CType(4, DrawTextFormat)
                                'End With
                                'Me.AddChild(CType(lblTitle, UIControl))
                                iLabelLeft += 15
                            End If

                            'Add Bar
                            Debug.Print("Line @ " & iWeek & "-" & iDay & "-" & iHour)
                            lnDiv1 = New UILine(oUILib)
                            With lnDiv1
                                .ControlName = "lnDiv1"
                                .Left = iLeft
                                .Top = iTop
                                .Width = 0
                                .Height = 20
                                .Enabled = True
                                .Visible = True
                                If moCalendarEntries(iEntry).bHasEntry = False Then
                                    .BorderColor = Color.Gray
                                ElseIf moCalendarEntries(iEntry).iWarable = False Then
                                    .BorderColor = Color.Green
                                Else
                                    .BorderColor = Color.Red
                                End If
                                .LineWidth = iLineWidth
                            End With
                            Me.AddChild(CType(lnDiv1, UIControl))
                            iLeft += iLineWidth


                            'Add Divider Bar
                            'lnDiv1 = New UILine(oUILib)
                            'With lnDiv1
                            '    .ControlName = "lnDiv1"
                            '    .Left = iLeft
                            '    .Top = iTop
                            '    .Width = 2
                            '    .Height = 20
                            '    .Enabled = True
                            '    .Visible = False
                            '    .BorderColor = muSettings.InterfaceBorderColor
                            'End With
                            'Me.AddChild(CType(lnDiv1, UIControl))
                            'iLeft += 2


                            'Debug.Print("R=" & .iRow.ToString & " C=" & .iHour.ToString)
                        End With
                        iEntry += 1
                    End If 'Valid Timeslot
                Next iHour
            Next iDay
        Next iWeek

        'mfraHeader Initial Props
        mfraEdit = New fraEdit(oUILib)
        With mfraEdit
            .Left = 10
            .Top = 120
            .Height = 265
            .Visible = True
            .Enabled = True
        End With
        Me.AddChild(CType(mfraEdit, UIControl))

        'Debug.Print("E=" & iEntry.ToString)

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
    End Sub

    Private Sub DoMultiColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point, ByRef oSpr As Sprite)
        Dim rcSrc As Rectangle

        Dim fX As Single
        Dim fY As Single

        If rcDest.Width = 0 OrElse rcDest.Height = 0 Then Exit Sub

        rcSrc.Location = New Point(192, 0)
        rcSrc.Width = 62
        rcSrc.Height = 64

        'Now, draw it...
        With oSpr
            fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
            fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))
            .Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFill)
        End With
    End Sub

    Private Sub frmWarCalendar_OnRender() Handles Me.OnRender
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        If miDayEditing <> -1 Then
            Using oFillSprite As New Sprite(MyBase.moUILib.oDevice)
                oFillSprite.Begin(SpriteFlags.AlphaBlend)
                Dim oSelColor As System.Drawing.Color
                oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)
                DoMultiColorFill(rcDays(miDayEditing), oSelColor, rcDays(miDayEditing).Location, oFillSprite)
                oFillSprite.End()
                oFillSprite.Dispose()
            End Using
        End If
        For x As Int32 = 0 To 7
            Using oBoldFont As Font = New Font(MyBase.moUILib.oDevice, moSysFontBold)
                Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                    MyBase.RenderRoundedBorder(rcDays(x), 1, muSettings.InterfaceBorderColor)
                    oTextSpr.Begin(SpriteFlags.AlphaBlend)
                    Try
                        oBoldFont.DrawText(oTextSpr, GetDayOfWeekName(miDayLabels(x)), rcDays(x), DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                    Catch
                    End Try
                    oTextSpr.End()
                    oTextSpr.Dispose()
                End Using
            End Using
        Next x
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        If mbDataHasChanged = True Then
            Dim oFrm As New frmMsgBox(goUILib, "You have made changes to your war calendar, and should save.  Are you sure you wish to close the window?", MsgBoxStyle.YesNo, "Confirm Close")
            oFrm.Visible = True
            AddHandler oFrm.DialogClosed, AddressOf frm_CloseWindow
        Else : MyBase.moUILib.RemoveWindow(Me.ControlName)
        End If
    End Sub

    Private Sub frm_CloseWindow(ByVal lResult As MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            MyBase.moUILib.RemoveWindow(Me.ControlName)
        End If
    End Sub

    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
        goUILib.AddNotification("War Calendar Tutorial is not yet implemented", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub frmWarCalendar_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        If lButton <> MouseButtons.Left Then Exit Sub
        Dim lX As Int32 = lMouseX - Me.GetAbsolutePosition.X
        Dim lY As Int32 = lMouseY - Me.GetAbsolutePosition.Y
        For x As Int32 = 0 To 6
            If rcDays(x).Contains(lX, lY) = True Then
                miDayEditing = x

                LoadDay(x)
                Me.IsDirty = True
            End If
        Next
    End Sub

    Private Sub frmWarCalendar_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim xxx As String = ""
        If mfraEdit.miSetting <> -1 Then
            Me.Moveable = False
        Else
            Me.Moveable = True
        End If
    End Sub

    Private Sub frmWarCalendar_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseUp
        Dim xxx As String = ""
    End Sub

    Private Sub frmWarCalendar_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.WarCalendarX = Me.Left
            muSettings.WarCalendarY = Me.Top
        End If
    End Sub

    Private Sub LoadDay(ByVal iDay As Int32)
        mfraEdit.LoadDay(iDay)
    End Sub

    Public Class fraEdit
        Inherits UIWindow

        Private lblTitle As UILabel
        Private lblHours(23) As UILabel
        Private btnHours(23) As UIButton
        Private WithEvents btnClear As UIButton
        Private WithEvents btnSet As UIButton
        Private WithEvents btnReset As UIButton
        Private WithEvents btnSave As UIButton
        Private Shared moTodaysEntries(24) As CalendarEntry
        Private rcHours(23) As Rectangle
        Public miSetting As Int32 = -1

        Private miDay As Int32 = -1

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraHeader initial props
            With Me
                .ControlName = "fraHeader"
                '.Left = 119
                '.Top = 213
                .Width = 680
                '.Height = 70
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 1
                .Moveable = False
            End With

            'lblFrom initial props
            lblTitle = New UILabel(oUILib)
            With lblTitle
                .ControlName = "lblFrom"
                .Left = 5
                .Top = 5
                .Width = 330
                .Height = 18
                .Enabled = True
                .Visible = False
                .Caption = "Configure Settings for"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTitle, UIControl))

            Dim iHH As Int32 = 0
            For iColumn As Int32 = 0 To 1
                For iHour As Int32 = 0 To 11
                    'lblFrom initial props

                    lblHours(iHH) = New UILabel(oUILib)
                    With lblHours(iHH)
                        .ControlName = "lblHours"
                        .Left = 5 + iColumn * 100
                        .Top = 25 + iHour * 20
                        .Width = 50
                        .Height = 22
                        .Enabled = True
                        .Visible = False
                        If iHour = 0 Then
                            .Caption = "12"
                        Else
                            .Caption = iHour.ToString
                        End If
                        If iColumn = 0 Then
                            .Caption &= " AM"
                        Else
                            .Caption &= " PM"
                        End If
                        .ForeColor = muSettings.InterfaceBorderColor
                        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                        .DrawBackImage = False
                        .FontFormat = CType(2, DrawTextFormat)
                    End With
                    Me.AddChild(CType(lblHours(iHH), UIControl))

                    btnHours(iHH) = New UIButton(oUILib)
                    With btnHours(iHH)
                        .ControlName = "btnHours"
                        .Left = 25 + iColumn * 100
                        .Top = 50 + iHH * 22
                        .Width = 17
                        .Height = 17
                        .Enabled = True
                        .Visible = False
                        .ForeColor = muSettings.InterfaceBorderColor
                        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                        .DrawBackImage = False
                        .FontFormat = CType(4, DrawTextFormat)
                    End With
                    Me.AddChild(CType(btnHours(iHH), UIControl))
                    iHH += 1
                Next iHour
            Next iColumn

            'btnClear initial props
            btnClear = New UIButton(oUILib)
            With btnClear
                .ControlName = "btnClear"
                .Left = 275
                .Top = (265 \ 2) - 100
                .Width = 100
                .Height = 50
                .Enabled = True
                .Visible = False
                .Caption = "Clear All"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnClear, UIControl))

            'btnSet initial props
            btnSet = New UIButton(oUILib)
            With btnSet
                .ControlName = "btnSet"
                .Left = btnClear.Left
                .Top = btnClear.Top + 55
                .Width = 100
                .Height = 50
                .Enabled = True
                .Visible = False
                .Caption = "Set All"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSet, UIControl))

            'btnReset initial props
            btnReset = New UIButton(oUILib)
            With btnReset
                .ControlName = "btnReset"
                .Left = btnClear.Left
                .Top = btnSet.Top + 55
                .Width = 100
                .Height = 50
                .Enabled = True
                .Visible = False
                .Caption = "Reset"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnReset, UIControl))

            'btnSave initial props
            btnSave = New UIButton(oUILib)
            With btnSave
                .ControlName = "btnSave"
                .Left = btnClear.Left
                .Top = btnReset.Top + 55
                .Width = 100
                .Height = 50
                .Enabled = True
                .Visible = False
                .Caption = "Save"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSave, UIControl))

        End Sub

        Private Sub btnClear_Click(ByVal sName As String) Handles btnClear.Click
            For x As Int32 = 0 To 23
                With moTodaysEntries(x)
                    .iWarable = False
                End With
            Next
            Me.IsDirty = True
        End Sub

        Private Sub btnSet_Click(ByVal sName As String) Handles btnSet.Click
            For x As Int32 = 0 To 23
                With moTodaysEntries(x)
                    If .bHasEntry = True Then
                        .iWarable = True
                    End If
                End With
            Next
            Me.IsDirty = True
        End Sub

        Private Sub btnReset_Click(ByVal sName As String) Handles btnReset.Click
            For x As Int32 = 0 To 23
                With moTodaysEntries(x)
                    .iWarable = .iWarableOrig
                End With
            Next
            Me.IsDirty = True
        End Sub

        Public Sub LoadDay(ByVal iDay As Int32)
            lblTitle.Caption = "Configure Settings for " & GetDayOfWeekName(iDay)
            Dim iHH As Int32 = 0
            Dim iCur As Int32 = -1
            For x As Int32 = 0 To 24
                moTodaysEntries(x) = moCalendarEntries(0)
            Next
            For iEntry As Int32 = 169 To 335
                If iCur <> -1 AndAlso moCalendarEntries(iEntry).iDay <> iCur Then Exit For
                If moCalendarEntries(iEntry).iDay = iDay Then

                    moTodaysEntries(iHH) = moCalendarEntries(iEntry)
                    iHH += 1
                    iCur = iDay
                End If
            Next
            lblTitle.Visible = True
            For x As Int32 = 0 To 23
                btnHours(x).Visible = True
                lblHours(x).Visible = True
            Next
            btnClear.Visible = True
            btnSet.Visible = True
            btnReset.Visible = True
            btnSave.Visible = True
            miDay = iDay
        End Sub

        Private Sub fraEdit_OnRender() Handles Me.OnRender
            If miDay = -1 Then Return
            Dim iDayWidth As Int32 = 22
            Dim iDayHeight As Int32 = 22
            Dim iIdx As Int32 = -1
            Dim clrVal As System.Drawing.Color
            Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                oSprite.Begin(SpriteFlags.AlphaBlend)
                Dim iHH As Int32 = 0
                For iColumn As Int32 = 0 To 1
                    For iHour As Int32 = 0 To 11
                        rcHours(iHH) = New Rectangle(460, 5, 17, 17)
                        With rcHours(iHH)
                            clrVal = Color.Red
                            Dim rcIcon As Rectangle = New Rectangle(127, 143, 17, 17)
                            .X = 75 + 100 * iColumn
                            .Y = 145 + iHour * 20
                            For x As Int32 = 0 To 23
                                If moTodaysEntries(x).iHour = (iHour + 12 * iColumn) Then
                                    iIdx = x
                                    Exit For
                                End If
                            Next

                            If moTodaysEntries(iIdx).bHasEntry = False Then
                                clrVal = Color.Gray
                            ElseIf moTodaysEntries(iIdx).iWarable = False Then
                                clrVal = Color.Green
                            Else
                                clrVal = Color.Red
                            End If
                            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcIcon, rcHours(iIdx), Point.Empty, 0, .Location, clrVal)

                        End With
                        iHH += 1
                    Next
                Next
                oSprite.End()
                oSprite.Dispose()
            End Using
        End Sub

        Private Function TestCalendarHit(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal bSet As Boolean) As Int32
            Dim lX As Int32 = lMouseX - Me.GetAbsolutePosition.X + Me.Left
            Dim lY As Int32 = lMouseY - Me.GetAbsolutePosition.Y + Me.Top
            Debug.Print("(" & lMouseX & "," & lMouseY & ") (" & lX & "," & lY & ")")
            For x As Int32 = 0 To 23
                If rcHours(x).Contains(lX, lY) = True Then
                    For y As Int32 = 0 To 23
                        If moTodaysEntries(y).iHour = x Then
                            If bSet = True Then
                                If moTodaysEntries(y).bHasEntry = True Then
                                    If moTodaysEntries(y).iWarable = True Then
                                        miSetting = 0
                                    Else
                                        miSetting = 1
                                    End If
                                    Debug.Print("Clicked Hour " & y.ToString)
                                End If
                            End If
                            Me.IsDirty = True
                            Return y
                        End If
                    Next
                End If
            Next
            Return -1
        End Function

        Private Sub fraEdit_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
            If lButton <> MouseButtons.Left Then Exit Sub
            Dim iIdx As Int32 = TestCalendarHit(lMouseX, lMouseY, True)
            If iIdx = -1 Then Return
            If moTodaysEntries(iIdx).iWarable = CBool(miSetting) Then
                mbDataHasChanged = True
                moTodaysEntries(iIdx).iWarable = CBool(miSetting)
            End If
            moTodaysEntries(iIdx).iWarable = CBool(miSetting)
            'Me.Moveable = False
        End Sub

        Private Sub fraEdit_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
            If miSetting = -1 Then Return
            If lButton <> MouseButtons.Left Then Exit Sub
            Dim iIdx As Int32 = TestCalendarHit(lMouseX, lMouseY, False)
            If iIdx = -1 Then Return
            If moTodaysEntries(iIdx).bHasEntry = True AndAlso moTodaysEntries(iIdx).iWarable <> CBool(miSetting) Then
                moTodaysEntries(iIdx).iWarable = CBool(miSetting)
                mbDataHasChanged = True
            End If
        End Sub

        Private Sub fraEdit_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseUp
            If lButton <> MouseButtons.Left Then Exit Sub
            'Me.Moveable = True
            miSetting = -1
        End Sub

        Private Sub btnSave_Click(ByVal sName As String) Handles btnSave.Click
            mbDataHasChanged = False
            Dim iHH As Int32 = 0
            For iEntry As Int32 = 169 To 335
                If moCalendarEntries(iEntry).iDay = miDay Then
                    moCalendarEntries(iEntry) = moTodaysEntries(iHH)
                    iHH += 1
                End If
            Next
            Me.IsDirty = True

        End Sub
    End Class
End Class

'End Class

