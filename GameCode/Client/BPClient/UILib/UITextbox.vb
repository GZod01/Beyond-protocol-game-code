Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UITextBox
    Inherits UILabel

	Private mlSelStart As Int32 = 0		 'current position in the string that the selection starts
    Public Property SelStart() As Int32
        Get
            Return mlSelStart
        End Get
        Set(ByVal Value As Int32)
            mlSelStart = Value
            IsDirty = True
        End Set
    End Property
    Private mlSelLength As Int32       'length of the selection
    Public Property SelLength() As Int32
        Get
            Return mlSelLength
        End Get
        Set(ByVal Value As Int32)
            mlSelLength = Value
            IsDirty = True
        End Set
	End Property
	Private mlCursorPos As Int32 = 0		'current position in the string where the cursor is located
	Public Property CursorPos() As Int32
		Get
			Return mlCursorPos
		End Get
		Set(ByVal value As Int32)
			mlCursorPos = value
			IsDirty = True
		End Set
	End Property


	Public BackColorEnabled As System.Drawing.Color = muSettings.InterfaceTextBoxFillColor
	Public BackColorDisabled As System.Drawing.Color = System.Drawing.Color.DimGray

	Public MaxLength As Int32

    'Private moBorderLine As Line
    'Private moCursorLine As Line
	Private moCursorVerts(1) As Vector2

	'For password functionality
	Public PasswordChar As String = ""

	Public Locked As Boolean = False

	Public BorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 255)
	Public BorderLineWidth As Int32 = 1		'was 3

	Private moHighlightColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 255)
	Private moHighlightTextColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
	Public Property HighlightColor() As System.Drawing.Color
		Get
			Return moHighlightColor
		End Get
		Set(ByVal Value As System.Drawing.Color)
			moHighlightColor = Value
			moHighlightTextColor = Color.FromArgb(255, 255 - moHighlightColor.R, 255 - moHighlightColor.G, 255 - moHighlightColor.B)
			IsDirty = True
		End Set
	End Property

    Private mbMultiLine As Boolean = False
	Private mbFormatted As Boolean = False

	Private moScrollBar As UIScrollBar = Nothing

	Public bNumericOnly As Boolean = False

	Public bCanGetFocus As Boolean = True

    Public Enum DoNotRenderSetting As Integer
        eRenderAll = 0
        eBorder = 1
        eFillColor = 2
        eText = 4
        eCaret = 8
        eDoNotRenderControl = 15
    End Enum
	Public DoNotRender As DoNotRenderSetting = DoNotRenderSetting.eRenderAll

	Public Property MultiLine() As Boolean
		Get
			Return mbMultiLine
		End Get
		Set(ByVal value As Boolean)
			mbMultiLine = value
			Me.IsDirty = True
		End Set
	End Property

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		MyBase.mbRaiseTextChangeEvent = True

		'Me.ControlImage_Normal = oUILib.oResMgr.GetTexture("Textbox_Normal.bmp")
		'Me.ControlImage_Disabled = oUILib.oResMgr.GetTexture("Textbox_Disabled.bmp")
		Me.ForeColor = System.Drawing.Color.Black
		Me.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'Device.IsUsingEventHandlers = False
        '      moBorderLine = New Line(oUILib.oDevice)
        '      moCursorLine = New Line(oUILib.oDevice)
        'Device.IsUsingEventHandlers = True
		MyBase.bAcceptEvents = True
	End Sub

	Private Sub UITextBox_OnGotFocus() Handles Me.OnGotFocus
		If Me.Locked = False AndAlso Me.Enabled = True AndAlso Me.MultiLine = False Then
			If Me.Caption Is Nothing = False AndAlso Me.Caption <> "" Then
				Me.SelStart = 0
				CursorPos = 0
				Me.SelLength = Me.Caption.Length
			End If
		End If
	End Sub

#Region "  Bulk Rendering  "
    'This region of code executes the contents of Control_OnRender but is intended for multiple UITextbox controls in succession
    '  This is mainly for the frmBuilderWrksht to increase its rendering performance, however, for expandability, be sure changes
    '  done to the Control_OnRender method are taken into account here as well. This does not handle multiline textboxes
    Private msBulkRenderedText As String = ""
    Private moBulkRenderedTextRect As Rectangle
    Private mlBulkLineStart As Int32
    Private mlBulkTrimedLen As Int32 = 0
    Public Sub PrepareForBulkRender()
        If PasswordChar <> "" Then
            msBulkRenderedText = StrDup(Caption.Length, PasswordChar)
        Else
            msBulkRenderedText = Caption
        End If
        If msBulkRenderedText Is Nothing Then Return
        If SelStart > msBulkRenderedText.Length Then SelStart = msBulkRenderedText.Length
        If CursorPos > msBulkRenderedText.Length Then CursorPos = msBulkRenderedText.Length


        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()
        With moBulkRenderedTextRect
            .Location() = oLoc
            .Width = Width
            .Height = Height

            'Draw the caption, but we need to scoot the rect over a bit
            .X += 5
            .Width -= 5
        End With

        mlBulkLineStart = 0

        'Here, we measure our stuff and stuff
        '  get all of the calculation stuff done
        If MultiLine = True Then
            msBulkRenderedText = FormatMultiLineString(Caption, mlBulkLineStart)
        Else
            'Ok, not multiline... see if our text scrolls off the side
            Dim rcTemp As Rectangle = BPFont.MeasureString(moSysFont, msBulkRenderedText, FontFormat)  'BPFont.CFFont.GetTextDimensions(sRenderedText, moSysFont.Size) 'Dim rcTemp As Rectangle = moFont.MeasureString(Nothing, sRenderedText, FontFormat, ForeColor)

            mlBulkTrimedLen = 0
            While rcTemp.Width > moBulkRenderedTextRect.Width
                If (FontFormat And DrawTextFormat.Right) <> 0 Then
                    'ok, right aligned so we want the left side of the string
                    If msBulkRenderedText.Length > 0 Then
                        msBulkRenderedText = msBulkRenderedText.Substring(0, msBulkRenderedText.Length - 2)
                        mlBulkTrimedLen += 1
                    Else : Exit While
                    End If
                Else
                    'ok, left aligned or Center aligned, so we want the right side of the string
                    If msBulkRenderedText.Length > 0 Then
                        msBulkRenderedText = msBulkRenderedText.Substring(1)
                        mlBulkTrimedLen += 1
                    Else : Exit While
                    End If
                End If
                rcTemp = BPFont.MeasureString(moSysFont, msBulkRenderedText, FontFormat) 'BPFont.CFFont.GetTextDimensions(sRenderedText, moSysFont.Size) 'rcTemp = moFont.MeasureString(Nothing, sRenderedText, FontFormat, ForeColor)
            End While


        End If
    End Sub
    Public Sub BulkBackColorFill(ByRef oSprite As Sprite)
        'render the alpha blended color fill if it is to be rendered

        If (DoNotRender And DoNotRenderSetting.eFillColor) <> 0 Then Return

        Dim lFillColor As System.Drawing.Color
        If Enabled = True Then
            lFillColor = BackColorEnabled
        Else
            lFillColor = BackColorDisabled
        End If

        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()
        Dim oRect As Rectangle
        With oRect
            .Location() = oLoc
            .Width = Width
            .Height = Height

            If (DoNotRender And DoNotRenderSetting.eFillColor) = 0 Then DoMultiColorFill(oRect, lFillColor, oLoc, oSprite)
        End With

        'render the alpha blended color fill of the selected text
        'now, render the selection stuff...
        If Me.SelLength > 0 AndAlso Me.HasFocus = True AndAlso MyBase.moUILib.FocusedControl Is Nothing = False AndAlso MyBase.moUILib.FocusedControl.ControlName = Me.ControlName Then
            'ok, determine what is being rendered
            If (FontFormat And DrawTextFormat.Right) <> 0 Then
                'right aligned
                If SelStart + SelLength > mlBulkTrimedLen Then
                    'ok, find our starting spot
                    Dim lActStart As Int32 = SelStart
                    If SelStart < mlBulkTrimedLen Then
                        lActStart = mlBulkTrimedLen
                    End If
                    Dim lLeft As Int32 = BPFont.MeasureString(moSysFont, msBulkRenderedText.Substring(lActStart - mlBulkTrimedLen), FontFormat).Width 'Dim lLeft As Int32 = moFont.MeasureString(Nothing, sRenderedText.Substring(lActStart - lTrimedLen), FontFormat, ForeColor).Width
                    Dim lWidth As Int32 = BPFont.MeasureString(moSysFont, msBulkRenderedText.Substring(lActStart - mlBulkTrimedLen, SelLength), FontFormat).Width 'Dim lWidth As Int32 = moFont.MeasureString(Nothing, sRenderedText.Substring(lActStart - lTrimedLen, SelLength), FontFormat, ForeColor).Width

                    If lLeft + lWidth > oRect.Width Then
                        lWidth = oRect.Width - lLeft
                    End If

                    DoMultiColorFill(New Rectangle(lLeft, oRect.Top, lWidth, oRect.Height), System.Drawing.Color.FromArgb(255, lFillColor.R, lFillColor.G, lFillColor.B), New Point(lLeft, oRect.Top), oSprite)
                End If
            Else
                'left aligned
                If SelStart + SelLength > mlBulkTrimedLen Then
                    'ok, find our starting spot
                    Dim lActStart As Int32 = SelStart
                    If SelStart < mlBulkTrimedLen Then
                        lActStart = mlBulkTrimedLen
                    End If
                    Dim lActLength As Int32 = SelLength
                    If (lActStart - mlBulkTrimedLen) + lActLength > msBulkRenderedText.Length Then
                        lActLength = msBulkRenderedText.Length - (lActStart - mlBulkTrimedLen)
                    End If
                    Dim lLeft As Int32 = BPFont.MeasureString(moSysFont, msBulkRenderedText.Substring(0, lActStart - mlBulkTrimedLen), FontFormat).Width 'Dim lLeft As Int32 = moFont.MeasureString(Nothing, sRenderedText.Substring(0, lActStart - lTrimedLen), FontFormat, ForeColor).Width
                    Dim lWidth As Int32 = BPFont.MeasureString(moSysFont, msBulkRenderedText.Substring(lActStart - mlBulkTrimedLen, lActLength), FontFormat).Width 'Dim lWidth As Int32 = moFont.MeasureString(Nothing, sRenderedText.Substring(lActStart - lTrimedLen, lActLength), FontFormat, ForeColor).Width

                    If lLeft + lWidth > oRect.Width Then
                        lWidth = oRect.Width - lLeft
                    End If

                    If lActStart <> 0 Then lLeft += oLoc.X + 5 Else lLeft += oLoc.X

                    DoMultiColorFill(New Rectangle(lLeft, oLoc.Y + 1, lWidth, oRect.Height - 2), System.Drawing.Color.FromArgb(255, Math.Min(255, lFillColor.R * 2), Math.Min(255, lFillColor.G * 2), Math.Min(255, lFillColor.B * 2)), New Point(lLeft, oLoc.Y + 1), oSprite)
                End If
            End If
        End If
    End Sub
    Public Sub BulkRenderBorder(ByRef oBLine As Line)
        'render the border using oBLine
        RenderBorder(0, 0, oBLine)
    End Sub
    Public Sub BulkRenderText(ByRef oFont As Font)
        'Bulk render the text in the textbox
        If (DoNotRender And DoNotRenderSetting.eText) = 0 Then oFont.DrawText(Nothing, msBulkRenderedText, moBulkRenderedTextRect, FontFormat, ForeColor)

        'If this textbox is active and has focus, render the cursor line?
        If (DoNotRender And DoNotRenderSetting.eCaret) = 0 AndAlso HasFocus = True AndAlso Locked = False Then
            Dim bSkipCursorRender As Boolean = False    'used when the cursor is not visible
            Dim lTempValue As Int32 = CursorPos - mlBulkLineStart
            Dim sTemp As String = ""
            If lTempValue > 0 Then
                sTemp = Trim$(Mid$(msBulkRenderedText, 1, CursorPos - mlBulkLineStart)) ' SelStart))
            End If
            'oRect = moFont.MeasureString(Nothing, Mid$(sRenderedText, 1, CursorPos), FontFormat, ForeColor)	'SelStart), FontFormat, ForeColor)
            Dim oRect As Rectangle = oFont.MeasureString(Nothing, Mid$(msBulkRenderedText, 1, CursorPos), FontFormat, ForeColor)  'SelStart), FontFormat, ForeColor)
            If sTemp.Length < CursorPos Then ' SelStart Then
                oRect.Width += (5 * (CursorPos - sTemp.Length)) ' SelStart - sTemp.Length))
            End If

            Dim oLoc As System.Drawing.Point = GetAbsolutePosition()

            If MultiLine = False Then
                If (FontFormat And DrawTextFormat.Right) <> 0 Then
                    'ok, now, draw a line there
                    If Caption = "" Then
                        moCursorVerts(0).X = oLoc.X + Me.Width - 5
                        moCursorVerts(0).Y = oLoc.Y + 2
                        moCursorVerts(1).X = moCursorVerts(0).X
                        moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
                    Else
                        moCursorVerts(0).X = (oLoc.X + Me.Width) - (oRect.Width + 5)
                        moCursorVerts(0).Y = oLoc.Y + 2
                        moCursorVerts(1).X = moCursorVerts(0).X
                        moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
                    End If
                Else
                    'ok, now, draw a line there
                    If Caption = "" OrElse sTemp = "" Then
                        moCursorVerts(0).X = oLoc.X + 5
                        moCursorVerts(0).Y = oLoc.Y + 2
                        moCursorVerts(1).X = moCursorVerts(0).X
                        moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
                    Else
                        moCursorVerts(0).X = oRect.Left + oRect.Width + oLoc.X + 5
                        moCursorVerts(0).Y = oLoc.Y + 2
                        moCursorVerts(1).X = moCursorVerts(0).X
                        moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
                    End If
                End If
            Else
                If msBulkRenderedText = "" Then
                    moCursorVerts(0).X = oLoc.X + Me.Width - 5
                    moCursorVerts(0).Y = oLoc.Y + 2
                    moCursorVerts(1).X = moCursorVerts(0).X
                    moCursorVerts(1).Y = oLoc.Y + 18 - 2
                Else
                    Dim ptPos As Point = GetCursorPos()
                    moCursorVerts(0).X = ptPos.X
                    moCursorVerts(0).Y = ptPos.Y + 2
                    moCursorVerts(1).X = moCursorVerts(0).X
                    moCursorVerts(1).Y = ptPos.Y + 16

                    Dim lMeTop As Int32 = Me.GetRelativeTop()
                    If moCursorVerts(0).Y > lMeTop + Me.Height Then
                        bSkipCursorRender = True
                    ElseIf moCursorVerts(1).Y > lMeTop + Me.Height Then
                        moCursorVerts(1).Y = lMeTop + Me.Height
                    ElseIf moCursorVerts(1).Y < lMeTop Then
                        bSkipCursorRender = True
                    ElseIf moCursorVerts(0).Y < lMeTop Then
                        moCursorVerts(0).Y = lMeTop
                    End If
                End If
            End If

            If moCursorVerts Is Nothing = False AndAlso bSkipCursorRender = False Then
                Using moCursorLine As New Line(MyBase.moUILib.oDevice)
                    moCursorLine.Begin()
                    moCursorLine.Draw(moCursorVerts, ForeColor)
                    moCursorLine.End()
                End Using
            End If
        End If
    End Sub
    Private Sub DoMultiColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As System.Drawing.Point, ByRef oSprite As Sprite)
        Dim rcSrc As Rectangle

        Dim fX As Single
        Dim fY As Single

        If rcDest.Width = 0 OrElse rcDest.Height = 0 OrElse moUILib.oInterfaceTexture Is Nothing Then Return

        'rcSrc = System.Drawing.Rectangle.FromLTRB(225, 0, 255, 30)
        rcSrc.Location = New Point(192, 0)
        rcSrc.Width = 62
        rcSrc.Height = 64

        'Now, draw it...
        With oSprite
            fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
            fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))
            .Draw2D(moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), System.Drawing.Color.FromArgb(255, 0, 0, 0))
            .Draw2D(moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFill)
        End With
    End Sub
#End Region


	'Public Overrides Sub Control_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles MyBase.OnRender
    Public Overrides Sub Control_OnRender() 'Handles MyBase.OnRender
        If DoNotRender = DoNotRenderSetting.eDoNotRenderControl Then Return

        Dim oRect As Rectangle
        Dim oLoc As System.Drawing.Point
        Dim lFillColor As System.Drawing.Color

        Dim sRenderedText As String

        If Enabled Then
            lFillColor = BackColorEnabled
        Else
            lFillColor = BackColorDisabled
        End If

        If PasswordChar <> "" Then
            sRenderedText = StrDup(Caption.Length, PasswordChar)
        Else
            sRenderedText = Caption
        End If
        If sRenderedText Is Nothing Then Return
        If SelStart > sRenderedText.Length Then SelStart = sRenderedText.Length
        If CursorPos > sRenderedText.Length Then CursorPos = sRenderedText.Length

        If moScrollBar Is Nothing = False Then
            If moScrollBar.Height <> Me.Height Then moScrollBar.Height = Me.Height
        End If

        oLoc = GetAbsolutePosition()
        With oRect
            .Location() = oLoc
            .Width = Width
            .Height = Height

            'Do a color fill 
            'moUILib.DoAlphaBlendColorFill_EX(oRect, lFillColor, oLoc, oImgSprite)
            If (DoNotRender And DoNotRenderSetting.eFillColor) = 0 Then moUILib.DoAlphaBlendColorFill(oRect, lFillColor, oLoc)
            'end of color fill

            ''Draw a box border around our window...
            'With moBorderLineVerts(0)
            '    .X = oLoc.X
            '    .Y = oLoc.Y
            'End With
            'With moBorderLineVerts(1)
            '    .X = oLoc.X + Width
            '    .Y = oLoc.Y
            'End With
            'With moBorderLineVerts(2)
            '    .X = oLoc.X + Width
            '    .Y = oLoc.Y + Height
            'End With
            'With moBorderLineVerts(3)
            '    .X = oLoc.X
            '    .Y = oLoc.Y + Height
            'End With
            'With moBorderLineVerts(4)
            '    .X = oLoc.X
            '    .Y = oLoc.Y
            'End With

            'moBorderLine.Antialias = True
            'moBorderLine.Width = BorderLineWidth
            'moBorderLine.Begin()
            'moBorderLine.Draw(moBorderLineVerts, BorderColor)
            'moBorderLine.End()
            ''End of the border drawing
            If (DoNotRender And DoNotRenderSetting.eBorder) = 0 Then RenderBorder(0, 0, Nothing)

            'Draw the caption, but we need to scoot the rect over a bit
            oRect.X += 5
            oRect.Width -= 5
            'Measure what a W would take
            'Dim rcMsr As Rectangle = moFont.MeasureString(Nothing, "W", FontFormat, ForeColor)
            'Dim lMaxLen As Int32 = oRect.Width \ rcMsr.Width
            'moFont.DrawText(Nothing, Mid$(sRenderedText, Math.Max(1, sRenderedText.Length - lMaxLen + 1), lMaxLen), oRect, FontFormat, ForeColor)

            'If MultiLine = True AndAlso mbFormatted = False Then sRenderedText = FormatMultiLineString(sRenderedText)
            Dim lLineStart As Int32 = 0
            If MultiLine = True Then
                sRenderedText = FormatMultiLineString(Caption, lLineStart)
                'TODO: Needs testing, adjustment.
                'TODO:  Without this EnsureCursorVisible when CR'ing on the last line when no scroller present, the scroller doens't appear
                'TODO:  untill the next render.  Once the scoller is present the EnsureCursorVisible needs to be fired.
                'EnsureCursorVisible
            Else
                'Ok, not multiline... see if our text scrolls off the side
                Dim rcTemp As Rectangle = BPFont.MeasureString(moSysFont, sRenderedText, FontFormat)  'BPFont.CFFont.GetTextDimensions(sRenderedText, moSysFont.Size) 'Dim rcTemp As Rectangle = moFont.MeasureString(Nothing, sRenderedText, FontFormat, ForeColor)

                Dim lTrimedLen As Int32 = 0
                While rcTemp.Width > oRect.Width
                    If (FontFormat And DrawTextFormat.Right) <> 0 Then
                        'ok, right aligned so we want the left side of the string
                        If sRenderedText.Length > 0 Then
                            sRenderedText = sRenderedText.Substring(0, sRenderedText.Length - 2)
                            lTrimedLen += 1
                        Else : Exit While
                        End If
                    Else
                        'ok, left aligned or Center aligned, so we want the right side of the string
                        If sRenderedText.Length > 0 Then
                            sRenderedText = sRenderedText.Substring(1)
                            lTrimedLen += 1
                        Else : Exit While
                        End If
                    End If
                    rcTemp = BPFont.MeasureString(moSysFont, sRenderedText, FontFormat) 'BPFont.CFFont.GetTextDimensions(sRenderedText, moSysFont.Size) 'rcTemp = moFont.MeasureString(Nothing, sRenderedText, FontFormat, ForeColor)
                End While

                'now, render the selection stuff...
                If Me.SelLength > 0 AndAlso Me.HasFocus = True AndAlso MyBase.moUILib.FocusedControl Is Nothing = False AndAlso MyBase.moUILib.FocusedControl.ControlName = Me.ControlName Then
                    'ok, determine what is being rendered
                    If (FontFormat And DrawTextFormat.Right) <> 0 Then
                        'right aligned
                        If SelStart + SelLength > lTrimedLen Then
                            'ok, find our starting spot
                            Dim lActStart As Int32 = SelStart
                            If SelStart < lTrimedLen Then
                                lActStart = lTrimedLen
                            End If
                            Dim lLeft As Int32 = BPFont.MeasureString(moSysFont, sRenderedText.Substring(lActStart - lTrimedLen), FontFormat).Width 'Dim lLeft As Int32 = moFont.MeasureString(Nothing, sRenderedText.Substring(lActStart - lTrimedLen), FontFormat, ForeColor).Width
                            Dim lWidth As Int32 = BPFont.MeasureString(moSysFont, sRenderedText.Substring(lActStart - lTrimedLen, SelLength), FontFormat).Width 'Dim lWidth As Int32 = moFont.MeasureString(Nothing, sRenderedText.Substring(lActStart - lTrimedLen, SelLength), FontFormat, ForeColor).Width

                            If lLeft + lWidth > oRect.Width Then
                                lWidth = oRect.Width - lLeft
                            End If

                            MyBase.moUILib.DoAlphaBlendColorFill(New Rectangle(lLeft, oRect.Top, lWidth, oRect.Height), System.Drawing.Color.FromArgb(255, lFillColor.R, lFillColor.G, lFillColor.B), New Point(lLeft, oRect.Top))
                        End If
                    Else
                        'left aligned
                        If SelStart + SelLength > lTrimedLen Then
                            'ok, find our starting spot
                            Dim lActStart As Int32 = SelStart
                            If SelStart < lTrimedLen Then
                                lActStart = lTrimedLen
                            End If
                            Dim lActLength As Int32 = SelLength
                            If (lActStart - lTrimedLen) + lActLength > sRenderedText.Length Then
                                lActLength = sRenderedText.Length - (lActStart - lTrimedLen)
                            End If
                            Dim lLeft As Int32 = BPFont.MeasureString(moSysFont, sRenderedText.Substring(0, lActStart - lTrimedLen), FontFormat).Width 'Dim lLeft As Int32 = moFont.MeasureString(Nothing, sRenderedText.Substring(0, lActStart - lTrimedLen), FontFormat, ForeColor).Width
                            Dim lWidth As Int32 = BPFont.MeasureString(moSysFont, sRenderedText.Substring(lActStart - lTrimedLen, lActLength), FontFormat).Width 'Dim lWidth As Int32 = moFont.MeasureString(Nothing, sRenderedText.Substring(lActStart - lTrimedLen, lActLength), FontFormat, ForeColor).Width

                            If lLeft + lWidth > oRect.Width Then
                                lWidth = oRect.Width - lLeft
                            End If

                            If lActStart <> 0 Then lLeft += oLoc.X + 5 Else lLeft += oLoc.X

                            MyBase.moUILib.DoAlphaBlendColorFill(New Rectangle(lLeft, oLoc.Y + 1, lWidth, oRect.Height - 2), System.Drawing.Color.FromArgb(255, Math.Min(255, lFillColor.R * 2), Math.Min(255, lFillColor.G * 2), Math.Min(255, lFillColor.B * 2)), New Point(lLeft, oLoc.Y + 1))
                        End If
                    End If
                End If
            End If
            'If moUILib.bTextPenBegun = False Then moUILib.TextPen.Begin(SpriteFlags.AlphaBlend)
            If (DoNotRender And DoNotRenderSetting.eText) = 0 Then BPFont.DrawText(moSysFont, sRenderedText, oRect, FontFormat, ForeColor) 'moFont.DrawText(Nothing, sRenderedText, oRect, FontFormat, ForeColor)
            'If moUILib.bTextPenBegun = False Then moUILib.TextPen.End()

            If (DoNotRender And DoNotRenderSetting.eCaret) = 0 AndAlso HasFocus = True AndAlso Locked = False Then
                Dim bSkipCursorRender As Boolean = False    'used when the cursor is not visible
                Dim lTempValue As Int32 = CursorPos - lLineStart
                Dim sTemp As String = ""
                If lTempValue > 0 Then
                    sTemp = Trim$(Mid$(sRenderedText, 1, CursorPos - lLineStart)) ' SelStart))
                End If
                'oRect = moFont.MeasureString(Nothing, Mid$(sRenderedText, 1, CursorPos), FontFormat, ForeColor)	'SelStart), FontFormat, ForeColor)
                oRect = BPFont.MeasureString(moSysFont, Mid$(sRenderedText, 1, CursorPos), FontFormat)  'SelStart), FontFormat, ForeColor)
                If sTemp.Length < CursorPos Then ' SelStart Then
                    oRect.Width += (5 * (CursorPos - sTemp.Length)) ' SelStart - sTemp.Length))
                End If

                If MultiLine = False Then
                    If (FontFormat And DrawTextFormat.Right) <> 0 Then
                        'ok, now, draw a line there
                        If Caption = "" Then
                            moCursorVerts(0).X = oLoc.X + Me.Width - 5
                            moCursorVerts(0).Y = oLoc.Y + 2
                            moCursorVerts(1).X = moCursorVerts(0).X
                            moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
                        Else
                            moCursorVerts(0).X = (oLoc.X + Me.Width) - (oRect.Width + 5)
                            moCursorVerts(0).Y = oLoc.Y + 2
                            moCursorVerts(1).X = moCursorVerts(0).X
                            moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
                        End If
                    Else
                        'ok, now, draw a line there
                        If Caption = "" OrElse sTemp = "" Then
                            moCursorVerts(0).X = oLoc.X + 5
                            moCursorVerts(0).Y = oLoc.Y + 2
                            moCursorVerts(1).X = moCursorVerts(0).X
                            moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
                        Else
                            moCursorVerts(0).X = oRect.Left + oRect.Width + oLoc.X + 5
                            moCursorVerts(0).Y = oLoc.Y + 2
                            moCursorVerts(1).X = moCursorVerts(0).X
                            moCursorVerts(1).Y = oLoc.Y + Me.Height - 2
                        End If
                    End If
                Else
                    If sRenderedText = "" Then
                        moCursorVerts(0).X = oLoc.X + Me.Width - 5
                        moCursorVerts(0).Y = oLoc.Y + 2
                        moCursorVerts(1).X = moCursorVerts(0).X
                        moCursorVerts(1).Y = oLoc.Y + 18 - 2
                    Else
                        Dim ptPos As Point = GetCursorPos()
                        moCursorVerts(0).X = ptPos.X
                        moCursorVerts(0).Y = ptPos.Y + 2
                        moCursorVerts(1).X = moCursorVerts(0).X
                        moCursorVerts(1).Y = ptPos.Y + 16

                        Dim lMeTop As Int32 = Me.GetRelativeTop()
                        If moCursorVerts(0).Y > lMeTop + Me.Height Then
                            bSkipCursorRender = True
                        ElseIf moCursorVerts(1).Y > lMeTop + Me.Height Then
                            moCursorVerts(1).Y = lMeTop + Me.Height
                        ElseIf moCursorVerts(1).Y < lMeTop Then
                            bSkipCursorRender = True
                        ElseIf moCursorVerts(0).Y < lMeTop Then
                            moCursorVerts(0).Y = lMeTop
                        End If
                    End If
                End If

                If moCursorVerts Is Nothing = False AndAlso bSkipCursorRender = False Then
                    Using moCursorLine As New Line(MyBase.moUILib.oDevice)
                        moCursorLine.Begin()
                        moCursorLine.Draw(moCursorVerts, ForeColor)
                        moCursorLine.End()
                    End Using
                End If
            End If

        End With

    End Sub

	Private Sub UITextBox_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
		If bCanGetFocus = False Then Return

		If moUILib.FocusedControl Is Nothing = False Then
			moUILib.FocusedControl.HasFocus = False
		End If
		HasFocus = True
		moUILib.FocusedControl = Me
	End Sub

	Private Sub UITextBox_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.OnKeyPress
		Dim sTemp As String

		If Char.IsControl(e.KeyChar) = False AndAlso Locked = False Then
            If MaxLength = 0 OrElse Caption.Length + 1 - SelLength <= MaxLength Then
                If bNumericOnly = True Then
                    Dim lIdx As Int32 = "-1234567890.".IndexOf(e.KeyChar)
                    If lIdx = -1 Then Return
                End If

                sTemp = Mid$(Caption, 1, SelStart) & e.KeyChar & Mid$(Caption, SelStart + Math.Max(SelLength + 1, 1))

                If NewTutorialManager.TutorialOn = True Then
                    Dim sParmList(0) As String
                    sParmList(0) = sTemp
                    If MyBase.moUILib.CommandAllowedWithParms(True, GetFullName(), sParmList, True) = False Then
                        Return
                    End If
                End If

                SelStart += 1
                CursorPos = SelStart
                SelLength = 0
                Caption = sTemp
                mbFormatted = False
            Else
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            End If
			'Else
			'If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, EpicaSound.SoundUsage.eUsedButNot3D, Vector3.Empty, Vector3.Empty)
		End If

	End Sub

	Private Sub UITextBox_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.OnKeyDown
		If Locked = True Then Return

		If NewTutorialManager.TutorialOn = True Then
			If MyBase.moUILib.CommandAllowed(True, GetFullName) = False Then Return
		End If

		If e.KeyCode = Keys.Back Then
			If CursorPos > 0 Then ' SelStart > 0 Then
				If Mid$(Caption, 1, CursorPos).EndsWith(vbCrLf) = True Then	'If Mid$(Caption, 1, SelStart).EndsWith(vbCrLf) = True Then
					Caption = Mid$(Caption, 1, CursorPos - 2) & Mid$(Caption, CursorPos + 1) 'Caption = Mid$(Caption, 1, SelStart - 2) & Mid$(Caption, SelStart + 1)
					CursorPos -= 2
					SelStart = CursorPos
				Else
					If SelLength > 0 Then
						Caption = Mid$(Caption, 1, SelStart) & Mid$(Caption, SelStart + Math.Max(1, SelLength + 1))
					Else
						Caption = Mid$(Caption, 1, CursorPos - 1) & Mid$(Caption, CursorPos + 1)
					End If

					If SelLength < 1 Then
						CursorPos -= 1
					Else : CursorPos = SelStart
					End If
					SelStart = CursorPos
				End If
			ElseIf SelLength > 0 Then
				Caption = Mid$(Caption, 1, SelStart) & Mid$(Caption, SelStart + Math.Max(1, SelLength + 1))
				SelStart = CursorPos
			End If
			SelLength = 0
		ElseIf e.KeyCode = Keys.Delete Then
			If Mid$(Caption, 1, CursorPos).StartsWith(vbCrLf) = True Then	'If Mid$(Caption, 1, SelStart).StartsWith(vbCrLf) = True Then
				'Caption = Mid$(Caption, 1, SelStart) & Mid$(Caption, SelStart + 3)
				Caption = Mid$(Caption, 1, CursorPos) & Mid$(Caption, CursorPos + 3)
				SelStart = CursorPos
			Else
				If SelLength > 0 Then
					Caption = Mid$(Caption, 1, SelStart) & Mid$(Caption, SelStart + SelLength + 1)
				Else
					Caption = Mid$(Caption, 1, CursorPos) & Mid$(Caption, CursorPos + 2)
				End If
				If SelLength > 0 Then
					CursorPos = SelStart
				End If
				SelStart = CursorPos
			End If
			SelLength = 0
		ElseIf e.KeyCode = Keys.Right Then
			If e.Shift = True AndAlso MultiLine = False Then
				If SelStart + SelLength < Caption.Length Then
					CursorPos += 1
					If CursorPos > SelStart + SelLength Then
						SelLength += 1
					ElseIf CursorPos > SelStart Then
						SelStart += 1
						SelLength -= 1
					End If
				End If
			Else
				If CursorPos < Caption.Length Then
                    If Mid$(Caption, CursorPos + 1).StartsWith(vbCrLf) = True Then CursorPos += 2 Else CursorPos += 1
				End If
				SelStart = CursorPos
				SelLength = 0
			End If
		ElseIf e.KeyCode = Keys.Left Then
			If CursorPos > 0 Then
				If e.Shift = True AndAlso MultiLine = False Then
					CursorPos -= 1
					If CursorPos < SelStart Then
						SelStart = CursorPos
						SelLength += 1
					ElseIf CursorPos < SelStart + SelLength Then
						SelLength -= 1
					End If
				Else
					If Mid$(Caption, 1, CursorPos + 1).EndsWith(vbCrLf) = True Then CursorPos -= 2 Else CursorPos -= 1
					SelStart = CursorPos
					SelLength = 0
				End If
			End If
		ElseIf e.KeyCode = Keys.Home Then
			Dim lOrigSelStart As Int32 = CursorPos
            If MultiLine = True AndAlso Me.Locked = False AndAlso e.Control = False Then
                MoveToStartOfLine()
            Else
                CursorPos = 0
                SelStart = 0
                If moScrollBar Is Nothing = False Then moScrollBar.Value = 0
            End If
			If e.Shift = True AndAlso MultiLine = False Then
				SelLength = Math.Abs(CursorPos - lOrigSelStart)
			Else : SelLength = 0
			End If
			SelStart = CursorPos
		ElseIf e.KeyCode = Keys.End Then
			Dim lOrigSelStart As Int32 = CursorPos
            If MultiLine = True AndAlso Me.Locked = False AndAlso e.Control = False Then
                MoveToEndOfLine()
            Else
                CursorPos = Caption.Length
                SelStart = CursorPos
                EnsureCursorVisible()
            End If
			If e.Shift = True AndAlso MultiLine = False Then
                SelLength = Math.Abs(CursorPos - lOrigSelStart)
                SelStart = lOrigSelStart
            Else
                SelLength = 0
                SelStart = CursorPos
            End If
		ElseIf e.KeyCode = Keys.Enter AndAlso Locked = False AndAlso MultiLine = True Then
			Dim sTemp As String = Mid$(Caption, 1, SelStart) & vbCrLf & Mid$(Caption, SelStart + SelLength + 1)
			SelStart += 2
			CursorPos = SelStart
			Caption = sTemp
			SelLength = 1
        ElseIf e.KeyCode = Keys.Up Then
            If Me.MultiLine = True OrElse bNumericOnly = False Then
                'Ok, move up a line
                MoveUpALine()
                SelLength = 0
            ElseIf e.Shift = True Then
                NumericCount(100)
                SelStart = 0
                SelLength = Caption.Length
            Else
                NumericCount(1)
                SelStart = 0
                SelLength = Caption.Length
            End If
        ElseIf e.KeyCode = Keys.Down Then
            If Me.MultiLine = True OrElse bNumericOnly = False Then
                MoveDownALine()
                SelLength = 0
            ElseIf e.Shift = True Then
                NumericCount(-100)
                SelStart = 0
                SelLength = Caption.Length
            Else
                NumericCount(-1)
                SelStart = 0
                SelLength = Caption.Length
            End If
		ElseIf e.KeyCode = Keys.PageDown AndAlso Me.MultiLine = True Then
			MoveDownALine()
			MoveDownALine()
			MoveDownALine()
			SelLength = 0
		ElseIf e.KeyCode = Keys.PageUp AndAlso Me.MultiLine = True Then
			MoveUpALine()
			MoveUpALine()
			MoveUpALine()
			SelLength = 0
        End If
        If Me.MultiLine = True Then
            Caption = Caption.Replace(Chr(13) & Chr(13), Chr(13))
            Caption = Caption.Replace(Chr(10) & Chr(10), Chr(10))
        End If
        mbFormatted = False
        EnsureCursorVisible()
    End Sub

    Private Sub NumericCount(ByVal blAdd As Int64)
        If Caption = "" AndAlso blAdd > 0 Then
            Caption = blAdd.ToString
        ElseIf Caption <> "" Then
            Dim blVal As Int64 = CLng(Caption) + blAdd
            If blVal < 0 Then Caption = "" Else Caption = blVal.ToString
        End If
    End Sub

	Private Sub MoveToEndOfLine()
        Try
            Dim sRenderedText As String = JustDoFormattingNoSizeConstraints(Caption)
            'sRenderedText now contains the text exactly as it is rendered to the screen
            Dim lAdjustCursorPos As Int32 = GetAdjustedCursorLoc(sRenderedText)
            'lAdjustCursorPos is the cursor position within the adjusted rendered text
            Dim lDiff As Int32 = lAdjustCursorPos - CursorPos
            'lDiff is the difference for adjustment's sake
            Dim lNextVBCRLF As Int32 = sRenderedText.IndexOf(vbCrLf, lAdjustCursorPos)
            'lNextVBCRLF is the position of the next VBCRLF in the list
            lNextVBCRLF -= lDiff


            If lNextVBCRLF > 0 Then
                CursorPos = lNextVBCRLF
            Else : CursorPos = Caption.Length
            End If
            SelStart = CursorPos
            SelLength = 0
            EnsureCursorVisible()
        Catch ex As Exception
            'do nothing for now
        End Try
	End Sub

	Private Sub MoveToStartOfLine()
        Try
            If MultiLine = True Then
                Dim sRenderedText As String = JustDoFormattingNoSizeConstraints(Caption)
                'sRenderedText now contains the text exactly as it is rendered to the screen
                Dim lAdjustCursorPos As Int32 = GetAdjustedCursorLoc(sRenderedText)
                'lAdjustCursorPos is the cursor position within the adjusted rendered text
                Dim lDiff As Int32 = lAdjustCursorPos - CursorPos
                'lDiff is the difference for adjustment's sake
                Dim lPrevVBCRLF As Int32 = sRenderedText.LastIndexOf(vbCrLf, lAdjustCursorPos)
                'lPrevVBCRLF is the position of the next VBCRLF in the list
                lPrevVBCRLF -= lDiff

                If lPrevVBCRLF > 0 Then
                    'Now, adjust lPrevVBCRLF by 2 because of vbcrlf
                    CursorPos = lPrevVBCRLF + 2
                Else : CursorPos = 0
                End If
                SelStart = CursorPos
                SelLength = 0
                EnsureCursorVisible()

            Else
                CursorPos = 0
            End If
            SelStart = CursorPos
            EnsureCursorVisible()
        Catch ex As Exception
            'do nothing for now
        End Try
	End Sub

    Private Sub MoveUpALine()
        Try
            If Me.MultiLine = True Then
                'Ok... get our formatted string
                Dim sFormatted As String = JustDoFormattingNoSizeConstraints(Caption)
                'Ok, get the adjusted caret pos
                Dim lAdjCaret As Int32 = GetAdjustedCursorLoc(sFormatted)
                'Ok, now, determine the previous VBCRLF
                Dim lPrevEndLine As Int32 = sFormatted.LastIndexOf(vbCrLf, lAdjCaret)
                'Ok, get the width of the current line
                Dim lCurrLineWidth As Int32 = 0
                If lPrevEndLine > 0 Then
                    lCurrLineWidth = BPFont.MeasureString(moSysFont, sFormatted.Substring(lPrevEndLine, lAdjCaret - lPrevEndLine), FontFormat).Width
                Else
                    CursorPos = 0
                    Return
                End If

                'Get the diff from adjust cursor to normal cursor
                Dim lDiff As Int32 = lAdjCaret - CursorPos
                'Now, get the previous to that one
                lPrevEndLine = sFormatted.LastIndexOf(vbCrLf, lPrevEndLine - 2)
                'Now, get the length of the new line
                Dim sNewLine As String = sFormatted.Substring(lPrevEndLine + 2)
                Dim lEndOfNewLine As Int32 = sFormatted.IndexOf(vbCrLf, lPrevEndLine + 1) - lPrevEndLine
                Dim lClosestLen As Int32 = 100000
                Dim lClosestWidth As Int32 = 100000
                Try
                    For X As Int32 = 1 To lEndOfNewLine
                        Dim lTmpWidth As Int32 = BPFont.MeasureString(moSysFont, sNewLine.Substring(0, X), FontFormat).Width
                        If Math.Abs(lTmpWidth - lCurrLineWidth) < lClosestWidth Then
                            lClosestLen = X
                            lClosestWidth = Math.Abs(lTmpWidth - lCurrLineWidth)
                        End If
                    Next X
                Catch
                End Try
                CursorPos = (lPrevEndLine + lClosestLen + 3) - lDiff        '+2 for the vbcrlf and +1 for the character to move past
            Else : MoveToStartOfLine()
            End If
            SelStart = CursorPos
            EnsureCursorVisible()
        Catch
        End Try
    End Sub

	Private Sub MoveDownALine()
        Try
            If Me.MultiLine = True Then
                'Ok... get our formatted string
                Dim sFormatted As String = JustDoFormattingNoSizeConstraints(Caption)
                'Ok, get the adjusted caret pos
                Dim lAdjCaret As Int32 = GetAdjustedCursorLoc(sFormatted)
                'Ok, now, determine the previous VBCRLF
                Dim lNextEndLine As Int32 = sFormatted.IndexOf(vbCrLf, lAdjCaret)
                'Ok, get the width of the current line
                Dim lPrevLine As Int32 = sFormatted.LastIndexOf(vbCrLf, lNextEndLine - 1)

                Dim lCurrLineWidth As Int32 = 0
                If lPrevLine = -1 Then
                    'Ok, no lines before me...
                    If lAdjCaret = 0 Then
                        CursorPos = lNextEndLine + 2
                        SelStart = CursorPos
                        EnsureCursorVisible()
                        Return
                    Else
                        lCurrLineWidth = BPFont.MeasureString(moSysFont, sFormatted.Substring(0, lAdjCaret), FontFormat).Width
                    End If
                Else
                    lCurrLineWidth = BPFont.MeasureString(moSysFont, sFormatted.Substring(lPrevLine, lAdjCaret - lPrevLine), FontFormat).Width
                End If

                'Get the diff from adjust cursor to normal cursor
                Dim lDiff As Int32 = lAdjCaret - CursorPos

                'Now, get the length of the new line
                Dim sNewLine As String = sFormatted.Substring(lNextEndLine)
                Dim lEndOfNewLine As Int32 = sFormatted.IndexOf(vbCrLf, lNextEndLine + 2) - lNextEndLine
                Dim lClosestLen As Int32 = 100000
                Dim lClosestWidth As Int32 = 100000

                If lCurrLineWidth <> 0 Then
                    Try
                        For X As Int32 = 1 To lEndOfNewLine
                            Dim lTmpWidth As Int32 = BPFont.MeasureString(moSysFont, sNewLine.Substring(0, X), FontFormat).Width
                            If Math.Abs(lTmpWidth - lCurrLineWidth) < lClosestWidth Then
                                lClosestLen = X
                                lClosestWidth = Math.Abs(lTmpWidth - lCurrLineWidth)
                            End If
                        Next X
                    Catch
                    End Try
                Else
                    lClosestLen = 2 'for vbcrlf
                End If

                CursorPos = (lNextEndLine + lClosestLen) - lDiff

            Else : MoveToEndOfLine()
            End If
            SelStart = CursorPos
            EnsureCursorVisible()
 
        Catch ex As Exception
            'do nothing for now
        End Try
	End Sub

    Private Sub EnsureCursorVisible()
        Try
            If moScrollBar Is Nothing OrElse moScrollBar.Visible = False Then Return

            Dim bDone As Boolean = False
            While bDone = False
                bDone = True
                Dim ptPos As Point = GetCursorPos()
                moCursorVerts(0).X = ptPos.X
                moCursorVerts(0).Y = ptPos.Y + 2
                moCursorVerts(1).X = moCursorVerts(0).X
                moCursorVerts(1).Y = ptPos.Y + 16

                If moCursorVerts(0).Y > Me.Top + Me.Height OrElse moCursorVerts(1).Y > Me.Top + Me.Height Then
                    'need to increase
                    If moScrollBar.MaxValue = moScrollBar.Value Then Exit While
                    moScrollBar.Value += 1
                    bDone = False
                ElseIf moCursorVerts(1).Y < Me.Top OrElse moCursorVerts(0).Y < Me.Top Then
                    If moScrollBar.MinValue = moScrollBar.Value Then Exit While
                    moScrollBar.Value -= 1
                    bDone = False
                End If
                'TODO: Needs testing, adjustment.
                '    Dim lMeTop As Int32 = Me.GetRelativeTop()
                '    If moCursorVerts(0).Y >= lMeTop + Me.Height OrElse moCursorVerts(1).Y >= lMeTop + Me.Height Then
                '        'need to increase
                '        If moScrollBar.MaxValue = moScrollBar.Value Then Exit While
                '        moScrollBar.Value += 1
                '        bDone = False
                '    ElseIf moCursorVerts(1).Y < lMeTop OrElse moCursorVerts(0).Y < lMeTop Then
                '        If moScrollBar.MinValue = moScrollBar.Value Then Exit While
                '        moScrollBar.Value -= 1
                '        bDone = False
                '    End If
            End While


        Catch
            'do nothing
        End Try
    End Sub

    Public Sub ForceEndVisible()
        Try
            If MultiLine = True Then
                FormatMultiLineString(Me.Caption, 0)
                If moScrollBar Is Nothing = False AndAlso moScrollBar.Visible = True Then
                    moScrollBar.Value = moScrollBar.MaxValue
                End If
            End If
        Catch
        End Try
    End Sub
    Public Sub ForceStartVisible()
        Try
            If MultiLine = True Then
                FormatMultiLineString(Me.Caption, 0)
                If moScrollBar Is Nothing = False AndAlso moScrollBar.Visible = True Then
                    moScrollBar.Value = 0
                End If
            End If
        Catch
        End Try
    End Sub

    'Protected Overrides Sub Finalize()
    '       'If moBorderLine Is Nothing = False Then moBorderLine.Dispose()
    '       'moBorderLine = Nothing
    '       'If moCursorLine Is Nothing = False Then moCursorLine.Dispose()
    '       'moCursorLine = Nothing
    '	MyBase.Finalize()
    'End Sub

	Private Function FormatMultiLineString(ByVal sVal As String, ByRef lLineStart As Int32) As String
		'Ok... split our value into groups of vbCRLF
		Dim sLines() As String = Split(sVal, vbCrLf)
		Dim sFinal As String = ""
		Dim oRect As Rectangle
		Dim fOverage As Single
		'Dim lChars As Int32
		'Dim lFinalLen As Int32

		Dim lWidthMax As Int32 = Me.Width - 10
		If moScrollBar Is Nothing = False AndAlso moScrollBar.Visible = True Then lWidthMax -= moScrollBar.Width

		mbFormatted = True

		sFinal = JustDoFormattingNoSizeConstraints(sVal)

		'Now... figure out how many lines can fit inside the current rect
		sLines = Split(sFinal, vbCrLf)
        oRect = BPFont.MeasureString(moSysFont, sFinal, FontFormat) 'oRect = moFont.MeasureString(Nothing, sFinal, FontFormat, Color.White)
		'Now... the rect's total size
		fOverage = CSng(Me.Height / oRect.Height)
		If fOverage < 1.0F Then
			Dim lLineCnt As Int32 = CInt(Math.Floor(sLines.Length * fOverage))
			Dim bNeedRedo As Boolean = False
			'Ok, now that's how many lines we can see...
			If moScrollBar Is Nothing Then
				moScrollBar = New UIScrollBar(MyBase.moUILib, True)
				With moScrollBar
					.Width = 24
					.Left = Me.Width - .Width
					.Height = Me.Height
					.Top = 0
					.Value = 0
					.Visible = True
					.ReverseDirection = True
				End With
				Me.AddChild(CType(moScrollBar, UIControl))

				AddHandler moScrollBar.ValueChange, AddressOf ScrollBarValChange
				bNeedRedo = True
			End If
			If moScrollBar.Visible = False Then moScrollBar.Visible = True

			If bNeedRedo = True Then
				sFinal = FormatMultiLineString(sVal, lLineStart)
			Else
				'Ok, we can go ahead and do it now...
				If moScrollBar.MaxValue <> sLines.Length - lLineCnt Then moScrollBar.MaxValue = sLines.Length - lLineCnt
			End If
		ElseIf moScrollBar Is Nothing = False AndAlso moScrollBar.Visible = True Then
			moScrollBar.Visible = False
		End If

		'Now, one last thing to do...
		If moScrollBar Is Nothing = False AndAlso moScrollBar.Visible = True Then
			Dim lLineSkips As Int32 = moScrollBar.Value

			'Ok, line skips mean that we need to move the text that way...
			Dim lVal As Int32 = 0 'sFinal.IndexOf(vbCrLf, 0, lLineSkips, StringComparison.Ordinal)

			While lLineSkips <> 0 And lVal <> -1
				lVal = sFinal.IndexOf(vbCrLf, lVal + 1)
				lLineSkips -= 1
			End While

			If lVal = -1 Then lVal = sFinal.LastIndexOf(vbCrLf)
			If lVal <> -1 AndAlso lVal <> 0 Then
				lLineStart = lVal + 2
				sFinal = sFinal.Substring(lVal + 2)
			End If
		Else
			lLineStart = 0
		End If

		Return sFinal
	End Function

	Private Function JustDoFormattingNoSizeConstraints(ByVal sVal As String) As String
		'Ok... split our value into groups of vbCRLF
		Dim sLines() As String = Split(sVal, vbCrLf)
		Dim sFinal As String = ""
		Dim oRect As Rectangle
		Dim fOverage As Single
		Dim lChars As Int32
		Dim lFinalLen As Int32

		Dim lWidthMax As Int32 = Me.Width - 10
		If moScrollBar Is Nothing = False AndAlso moScrollBar.Visible = True Then lWidthMax -= moScrollBar.Width

		Dim lLenAdded As Int32 = 0

		For X As Int32 = 0 To sLines.GetUpperBound(0)
            oRect = BPFont.MeasureString(moSysFont, sLines(X), FontFormat) 'oRect = moFont.MeasureString(Nothing, sLines(X), FontFormat, Color.White)
			If oRect.Width > lWidthMax Then
				'get the overage multiplier...
				fOverage = CSng(lWidthMax / oRect.Width)	  '-10 for the left-right -ffsets of 5
				'Take the overage and multiply it by the length of the string...
				lChars = CInt(Math.Floor(fOverage * sLines(X).Length))
				If lChars <> 0 Then
					'Now, the number of chars will fit on the line... (hopefully)
					'  find the next space in the line before lchars
					lFinalLen = sLines(X).LastIndexOf(" ", lChars)
					If lFinalLen = -1 Then
						'Just append the line regardless...
						sFinal &= sLines(X).Substring(0, lChars) & vbCrLf '.Trim & vbCrLf
						sLines(X) = sLines(X).Substring(lChars + 1) & vbCrLf '.Trim
						If sLines(X).Trim.Length <> 0 Then
							X -= 1
						End If
					Else
						'Do a quick check for the max width....
						Dim sNewLine As String = sLines(X).Substring(0, lFinalLen) '.Trim
                        oRect = BPFont.MeasureString(moSysFont, sNewLine, FontFormat) 'oRect = moFont.MeasureString(Nothing, sNewLine, FontFormat, Color.White)
						If oRect.Width > lWidthMax Then
							Dim lTemp As Int32 = sNewLine.LastIndexOf(" ", lFinalLen - 1)
							If lTemp <> -1 Then lFinalLen = lTemp
						End If

						'Ok, in the end, let's just do it...
						sFinal &= sLines(X).Substring(0, lFinalLen) & vbCrLf '.Trim & vbCrLf
						sLines(X) = sLines(X).Substring(lFinalLen + 1) '.Trim
						If sLines(X).Trim.Length <> 0 Then
							X -= 1
						End If
					End If
				Else
					sFinal &= sLines(X) & vbCrLf '.Trim & vbCrLf
				End If
			Else : sFinal &= sLines(X) & vbCrLf	'.Trim & vbCrLf
			End If
		Next X
		Return sFinal
	End Function

	Private Function GetAdjustedCursorLoc(ByVal sFormatted As String) As Int32
		Dim sCaption As String = Caption
        Dim lNewLoc As Int32 = 0
        Try
            For X As Int32 = 0 To CursorPos - 1
                Dim sOrig As String = sCaption.Substring(X, 1)
                Dim sNew As String = sFormatted.Substring(lNewLoc, 1)
                If sOrig <> sNew Then
                    sNew = sFormatted.Substring(lNewLoc, 2)
                    If sNew = vbCrLf Then lNewLoc += 1
                End If
                lNewLoc += 1
            Next X
        Catch
        End Try
        Return lNewLoc
	End Function

	Private Sub ScrollBarValChange()
		Me.IsDirty = False
		Me.IsDirty = True
	End Sub

	Public bRoundedBorder As Boolean = False
    Protected Sub RenderBorder(ByVal lTopLineX As Int32, ByVal lAddToHeight As Int32, ByRef oBLine As Line)
        Dim v2Lines() As Vector2
        Dim oLoc As Point = GetAbsolutePosition()
        Dim lOffset As Int32 = BorderLineWidth \ 2

        If bRoundedBorder = True Then
            ReDim v2Lines(8)

            With v2Lines(0)
                .X = oLoc.X + lTopLineX + lOffset + 3
                If lTopLineX <> 0 Then lTopLineX += 10
                .Y = oLoc.Y + lOffset
            End With
            With v2Lines(1)
                .X = oLoc.X + Width - 3 - lOffset
                .Y = oLoc.Y + lOffset
            End With
            With v2Lines(2)
                .X = oLoc.X + Width - lOffset
                .Y = oLoc.Y + 3 + lOffset
            End With
            With v2Lines(3)
                .X = oLoc.X + Width - lOffset
                .Y = oLoc.Y + Height - 3 - lOffset + lAddToHeight
            End With
            With v2Lines(4)
                .X = oLoc.X + Width - 3 - lOffset
                .Y = oLoc.Y + Height - lOffset + lAddToHeight
            End With
            With v2Lines(5)
                .X = oLoc.X + 3 + lOffset
                .Y = oLoc.Y + Height - lOffset + lAddToHeight
            End With
            With v2Lines(6)
                .X = oLoc.X + lOffset
                .Y = oLoc.Y + Height - 3 - lOffset + lAddToHeight
            End With
            With v2Lines(7)
                .X = oLoc.X + lOffset
                .Y = oLoc.Y + 3 + lOffset
            End With
            With v2Lines(8)
                .X = oLoc.X + 3 + lOffset
                .Y = oLoc.Y + lOffset
            End With
        Else
            ReDim v2Lines(4)

            With v2Lines(0)
                .X = oLoc.X + lTopLineX + lOffset
                If lTopLineX <> 0 Then .X += 10
                .Y = oLoc.Y + lOffset
            End With
            With v2Lines(1)
                .X = oLoc.X + Width - lOffset
                .Y = oLoc.Y + lOffset
            End With
            With v2Lines(2)
                .X = oLoc.X + Width - lOffset
                .Y = oLoc.Y + Height - lOffset + lAddToHeight
            End With
            With v2Lines(3)
                .X = oLoc.X + lOffset
                .Y = oLoc.Y + Height - lOffset + lAddToHeight
            End With
            With v2Lines(4)
                .X = oLoc.X + lOffset
                .Y = oLoc.Y + lOffset
            End With
        End If

        'Draw a box border around our window...
        If oBLine Is Nothing = False Then
            Try
                With oBLine
                    .Draw(v2Lines, BorderColor)
                End With
            Catch
            End Try
        Else

            Try
                Using moBorderLine As New Line(MyBase.moUILib.oDevice)
                    With moBorderLine
                        .Antialias = True
                        .Width = BorderLineWidth
                        .Begin()
                        .Draw(v2Lines, BorderColor)
                        .End()
                    End With
                End Using
            Catch
            End Try
        End If
        

        'MyBase.moUILib.AddLineListItem(v2Lines, True, BorderLineWidth, BorderColor)

    End Sub

	Private Function GetInvisibleTextLen(ByVal sFormatted As String) As Int32
		'Scroll Bar Value is number of VBCRLF's to exclude
		'  Loop thru string finding loc of VBCRLF's until you find the one needed
		'That position, store it as InvisibleTextLen
		Dim lInvisibleTextLen As Int32 = 0
		If moScrollBar Is Nothing = False AndAlso moScrollBar.Value <> 0 Then
			Dim lVBCRLFs As Int32 = 0
			While lVBCRLFs <> moScrollBar.Value
				lInvisibleTextLen = sFormatted.IndexOf(vbCrLf, lInvisibleTextLen + 1)
				If lInvisibleTextLen = -1 Then
					lInvisibleTextLen = 0
					Exit While
				Else
					lVBCRLFs += 1
				End If
			End While
		End If
		If lInvisibleTextLen <> 0 Then lInvisibleTextLen += 2
		Return lInvisibleTextLen
	End Function
	Private Function GetCursorPos() As Point
        'sFormatted = Format Text to be wrap-safe (insert VBCRLF's as needed)
        Dim sTempCaption As String = Caption.Replace(vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)
        Dim sFormatted As String = JustDoFormattingNoSizeConstraints(sTempCaption)
        sFormatted = sFormatted.Replace(vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)
		Dim lInvisibleTextLen As Int32 = GetInvisibleTextLen(sFormatted)

		'Take the caret position and subtract the InvisibleTextLen, this is the position within visible text or higher of the caret
		Dim lCaretPos As Int32 = GetAdjustedCursorLoc(sFormatted) ' SelStart
		lCaretPos -= lInvisibleTextLen
		'Dim lMultiLineRemoval As Int32 = 0
		'If moScrollBar Is Nothing = False Then lMultiLineRemoval = (moScrollBar.Value * 2)
		'lCaretPos -= lMultiLineRemoval
		Dim sVisibleText As String = sFormatted.Substring(lInvisibleTextLen) ' + lMultiLineRemoval)

		''find the preceding vbCRLF's Index
		'Dim lCaretPosY As Int32 = 0
		'Dim lPreceding As Int32 = sVisibleText.LastIndexOf(vbCrLf, lCaretPos)
		'If lPreceding <> -1 Then
		'    'sPrecedingLINES = the visible text up to the preceding vbCRLF not including the preceding vbCRLF
		'    Dim sPrecedingLines As String = sVisibleText.Substring(0, lPreceding)
		'    'Measure that text as rcPreceding
		'    Dim rcPreceding As Rectangle = moFont.MeasureString(Nothing, sPrecedingLines, FontFormat, ForeColor)
		'    'Caret position Y = rcPreceding.Top + rcPreceding.Height (plus any offset from being within the textbox)
		'    lCaretPosY = rcPreceding.Top + rcPreceding.Height

		'    'Take the remaining text line from the preceding vbCRLF to the caret position (not including the preceding vbCRLF)
		'    sVisibleText = sVisibleText.Substring(lPreceding + 2)
		'End If

		''Measure that text as rcLine
		'Dim rcLine As Rectangle = moFont.MeasureString(Nothing, sVisibleText, FontFormat, ForeColor)
		''Caret Position X = rcLine.Left + rcLine.Width (plus any offset from being within the textbox)
		'Dim lCaretPosX As Int32 = rcLine.Left + rcLine.Width + 5

		''NOTE: Each vbCRLF is 2 characters
		'Return New Point(lCaretPosX + Me.Left, lCaretPosY + Me.Top)

		Dim lWidthMax As Int32 = Me.Width '- 5
		If moScrollBar Is Nothing = False AndAlso moScrollBar.Visible = True Then lWidthMax -= moScrollBar.Width

        'InitializeTextBox(True, (moScrollBar Is Nothing = False AndAlso moScrollBar.Visible = True), sVisibleText, lWidthMax, Me.Height, moSysFont)
        'Dim ptRes As Point = moTextBox.GetPositionFromCharIndex(lCaretPos)
        Dim lCurrLineVBCRLF As Int32 = -1
        If lCaretPos > 0 Then
            lCurrLineVBCRLF = sVisibleText.LastIndexOf(vbCrLf, lCaretPos)
        End If
        Dim sLinesBefore As String = ""
        Dim rcLinesBefore As Rectangle
        Dim ptRes As Point

        If lCurrLineVBCRLF <> -1 Then
            sLinesBefore = sVisibleText.Substring(0, lCurrLineVBCRLF)
            rcLinesBefore = BPFont.MeasureString(moSysFont, sLinesBefore, FontFormat Or DrawTextFormat.NoClip)
            ptRes = New Point(0, rcLinesBefore.Height)
        Else
            ptRes = New Point(0, 0)
        End If

        If lCurrLineVBCRLF + 2 > lCaretPos Then
            If sLinesBefore <> "" Then
                sLinesBefore = ""
                If lCurrLineVBCRLF + 1 <> lCaretPos Then
                    ptRes.Y += BPFont.MeasureString(moSysFont, "MWPmwpq", FontFormat).Height
                Else
                    lCurrLineVBCRLF = sVisibleText.LastIndexOf(vbCrLf, lCurrLineVBCRLF - 1)
                    ptRes.Y -= BPFont.MeasureString(moSysFont, "MWPmwpq", FontFormat).Height

                    Dim lAdjPos As Int32
                    If lCurrLineVBCRLF = -1 Then lAdjPos = 0 Else lAdjPos = lCurrLineVBCRLF + 2

                    sLinesBefore = sVisibleText.Substring(lAdjPos, lCaretPos - lAdjPos)
                    If sLinesBefore.EndsWith(vbCr) = True Then sLinesBefore = sLinesBefore.Substring(0, sLinesBefore.Length - 1)
                    If sLinesBefore.EndsWith(vbLf) = True Then sLinesBefore = sLinesBefore.Substring(0, sLinesBefore.Length - 1)
                    rcLinesBefore = BPFont.MeasureString(moSysFont, sLinesBefore, FontFormat Or DrawTextFormat.NoClip)
                    ptRes.X += rcLinesBefore.Width
                End If
            End If
        Else
            Dim lAdjPos As Int32
            If lCurrLineVBCRLF = -1 Then lAdjPos = 0 Else lAdjPos = lCurrLineVBCRLF + 2

            sLinesBefore = sVisibleText.Substring(lAdjPos, lCaretPos - lAdjPos)
            rcLinesBefore = BPFont.MeasureString(moSysFont, sLinesBefore, FontFormat Or DrawTextFormat.NoClip)
            ptRes.X += rcLinesBefore.Width
        End If

        Dim lWSpace As Int32 = BPFont.MeasureString(moSysFont, " W", DrawTextFormat.Left Or DrawTextFormat.VerticalCenter).Width
        Dim lWNoSpace As Int32 = BPFont.MeasureString(moSysFont, "W", DrawTextFormat.Left Or DrawTextFormat.VerticalCenter).Width
        Dim lPerSpace As Int32 = lWSpace - lWNoSpace
        If sLinesBefore = "" Then
            ptRes.X -= lPerSpace
        Else
            While sLinesBefore.EndsWith(" ") = True
                sLinesBefore = sLinesBefore.Substring(0, sLinesBefore.Length - 1)
                ptRes.X += lPerSpace
            End While
        End If
        




		'goUILib.AddNotification(Me.ControlName & ", CaretPosIdx: " & lCaretPos, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		'Dim ptTemp As Point = Me.GetAbsolutePosition()

		'Ok, here's the trick, I need to get my absolute position..
		Dim ptTemp As Point = Me.GetAbsolutePosition()
		'but then, I need to find the top-level window
		Dim oParent As UIControl = Me.ParentControl
		Dim ptParent As Point = New Point(0, 0)
		While oParent Is Nothing = False
			ptParent.X = oParent.Left
			ptParent.Y = oParent.Top
			oParent = oParent.ParentControl
		End While
		ptTemp.X -= ptParent.X
		ptTemp.Y -= ptParent.Y

        'If ptTemp.Y > Me.Height \ 2 Then
        '    ptTemp.Y -= BPFont.MeasureString(moSysFont, "W", Me.FontFormat).Height
        'End If
        ptRes.X += ptTemp.X + 5
		ptRes.Y += ptTemp.Y
		Return ptRes
	End Function

    'Private Shared moTextBox As TextBox
    '  Private Shared Sub InitializeTextBox(ByVal bMultiline As Boolean, ByVal bScrollbar As Boolean, ByVal sText As String, ByVal lWidth As Int32, ByVal lHeight As Int32, ByVal oFont As System.Drawing.Font)
    '      If moTextBox Is Nothing Then moTextBox = New TextBox()
    'moTextBox.Width = lWidth
    '      moTextBox.Height = lHeight
    '      moTextBox.Font = oFont
    '      moTextBox.Multiline = bMultiline
    '      moTextBox.WordWrap = bMultiline
    '      If bScrollbar = True Then moTextBox.ScrollBars = ScrollBars.Vertical Else moTextBox.ScrollBars = ScrollBars.None
    '      moTextBox.Text = sText
    '  End Sub

    Private Sub UITextBox_ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color) Handles Me.ResetInterfaceColors
        Select Case yType
            Case 1      'border
                If BorderColor = clrPrev Then BorderColor = muSettings.InterfaceBorderColor
            Case 2      'fill
                'unused
            Case 3      'textboxfore
                If ForeColor = clrPrev Then ForeColor = muSettings.InterfaceTextBoxForeColor
            Case 4      'textboxfill
                'unused
                If BackColorEnabled = clrPrev Then BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            Case 5      'button
                'unused
        End Select
    End Sub

    Public Function SelectedText() As String
        Try
            If (Me.PasswordChar Is Nothing OrElse Me.PasswordChar = "") AndAlso SelLength > 0 Then
                Return Me.Caption.Substring(SelStart, Math.Min(SelLength, Me.Caption.Length))
            Else : Return ""
            End If
        Catch
        End Try
        Return ""
    End Function

    Public Sub PasteText(ByVal sVal As String)
        If Me.Locked = True Then Exit Sub
        Dim sValue As String = Mid$(Caption, 1, SelStart) & sVal & Mid$(Caption, SelStart + Math.Max(SelLength + 1, 1))
        If Me.MaxLength > 0 Then sValue = Mid$(sValue, 1, Me.MaxLength)
        SelStart = SelStart + sVal.Length
        CursorPos = SelStart
        Me.Caption = sValue
    End Sub

    Public Function CutText() As String
        Dim sResult As String = ""
        If (Me.PasswordChar Is Nothing OrElse Me.PasswordChar = "") AndAlso SelLength > 0 Then
            sResult = Me.Caption.Substring(SelStart, SelLength)
        End If

        Caption = Mid$(Caption, 1, SelStart) & Mid$(Caption, SelStart + Math.Max(1, SelLength + 1))
        SelStart = CursorPos

        Return sResult
    End Function

End Class


'        /// <summary>Set the caret to a character position, and adjust the scrolling if necessary</summary>
'        protected void PlaceCaret(int pos)
'        {
'            // Store caret position
'            caretPosition = pos;

'            // First find the first visible char
'            for (int i = 0; i < textData.Text.Length; i++)
'            {
'                System.Drawing.Point p = textData.GetPositionFromCharIndex(i);
'                if (p.X >= 0) 
'                {
'                    firstVisible = i; // This is the first visible character
'                    break;
'                }
'            }

'            // if the new position is smaller than the first visible char 
'            // we'll need to scroll
'            if (firstVisible > caretPosition)
'                firstVisible = caretPosition;
'        }



'        /// <summary>Copy the selected text to the clipboard</summary>
'        protected void CopyToClipboard()
'        {
'            // Copy the selection text to the clipboard
'            if (caretPosition != textData.SelectionStart)
'            {
'                int first = Math.Min(caretPosition, textData.SelectionStart);
'                int last = Math.Max(caretPosition, textData.SelectionStart);
'                // Set the text to the clipboard
'                System.Windows.Forms.Clipboard.SetDataObject(textData.Text.Substring(first, (last-first)));
'            }

'        }
'        /// <summary>Paste the clipboard data to the control</summary>
'        protected void PasteFromClipboard()
'        {
'            // Get the clipboard data
'            System.Windows.Forms.IDataObject clipData = System.Windows.Forms.Clipboard.GetDataObject();
'            // Does the clipboard have string data?
'            if (clipData.GetDataPresent(System.Windows.Forms.DataFormats.StringFormat))
'            {
'                // Yes, get that data
'                string clipString = clipData.GetData(System.Windows.Forms.DataFormats.StringFormat) as string;
'                // find any new lines, remove everything after that
'                int index;
'                if ((index = clipString.IndexOf("\n")) > 0)
'                {
'                    clipString = clipString.Substring(0, index-1);
'                }

'                // Insert that into the text data
'                textData.Text = textData.Text.Insert(caretPosition, clipString);
'                caretPosition += clipString.Length;
'                textData.SelectionStart = caretPosition;
'                FocusText();
'            }
'        }

'/// <summary>Render the control</summary>
'        public override void Render(Device device, float elapsedTime)
'        {
'            if (!IsVisible)
'                return; // Nothing to render

'            // Render the control graphics
'            for (int i = 0; i <= LowerRightBorder; ++i)
'            {
'                Element e = elementList[i] as Element;
'                e.TextureColor.Blend(ControlState.Normal,elapsedTime);
'                parentDialog.DrawSprite(e, elementRects[i]);
'            }
'            //
'            // Compute the X coordinates of the first visible character.
'            //
'            int xFirst = textData.GetPositionFromCharIndex(firstVisible).X;
'            int xCaret = textData.GetPositionFromCharIndex(caretPosition).X;
'            int xSel;

'            if (caretPosition != textData.SelectionStart)
'                xSel = textData.GetPositionFromCharIndex(textData.SelectionStart).X;
'            else
'                xSel = xCaret;

'            // Render the selection rectangle
'            System.Drawing.Rectangle selRect = System.Drawing.Rectangle.Empty;
'            if (caretPosition != textData.SelectionStart)
'            {
'                int selLeft = xCaret, selRight = xSel;
'                // Swap if left is beigger than right
'                if (selLeft > selRight)
'                {
'                    int temp = selLeft;
'                    selLeft = selRight;
'                    selRight = temp;
'                }
'                selRect = System.Drawing.Rectangle.FromLTRB(
'                    selLeft, textRect.Top, selRight, textRect.Bottom);
'                selRect.Offset(textRect.Left - xFirst, 0);
'                selRect.Intersect(textRect);
'                Parent.DrawRectangle(selRect, selectedBackColor);
'            }

'            // Render the text
'            Element textElement = elementList[TextLayer] as Element;
'            textElement.FontColor.Current = textColor;
'            parentDialog.DrawText(textData.Text.Substring(firstVisible), textElement, textRect);

'            // Render the selected text
'            if (caretPosition != textData.SelectionStart)
'            {
'                int firstToRender = Math.Max(firstVisible, Math.Min(textData.SelectionStart, caretPosition));
'                int numToRender = Math.Max(textData.SelectionStart, caretPosition) - firstToRender;
'                textElement.FontColor.Current = selectedTextColor;
'                parentDialog.DrawText(textData.Text.Substring(firstToRender, numToRender), textElement, selRect);
'            }

'            //
'            // Blink the caret
'            //
'            if(FrameworkTimer.GetAbsoluteTime() - lastBlink >= blinkTime)
'            {
'                isCaretOn = !isCaretOn;
'                lastBlink = FrameworkTimer.GetAbsoluteTime();
'            }

'            //
'            // Render the caret if this control has the focus
'            //
'            if( hasFocus && isCaretOn && !isHidingCaret )
'            {
'                // Start the rectangle with insert mode caret
'                System.Drawing.Rectangle caretRect = textRect;
'                caretRect.Width = 2;
'                caretRect.Location = new System.Drawing.Point(
'                    caretRect.Left - xFirst + xCaret -1, 
'                    caretRect.Top);

'                // If we are in overwrite mode, adjust the caret rectangle
'                // to fill the entire character.
'                if (!isInsertMode)
'                {
'                    // Obtain the X coord of the current character
'                    caretRect.Width = 4;
'                }

'                parentDialog.DrawRectangle(caretRect, caretColor);
'            }

'        }