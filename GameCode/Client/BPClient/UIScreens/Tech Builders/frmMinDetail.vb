Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmMinDetail
    Inherits UIWindow

    'Private btnClose As UIButton
    Private WithEvents lstMinerals As UIListBox
    Private WithEvents btnHelp As UIButton
	Private WithEvents btnAvailResources As UIButton

    Private WithEvents btnSelect As UIButton

    Private WithEvents chkFilterArchived As UICheckBox

    Private WithEvents chkSortByMatch As UICheckBox

    Private lblTitle As UILabel
    Private lblProps() As UILabel
    Private lblValue() As UILabel
    'Private lblKnown() As UILabel
	Private myHighlighted() As Byte			'0 = not highlighted, 1 = white, 2 = green, 3 = red, 4 = yellow
	Private myPerfectValue() As Byte

	Private mlCurrentMineralID As Int32 = -1
	Private mlLastUpdateMsg As Int32

    Private mlLastListCheck As Int32

    Public bRaiseSelectEvent As Boolean = False
    Public Event MineralSelected(ByVal lMineralID As Int32)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmMinDetail initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eMineralDetail
            .ControlName = "frmMinDetail"
            .Left = 154
            .Top = 60
            .Width = 512
            .Height = 387
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .mbAcceptReprocessEvents = True
            .Moveable = False
        End With

        ''btnClose initial props
        'btnClose = New UIButton(oUILib)
        'With btnClose
        '    .ControlName = "btnClose"
        '    .Left = Me.Width - 25
        '    .Top = 1
        '    .Width = 24
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "X"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = CType(5, DrawTextFormat)
        '    .ControlImageRect = New Rectangle(0, 0, 120, 32)
        'End With
        'Me.AddChild(CType(btnClose, UIControl))

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = Me.Width - 25 ' btnClose.Left - 25
            .Top = 1 'btnClose.Top
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 14.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
            .ToolTipText = "Click to begin the tutorial for this window."
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'btnSelect initial props
        btnSelect = New UIButton(oUILib)
        With btnSelect
            .ControlName = "btnSelect"
            .Left = Me.Width - 110 ' btnClose.Left - 25
            .Top = Me.Height - 30
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = False
            .Caption = "Select"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
            .ToolTipText = "Click to select the selected material and close the window."
        End With
        Me.AddChild(CType(btnSelect, UIControl))

        lblTitle = New UILabel(MyBase.moUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5 '207
            .Top = 10
            .Width = 350
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'chkSortByMatch initial props
        chkSortByMatch = New UICheckBox(oUILib)
        With chkSortByMatch
            .ControlName = "chkSortByMatch"
            .Left = 10
            .Top = (lblTitle.Top + lblTitle.Height + 5)
            .Width = 125
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Sort By Best Match"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkSortByMatch, UIControl))

        'lstMinerals initial props
        lstMinerals = New UIListBox(oUILib)
        With lstMinerals
            .ControlName = "lstMinerals"
            .Left = 5
            .Top = chkSortByMatch.Top + chkSortByMatch.Height + 5
            .Width = 240 '150
			.Height = 360
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
		Me.AddChild(CType(lstMinerals, UIControl))

		'chkFilterArchived initial props
		chkFilterArchived = New UICheckBox(oUILib)
		With chkFilterArchived
			.ControlName = "chkFilterArchived"
			.Left = lstMinerals.Left + 5
			.Top = lstMinerals.Top + lstMinerals.Height + 5
			.Width = 100
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Filter Archived"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = True
			.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
		End With
		Me.AddChild(CType(chkFilterArchived, UIControl))

        'btnAvailResources initial props
        btnAvailResources = New UIButton(oUILib)
        With btnAvailResources
            .ControlName = "btnAvailResources"
            .Left = 5
			.Top = chkFilterArchived.Top + chkFilterArchived.Height + 5
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Available Resources"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAvailResources, UIControl))

        ReDim lblProps(glMineralPropertyUB)
        ReDim lblValue(glMineralPropertyUB)
        'ReDim lblKnown(glMineralPropertyUB)
		ReDim myHighlighted(glMineralPropertyUB)
		ReDim myPerfectValue(glMineralPropertyUB)

        For lIdx As Int32 = 0 To glMineralPropertyUB
            lblProps(lIdx) = New UILabel(MyBase.moUILib)
            With lblProps(lIdx)
                .ControlName = "lblProps(" & lIdx.ToString & ")"
                .Left = lstMinerals.Width + 15 '165 '5
                .Top = (chkSortByMatch.Top + chkSortByMatch.Height + 5) + (lIdx * 18) '(lblTitle.Top + lblTitle.Height + 5) + (lIdx * 18)
                .Width = 175
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = goMineralProperty(lIdx).MineralPropertyName
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter
            End With
            'Me.AddChild(CType(lblProps(lIdx), UIControl))

            lblValue(lIdx) = New UILabel(MyBase.moUILib)
            With lblValue(lIdx)
                .ControlName = "lblValue(" & lIdx.ToString & ")"
                .Left = lstMinerals.Width + 190 ' 360 '345 '185
                .Top = lblProps(lIdx).Top + 4
                .Width = 63
                .Height = 9
                .Enabled = True
                .Visible = False
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.VerticalCenter
            End With
            'Me.AddChild(CType(lblValue(lIdx), UIControl))

            'lblKnown(lIdx) = New UILabel(MyBase.moUILib)
            'With lblKnown(lIdx)
            '    .ControlName = "lblKnown(" & lIdx.ToString & ")"
            '    .Left = 435 '275
            '    .Top = lblProps(lIdx).Top + 4
            '    .Width = 63
            '    .Height = 9
            '    .Enabled = True
            '    .Visible = False
            '    .Caption = ""
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            '    .DrawBackImage = True
            '    .FontFormat = DrawTextFormat.VerticalCenter
            'End With
            ''Me.AddChild(CType(lblKnown(lIdx), UIControl))
            'lblKnown(lIdx).Visible = goMineralProperty(lIdx).yKnowledgeLevel > eMinPropKnowledgeLevel.eExistence

			myHighlighted(lIdx) = 0
			myPerfectValue(lIdx) = 255
        Next

        FillMineralList()

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewMining) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else : Return
        End If

        'AddHandler btnClose.Click, AddressOf btnClose_Click
    End Sub

    ''' <summary>
    ''' Only to be used from the AlloyBuilder!!!!
    ''' </summary>
    ''' <param name="lLeft"></param>
    ''' <param name="lTop"></param>
    ''' <param name="lHeight"></param>
    ''' <param name="oMineralRef"></param>
    ''' <remarks></remarks>
    Public Sub ShowMineralDetail(ByVal lLeft As Int32, ByVal lTop As Int32, ByVal lHeight As Int32, ByRef oMineralRef As Mineral)
        Dim lTemp As Int32

        For X As Int32 = 0 To lblProps.GetUpperBound(0)
            lblProps(X).Visible = False
            lblValue(X).Visible = False
            'lblKnown(X).Visible = False
        Next X

        If oMineralRef Is Nothing Then Return

        lblTitle.Caption = "Mineral Properties for " & oMineralRef.MineralName

        For Y As Int32 = 0 To glMineralPropertyUB
            Dim lIdx As Int32 = oMineralRef.GetPropertyIndex(glMineralPropertyIdx(Y))
            If lIdx <> -1 Then
                'Set the tooltip
                lblProps(Y).ToolTipText = ""

                'Now, the actual value...
                lTemp = oMineralRef.MinPropValueScore(lIdx) '\ 10
                If lTemp < 4 Then
                    lblValue(Y).BackImageColor = Color.Red
                ElseIf lTemp < 7 Then
                    lblValue(Y).BackImageColor = Color.Yellow
                Else : lblValue(Y).BackImageColor = Color.Green
                End If

                lblValue(Y).ControlImageRect = grc_UI(elInterfaceRectangle.eMinBar_0 + lTemp)
                lblValue(Y).ToolTipText = "Value: " & lTemp
                lblProps(Y).Visible = True
                lblValue(Y).Visible = True
            End If
        Next Y

        Me.Visible = True
        Me.Left = lLeft
        Me.Top = lTop
        Me.Height = lHeight
    End Sub

    Public Sub ShowMineralDetail(ByVal lLeft As Int32, ByVal lTop As Int32, ByVal lHeight As Int32, ByVal lMineralID As Int32)
        Dim lTemp As Int32

        If lMineralID = -1 Then
            If lstMinerals.ListCount > 0 Then
                lMineralID = lstMinerals.ItemData(0)
            End If
        End If

        If mlCurrentMineralID = lMineralID AndAlso lMineralID <> -1 Then Return

        For X As Int32 = 0 To lblProps.GetUpperBound(0)
            lblProps(X).Visible = False
            lblValue(X).Visible = False
            'lblKnown(X).Visible = False
        Next X

		Dim lMinUB As Int32 = -1
		If glMineralIdx Is Nothing = False Then
			lMinUB = Math.Min(glMineralUB, glMineralIdx.GetUpperBound(0))
		End If
		For X As Int32 = 0 To lMinUB
			If glMineralIdx(X) = lMineralID Then

				mlLastUpdateMsg = goMinerals(X).lLastMsgUpdate
				If goMinerals(X).bRequestedProps = False Then
					mlLastUpdateMsg = 0
					goMinerals(X).lLastMsgUpdate = 0
					goMinerals(X).bRequestedProps = True
					Dim yMsg(5) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eRequestMineral).CopyTo(yMsg, 0)
					System.BitConverter.GetBytes(lMineralID).CopyTo(yMsg, 2)
					MyBase.moUILib.SendMsgToPrimary(yMsg)
				End If

				lblTitle.Caption = "Mineral Properties for " & goMinerals(X).MineralName

				For Y As Int32 = 0 To glMineralPropertyUB
					Try
						Dim lIdx As Int32 = goMinerals(X).GetPropertyIndex(glMineralPropertyIdx(Y))

						'Set the tooltip
						lblProps(Y).ToolTipText = goMinerals(X).SentenceDescription(lIdx)

						'Now, the actual value...
						lTemp = goMinerals(X).MinPropValueScore(lIdx) '\ 10
						If lTemp < 4 Then
							lblValue(Y).BackImageColor = Color.Red
						ElseIf lTemp < 7 Then
                            lblValue(Y).BackImageColor = System.Drawing.Color.FromArgb(255, 192, 192, 0)  'Color.Yellow
						Else : lblValue(Y).BackImageColor = Color.Green
						End If
						'lLeftOffset = ((lTemp \ 4) * 63) + 50
						'lTemp = lTemp Mod 4
						'lblValue(Y).ControlImageRect = Rectangle.FromLTRB(lLeftOffset, 215 + (lTemp * 9), lLeftOffset + 63, 215 + ((lTemp + 1) * 9))
                        lblValue(Y).ControlImageRect = grc_UI(elInterfaceRectangle.eMinBar_0 + lTemp)
                        lblValue(Y).ToolTipText = "Value: " & lTemp

                        ''Now, the Known Value...
                        'lTemp = goMinerals(X).MinPropKnownScore(lIdx) + 1 '\ 10
                        'If lTemp < 4 Then
                        '	lblKnown(Y).BackImageColor = Color.Red
                        'ElseIf lTemp < 7 Then
                        '	lblKnown(Y).BackImageColor = Color.Yellow
                        'Else : lblKnown(Y).BackImageColor = Color.Green
                        'End If
                        ''lLeftOffset = ((lTemp \ 4) * 63) + 50
                        ''lTemp = lTemp Mod 4
                        ''lblKnown(Y).ControlImageRect = Rectangle.FromLTRB(lLeftOffset, 215 + (lTemp * 9), lLeftOffset + 63, 215 + ((lTemp + 1) * 9))
                        'lblKnown(Y).ControlImageRect = Rectangle.FromLTRB(193, 157 + (lTemp * 9), 256, 157 + ((lTemp + 1) * 9))

						lblProps(Y).Visible = True
						lblValue(Y).Visible = True
                        'lblKnown(Y).Visible = goMineralProperty(Y).yKnowledgeLevel > eMinPropKnowledgeLevel.eExistence
					Catch
					End Try
				Next Y

				Exit For
			End If
		Next X

        Me.Visible = True
        Me.Left = lLeft
        Me.Top = lTop
        Me.Height = lHeight
		'lstMinerals.Height = Me.Height - lstMinerals.Top - 5

		mlCurrentMineralID = lMineralID
    End Sub

    'Private Sub btnClose_Click(ByVal sName As String)
    '	MyBase.moUILib.RemoveWindow(Me.ControlName)
    'End Sub

    Private Sub frmMinDetail_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        If lblProps Is Nothing = False Then
            Dim ptParent As Point = Me.GetAbsolutePosition()
            For X As Int32 = 0 To lblProps.GetUpperBound(0)
                Dim ptLoc As Point = lblProps(X).GetAbsolutePosition()
                ptLoc.X += ptParent.X : ptLoc.Y += ptParent.Y
                If lMouseX > ptLoc.X AndAlso lMouseX < ptLoc.X + lblProps(X).Width Then
                    If lMouseY > ptLoc.Y AndAlso lMouseY < ptLoc.Y + lblProps(X).Height Then

                        myHighlighted(X) += CByte(1)
						If myHighlighted(X) > 4 Then myHighlighted(X) = 0

                        Me.IsDirty = True
                        Exit For
                    End If
                End If
            Next X
        End If

    End Sub

    Private Sub frmMinDetail_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim ptLoc As Point = Me.GetAbsolutePosition
        Dim lX As Int32 = lMouseX
        Dim lY As Int32 = lMouseY
        lX -= ptLoc.X
        lY -= ptLoc.Y

        If lblValue Is Nothing = False AndAlso lblValue.GetUpperBound(0) > -1 AndAlso lblValue(0) Is Nothing = False Then
            If lX > lblValue(0).Left AndAlso lX < lblValue(0).Left + lblValue(0).Width Then
                Dim lLastIdx As Int32 = -1
                For X As Int32 = 0 To lblValue.GetUpperBound(0)
                    If lblValue(X) Is Nothing = False AndAlso lY > lblValue(X).Top Then
                        lLastIdx = X
                    Else
                        Exit For
                    End If
                Next X

                If lY > lblValue(lblValue.GetUpperBound(0)).Top + lblValue(lblValue.GetUpperBound(0)).Height Then lLastIdx = -1

                If lLastIdx <> -1 Then
                    Dim sText As String = lblValue(lLastIdx).ToolTipText
                    If myPerfectValue(lLastIdx) <> 255 Then
                        sText &= vbCrLf & "Desired: " & myPerfectValue(lLastIdx)
                    End If
                    MyBase.moUILib.SetToolTip(sText, lMouseX, lMouseY)
                End If
            End If

        End If
    End Sub

	Private Sub frmMinDetail_OnNewFrame() Handles Me.OnNewFrame
		If glCurrentCycle - mlLastListCheck > 30 Then
			mlLastListCheck = glCurrentCycle
			Dim bFilter As Boolean = chkFilterArchived.Value
			Dim lCnt As Int32 = 0
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) <> -1 AndAlso goMinerals(X).bDiscovered = True Then
					If goMinerals(X).bArchived = False OrElse bFilter = False Then
						lCnt += 1
					End If
				End If
			Next X
			If lCnt <> lstMinerals.ListCount Then FillMineralList()
		End If

		For X As Int32 = 0 To glMineralUB
			If glMineralIdx(X) = mlCurrentMineralID Then
				If mlLastUpdateMsg <> goMinerals(X).lLastMsgUpdate Then
					Dim lMinID As Int32 = mlCurrentMineralID
					mlCurrentMineralID = -1
					ShowMineralDetail(Me.Left, Me.Top, Me.Height, lMinID)
				End If
				Exit For
			End If
        Next X
        If btnSelect.Visible = True Then
            Dim lTop As Int32 = Me.Height - 30
            If btnSelect.Top <> lTop Then btnSelect.Top = lTop
        End If
    End Sub

    Private mfBlinkAlpha As Single = 1.0F
    Private mfAlphaChg As Single = -0.01F

    Private Sub frmMinDetail_OnNewFrameEnd() Handles Me.OnNewFrameEnd
        'Now, the other items
        mfBlinkAlpha += mfAlphaChg
        If mfBlinkAlpha < 0 Then
            mfBlinkAlpha = 0.0F
            mfAlphaChg = 0.1F
        ElseIf mfBlinkAlpha > 1.0F Then
            mfBlinkAlpha = 1.0F
            mfAlphaChg = -0.1F
        End If
        Dim lColor As Int32 = CInt(mfBlinkAlpha * 255)
        If lColor < 0 Then lColor = 0
        If lColor > 255 Then lColor = 255
        If lblValue Is Nothing = False Then
            'MyBase.moUILib.Pen.Begin(SpriteFlags.AlphaBlend)
            MyBase.moUILib.BeginPenSprite(SpriteFlags.AlphaBlend)
            Dim lBaseLeft As Int32 = Me.Left
            Dim lBaseTop As Int32 = Me.Top
            For X As Int32 = 0 To lblValue.GetUpperBound(0)
                'lblValue(X).ManualDrawSprite(MyBase.moUILib.Pen)
                'lblKnown(X).ManualDrawSprite(MyBase.moUILib.Pen)

                If myPerfectValue(X) <> 255 Then
                    If myPerfectValue(X) > 10 Then myPerfectValue(X) = 10
                    Dim lNewLeft As Int32 = lBaseLeft + lblValue(X).Left + (CInt(myPerfectValue(X)) * 6)
                    Dim lNewTop As Int32 = lBaseTop + lblValue(X).Top + 2
                    If myPerfectValue(X) <> 0 Then lNewLeft -= 4 Else lNewLeft -= 6

                    Dim rcSrc As Rectangle = New Rectangle(195, 168, 5, 5) ' Rectangle(27, 237, 9, 6)
                    Dim rcDest As Rectangle = New Rectangle(lNewLeft, lNewTop, 5, 5) 'Rectangle(lNewLeft, lNewTop, 9, 6)
                    With MyBase.moUILib.Pen
                        .Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(lNewLeft, lNewTop), System.Drawing.Color.FromArgb(lColor, 255, 255, 255))
                    End With
                End If
            Next X
            'MyBase.moUILib.Pen.End()
            MyBase.moUILib.EndPenSprite()
        End If
    End Sub

    Private Sub frmMinDetail_OnRenderEnd() Handles Me.OnRenderEnd

        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

        'Now, the other items
        If lblValue Is Nothing = False Then
            'MyBase.moUILib.Pen.Begin(SpriteFlags.AlphaBlend)
            MyBase.moUILib.BeginPenSprite(SpriteFlags.AlphaBlend)
            For X As Int32 = 0 To lblValue.GetUpperBound(0)
                lblValue(X).ManualDrawSprite(MyBase.moUILib.Pen)
                'lblKnown(X).ManualDrawSprite(MyBase.moUILib.Pen)

				If myPerfectValue(X) <> 255 Then
					'each box is 5x5
					'there is a 2 pixel left side
					'1 pixel between each box

					If myPerfectValue(X) > 10 Then myPerfectValue(X) = 10
					Dim lNewLeft As Int32 = lblValue(X).Left + (CInt(myPerfectValue(X)) * 6)
					Dim lNewTop As Int32 = lblValue(X).Top - 3
                    'If myPerfectValue(X) <> 0 Then lNewLeft -= 6 Else lNewLeft -= 8
                    lNewLeft -= 6

                    Dim rcSrc As Rectangle = New Rectangle(27, 237, 9, 6)
                    Dim rcDest As Rectangle = New Rectangle(lNewLeft, lNewTop, 9, 6)
					With MyBase.moUILib.Pen
                        .Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(lNewLeft, lNewTop), Color.White)
                    End With
                End If

                If myHighlighted(X) <> 0 Then
                    Dim rcSrc As Rectangle
                    Dim rcDest As Rectangle = New Rectangle(lblProps(X).Left, lblProps(X).Top, lblProps(X).Width, lblProps(X).Height)

                    Dim fX As Single
                    Dim fY As Single
                    Dim ptLoc As Point = lblProps(X).GetAbsolutePosition

                    Dim clrHiLite As System.Drawing.Color = System.Drawing.Color.FromArgb(128, 220, 220, 255)
                    If myHighlighted(X) = 2 Then
                        clrHiLite = System.Drawing.Color.FromArgb(128, 0, 255, 0)
                    ElseIf myHighlighted(X) = 3 Then
						clrHiLite = System.Drawing.Color.FromArgb(192, 255, 0, 0)
					ElseIf myHighlighted(X) = 4 Then
						clrHiLite = System.Drawing.Color.FromArgb(192, 255, 255, 0)
					End If

                    If rcDest.Width = 0 OrElse rcDest.Height = 0 Then Exit Sub

                    'rcSrc = System.Drawing.Rectangle.FromLTRB(225, 0, 255, 30)
                    rcSrc.Location = New Point(192, 0)
                    rcSrc.Width = 62
                    rcSrc.Height = 64

                    'Now, draw it...
                    With MyBase.moUILib.Pen
                        fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
                        fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))
                        .Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrHiLite)
                    End With
				End If
			Next X
            'MyBase.moUILib.Pen.End()
            MyBase.moUILib.EndPenSprite()

            Using oSpr As New Sprite(MyBase.moUILib.oDevice)
                oSpr.Begin(SpriteFlags.AlphaBlend)
                For X As Int32 = 0 To lblProps.GetUpperBound(0)
                    If lblProps(X).Visible = True Then
                        lblProps(X).ManualDrawText(oSpr)
                    End If
                Next X

                oSpr.End()
            End Using


        End If

    End Sub

    Private Sub lstMinerals_ItemClick(ByVal lIndex As Integer) Handles lstMinerals.ItemClick
		If lIndex <> -1 Then
			mlCurrentMineralID = -1
            ShowMineralDetail(Me.Left, Me.Top, Me.Height, lstMinerals.ItemData(lIndex)) '

            If bRaiseSelectEvent = True AndAlso btnSelect.Visible = False Then
                RaiseEvent MineralSelected(lstMinerals.ItemData(lIndex))
                bRaiseSelectEvent = False
            End If
		End If
    End Sub

    Private Sub FillMineralList()
        On Error Resume Next

        lstMinerals.Clear()

        Dim lSorted() As Int32
        Dim lSortedUB As Int32 = -1
        Dim lSortedVals() As Int32
        ReDim lSorted(glMineralUB)
        ReDim lSortedVals(glMineralUB)

        For X As Int32 = 0 To lSorted.GetUpperBound(0)
            lSorted(X) = -1
            lSortedVals(X) = Int32.MaxValue
        Next X

		Dim bFilter As Boolean = chkFilterArchived.Value

        For X As Int32 = 0 To glMineralUB
			If glMineralIdx(X) <> -1 AndAlso goMinerals(X).bDiscovered = True Then

				If bFilter = False OrElse goMinerals(X).bArchived = False Then
                    Dim lIdx As Int32 = -1

                    'determine our mineral score
                    Dim lDiff As Int32 = 0
                    If chkSortByMatch.Value = True Then
                        For Y As Int32 = 0 To glMineralPropertyUB
                            If myPerfectValue(Y) <> 255 Then
                                Dim lMinVal As Int32 = goMinerals(X).MinPropValueScore(goMinerals(X).GetPropertyIndex(goMineralProperty(Y).ObjectID))
                                lMinVal = Math.Abs(lMinVal - CInt(myPerfectValue(Y)))
                                lMinVal *= lMinVal
                                lDiff += lMinVal
                            End If
                        Next Y

                        For Y As Int32 = 0 To lSortedUB
                            If lSortedVals(Y) > lDiff Then
                                lIdx = Y
                                Exit For
                            ElseIf lSortedVals(Y) = lDiff Then
                                If goMinerals(lSorted(Y)).MineralName.ToUpper > goMinerals(X).MineralName.ToUpper Then
                                    'Ok, found a place...
                                    lIdx = Y
                                    Exit For
                                End If
                            End If
                        Next Y
                    Else
                        For Y As Int32 = 0 To lSortedUB
                            If goMinerals(lSorted(Y)).MineralName.ToUpper > goMinerals(X).MineralName.ToUpper Then
                                'Ok, found a place...
                                lIdx = Y
                                Exit For
                            End If
                        Next Y
                    End If

					If lIdx = -1 Then
						lSortedUB += 1
						lIdx = lSortedUB
					Else
                        'need to shift
                        For Y As Int32 = lSortedUB To lIdx Step -1
                            lSorted(Y + 1) = lSorted(Y)
                            lSortedVals(Y + 1) = lSortedVals(Y)
                        Next Y
                        lSortedUB += 1
					End If
                    lSorted(lIdx) = X
                    lSortedVals(lIdx) = lDiff
				End If

			End If
        Next X

        For X As Int32 = 0 To lSorted.GetUpperBound(0)
            If lSorted(X) <> -1 Then
                lstMinerals.AddItem(goMinerals(lSorted(X)).MineralName)
                lstMinerals.ItemData(lstMinerals.NewIndex) = goMinerals(lSorted(X)).ObjectID
            End If
        Next X

        'lstMinerals.ListIndex = 0

        'For X As Int32 = 0 To glMineralUB
        '    If glMineralIdx(X) <> -1 AndAlso goMinerals(X).bDiscovered = True Then
        '        lstMinerals.AddItem(goMinerals(X).MineralName)
        '        lstMinerals.ItemData(lstMinerals.NewIndex) = goMinerals(X).ObjectID
        '    End If
        'Next X
    End Sub

    Public Sub ForceHighlightFirstItem()
        If lstMinerals.ListCount > 0 Then lstMinerals.ListIndex = 0
    End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eMinPropHlpr)
	End Sub

	Public Sub ClearHighlights()
		If myHighlighted Is Nothing = False Then
			For X As Int32 = 0 To myHighlighted.GetUpperBound(0)
				myHighlighted(X) = 0
				myPerfectValue(X) = 255
            Next X
            Me.IsDirty = True
		End If
	End Sub

	Public Sub HighlightProperty(ByVal lPropID As Int32, ByVal ySetting As Byte, ByVal yPerfectValue As Byte)
		Try
			For X As Int32 = 0 To glMineralPropertyUB
				If glMineralPropertyIdx(X) = lPropID Then
					myHighlighted(X) = ySetting
					myPerfectValue(X) = yPerfectValue
					Exit For
				End If
			Next X
			Me.IsDirty = True
		Catch
        End Try
        FillMineralList()
	End Sub

	Private Sub btnAvailResources_Click(ByVal sName As String) Handles btnAvailResources.Click

		Dim lIndex As Int32 = -1
		If goCurrentEnvir Is Nothing = False Then
			For X As Int32 = 0 To goCurrentEnvir.lEntityUB
				If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eResearch AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
					lIndex = X
					Exit For
				End If
			Next X
		End If
		If lIndex = -1 Then Return

		Dim ofrm As frmAvailRes = CType(MyBase.moUILib.GetWindow("frmAvailRes"), frmAvailRes)
		If ofrm Is Nothing Then ofrm = New frmAvailRes(MyBase.moUILib, False, CType(goCurrentEnvir.oEntity(lIndex), Base_GUID))
		ofrm.Visible = True
	End Sub

	Private Sub chkFilterArchived_Click() Handles chkFilterArchived.Click
		MyBase.moUILib.GetMsgSys.LoadArchived()
		FillMineralList()
	End Sub

    Private Sub chkSortByMatch_Click() Handles chkSortByMatch.Click
        FillMineralList()
    End Sub

    Public Sub SetAsHullBuilder()
        btnSelect.Visible = True
    End Sub
    Private Sub btnSelect_Click(ByVal sName As String) Handles btnSelect.Click
        If lstMinerals.ListIndex > -1 Then
            If bRaiseSelectEvent = True Then
                RaiseEvent MineralSelected(lstMinerals.ItemData(lstMinerals.ListIndex))
                bRaiseSelectEvent = False
            End If
            MyBase.moUILib.RemoveWindow(Me.ControlName)
        Else
            MyBase.moUILib.AddNotification("Select a mineral in the list first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If 
    End Sub
End Class
