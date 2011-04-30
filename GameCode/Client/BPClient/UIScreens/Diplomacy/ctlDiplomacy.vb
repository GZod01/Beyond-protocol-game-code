Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class ctlDiplomacy
    Inherits UIWindow

    Public Const ml_ITEM_HEIGHT As Int32 = 71

    Private fraPlayerIcon As UIWindow
    Private lblPlayerName As UILabel
    Private lblGuild As UILabel
	'Private lblSenate As UILabel
    Private lblHomeworld As UILabel
    Private lblThreat As UILabel
    Public lblThreatVal As UILabel
	Private lblRelationship As UILabel 

	Private WithEvents btnPM As UIButton
    Private WithEvents btnEmail As UIButton
    Private lblActualTitle As UILabel

    Private mbSelected As Boolean = False

    Public Shared rcRelBarLoc As Rectangle = Rectangle.FromLTRB(0, 0, 255, 15)
    Public Shared rcRelBarSrc As Rectangle = Rectangle.FromLTRB(0, 0, 255, 15)
    Public Shared ptRelBarLoc As Point = New Point(450, 25)

    Public rcBack As Rectangle
    Public clrBack As Color
    Public rcFore1 As Rectangle
    Public clrFore1 As Color
    Public rcFore2 As Rectangle
    Public clrFore2 As Color

    Public Event ItemClicked(ByVal lPlayerID As Int32)
    Public Event ItemRelChanged(ByVal lPlayerID As Int32, ByVal lRel As Int32)

    Private moRel As PlayerRel = Nothing

    Public yMyRelScore As Byte
    Private mySetRelAfterRender As Byte = 0
    Public yTheirRelScore As Byte

    Private mbHasUnknowns As Boolean = True

    Private mlLastUpdate As Int32 = Int32.MinValue

    Public Shared moSprite As Sprite = Nothing
    Public Shared moRelBarImg As Texture = Nothing
    Public Shared moIconBack As Texture = Nothing
    Public Shared moIconFore As Texture = Nothing

    Private mbMouseDown As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'ctlDiplomacy initial props
        With Me
            .ControlName = "ctlDiplomacy"
            .Left = 0
            .Top = 0
            .Width = 705
            .Height = 71
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .BorderLineWidth = 1
        End With

        'fraPlayerIcon initial props
        fraPlayerIcon = New UIWindow(oUILib)
        With fraPlayerIcon
            .ControlName = "fraPlayerIcon"
            .Left = 2
            .Top = 2
            .Width = 67
            .Height = 67
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .BorderLineWidth = 1
        End With
		Me.AddChild(CType(fraPlayerIcon, UIControl))

		'btnPM initial props
		btnPM = New UIButton(oUILib)
		With btnPM
			.ControlName = "btnPM"
			.Left = 72
			.Top = 48
			.Width = 24
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "PM"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnPM, UIControl))

		'btnEmail initial props
		btnEmail = New UIButton(oUILib)
		With btnEmail
			.ControlName = "btnEmail"
			.Left = 99
			.Top = 48
			.Width = 24
			.Height = 17
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eEmailIcon)
			.ControlImageRect_Pressed = .ControlImageRect
			.ControlImageRect_Disabled = .ControlImageRect
			.ControlImageRect_Normal = .ControlImageRect
		End With
		Me.AddChild(CType(btnEmail, UIControl))

        'lblPlayerName initial props
        lblPlayerName = New UILabel(oUILib)
        With lblPlayerName
            .ControlName = "lblPlayerName"
            .Left = 72
            .Top = 2
            .Width = 275
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Magistrate Enoch Dagor and Company"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPlayerName, UIControl))

        'lblGuild initial props
        lblGuild = New UILabel(oUILib)
        With lblGuild
            .ControlName = "lblGuild"
            .Left = 72
            .Top = 24
            .Width = 275
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Magistrate Enoch Dagor and Company"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblGuild, UIControl))

        'lblActualTitle initial props
        lblActualTitle = New UILabel(oUILib)
        With lblActualTitle
            .ControlName = "lblActualTitle"
            .Left = btnEmail.Left + btnEmail.Width + 5
            .Top = 48
            .Width = lblGuild.Width - lblGuild.Left
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Senate Title: "
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblActualTitle, UIControl))

        'lblHomeworld initial props
        lblHomeworld = New UILabel(oUILib)
        With lblHomeworld
            .ControlName = "lblHomeworld"
            .Left = 350
            .Top = 2
            .Width = 205
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Homeworld: Unknown"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHomeworld, UIControl))

        'lblThreat initial props
        lblThreat = New UILabel(oUILib)
        With lblThreat
            .ControlName = "lblThreat"
            .Left = 555
            .Top = 2
            .Width = 45
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Threat:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblThreat, UIControl))

        'lblThreatVal initial props
        lblThreatVal = New UILabel(oUILib)
        With lblThreatVal
            .ControlName = "lblThreatVal"
            .Left = 600
            .Top = 5
            .Width = 100
            .Height = 12
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(4, DrawTextFormat)
            .ControlImageRect = New Rectangle(192, 211, 64, 9)
        End With
        Me.AddChild(CType(lblThreatVal, UIControl))

        'lblRelationship initial props
        lblRelationship = New UILabel(oUILib)
        With lblRelationship
            .ControlName = "lblRelationship"
            .Left = 350
            .Top = 24
            .Width = 80
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Relationship:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblRelationship, UIControl))

        Device.IsUsingEventHandlers = False
        If moRelBarImg Is Nothing OrElse moRelBarImg.Disposed = True Then moRelBarImg = goResMgr.GetTexture("Diplomacy.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
        If moIconBack Is Nothing OrElse moIconBack.Disposed = True Then moIconBack = goResMgr.GetTexture("DipPlayerBack.bmp", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
        If moIconFore Is Nothing OrElse moIconFore.Disposed = True Then moIconFore = goResMgr.GetTexture("DipPlayerFore.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
        If moSprite Is Nothing OrElse moSprite.Disposed = True Then moSprite = New Sprite(oUILib.oDevice)
        Device.IsUsingEventHandlers = True
    End Sub

    Public Property Selected() As Boolean
        Get
            Return mbSelected
        End Get
        Set(ByVal value As Boolean)
            If mbSelected <> value Then
                Dim lFontClr As Color
                If value = True Then
                    Me.FillColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
                    lFontClr = Color.FromArgb(255, 255 - FillColor.R, 255 - FillColor.G, 255 - FillColor.B)
                Else
                    Me.FillColor = muSettings.InterfaceFillColor
                    lFontClr = muSettings.InterfaceBorderColor
                End If

                lblPlayerName.ForeColor = lFontClr
                lblGuild.ForeColor = lFontClr
				'lblSenate.ForeColor = lFontClr
                lblHomeworld.ForeColor = lFontClr
                lblThreat.ForeColor = lFontClr
                lblRelationship.ForeColor = lFontClr
            End If
            mbSelected = value
        End Set
    End Property

    Public Sub ClearTempScore()
        mySetRelAfterRender = 0
    End Sub
    Public Sub SetValues(ByRef oPlayerRel As PlayerRel)
        moRel = oPlayerRel
        mbHasUnknowns = True
        mlLastUpdate = Int32.MinValue

        If moRel Is Nothing = False Then

            If moRel.lPlayerRegards = glPlayerID Then
                Me.yMyRelScore = moRel.WithThisScore
                'If mySetRelAfterRender <> 0 Then
                '    Me.yMyRelScore = mySetRelAfterRender
                '    'mySetRelAfterRender = 0
                'End If
            End If

            Dim oIntel As PlayerIntel = moRel.oPlayerIntel
            Dim yBackImg As Byte
            Dim yBackClr As Byte
            Dim yFore1Img As Byte
            Dim yFore1Clr As Byte
            Dim yFore2Img As Byte
            Dim yFore2Clr As Byte

            If oIntel Is Nothing = False Then

                PlayerIconManager.FillIconValues(oIntel.lPlayerIcon, yBackImg, yBackClr, yFore1Img, yFore1Clr, yFore2Img, yFore2Clr)

                rcBack = PlayerIconManager.ReturnImageRectangle(yBackImg, True)
                rcFore1 = PlayerIconManager.ReturnImageRectangle(yFore1Img, False)
                rcFore2 = PlayerIconManager.ReturnImageRectangle(yFore2Img, False)

                clrBack = PlayerIconManager.GetColorValue(yBackClr)
                clrFore1 = PlayerIconManager.GetColorValue(yFore1Clr)
                clrFore2 = PlayerIconManager.GetColorValue(yFore2Clr)
            End If
        End If
    End Sub

    Public Function HasUnknowns() As Boolean
        If moRel Is Nothing Then Return False
        Dim oPlayerIntel As PlayerIntel = moRel.oPlayerIntel
        If oPlayerIntel Is Nothing Then Return False

        If mbHasUnknowns = True OrElse mlLastUpdate <> oPlayerIntel.lLastUpdate Then
            mbHasUnknowns = False
            With oPlayerIntel
                mlLastUpdate = .lLastUpdate

                yTheirRelScore = .yRegardsCurrentPlayer

                If .PlayerName Is Nothing OrElse .PlayerName.ToUpper = "UNKNOWN" Then
                    .PlayerName = GetCacheObjectValue(.ObjectID, ObjectType.ePlayer)
                    mbHasUnknowns = True
                End If
                Dim yTitleToUse As Byte = .yCustomTitle
                If yTitleToUse = 0 Then yTitleToUse = .yPlayerTitle
                Dim sTemp As String = Player.GetPlayerNameWithTitle(yTitleToUse, .PlayerName, .bIsMale)
                If lblPlayerName.Caption <> sTemp Then lblPlayerName.Caption = sTemp

                If .lGuildID < 1 Then
                    sTemp = "Not In A Guild/Alliance"
                Else
                    sTemp = "Guild: " & GetCacheObjectValue(.lGuildID, ObjectType.eGuild)
                    If sTemp.ToUpper = "GUILD: UNKNOWN" OrElse sTemp = "" Then mbHasUnknowns = True
                End If
                If lblGuild.Caption <> sTemp Then lblGuild.Caption = sTemp

                sTemp = "Senate Title: " & Player.GetPlayerNameWithTitle(.yPlayerTitle, "", .bIsMale) & " (" & .lTotalVoteStrength '& " votes)"
                If .lTotalVoteStrength <> 1 Then sTemp &= " votes)" Else sTemp &= " vote)"
                If lblActualTitle.Caption <> sTemp Then lblActualTitle.Caption = sTemp

                If .lStartID < 1 OrElse .iStartTypeID < 1 Then
                    sTemp = "Homeworld: Unknown"
                Else
                    sTemp = "Homeworld: " & GetCacheObjectValue(.lStartID, .iStartTypeID)
                    If sTemp.ToUpper = "HOMEWORLD: UNKNOWN" Then mbHasUnknowns = True
                    sTemp = GetCacheObjectValue(.lStartID, .iStartTypeID)
                End If
                If lblHomeworld.Caption <> sTemp Then lblHomeworld.Caption = sTemp

                'Dim fTemp As Single = CSng(.lTotalScore / (goCurrentPlayer.lTotalScore + .lTotalScore)) - 0.5F

                Dim lModTheirs As Int32 = .lTotalScore
                Dim lModMine As Int32 = goCurrentPlayer.lTotalScore

                If .lTotalScore < 0 OrElse goCurrentPlayer.lTotalScore < 0 Then
                    Dim lMod As Int32 = Math.Abs(Math.Min(lModTheirs, lModMine))
                    lModTheirs += lMod
                    lModMine += lMod
                End If

                Dim fTemp As Single = 0.0F
                If lModTheirs + lModMine = 0 Then
                    fTemp = 0.5F
                Else : fTemp = CSng(lModTheirs / (lModTheirs + lModMine))
                End If

                Dim lValue As Int32 = CInt(fTemp * 10)
                If lValue < 0 Then lValue = 0
                If lValue > 9 Then lValue = 9

                lblThreatVal.ControlImageRect = Rectangle.FromLTRB(192, 157 + ((lValue + 1) * 9), 256, 157 + ((lValue + 2) * 9))

                Select Case lValue
                    Case 0
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                        lblThreatVal.ToolTipText = "We would easily crush this empire."
                    Case 1
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 0, 128, 0)
                        lblThreatVal.ToolTipText = "This empire would pose little problems for us."
                    Case 2
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                        lblThreatVal.ToolTipText = "This empire may be a nuisance but it is doubtful" & vbCrLf & "that they could do any real harm."
                    Case 3
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
                        lblThreatVal.ToolTipText = "This empire may prove difficult to defeat."
                    Case 4
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 0, 0, 255)
                        lblThreatVal.ToolTipText = "We are only slightly more powerful than this empire."
                    Case 5
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 255, 255, 0)
                        lblThreatVal.ToolTipText = "This empire is only slightly more powerful than us."
                    Case 6
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 255, 128, 0)
                        lblThreatVal.ToolTipText = "Defeating this empire would be a difficult task."
                    Case 7
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 128, 0, 0)
                        lblThreatVal.ToolTipText = "We would suffer heavy losses in a conflict with this empire."
                    Case 8
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                        lblThreatVal.ToolTipText = "We pale greatly in comparison to this empire."
                    Case Else
                        lblThreatVal.BackImageColor = System.Drawing.Color.FromArgb(255, 255, 0, 255)
                        lblThreatVal.ToolTipText = "This empire would easily crush us."
                End Select

            End With
        End If

        Return Not mbHasUnknowns
    End Function

    Private Sub ctlBudgetDetailItem_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        mbMouseDown = True 
    End Sub

    Public Sub ctlDiplomacy_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        'If lButton = MouseButtons.Left AndAlso moRel Is Nothing = False AndAlso moRel.oPlayerIntel Is Nothing = False Then
        '    'ok, see if we are already scrolling...
        '    If MyBase.moUILib.lUISelectState <> UILib.eSelectState.eScrollBarDrag AndAlso mbMouseDown = True Then
        '        MyBase.moUILib.oScrollingRel = Me
        '        MyBase.moUILib.lUISelectState = UILib.eSelectState.eScrollBarDrag
        '        RaiseEvent ItemClicked(moRel.oPlayerIntel.ObjectID)
        '    End If

        '    If MyBase.moUILib.lUISelectState = UILib.eSelectState.eScrollBarDrag Then
        '        If Object.Equals(Me, MyBase.moUILib.oScrollingRel) = True Then
        '            Dim lAbsPosX As Int32 = Me.GetAbsolutePosition.X
        '            lAbsPosX += ptRelBarLoc.X
        '            lMouseX -= lAbsPosX

        '            Dim decValue As Decimal = 0
        '            If lMouseX < 10 Then
        '                decValue = 0D
        '            ElseIf lMouseX > rcRelBarLoc.Width + 10 Then
        '                decValue = 1D
        '            Else
        '                decValue = CDec(lMouseX / rcRelBarLoc.Width)
        '            End If

        '            decValue *= 254 '255 - 1
        '            decValue += 1

        '            Dim lNewVal As Int32 = CInt(decValue)
        '            If lNewVal < 1 Then lNewVal = 1
        '            If lNewVal > 255 Then lNewVal = 255

        '            mySetRelAfterRender = CByte(lNewVal)
        '            Me.IsDirty = True
        '            RaiseEvent ItemRelChanged(moRel.oPlayerIntel.ObjectID, lNewVal)

        '        End If
        '    End If

        'End If
    End Sub

    Private Sub ctlBudgetDetailItem_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseUp
        If mbMouseDown = True Then
            If moRel Is Nothing = False AndAlso moRel.oPlayerIntel Is Nothing = False Then
                RaiseEvent ItemClicked(moRel.oPlayerIntel.ObjectID)
            End If
        End If
        mbMouseDown = False
    End Sub
 
	Private Sub btnPM_Click(ByVal sName As String) Handles btnPM.Click
		If moRel Is Nothing Then Return
		Dim oPlayerIntel As PlayerIntel = moRel.oPlayerIntel
		If oPlayerIntel Is Nothing Then Return

		Dim oWin As frmChat = CType(MyBase.moUILib.GetWindow("frmChat"), frmChat)
		If oWin Is Nothing = False Then
			Dim sTmpName As String = GetCacheObjectValue(oPlayerIntel.ObjectID, ObjectType.ePlayer)
            oWin.SetNewMsgText("/pm " & sTmpName & " ")
			For lCtrlIdx As Int32 = 0 To oWin.ChildrenUB
				If oWin.moChildren(lCtrlIdx).ControlName = "txtNew" Then
					If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
					goUILib.FocusedControl = oWin.moChildren(lCtrlIdx)
					oWin.moChildren(lCtrlIdx).HasFocus = True
					Exit For
				End If
			Next
		End If
		oWin = Nothing
	End Sub

	Private Sub btnEmail_Click(ByVal sName As String) Handles btnEmail.Click
		If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
			MyBase.moUILib.AddNotification("You lack rights to create emails.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		If moRel Is Nothing Then Return
		Dim oPlayerIntel As PlayerIntel = moRel.oPlayerIntel
		If oPlayerIntel Is Nothing Then Return

		Dim ofrm As frmNewEmail = CType(MyBase.moUILib.GetWindow("frmNewEmail"), frmNewEmail)
		If ofrm Is Nothing = True Then ofrm = New frmNewEmail(goUILib)
		Dim sTmpName As String = GetCacheObjectValue(oPlayerIntel.ObjectID, ObjectType.ePlayer)
		ofrm.SetToText(sTmpName)
		ofrm = Nothing
	End Sub
End Class