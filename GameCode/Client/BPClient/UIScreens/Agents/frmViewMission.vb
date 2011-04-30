Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmViewMission
	Inherits UIWindow

	Private lblTitle As UILabel
	Private lblMission As UILabel
	Private lblTarget As UILabel
    Private lblMethod As UILabel
    Private lblInfiltrationType As UILabel
	Private lnDiv1 As UILine
	Private lblMethodName As UILabel
	Private lblTargetValue As UILabel
    Private lblMissionName As UILabel
    Private lblInfiltrationName As UILabel

	Private WithEvents btnClose As UIButton
	Private WithEvents btnHelp As UIButton
	Private WithEvents btnCancel As UIButton
	Private WithEvents btnPause As UIButton
	Private WithEvents mfraGoalList As fraGoalList

	Private mlLastRequest As Int32 = -1000000
	Private Const ml_REQUEST_DELAY As Int32 = 300			'mission updates are every 30 seconds... we'll do an update request every 10

    Private moPM As PlayerMission = Nothing
    Private mbLoading As Boolean = True
    Private mbPausing As Boolean = False

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmViewMission initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eViewMission
            .ControlName = "frmViewMission"

            .Width = 800
            .Height = 600
            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1
            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.AgentMissionDetailsX
                lTop = muSettings.AgentMissionDetailsY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 300
            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Left = lLeft
            .Top = lTop

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 3
			.Width = 175
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Mission Management"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'lblMission initial props
		lblMission = New UILabel(oUILib)
		With lblMission
			.ControlName = "lblMission"
			.Left = 10
			.Top = 35
			.Width = 70
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Mission:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMission, UIControl))

		'lblTarget initial props
		lblTarget = New UILabel(oUILib)
		With lblTarget
			.ControlName = "lblTarget"
            .Left = 300 '310
			.Top = 35
			.Width = 70
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Target:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTarget, UIControl))

		'lblMethod initial props
		lblMethod = New UILabel(oUILib)
		With lblMethod
			.ControlName = "lblMethod"
            .Left = 590 '580
			.Top = 35
			.Width = 70
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Method:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMethod, UIControl))

		'lblMethodName initial props
		lblMethodName = New UILabel(oUILib)
		With lblMethodName
			.ControlName = "lblMethodName"
			.Left = 650
			.Top = 35
			.Width = 110
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Seductively"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMethodName, UIControl))

		'lnDiv1 initial props
		lnDiv1 = New UILine(oUILib)
		With lnDiv1
			.ControlName = "lnDiv1"
			.Left = 1
			.Top = 27
			.Width = 799
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

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

		'lblTargetValue initial props
		lblTargetValue = New UILabel(oUILib)
		With lblTargetValue
			.ControlName = "lblTargetValue"
            .Left = 355 '380
			.Top = 35
            .Width = 230 '200
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Target List goes here...."
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTargetValue, UIControl))

		'lblMissionName initial props
		lblMissionName = New UILabel(oUILib)
		With lblMissionName
			.ControlName = "lblMissionName"
            .Left = 70 '80
			.Top = 35
			.Width = 230
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Mission Value goes here..."
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMissionName, UIControl))

		'btnCancel initial props
		btnCancel = New UIButton(oUILib)
		With btnCancel
			.ControlName = "btnCancel"
			.Left = 675
			.Top = 570
			.Width = 120
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Abort Mission"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
			.ToolTipText = "Abort the mission. Missions can only be aborted during the preparation phase."
		End With
		Me.AddChild(CType(btnCancel, UIControl))

        lblInfiltrationType = New UILabel(oUILib)
        With lblInfiltrationType
            .ControlName = "lblInfiltrationType"
            .Left = 15
            .Top = 570
            .Width = 160
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Required Infiltration:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblInfiltrationType, UIControl))

        'lblInfiltrationName initial props
        lblInfiltrationName = New UILabel(oUILib)
        With lblInfiltrationName
            .ControlName = "lblInfiltrationName"
            .Left = 160
            .Top = 570
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Infiltration type goes here...."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblInfiltrationName, UIControl))

		'btnPause initial props
		btnPause = New UIButton(oUILib)
		With btnPause
			.ControlName = "btnPause"
			.Left = 550
			.Top = 570
			.Width = 120
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Pause Mission"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnPause, UIControl))

		mfraGoalList = New fraGoalList(oUILib)
		With mfraGoalList
			.ControlName = "mfraGoalList"
			.Left = 10
			.Top = 65
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
			.bRoundedBorder = True
		End With
		Me.AddChild(CType(mfraGoalList, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
	End Sub

	Public Sub SetFromMission(ByRef oPM As PlayerMission)
		moPM = oPM
        lblMissionName.Caption = oPM.oMission.sMissionName
        lblInfiltrationName.Caption = frmAgent.fraInfiltration.GetInfTypeText(oPM.oMission.lInfiltrationType)
        For X As Int32 = 0 To glMissionMethodUB
            If glMissionMethodIdx(X) = oPM.lMethodID Then
                lblMethodName.Caption = gsMissionMethods(X)
                Exit For
            End If
        Next X
		mfraGoalList.SetPlayerMission(moPM)
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub frmViewMission_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        'New Point(lblMethodName.Left + lblMethodName.Width + 5, lblMethodName.Top + lblMethodName.Height + 5)
        Dim pt As Point = Me.GetAbsolutePosition()
        Dim lTmpX As Int32 = lMouseX - pt.X
        Dim lTmpY As Int32 = lMouseY - pt.Y

        Dim lLeft As Int32 = (lblMethodName.Left + lblMethodName.Width + 5)
        Dim lTop As Int32 = (lblMethodName.Top + lblMethodName.Height + 5)
        If lTmpX > lLeft AndAlso lTmpX < lLeft + 16 AndAlso lTmpY > lTop AndAlso lTmpY < lTop + 16 Then
            If moPM Is Nothing = False Then
                If moPM.bAlarmThrown = True Then
                    MyBase.moUILib.SetToolTip("Alarm has been thrown. Penalties are" & vbCrLf & "applied to all tests including infiltration.", lMouseX, lMouseY)
                Else
                    MyBase.moUILib.SetToolTip("Alarm has not been thrown. Your team is undetected.", lMouseX, lMouseY)
                End If
            End If
        End If
    End Sub

	Private Sub frmViewMission_OnNewFrame() Handles Me.OnNewFrame
		If moPM Is Nothing = False Then

			If glCurrentCycle - mlLastRequest > ml_REQUEST_DELAY Then
				mlLastRequest = glCurrentCycle
				If moPM Is Nothing = False Then
					Dim yMsg(5) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eGetPMUpdate).CopyTo(yMsg, 0)
					System.BitConverter.GetBytes(moPM.PM_ID).CopyTo(yMsg, 2)
					MyBase.moUILib.SendMsgToPrimary(yMsg)
				End If
			End If

			Dim sText As String = ""
			If moPM.lTargetPlayerID > -1 Then
				sText = GetCacheObjectValue(moPM.lTargetPlayerID, ObjectType.ePlayer)
			End If
            If moPM.lTargetID > -1 AndAlso moPM.iTargetTypeID > -1 Then
                If moPM.iTargetTypeID = 0 AndAlso moPM.oMission Is Nothing = False AndAlso moPM.oMission.ObjectID = eMissionResult.eGetFacilityList Then
                    If sText <> "" Then sText &= " - "
                    'mopm.oMission.objectid
                    Select Case moPM.lTargetID
                        Case ProductionType.eEnlisted
                            sText &= "Barracks"
                        Case ProductionType.eCommandCenterSpecial
                            sText &= "Command Center"
                        Case ProductionType.eProduction
                            sText &= "Factory"
                        Case ProductionType.eProduction
                            sText &= "Land Production"
                        Case ProductionType.eMining
                            sText &= "Mining"
                        Case ProductionType.eNavalProduction
                            sText &= "Naval Production"
                        Case ProductionType.eOfficers
                            sText &= "Officers"
                        Case ProductionType.ePowerCenter
                            sText &= "Power Center"
                        Case ProductionType.eRefining
                            sText &= "Refining"
                        Case ProductionType.eResearch
                            sText &= "Research"
                        Case ProductionType.eColonists
                            sText &= "Residence"
                        Case ProductionType.eAerialProduction
                            sText &= "Spaceport"
                        Case ProductionType.eSpaceStationSpecial
                            sText &= "Space Station"
                        Case ProductionType.eTradePost
                            sText &= "Tradepost"
                        Case ProductionType.eWareHouse
                            sText &= "Warehouse"
                        Case Else
                            sText &= "Unknown"
                    End Select
                ElseIf moPM.iTargetTypeID = 0 Then
                    'sText &= "Any"
                Else
                    If sText <> "" Then sText &= " - "
                    sText &= GetCacheObjectValue(moPM.lTargetID, moPM.iTargetTypeID)
                    For X As Int32 = 0 To glItemIntelUB
                        If glItemIntelIdx(X) <> -1 AndAlso goItemIntel(X).lOtherPlayerID = moPM.lTargetPlayerID AndAlso goItemIntel(X).iItemTypeID = moPM.iTargetTypeID AndAlso goItemIntel(X).lItemID = moPM.lTargetID AndAlso goItemIntel(X).EnvirTypeID > 0 AndAlso goItemIntel(X).EnvirID > 0 Then
                            sText &= " - " & GetCacheObjectValue(goItemIntel(X).EnvirID, goItemIntel(X).EnvirTypeID)
                        End If
                    Next X
                End If
            End If
			If moPM.lTargetID2 > -1 AndAlso moPM.iTargetTypeID2 > -1 Then
				If sText <> "" Then sText &= " - "
				sText &= GetCacheObjectValue(moPM.lTargetID2, moPM.iTargetTypeID2)
            End If
			If lblTargetValue.Caption <> sText Then lblTargetValue.Caption = sText

            Dim lPhase As Int32 = moPM.yCurrentPhase
            If mbPausing = False Then

                'If (lPhase And (eMissionPhase.eCancelled Or eMissionPhase.eMissionOverFailure Or eMissionPhase.eMissionOverSuccess)) <> 0 Then
                If lPhase > eMissionPhase.eWaitingToExecute AndAlso lPhase < eMissionPhase.eMissionPaused Then
                    sText = "Repeat Mission"
                Else
                    If (moPM.yCurrentPhase And eMissionPhase.eMissionPaused) <> 0 Then
                        sText = "Resume Mission"
                        lPhase = lPhase Xor eMissionPhase.eMissionPaused
                    Else : sText = "Pause Mission"
                    End If
                End If

                If btnPause.Caption <> sText Then btnPause.Caption = sText
            End If

            'If lPhase > eMissionPhase.ePreparationTime AndAlso lPhase <> eMissionPhase.eWaitingToExecute Then
            '	If btnCancel.Enabled = True Then btnCancel.Enabled = False
            'ElseIf btnCancel.Enabled = False Then
            '	btnCancel.Enabled = True
            'End If
            'If btnCancel.Enabled = False Then btnCancel.Enabled = True
            If lPhase >= eMissionPhase.ePreparationTime AndAlso lPhase <= eMissionPhase.eWaitingToExecute Then
                If btnCancel.Enabled = False Then btnCancel.Enabled = True
            ElseIf btnCancel.Enabled = True Then
                btnCancel.Enabled = False
            End If


            mfraGoalList.fraGoalList_OnNewFrame()
        End If

	End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		If btnCancel.Caption.ToUpper = "CONFIRM" Then
			btnCancel.Enabled = False
			Dim yMsg(9) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eSetSkipStatus).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(moPM.PM_ID).CopyTo(yMsg, 2)
			System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 6)
			MyBase.moUILib.SendMsgToPrimary(yMsg)

			MyBase.moUILib.RemoveWindow(Me.ControlName)
		Else : btnCancel.Caption = "Confirm"
		End If
	End Sub

	Private Sub btnPause_Click(ByVal sName As String) Handles btnPause.Click
        If btnPause.Caption.ToUpper = "CONFIRM" OrElse btnPause.Caption = "Repeat Mission" Then
            mbPausing = False
            btnPause.Enabled = False

            Dim lPhase As Int32 = moPM.yCurrentPhase
            If lPhase = eMissionPhase.eCancelled OrElse lPhase = eMissionPhase.eMissionOverFailure OrElse lPhase = eMissionPhase.eMissionOverSuccess Then
                Dim oFrm As New frmMission(goUILib)
                oFrm.SetFromMission(moPM, True)
                oFrm = Nothing
                MyBase.moUILib.RemoveWindow(Me.ControlName)
            Else
                Dim yMsg(9) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eSetSkipStatus).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(moPM.PM_ID).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(-2I).CopyTo(yMsg, 6)
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
            btnPause.Enabled = True
        Else
            mbPausing = True
            btnPause.Caption = "Confirm"
        End If
	End Sub

    Private Sub frmViewMission_OnRenderEnd() Handles Me.OnRenderEnd
        Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        If moPM.bAlarmThrown = True Then
            clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        End If
        Dim ptLoc As Point = New Point(lblMethodName.Left + lblMethodName.Width + 5, lblMethodName.Top + lblMethodName.Height + 5)
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        Using oSprite As New Sprite(MyBase.moUILib.oDevice)
            oSprite.Begin(SpriteFlags.AlphaBlend)
            Try
                oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eSphere), New Rectangle(0, 0, 16, 16), Point.Empty, 0, ptLoc, clrVal)
            Catch
            End Try
            oSprite.End()
        End Using
    End Sub

    Private Sub frmViewMission_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.AgentMissionDetailsX = Me.Left
            muSettings.AgentMissionDetailsY = Me.Top
        End If
    End Sub
End Class