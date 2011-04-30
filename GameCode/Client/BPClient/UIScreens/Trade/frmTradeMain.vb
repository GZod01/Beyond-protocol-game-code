Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmTradeMain
    Inherits UIWindow

    Private rcBuy As Rectangle = New Rectangle(5, 30, 100, 24)
    Private rcSell As Rectangle = New Rectangle(105, 30, 100, 24)
    Private rcDirect As Rectangle = New Rectangle(205, 30, 100, 24)
    'Private rcBuyOrder As Rectangle = New Rectangle(305, 30, 100, 24)
    'Private rcHistory As Rectangle = New Rectangle(405, 30, 100, 24)
    'Private rcInProgress As Rectangle = New Rectangle(505, 30, 115, 24)
    Private rcHistory As Rectangle = New Rectangle(305, 30, 100, 24)
    Private rcInProgress As Rectangle = New Rectangle(405, 30, 115, 24)

    Private myButtonState As Byte = 0

    Private WithEvents btnClose As UIButton

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblSelectTradepost As UILabel
    Private WithEvents cboTradePost As UIComboBox
    Private WithEvents btnHelp As UIButton
    Private lblEmpty As UILabel

    Private mlTradepostID As Int32 = -1

    Public fraContents As UIWindow

    Private mbLoading As Boolean = True

    Private mlLastForcefulRender As Int32 = 0

    Public Sub New(ByRef oUILib As UILib, ByVal lTradePostID As Int32)
        MyBase.New(oUILib)

        'frmTradeMain initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTradeMain
            .ControlName = "frmTradeMain"
            '.Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
            '.Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 300

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.TradeMainX
                lTop = muSettings.TradeMainY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 300
            If lLeft + 800 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 800
            If lTop + 600 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 600

            .Left = lLeft
            .Top = lTop

            .Width = 800
            .Height = 600
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .mbAcceptReprocessEvents = True
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

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 280
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "GALACTIC TRADE INTERFACE"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'NewControl5 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = Me.BorderLineWidth \ 2
            .Top = 25
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblSelectTradepost initial props
        lblSelectTradepost = New UILabel(oUILib)
        With lblSelectTradepost
            .ControlName = "lblSelectTradepost"
            .Left = lblTitle.Left + lblTitle.Width + 5
            .Top = 5
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select a Tradepost:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
        End With
        Me.AddChild(CType(lblSelectTradepost, UIControl))

        'cboTradePost initial props
        cboTradePost = New UIComboBox(oUILib)
        With cboTradePost
            .ControlName = "cboTradePost"
            .Left = lblSelectTradepost.Left + lblSelectTradepost.Width + 5
            .Top = 5
            .Width = btnHelp.Left - 15 - .Left
            .Height = 18
            .Enabled = True
            .Visible = True
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .l_ListBoxHeight = 256

            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboTradePost, UIControl))

        'lblEmpty initial props
        lblEmpty = New UILabel(oUILib)
        With lblEmpty
            .ControlName = "lblEmpty"
            .Left = Me.Width \ 2 - 200
            .Top = Me.Height \ 2 - 9
            .Width = 400
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "SELECT OR BUILD A TRADEPOST"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblEmpty, UIControl))

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If lTradePostID = -1 Then
            For X As Int32 = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(X) <> -1 AndAlso oEnvir.oEntity(X).OwnerID = glPlayerID AndAlso oEnvir.oEntity(X).yProductionType = ProductionType.eTradePost Then
                    lTradePostID = oEnvir.oEntity(X).ObjectID
                    Exit For
                End If
            Next X
        End If
        mlTradepostID = lTradePostID

        'clear our tradepost
        cboTradePost.Clear()

        If lTradePostID = -1 Then
            myButtonState = 3
        Else
            For X As Int32 = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(X) = lTradePostID AndAlso oEnvir.oEntity(X).ObjTypeID = ObjectType.eFacility Then
                    Dim sEnvirName As String = ""
                    If oEnvir.oGeoObject Is Nothing = False Then
                        Select Case CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID
                            Case ObjectType.ePlanet
                                sEnvirName = " (" & CType(oEnvir.oGeoObject, Planet).PlanetName & ")"
                            Case ObjectType.eSolarSystem
                                sEnvirName = " (" & CType(oEnvir.oGeoObject, SolarSystem).SystemName & ")"
                        End Select
                    End If
                    cboTradePost.AddItem(oEnvir.oEntity(X).EntityName & sEnvirName)
                    cboTradePost.ItemData(cboTradePost.NewIndex) = lTradePostID
                    cboTradePost.FindComboItemData(lTradePostID)
                    Exit For
                End If
            Next X
            myButtonState = 0
        End If

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewTrades) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack rights to view the trades interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        UpdateWindowContents()

        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetTradePostList).CopyTo(yMsg, 0)
        MyBase.moUILib.SendMsgToPrimary(yMsg)


		glCurrentEnvirView = CurrentView.eGTCScreen

		Dim oTmpWin As frmQuickBar = CType(MyBase.moUILib.GetWindow("frmQuickBar"), frmQuickBar)
		If oTmpWin Is Nothing = False Then
			oTmpWin.ClearFlashState(3)
		End If
		oTmpWin = Nothing

        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
        'goFullScreenBackground = goResMgr.GetTexture(".bmp", EpicaResourceManager.eGetTextureType.NoSpecifics, "TechBacks.pak")
        mbLoading = False
    End Sub

    Private Sub frmTradeMain_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        If cboTradePost Is Nothing Then Return

        Dim lTempX As Int32 = lMouseX
        Dim lTempY As Int32 = lMouseY
        With Me.GetAbsolutePosition()
            lTempX -= .X
            lTempY -= .Y
        End With

        Dim yNewState As Byte = myButtonState
        Dim bIgnoreMouseMove As Boolean = True
        If rcBuy.Contains(lTempX, lTempY) = True AndAlso cboTradePost.ListIndex <> -1 Then
            yNewState = 0
        ElseIf rcSell.Contains(lTempX, lTempY) = True AndAlso cboTradePost.ListIndex <> -1 Then
            yNewState = 1
        ElseIf rcDirect.Contains(lTempX, lTempY) = True AndAlso cboTradePost.ListIndex <> -1 Then
            yNewState = 2
        ElseIf rcHistory.Contains(lTempX, lTempY) = True Then
            yNewState = 3
            'ElseIf rcBuyOrder.Contains(lTempX, lTempY) = True Then
            '   yNewState = 4
        ElseIf rcInProgress.Contains(lTempX, lTempY) = True Then
            yNewState = 5
        Else
            bIgnoreMouseMove = False
        End If

        If yNewState <> myButtonState Then
            If bIgnoreMouseMove = True Then frmMain.mbIgnoreMouseMove = True 'Only ignore if button click is on existing button, not different one
            myButtonState = yNewState
            UpdateWindowContents()
            Me.IsDirty = True
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
        End If

    End Sub

    Private Sub frmTradeMain_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim lTempX As Int32 = lMouseX
        Dim lTempY As Int32 = lMouseY
        With Me.GetAbsolutePosition()
            lTempX -= .X
            lTempY -= .Y
        End With

        MyBase.moUILib.SetToolTip(False)

        If rcBuy.Contains(lTempX, lTempY) = True Then
            If cboTradePost.ListIndex <> -1 Then
                MyBase.moUILib.SetToolTip("Click to view the Buy Window.", lMouseX, lMouseY)
            Else : MyBase.moUILib.SetToolTip("Select a Tradepost first.", lMouseX, lMouseY)
            End If
        ElseIf rcSell.Contains(lTempX, lTempY) = True AndAlso cboTradePost.ListIndex <> -1 Then
            If cboTradePost.ListIndex <> -1 Then
                MyBase.moUILib.SetToolTip("Click to view the Sell Window.", lMouseX, lMouseY)
            Else : MyBase.moUILib.SetToolTip("Select a Tradepost first.", lMouseX, lMouseY)
            End If
        ElseIf rcDirect.Contains(lTempX, lTempY) = True AndAlso cboTradePost.ListIndex <> -1 Then
            If cboTradePost.ListIndex <> -1 Then
                MyBase.moUILib.SetToolTip("Click to view the Direct Trade Window.", lMouseX, lMouseY)
            Else : MyBase.moUILib.SetToolTip("Select a Tradepost first.", lMouseX, lMouseY)
            End If
        ElseIf rcHistory.Contains(lTempX, lTempY) = True Then
            MyBase.moUILib.SetToolTip("Click to view all trade history.", lMouseX, lMouseY)
            'ElseIf rcBuyOrder.Contains(lTempX, lTempY) = True Then
            '    If cboTradePost.ListIndex <> -1 Then
            '        MyBase.moUILib.SetToolTip("Click to view the Buy Order Window.", lMouseX, lMouseY)
            '    Else : MyBase.moUILib.SetToolTip("Select a Tradepost first.", lMouseX, lMouseY)
            '    End If
        ElseIf rcInProgress.Contains(lTempX, lTempY) = True Then
            MyBase.moUILib.SetToolTip("Click to view trades in progress.", lMouseX, lMouseY)
        End If

    End Sub

    Private Sub frmTradeMain_OnNewFrame() Handles Me.OnNewFrame
        If glCurrentCycle - mlLastForcefulRender > 45 Then
            Me.IsDirty = True
            mlLastForcefulRender = glCurrentCycle
        End If
        If cboTradePost Is Nothing = False AndAlso cboTradePost.ListIndex > -1 Then
            If lblEmpty.Visible = True Then
                lblEmpty.Visible = False
                Me.IsDirty = True
            End If
        Else
            If lblEmpty.Visible = False Then
                lblEmpty.Visible = True
                Me.IsDirty = True
            End If
        End If
        If fraContents Is Nothing = False Then
            Try
                If fraContents.ControlName = "fraTradeContents" Then CType(fraContents, fraTradeContents).fraTradeContents_OnNewFrame()
                If fraContents.ControlName = "fraTradeHistory" Then CType(fraContents, fraTradeHistory).fraTradeHistory_OnNewFrame()
                If fraContents.ControlName = "fraTradeInProgress" Then CType(fraContents, fraTradeInProgress).fraTradeInProgress_OnNewFrame()
            Catch
            End Try
        End If
    End Sub

    Private Sub frmTradeMain_OnRenderEnd() Handles Me.OnRenderEnd
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        Dim clrTPDepend As System.Drawing.Color = muSettings.InterfaceBorderColor

		If cboTradePost.ListIndex = -1 OrElse cboTradePost.ItemData(cboTradePost.ListIndex) < 1 Then
			clrTPDepend = System.Drawing.Color.FromArgb(clrTPDepend.A, clrTPDepend.R \ 2, clrTPDepend.G \ 2, clrTPDepend.B \ 2)
			myButtonState = 3
		End If

        If cboTradePost.IsExpanded = False OrElse myButtonState <> 5 Then
            Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                Dim oSelColor As System.Drawing.Color
                oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)

                oSprite.Begin(SpriteFlags.AlphaBlend)
                Dim rcTemp As Rectangle

                Select Case myButtonState
                    Case 0
                        rcTemp = rcBuy
                    Case 1
                        rcTemp = rcSell
                    Case 2
                        rcTemp = rcDirect
                    Case 3
                        rcTemp = rcHistory
                    Case 4
                        'rcTemp = rcBuyOrder
                    Case 5
                        rcTemp = rcInProgress
                    Case Else
                        rcTemp = rcBuy
                        myButtonState = 0
                End Select

                If cboTradePost.IsExpanded = True Then
                    If rcTemp.Right > cboTradePost.Left Then
                        rcTemp.Width = cboTradePost.Left - rcTemp.Left
                    End If
                End If

                DoMultiColorFill(rcTemp, oSelColor, rcTemp.Location, oSprite)
                oSprite.End()
            End Using
        End If

        'Now, draw our borders around the buttons
        MyBase.RenderRoundedBorder(rcBuy, 1, clrTPDepend)
        MyBase.RenderRoundedBorder(rcSell, 1, clrTPDepend)
        MyBase.RenderRoundedBorder(rcDirect, 1, clrTPDepend)
        'MyBase.RenderRoundedBorder(rcBuyOrder, 1, clrTPDepend)

        'Dim rcFinalHistory As Rectangle = rcHistory
        'If cboTradePost.IsExpanded = True Then
        '    If rcFinalHistory.Right > cboTradePost.Left Then
        '        rcFinalHistory.Width = cboTradePost.Left - rcFinalHistory.Left
        '    End If
        'End If
        'MyBase.RenderRoundedBorder(rcFinalHistory, 1, muSettings.InterfaceBorderColor)
        MyBase.RenderRoundedBorder(rcHistory, 1, muSettings.InterfaceBorderColor)
        'If cboTradePost.IsExpanded = False Then MyBase.RenderRoundedBorder(rcInProgress, 1, muSettings.InterfaceBorderColor)
        Dim rcFinalIP As Rectangle = rcInProgress
        If cboTradePost.IsExpanded = True Then
            If rcFinalIP.Right > cboTradePost.Left Then
                rcFinalIP.Width = cboTradePost.Left - rcFinalIP.Left
            End If
        End If
        MyBase.RenderRoundedBorder(rcFinalIP, 1, muSettings.InterfaceBorderColor)
        'If cboTradePost.IsExpanded = False Then MyBase.RenderRoundedBorder(rcInProgress, 1, muSettings.InterfaceBorderColor)

        'Now, render our text...
        Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                oTextSpr.Begin(SpriteFlags.AlphaBlend)

                oFont.DrawText(oTextSpr, "BUY", rcBuy, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrTPDepend)
                oFont.DrawText(oTextSpr, "SELL", rcSell, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrTPDepend)
                oFont.DrawText(oTextSpr, "DIRECT", rcDirect, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrTPDepend)
                oFont.DrawText(oTextSpr, "HISTORY", rcHistory, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                'oFont.DrawText(oTextSpr, "BUY ORDER", rcBuyOrder, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrTPDepend)
                'If cboTradePost.IsExpanded = False Then oFont.DrawText(oTextSpr, "IN PROGRESS", rcInProgress, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                If cboTradePost.IsExpanded = False Then
                    oFont.DrawText(oTextSpr, "  IN PROGRESS", rcFinalIP, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                Else
                    oFont.DrawText(oTextSpr, "  IN PROGRE  ", rcFinalIP, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                End If

                oTextSpr.End()
                oTextSpr.Dispose()
            End Using
            oFont.Dispose()
        End Using
    End Sub

    Private Sub UpdateWindowContents()
        If cboTradePost Is Nothing = False Then
            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X).ControlName.ToUpper = cboTradePost.ControlName.ToUpper Then
                    Me.RemoveChild(X)
                    Exit For
                End If
            Next X
        End If

        If fraContents Is Nothing = False Then
            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X).ControlName.ToUpper = fraContents.ControlName.ToUpper Then
                    Me.RemoveChild(X)
                    Exit For
                End If
            Next X
        End If
        fraContents = Nothing

		If cboTradePost Is Nothing = False AndAlso cboTradePost.ListIndex <> -1 AndAlso cboTradePost.ItemData(cboTradePost.ListIndex) > 0 Then
			Dim lTradePostID As Int32 = cboTradePost.ItemData(cboTradePost.ListIndex)

			'fraContents initial props
			Select Case myButtonState
				Case 0, 1, 4
					fraContents = New fraTradeContents(MyBase.moUILib)
					CType(fraContents, fraTradeContents).SetFrameDetails(myButtonState, lTradePostID)

					If myButtonState = 1 Then
						Dim yMsg(5) As Byte
						System.BitConverter.GetBytes(GlobalMessageCode.eGetTradePostTradeables).CopyTo(yMsg, 0)
						System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, 2)
						MyBase.moUILib.SendMsgToPrimary(yMsg)
					End If
				Case 2
					fraContents = New fraDirectTrade(MyBase.moUILib)
					CType(fraContents, fraDirectTrade).SetTradePost(lTradePostID)

					Dim yMsg(5) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eGetTradePostTradeables).CopyTo(yMsg, 0)
					System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, 2)
					MyBase.moUILib.SendMsgToPrimary(yMsg)
				Case 3
					fraContents = New fraTradeHistory(MyBase.moUILib)

					'Send off a request for the player's trade history
					Dim yMsg(5) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eGetTradeHistory).CopyTo(yMsg, 0)
					System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
					MyBase.moUILib.SendMsgToPrimary(yMsg)
				Case 5
					fraContents = New fraTradeInProgress(MyBase.moUILib)
					TradeDeliveryPackage.lTradeDeliveryUB = -1
					'Send off a request for the player's trades in progress 
					Dim yMsg(5) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eGetTradeDeliveries).CopyTo(yMsg, 0)
					System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
					MyBase.moUILib.SendMsgToPrimary(yMsg)
			End Select
			If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
				Select Case myButtonState
					Case 0		'buy window
						goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.BuyScreenSelected)
					Case 1		'sell window
						goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.SellScreenSelected)
					Case 2		'Direct trade
						goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.DirectTradeScreenSelected)
					Case 3		'Trade History
						goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.TradeHistoryScreenSelected)
					Case 4		'Buy Order
						goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.BuyOrderScreenSelected)
					Case 5
						goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.TradeInProgressScreenSelected)
				End Select
			End If
		End If

        If fraContents Is Nothing = False Then
            With fraContents
                '.ControlName = "fraContents"
                .Left = 5
                .Top = 55
                .Width = 790
                .Height = 540
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .BorderLineWidth = 2
                .FillWindow = False
            End With
            Me.AddChild(CType(fraContents, UIControl))
        End If

        If cboTradePost Is Nothing = False Then Me.AddChild(CType(cboTradePost, UIControl))
    End Sub

    Private Sub DoMultiColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point, ByRef oSpr As Sprite)
        Dim rcSrc As Rectangle

        Dim fX As Single
        Dim fY As Single

        If rcDest.Width = 0 OrElse rcDest.Height = 0 OrElse MyBase.moUILib.oInterfaceTexture Is Nothing Then Exit Sub

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

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		ReturnToPreviousViewAndReleaseBackground()
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Public Sub SetTradePostID(ByVal lID As Int32)
		mlTradepostID = lID
		cboTradePost.FindComboItemData(lID)
	End Sub

	Private Sub cboTradePost_ItemSelected(ByVal lItemIndex As Integer) Handles cboTradePost.ItemSelected
		If mbLoading = True Then Return
        UpdateWindowContents()
        lblEmpty.Visible = lItemIndex = -1
	End Sub

	Public Sub HandleGetTradePostList(ByRef yData() As Byte)
		Dim lPos As Int32 = 2		'for msgcode
		Dim lUB As Int32 = System.BitConverter.ToInt16(yData, lPos) - 1 : lPos += 2
		Dim lPrevID As Int32 = -1
		If cboTradePost.ListIndex <> -1 Then lPrevID = cboTradePost.ItemData(cboTradePost.ListIndex)

		cboTradePost.Clear()

		For X As Int32 = 0 To lUB
			Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			Dim sColony As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

			cboTradePost.AddItem(sName & " (" & sColony & ")")
			cboTradePost.ItemData(cboTradePost.NewIndex) = lID
        Next X
        cboTradePost.SortList(True, False)
		If mlTradepostID <> -1 Then cboTradePost.FindComboItemData(mlTradepostID)
	End Sub
 
	Public Sub ReturnToPreviousViewAndReleaseBackground()
		If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True AndAlso goTutorial.CurrentStepType = TutorialManager.TutorialStepType.eTradeScreen Then goTutorial.CancelCustomStepType()
		'If glCurrentEnvirView = CurrentView.eGTCScreen Then ReturnToPreviousView()
		ReturnToPreviousView()
		If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
		goFullScreenBackground = Nothing
	End Sub

	Public Sub SwitchToCreateBuyOrder()
		If cboTradePost Is Nothing = False Then
			For X As Int32 = 0 To Me.ChildrenUB
				If Me.moChildren(X).ControlName.ToUpper = cboTradePost.ControlName.ToUpper Then
					Me.RemoveChild(X)
					Exit For
				End If
			Next X
		End If

		If fraContents Is Nothing = False Then
			For X As Int32 = 0 To Me.ChildrenUB
				If Me.moChildren(X).ControlName.ToUpper = fraContents.ControlName.ToUpper Then
					Me.RemoveChild(X)
					Exit For
				End If
			Next X
		End If
		fraContents = Nothing

		'Now, set fraContents
		If cboTradePost Is Nothing = False AndAlso cboTradePost.ListIndex <> -1 Then
			Dim lTradePostID As Int32 = cboTradePost.ItemData(cboTradePost.ListIndex)
			fraContents = New fraCreateBuyOrder(MyBase.moUILib)
			CType(fraContents, fraCreateBuyOrder).SetTradePostID(lTradePostID)
		End If

		If fraContents Is Nothing = False Then
			If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.CreateBuyOrderSelected)
			With fraContents
				.Left = 5
				.Top = 55
				.Width = 790
				.Height = 540
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
				.BorderLineWidth = 2
				.FillWindow = False
			End With
			Me.AddChild(CType(fraContents, UIControl))
		End If

		If cboTradePost Is Nothing = False Then Me.AddChild(CType(cboTradePost, UIControl))
	End Sub

	Public Sub GotoNewView(ByVal yViewVal As Byte)
		myButtonState = yViewVal
		UpdateWindowContents()
	End Sub

	Public Sub ReturnToCurrentView()
		UpdateWindowContents()
	End Sub

	Protected Overrides Sub Finalize()
		ReturnToPreviousViewAndReleaseBackground()
		MyBase.Finalize()
	End Sub

	Public Sub ViewDirectTradeDetails(ByRef oTrade As Trade)
		If cboTradePost.ListIndex = -1 Then
			Dim lTPID As Int32 = -1
			If oTrade.lPlayer1ID = glPlayerID Then lTPID = oTrade.lP1TradepostID Else lTPID = oTrade.lP2TradepostID
			If cboTradePost.FindComboItemData(lTPID) = False Then Return
		End If
		myButtonState = 2
		UpdateWindowContents()

		If fraContents Is Nothing = False Then
			If fraContents.ControlName = "fraDirectTrade" Then CType(fraContents, fraDirectTrade).SetFromTradeObject(oTrade)
		End If
	End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eTradeScreen)
    End Sub

    Private Sub frmTradeMain_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.TradeMainX = Me.Left
            muSettings.TradeMainY = Me.Top
        End If
    End Sub
End Class

