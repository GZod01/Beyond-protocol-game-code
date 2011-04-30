Option Strict On

'this class is responsible for all user interfaces...
Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Enum UILibMsgCode As Short
    eMouseDownMsgCode = 0
    eMouseMoveMsgCode
    eMouseUpMsgCode
    eKeyDownCode
    eKeyUpCode
    eKeyPressCode
End Enum

Public Class UILib
    Public Enum eSelectState As Integer
        eNoSelectState = 0
        eBombTargetSelection
        eDockWithObjectSelection
        eEvadeToLocation
        eMoveWindow
        eScrollBarDrag

        eNotification_ClickTo

        eSetFleetDest

        eEmailAttachmentJumpTo

        eRepairTarget
        eDismantleTarget

		eBattlegroupJumpTo

		eSelectRouteLoc

		eResizeWindow

		eSketchPad_SelectPoint
		eSketchPad_TextEntry

		eSelectRaceWaypoint
        ePlaceGuildFacility

        eJumpToEntity

        eSelectRouteTemplateLoc
    End Enum
    Public oDevice As Device        'POINTER!! DO NOT DISPOSE!

    Private mlWindowUB As Int32 = -1
    Private myWindowUsed() As Byte
    Private moWindows() As UIWindow

    Public Pen As Sprite

    Public FocusedControl As UIControl
	Private moInterfaceTexture As Texture
	Public ReadOnly Property oInterfaceTexture() As Texture
		Get
			If moInterfaceTexture Is Nothing OrElse moInterfaceTexture.Disposed = True Then
				If goResMgr Is Nothing = False Then
					moInterfaceTexture = goResMgr.GetTexture("Interface.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
				Else : moInterfaceTexture = Nothing
				End If
			End If
			Return moInterfaceTexture
		End Get
	End Property

	Private moToolTip As UIWindow
    Private moMsgBox As frmMsgBox = Nothing
    Public Sub SetMsgBox(ByRef oMsgBox As frmMsgBox)
        moMsgBox = oMsgBox
    End Sub

	Private moSelection As UIWindow		'POINTER!!! can be frmFacilityDisplay, frmUnitDisplay or frmMultiDisplay
	Private moSingleSel As frmSingleSelect
	Private moMultiSel As frmMultiDisplay
	Private moAdvDisplay As frmAdvanceDisplay

	Private mlSelections() As Int32
	Private mlSelectionUB As Int32 = -1

	Private moMsgSys As MsgSystem		'POINTER!!! TO BE ASSOCIATED USING MEMBER FUNCTIONS/METHODS

	Public BuildGhost As BaseMesh
	Public vecBuildGhostLoc As Vector3
	Public BuildGhostAngle As Int16		'measured in 10ths of a degree so 0 to 3600
	'GUID for the ghost object
	Public BuildGhostID As Int32
    Public BuildGhostTypeID As Int16
    Public bBuildGhostNaval As Boolean = False
    'arc direction ring for build ghost
    Public oBuildGhostFont As System.Drawing.Font = New System.Drawing.Font("MS Sans Serif", 30.0F, FontStyle.Bold, GraphicsUnit.Point)

	'For other selection states...
	Public lUISelectState As eSelectState
	Public yBombardType As BombardType
	Public vecBombardLoc As Vector3
    Public oMovingWindow As UIWindow = Nothing

    Public oScrollingBar As UIScrollBar = Nothing
    Public oTechProp As ctlTechProp = Nothing
    Public oScrollingRel As ctlDiplomacy = Nothing
    Public oButtonDown As UIButton = Nothing
    Public oOptionDown As UIOption = Nothing


	Public CurrentComboBoxSelected As UIControl = Nothing
	Public RetainTooltip As Boolean = False

	Private mbShiftDown As Boolean = False

    Public oFrameFont As System.Drawing.Font = New System.Drawing.Font("MS Sans Serif", 10.0F, FontStyle.Bold, GraphicsUnit.Point)

    Public yRenderUI As Byte = 255      '255 indicates all UI, 1 indicates only notification and chat, 0 = none

    Private mptToolTipMouse As Point = Point.Empty

    Public mfDistFromSelection As Single = Single.MinValue
    'Private moDistFont As Font
    Private moDistSysFont As System.Drawing.Font

    Public lJumpToExtended1 As Int32 = 0
    Public lJumpToExtended2 As Int32 = 0

	'Public lLineListUB As Int32 = -1
	'Public uLineList() As LineListItem
	'Public Structure LineListItem
	'    Public v2LineVecs() As Vector2
	'    Public bAntiAliased As Boolean
	'    Public LineWidth As Single
	'    Public LineColor As System.Drawing.Color
	'End Structure
	'Public Sub AddLineListItem(ByRef vecs() As Vector2, ByVal bAA As Boolean, ByVal fWidth As Single, ByVal clrVal As System.Drawing.Color)
	'    lLineListUB += 1
	'    ReDim Preserve uLineList(lLineListUB)
	'    With uLineList(lLineListUB)
	'        .v2LineVecs = vecs
	'        .bAntiAliased = bAA
	'        .LineWidth = fWidth
	'        .LineColor = clrVal
	'    End With
	'End Sub
	'Public Sub DrawLineList()
	'    If lLineListUB <> -1 Then
	'        Using oLine As New Line(oDevice)
	'            For X As Int32 = 0 To lLineListUB
	'                oLine.Antialias = uLineList(X).bAntiAliased
	'                oLine.Width = uLineList(X).LineWidth
	'                oLine.Begin()
	'                oLine.Draw(uLineList(X).v2LineVecs, uLineList(X).LineColor)
	'                oLine.End()
	'            Next X
	'        End Using
	'    End If

	'    lLineListUB = -1
	'    Erase uLineList
	'End Sub


#Region " Message System Accessibility "
	Public Sub SetMsgSystem(ByRef oMsgSys As MsgSystem)
		moMsgSys = Nothing
		moMsgSys = oMsgSys
	End Sub

	'HACK: not secure and not safe
	Public Sub SendMsgToOperator(ByVal yData() As Byte)
		If moMsgSys Is Nothing = False Then moMsgSys.SendToOperator(yData)
	End Sub

	Public Sub SendMsgToPrimary(ByVal yData() As Byte)
		If moMsgSys Is Nothing = False Then moMsgSys.SendToPrimary(yData)
	End Sub

    Public Sub SendLenAppendedMsgToPrimary(ByRef yData() As Byte)
        If moMsgSys Is Nothing = False Then moMsgSys.SendLenAppendMsgToPrimary(yData)
    End Sub

	Public Sub SendMsgToRegion(ByVal yData() As Byte)
		If moMsgSys Is Nothing = False Then moMsgSys.SendToRegion(yData)
	End Sub

	Public Sub SendLenAppendedMsgToRegion(ByRef yData() As Byte)
		If moMsgSys Is Nothing = False Then moMsgSys.SendLenAppendMsgToRegion(yData)
	End Sub
#End Region

	Public Sub SetBehaviorExtendedVals(ByVal lVal1 As Int32, ByVal lVal2 As Int32)
		Dim X As Int32
		Dim oTmpWin As frmBehavior

		For X = 0 To mlWindowUB
			If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmBehavior" Then
				oTmpWin = CType(moWindows(X), frmBehavior)
				oTmpWin.lExtendVal1 = lVal1
				oTmpWin.lExtendVal2 = lVal2
				oTmpWin.PrepareAndSendUpdateMsg()
				oTmpWin = Nothing
				Exit For
			End If
		Next X

	End Sub

	Private Sub HUDVisible(ByVal bVisible As Boolean, ByVal bMinMaxVisible As Boolean)
		'when the HUD is to not be displayed (because of no selection, for example)
		'If we are here, then the Minimize/Maximize button brought us here...
		Dim X As Int32
		'Dim lCnt As Int32

        If moSelection Is Nothing = False Then moSelection.Visible = bVisible
		If moAdvDisplay Is Nothing = False Then moAdvDisplay.Visible = bVisible
		'If moContents Is Nothing = False Then moContents.Visible = bVisible


		'If moCommands Is Nothing = False Then
		'    For X = 0 To mlSelectionUB
		'        If mlSelections(X) <> -1 Then lCnt += 1
		'    Next X
		'    If lCnt = 1 Then moCommands.Visible = bVisible Else moCommands.Visible = False
		'End If

		For X = 0 To mlWindowUB
            If myWindowUsed(X) > 0 AndAlso (moWindows(X).ControlName = "frmBuildWindow" OrElse moWindows(X).ControlName = "frmBehavior" _
               OrElse moWindows(X).ControlName = "frmSelectFac" OrElse moWindows(X).ControlName = "frmFacilityAdv" OrElse _
               moWindows(X).ControlName = "frmTransfer" OrElse moWindows(X).ControlName = "frmRouteConfig" OrElse moWindows(X).ControlName = "frmUnitGoto") Then
                moWindows(X).Visible = bVisible
            End If
		Next X

		'If moMinMax Is Nothing = False Then moMinMax.Visible = bMinMaxVisible

	End Sub

	Public Sub AddSelection(ByVal lEntityIndex As Int32)
		Dim lIdx As Int32 = -1
		Dim X As Int32
		Dim lCnt As Int32 = 0
		Dim bShowStatus As Boolean = True
		Dim bMyUnit As Boolean = False
        Static iLastEntityIndex As Int32 = -1

		Try
			Dim oEnvir As BaseEnvironment = goCurrentEnvir
			If oEnvir Is Nothing Then Return

			Dim oEntity As BaseEntity = oEnvir.oEntity(lEntityIndex)
			If oEntity Is Nothing Then Return

			For X = 0 To mlSelectionUB
				If mlSelections(X) = -1 AndAlso lIdx = -1 Then
					lIdx = X
				ElseIf mlSelections(X) = lEntityIndex Then
					lIdx = X
					lCnt += 1
				ElseIf mlSelections(X) <> -1 Then
					lCnt += 1
				End If
			Next X

			If lIdx = -1 Then
				mlSelectionUB += 1
				ReDim Preserve mlSelections(mlSelectionUB)
				lIdx = mlSelectionUB
				lCnt += 1
			End If
			mlSelections(lIdx) = lEntityIndex

			'Now, do the HUD, remove any existing Selection form...
			moSelection = Nothing

			'Now, create our new form
			If lCnt = 1 Then
				'Ok, single item selection
				If moSingleSel Is Nothing Then moSingleSel = New frmSingleSelect(Me)
                moSelection = moSingleSel
                If moMultiSel Is Nothing = False Then 'Clear on 1st new unit
                    moMultiSel.ClearItems()
                    moMultiSel.SetFromEntity(lEntityIndex)
                End If

                If oEntity.OwnerID = glPlayerID Then
                    'If moContents Is Nothing Then moContents = New frmContents(Me)
                    'moContents.Visible = True
                    'moContents.SetEntityRef(lEntityIndex)

                    'Dim oAIWin As frmBehavior = CType(GetWindow("frmBehavior"), frmBehavior)
                    'If oAIWin Is Nothing Then oAIWin = New frmBehavior(Me)
                    'oAIWin.MultiSelect = False
                    'oAIWin.SetFromEntity(lEntityIndex)
                ElseIf oEntity.OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                    moMsgSys.RequestEntityDetails(oEntity, oEnvir)
                    bShowStatus = False
                Else : bShowStatus = False
                End If

				moSingleSel.SetFromEntity(lEntityIndex)
			Else
				If moMultiSel Is Nothing Then moMultiSel = New frmMultiDisplay(Me)
                'moMultiSel.ClearItems()

                'Now, go thru and add em
                If lCnt = 2 AndAlso iLastEntityIndex > -1 Then
                    moMultiSel.SetFromEntity(iLastEntityIndex)
                    iLastEntityIndex = -1
                End If
                moMultiSel.SetFromEntity(lEntityIndex)
                'For X = 0 To mlSelectionUB
                'moMultiSel.SetFromEntity(mlSelections(X))
                'Next X

                'If moContents Is Nothing = False Then moContents.Visible = False
                RemoveWindow("frmContents")
                RemoveWindow("frmTransfer")

                'ensure our behavior form is not visible
                'RemoveWindow("frmBehavior")

                moSelection = moMultiSel

                If HasAliasedRights(AliasingRights.eChangeBehavior) = True Then
                    Dim oAIWin As frmBehavior = CType(GetWindow("frmBehavior"), frmBehavior)
                    If oAIWin Is Nothing Then oAIWin = New frmBehavior(Me)
                    If mlSelectionUB > -1 Then oAIWin.SetFromEntity(mlSelections(0))
                    oAIWin.MultiSelect = True       'when multiselect is on, the window does not care about multiple calls to SetFromEntity
                    '  it will populate based on what is currently selected in the envir
                End If
                bShowStatus = False
			End If
            iLastEntityIndex = lEntityIndex
			'Late bind function call, must ensure that all 3 display types have this function...
			bMyUnit = oEntity.OwnerID = glPlayerID
			If moSelection Is Nothing = False Then moSelection.Visible = True

			'Now, show our other forms...
			If bShowStatus = True Then
				If moAdvDisplay Is Nothing Then moAdvDisplay = New frmAdvanceDisplay(Me)
				moAdvDisplay.Visible = True
				moAdvDisplay.SetFromEntity(lEntityIndex)
			ElseIf moAdvDisplay Is Nothing = False Then
				moAdvDisplay.Visible = False
			End If

			'If moCommands Is Nothing Then moCommands = New frmCommands(Me)
			'moCommands.Visible = bMyUnit

			'If moMinMax Is Nothing Then moMinMax = New frmMinMax(Me)
			'moMinMax.Visible = bMyUnit

			'If moMinMax.Visible = True AndAlso moMinMax.HUDMinimized = True Then
			'    HUDVisible(False, True)
			'End If
		Catch
			'do nothing, we'll see what happens
		End Try

	End Sub

	Public Sub RemoveSelection(ByVal lEntityIndex As Int32)
		Dim X As Int32
		Dim bFound As Boolean = False
		Dim lCnt As Int32

		For X = 0 To mlSelectionUB
			If mlSelections(X) = lEntityIndex Then
				mlSelections(X) = -1
				If lCnt > 0 Then Exit Sub Else bFound = True
			ElseIf mlSelections(X) <> -1 Then
				lCnt += 1
				If bFound = True Then Exit Sub
			End If
		Next X

		'Ok, we are here, which means we have nothing selected unless bfound = false
		If bFound = True Then
			moSelection = Nothing	'clear our pointer
			mlSelectionUB = -1
			Erase mlSelections
		End If
	End Sub

    Public ReadOnly Property bEntitiesSelected() As Boolean
        Get
            If mlSelectionUB = -1 Then Return False
            For X As Int32 = 0 To mlSelectionUB
                If mlSelections(X) <> -1 Then
                    Return True
                End If
            Next
            Return False
        End Get
    End Property

	Public Sub ClearSelection()
		If moMultiSel Is Nothing = False Then moMultiSel.ClearItems()

		'set all windows to nothing now
		moSelection = Nothing


		RemoveWindow("frmContents")
		RemoveWindow("frmBuildWindow")
		RemoveWindow("frmSelectFac")
		RemoveWindow("frmResearchMain")
		RemoveWindow("frmBehavior")
		RemoveWindow("frmRouteConfig")
		RemoveWindow("frmFacilityAdv")
		RemoveWindow("frmProdStatus")
		RemoveWindow("frmTransfer")
		RemoveWindow("frmRepair")
		RemoveWindow("frmMatRes")
        RemoveWindow("frmUnitSelectFilter")

		BuildGhost = Nothing

		mlSelectionUB = -1
		Erase mlSelections
		HUDVisible(False, False)
	End Sub

	Public Sub RenderInterfaces(ByVal lEnvirView As Int32)
		Dim X As Int32

        glUI_Rendered = 0

        Try

            If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
            If Pen Is Nothing OrElse Pen.Disposed = True Then
                Device.IsUsingEventHandlers = False
                Pen = New Sprite(oDevice)
                'AddHandler Pen.Disposing, AddressOf SpriteDispose
                'AddHandler Pen.Lost, AddressOf SpriteLost
                'AddHandler Pen.Reset, AddressOf SpriteLost
                Device.IsUsingEventHandlers = True
            End If

            If yRenderUI = 0 Then
                Return
            ElseIf yRenderUI = 1 Then
                'render notification and chat only
                For X = 0 To mlWindowUB
                    If myWindowUsed(X) > 0 AndAlso moWindows(X) Is Nothing = False AndAlso moWindows(X).ControlName Is Nothing = False Then
                        If moWindows(X).ControlName = "frmChat" Then
                            moWindows(X).DrawControl()
                        ElseIf moWindows(X).ControlName = "frmNotification" Then
                            moWindows(X).DrawControl()
                        ElseIf moWindows(X).ControlName = "frmNoteHistory" Then
                            moWindows(X).DrawControl()
                        End If
                    End If
                Next X
                Return
            ElseIf yRenderUI = 2 Then
                'render chat, notification and sketchpad only
                For X = 0 To mlWindowUB
                    If myWindowUsed(X) > 0 AndAlso moWindows(X) Is Nothing = False AndAlso moWindows(X).ControlName Is Nothing = False Then
                        If moWindows(X).ControlName = "frmChat" Then
                            moWindows(X).DrawControl()
                        ElseIf moWindows(X).ControlName = "frmNotification" Then
                            moWindows(X).DrawControl()
                        ElseIf moWindows(X).ControlName = "frmNoteHistory" Then
                            moWindows(X).DrawControl()
                        ElseIf moWindows(X).ControlName = "frmSketchPad" Then
                            moWindows(X).DrawControl()
                        End If
                    End If
                Next X
                Return
            End If


            'If goTutorial Is Nothing = False Then goTutorial.HandleAlertControlUpdate()
            If NewTutorialManager.TutorialOn = True Then NewTutorialManager.HandleAlertControlUpdate()

            Dim lTutorialWinIdx As Int32 = -1
            If glCurrentEnvirView < CurrentView.eFullScreenInterface Then
                'draw our contents display
                'If moContents Is Nothing = False Then moContents.DrawControl()

                'draw our commands display
                'If moCommands Is Nothing = False Then moCommands.DrawControl()

                'draw our min/max button
                'If moMinMax Is Nothing = False Then moMinMax.DrawControl()

                'Now, render our other interfaces
                Dim lCurUB As Int32 = -1
                If myWindowUsed Is Nothing = False Then lCurUB = Math.Min(mlWindowUB, myWindowUsed.GetUpperBound(0))
                For X = 0 To lCurUB
                    If myWindowUsed(X) > 0 AndAlso moWindows(X) Is Nothing = False Then
                        If moWindows(X).ControlName Is Nothing = False AndAlso moWindows(X).ControlName = "frmTutorialStep" Then
                            lTutorialWinIdx = X
                        Else
                            moWindows(X).DrawControl()
                        End If
                    End If
                Next X

                'draw our selection box
                If moSelection Is Nothing = False Then moSelection.DrawControl()

                'draw our defense display
                If moAdvDisplay Is Nothing = False Then moAdvDisplay.DrawControl()
            Else
                'Now, render our other interfaces
                Dim lChatWinIdx As Int32 = -1
                Dim lNoteWinIdx As Int32 = -1

                Dim lCurUB As Int32 = -1
                If myWindowUsed Is Nothing = False Then lCurUB = Math.Min(mlWindowUB, myWindowUsed.GetUpperBound(0))
                For X = 0 To lCurUB
                    If myWindowUsed(X) > 0 AndAlso moWindows(X) Is Nothing = False AndAlso moWindows(X).FullScreen = True Then
                        If moWindows(X).ControlName Is Nothing = False Then
                            If moWindows(X).ControlName = "frmChat" Then
                                lChatWinIdx = X
                            ElseIf moWindows(X).ControlName = "frmNotification" Then
                                lNoteWinIdx = X
                            ElseIf moWindows(X).ControlName = "frmTutorialStep" Then
                                lTutorialWinIdx = X
                            Else
                                moWindows(X).DrawControl()
                            End If
                        Else
                            moWindows(X).DrawControl()
                        End If
                    End If
                Next X

                If lChatWinIdx <> -1 Then moWindows(lChatWinIdx).DrawControl()
                If lNoteWinIdx <> -1 Then moWindows(lNoteWinIdx).DrawControl()
            End If

            'Render our msgbox on top of everything (except the tooltip)
            If moMsgBox Is Nothing = False Then moMsgBox.DrawControl()

            ' Check to see if alert parent is invisible
            If NewTutorialManager.moAlertControl Is Nothing = False Then
                Dim bRenderAlert As Boolean = True
                Dim oAlertParent As UIControl = NewTutorialManager.moAlertControl.ParentControl
                While oAlertParent Is Nothing = False
                    If oAlertParent.Visible = False Then
                        'NewTutorialManager.moAlertControl = Nothing
                        bRenderAlert = False
                        Exit While
                    End If
                    oAlertParent = oAlertParent.ParentControl
                End While
                If bRenderAlert = True Then UpdateAlertedControl(NewTutorialManager.moAlertControl)
            End If

            'CSAJ 8/19/08 Grey out the tutorial next button if clicking it won't do anything
            'If lTutorialWinIdx <> -1 Then moWindows(lTutorialWinIdx).DrawControl()
            If lTutorialWinIdx <> -1 Then
                For X = 0 To moWindows(lTutorialWinIdx).ChildrenUB
                    If moWindows(lTutorialWinIdx).moChildren(X).ControlName.ToUpper = "BTNNEXT" Then
                        'If NewTutorialManager.TutorialStepFinished() = True Then NewTutorialManager.mbNextButtonClickable = True

                        If moWindows(lTutorialWinIdx).moChildren(X).Enabled <> NewTutorialManager.mbNextButtonClickable Then
                            moWindows(lTutorialWinIdx).moChildren(X).Enabled = NewTutorialManager.mbNextButtonClickable
                        End If
                        Exit For
                    End If
                Next X
                moWindows(lTutorialWinIdx).DrawControl()
            End If

            If GFXEngine.mbCaptureScreenshot = False AndAlso (glCurrentEnvirView = CurrentView.eSystemView OrElse glCurrentEnvirView = CurrentView.ePlanetView) Then
                If mfDistFromSelection <> Single.MinValue Then
                    Dim fDistInMax As Single = mfDistFromSelection * 0.25F
                    BPFont.DrawTextAtPoint(moDistSysFont, "Vis: " & mfDistFromSelection.ToString("#,##0.#0") & vbCrLf & "Det: " & fDistInMax.ToString("#,##0.#0"), frmMain.mlMouseX, frmMain.mlMouseY + 20, muSettings.InterfaceBorderColor)
                End If
            Else : mfDistFromSelection = Single.MinValue
            End If

            'Then, render our tooltip (to ensure it is on top of EVERYTHING)
            If moToolTip Is Nothing = False Then moToolTip.DrawControl()

            If GFXEngine.bRenderInProgress = True Then
                frmLoginDlg.RenderInProgress()
            End If
        Catch
        End Try
    End Sub

	Private Sub InitializeToolTip()
		moToolTip = New UIWindow(Me)
		Dim oLbl As UILabel = New UILabel(Me)

		moToolTip.Visible = False
		moToolTip.Left = 0
		moToolTip.Top = 0
		moToolTip.Width = 100
		moToolTip.Height = 22
		moToolTip.Enabled = False
		moToolTip.BorderLineWidth = 2
		moToolTip.BorderColor = muSettings.InterfaceBorderColor
		moToolTip.FillColor = muSettings.InterfaceFillColor
		moToolTip.bRoundedBorder = False

		oLbl.Top = 2
		oLbl.Left = 2
		oLbl.Width = 96
		oLbl.Height = 18
        oLbl.Caption = "Tooltip"
        oLbl.bFilterBadWords = False
		oLbl.FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
		'oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
		oLbl.ForeColor = muSettings.InterfaceBorderColor
		moToolTip.AddChild(CType(oLbl, UIControl))
		oLbl = Nothing
	End Sub

	Public Sub SetToolTip(ByVal sCaption As String, ByVal lLocX As Int32, ByVal lLocY As Int32)
		If moToolTip Is Nothing Then InitializeToolTip() 'ensure that it is ready

		mptToolTipMouse.X = lLocX : mptToolTipMouse.Y = lLocY

		'Ok, get our label object (should be item 0 in the window's collection)
		Dim oLbl As UILabel = CType(moToolTip.moChildren(0), UILabel)

		'Now, labels have a public function that returns a rectangle indicating the caption's dimensions...
		' so, let's set our caption
		oLbl.Caption = sCaption

		'Now, get our rectangle
		Dim oRct As Rectangle = oLbl.GetTextDimensions()

		moToolTip.Width = oRct.Width + 4
		moToolTip.Height = oRct.Height + 4

		If lLocX < 0 Then
			'ok, center the tooltip window in the middle X
			moToolTip.Left = CInt((oDevice.PresentationParameters.BackBufferWidth / 2) - (moToolTip.Width / 2))
		Else
			'Ok, check if the resulting rect will push us off the screen
			If lLocX + moToolTip.Width > oDevice.PresentationParameters.BackBufferWidth Then
				moToolTip.Left = lLocX - ((lLocX + moToolTip.Width) - oDevice.PresentationParameters.BackBufferWidth)
			Else : moToolTip.Left = lLocX
			End If
		End If
		If lLocY < 0 Then
			'ok, center the tooltip window in the middle Z
			moToolTip.Top = CInt((oDevice.PresentationParameters.BackBufferHeight / 2) - (moToolTip.Height / 2))
		Else
			If lLocY + moToolTip.Height + 10 > oDevice.PresentationParameters.BackBufferHeight Then
				moToolTip.Top = lLocY - ((lLocY + moToolTip.Height + 10) - oDevice.PresentationParameters.BackBufferHeight)
			Else : moToolTip.Top = lLocY + 10
			End If
		End If

		oLbl.Width = moToolTip.Width - 4
		oLbl.Height = moToolTip.Height - 4

		moToolTip.Visible = True

		'release our pointer to the label
		oLbl = Nothing

	End Sub

	Public Sub SetToolTip(ByVal bVisible As Boolean)
		If moToolTip Is Nothing Then InitializeToolTip() 'ensure that it is ready
		If bVisible = True OrElse Me.RetainTooltip = False Then moToolTip.Visible = bVisible
	End Sub

    Public Sub UpdateDistanceIndicator(ByVal fX As Single, ByVal fZ As Single)
        mfDistFromSelection = Single.MinValue
        If moSingleSel Is Nothing = False AndAlso moSingleSel.Visible = True AndAlso Object.Equals(moSingleSel, moSelection) = True Then
            mfDistFromSelection = moSingleSel.GetSelectedItemsDistance(fX, fZ)
        End If
    End Sub

	Public Function GetToolTipLoc() As Point
		If moToolTip Is Nothing = False Then
			If moToolTip.Visible = True Then
				Return mptToolTipMouse
			End If
		End If
		Return Point.Empty
	End Function

	Public Sub New(ByRef oDev As Device)
		oDevice = oDev
        BPFont.ClearAllFonts()

        CreateGlobalRectangleList()

        '        'Create a new Envirdisplay, because it is a singleton, it will work pretty much on its own...
        '		Dim oED As frmEnvirDisplay = New frmEnvirDisplay(Me)
        '		oED.Visible = True
        '		oED = Nothing			'remove our local pointer, but WindowList should contain it still

        moDistSysFont = New System.Drawing.Font("Courier New", 8.0F, FontStyle.Bold, GraphicsUnit.Point)

	End Sub

    'MSC - 08/18/08 - remarked these out to see if we can get rid of that annoying jit at close issue
    'Private Sub SpriteDispose(ByVal sender As Object, ByVal e As EventArgs)
    '	Pen = Nothing
    'End Sub

    'Private Sub SpriteLost(ByVal sender As Object, ByVal e As EventArgs)
    '	If Pen Is Nothing = False Then
    '		Pen.Dispose()
    '	End If
    '	Pen = Nothing
    'End Sub

	Protected Overrides Sub Finalize()
		oDevice = Nothing

        'If moDistFont Is Nothing = False Then moDistFont.Dispose()
        'moDistFont = Nothing

		BuildGhost = Nothing

		Erase moWindows
		Erase myWindowUsed
		mlWindowUB = -1

		If oFrameFont Is Nothing = False Then oFrameFont.Dispose()
		oFrameFont = Nothing

		If Pen Is Nothing = False Then
			Pen.Dispose()
		End If
		Pen = Nothing
		FocusedControl = Nothing

		If oInterfaceTexture Is Nothing = False Then
			oInterfaceTexture.Dispose()
		End If

        'If lnTemp Is Nothing = False Then
        '	lnTemp.Dispose()
        'End If
        'lnTemp = Nothing
		MyBase.Finalize()
	End Sub

	'Private Sub moMinMax_MaximizeAllWindows() Handles moMinMax.MaximizeAllWindows
	'    HUDVisible(True, True)
	'End Sub

	'Private Sub moMinMax_MinimizeAllWindows() Handles moMinMax.MinimizeAllWindows
	'    HUDVisible(False, True)
	'End Sub
	Public Sub WriteDebugWindows(ByRef oWriter As IO.StreamWriter)
		For X As Int32 = 0 To mlWindowUB
			If myWindowUsed(X) > 0 Then
				oWriter.WriteLine("     " & moWindows(X).ControlName & ". Visible=" & moWindows(X).Visible.ToString)
			End If
		Next X
		If moToolTip Is Nothing = False AndAlso moToolTip.Visible = True Then
			Dim sLblCap As String = ""
			Try
				If moToolTip.ChildrenUB > -1 Then
					If moToolTip.moChildren(0) Is Nothing = False Then
						sLblCap = CType(moToolTip.moChildren(0), UILabel).Caption
					End If
				End If
			Catch
			End Try
			oWriter.WriteLine("     " & moToolTip.ControlName & " (tooltip). Caption: " & sLblCap)
		End If
	End Sub

	Public Sub ReprocessInputs()
		Dim X As Int32
        Try
            For X = 0 To mlWindowUB
                If myWindowUsed(X) > 0 Then
                    If moWindows(X) Is Nothing = False Then moWindows(X).ReprocessInputs()
                End If
            Next X
        Catch
        End Try
    End Sub

	Public Function PostMessage(ByVal lMsgType As UILibMsgCode, ByVal e As System.Windows.Forms.KeyPressEventArgs) As Boolean
		'check all of our windows...
		Dim bRes As Boolean = False
		Dim X As Int32

		If lUISelectState = eSelectState.eSketchPad_TextEntry Then
			If lMsgType = UILibMsgCode.eKeyPressCode Then
				Dim oWin As frmSketchPad = CType(Me.GetWindow("frmSketchPad"), frmSketchPad)
				If oWin Is Nothing = False Then
					oWin.AddCharToText(e.KeyChar.ToString)
				End If
			End If
			Return True
		End If

		If e.KeyChar = vbTab Then
			If FocusedControl Is Nothing = False Then
				If mbShiftDown = True Then
					FocusedControl.ProcessShiftTabButton("")
				Else : FocusedControl.ProcessTabButton("")
				End If
				Return True
			Else : Return False	  '???
			End If
		End If

		If moMsgBox Is Nothing = False Then
			bRes = moMsgBox.PostMessage(lMsgType, e)
            'If bRes = True Then Return True Else Return False
            Return True
		End If

		If glCurrentEnvirView < CurrentView.eFullScreenInterface Then
			If NewTutorialManager.TutorialOn = True Then
				Dim lTutorialWinIdx As Int32 = -1
				For X = mlWindowUB To 0 Step -1
					If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmTutorialStep" Then
						lTutorialWinIdx = X
						Exit For
					End If
				Next X
				If lTutorialWinIdx > -1 AndAlso moWindows(lTutorialWinIdx).PostMessage(lMsgType, e) = True Then Return True
			End If

			For X = mlWindowUB To 0 Step -1
				If myWindowUsed(X) > 0 Then
					If moWindows(X).PostMessage(lMsgType, e) = True Then Return True
				End If
			Next X

			If moSelection Is Nothing = False Then bRes = moSelection.PostMessage(lMsgType, e)
			If bRes = True Then Return bRes
			If moAdvDisplay Is Nothing = False Then bRes = moAdvDisplay.PostMessage(lMsgType, e)
			If bRes = True Then Return bRes
			'If moContents Is Nothing = False Then bRes = moContents.PostMessage(lMsgType, e)
			'If bRes = True Then Return bRes
			'If moCommands Is Nothing = False Then bRes = moCommands.PostMessage(lMsgType, e)
			'If bRes = True Then Return bRes
			'If moMinMax Is Nothing = False Then bRes = moMinMax.PostMessage(lMsgType, e)
			'If bRes = True Then Return bRes
		Else
			If NewTutorialManager.TutorialOn = True Then
				Dim lTutorialWinIdx As Int32 = -1
				For X = mlWindowUB To 0 Step -1
					If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmTutorialStep" Then
						lTutorialWinIdx = X
						Exit For
					End If
				Next X
				If lTutorialWinIdx > -1 AndAlso moWindows(lTutorialWinIdx).PostMessage(lMsgType, e) = True Then Return True
			End If

			For X = mlWindowUB To 0 Step -1
				If myWindowUsed(X) > 0 AndAlso moWindows(X).FullScreen = True Then
					If moWindows(X).PostMessage(lMsgType, e) = True Then Return True
				End If
			Next X
		End If

		'If we are here, then clear any items with focus
		If FocusedControl Is Nothing = False Then
			FocusedControl.HasFocus = False
			FocusedControl = Nothing
		End If

		Return False

	End Function

	Public Function PostMessage(ByVal lMsgType As UILibMsgCode, ByVal e As System.Windows.Forms.KeyEventArgs) As Boolean
		Dim bRes As Boolean = False
		Dim X As Int32

		mbShiftDown = e.Shift

		If Me.lUISelectState = eSelectState.eSketchPad_TextEntry Then
			If e.KeyCode = Keys.Back Then
				Dim oWin As frmSketchPad = CType(GetWindow("frmSketchPad"), frmSketchPad)
				If oWin Is Nothing = False Then
					oWin.BackSpaceHit()
				End If
			ElseIf e.KeyCode = Keys.Enter Then
				Dim oWin As frmSketchPad = CType(GetWindow("frmSketchPad"), frmSketchPad)
				If oWin Is Nothing = False Then
					oWin.AddCharToText(vbCrLf)
				End If
			End If
			Return True
		End If

		If moMsgBox Is Nothing = False Then
			bRes = moMsgBox.PostMessage(lMsgType, e)
            'If bRes = True Then Return True
            Return True
		End If

		If glCurrentEnvirView < CurrentView.eFullScreenInterface Then
			If NewTutorialManager.TutorialOn = True Then
				Dim lTutorialWinIdx As Int32 = -1
				For X = mlWindowUB To 0 Step -1
					If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmTutorialStep" Then
						lTutorialWinIdx = X
						Exit For
					End If
				Next X
				If lTutorialWinIdx > -1 AndAlso moWindows(lTutorialWinIdx).PostMessage(lMsgType, e) = True Then Return True
			End If

			For X = mlWindowUB To 0 Step -1
				If myWindowUsed(X) > 0 Then
					If moWindows(X).PostMessage(lMsgType, e) = True Then Return True
				End If
			Next X

			If moSelection Is Nothing = False Then bRes = moSelection.PostMessage(lMsgType, e)
			If bRes = True Then Return bRes
			If moAdvDisplay Is Nothing = False Then bRes = moAdvDisplay.PostMessage(lMsgType, e)
			If bRes = True Then Return bRes
			'If moContents Is Nothing = False Then bRes = moContents.PostMessage(lMsgType, e)
			'If bRes = True Then Return bRes
			'If moCommands Is Nothing = False Then bRes = moCommands.PostMessage(lMsgType, e)
			'If bRes = True Then Return bRes
			'If moMinMax Is Nothing = False Then bRes = moMinMax.PostMessage(lMsgType, e)
			'If bRes = True Then Return bRes
		Else
			If NewTutorialManager.TutorialOn = True Then
				Dim lTutorialWinIdx As Int32 = -1
				For X = mlWindowUB To 0 Step -1
					If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmTutorialStep" Then
						lTutorialWinIdx = X
						Exit For
					End If
				Next X
				If lTutorialWinIdx > -1 AndAlso moWindows(lTutorialWinIdx).PostMessage(lMsgType, e) = True Then Return True
			End If

			For X = mlWindowUB To 0 Step -1
				If myWindowUsed(X) > 0 AndAlso moWindows(X).FullScreen = True Then
					If moWindows(X).PostMessage(lMsgType, e) = True Then Return True
				End If
			Next X
		End If

		'If we are here, then clear any items with focus
		If FocusedControl Is Nothing = False Then
			FocusedControl.HasFocus = False
			FocusedControl = Nothing
		End If

		Return False
	End Function

    Public Function PostMouseWheelEvent(ByVal e As System.Windows.Forms.MouseEventArgs) As Boolean
        If yRenderUI = 0 Then
            Return False
        ElseIf yRenderUI = 1 Then
            'only post to chat and notification
            Dim sNames() As String = {"frmChat", "frmNotification", "frmNoteHistory"}
            For X As Int32 = 0 To mlWindowUB
                If myWindowUsed(X) > 0 Then
                    For Y As Int32 = 0 To sNames.GetUpperBound(0)
                        If moWindows(X).ControlName = sNames(Y) Then
                            If moWindows(X).PostMouseWheel(e) = True Then Return True
                        End If
                    Next
                 End If
            Next X
            Return False
        ElseIf yRenderUI = 2 Then
            Dim sNames() As String = {"frmChat", "frmNotification", "frmNoteHistory", "frmSketchPad"}
            For X As Int32 = 0 To mlWindowUB
                If myWindowUsed(X) > 0 Then
                    For Y As Int32 = 0 To sNames.GetUpperBound(0)
                        If moWindows(X).ControlName = sNames(Y) Then
                            If moWindows(X).PostMouseWheel(e) = True Then Return True
                        End If
                    Next Y
                End If
            Next X
            Return False
        End If

        If glCurrentEnvirView = CurrentView.eHullResearch Then Return False

        If glCurrentEnvirView < CurrentView.eFullScreenInterface Then
            If NewTutorialManager.TutorialOn = True Then
                Dim lTutorialWinIdx As Int32 = -1
                For X As Int32 = mlWindowUB To 0 Step -1
                    If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmTutorialStep" Then
                        lTutorialWinIdx = X
                        Exit For
                    End If
                Next X
                If lTutorialWinIdx > -1 AndAlso moWindows(lTutorialWinIdx).PostMouseWheel(e) = True Then Return True
            End If

            For X As Int32 = mlWindowUB To 0 Step -1
                If myWindowUsed(X) > 0 Then
                    If moWindows(X).PostMouseWheel(e) = True Then Return True
                End If
            Next X

            Dim bRes As Boolean = False
            If moSelection Is Nothing = False Then bRes = moSelection.PostMouseWheel(e)
            If bRes = True Then Return bRes
            If moAdvDisplay Is Nothing = False Then bRes = moAdvDisplay.PostMouseWheel(e)
            If bRes = True Then Return bRes
        Else
            If NewTutorialManager.TutorialOn = True Then
                Dim lTutorialWinIdx As Int32 = -1
                For X As Int32 = mlWindowUB To 0 Step -1
                    If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmTutorialStep" Then
                        lTutorialWinIdx = X
                        Exit For
                    End If
                Next X
                If lTutorialWinIdx > -1 AndAlso moWindows(lTutorialWinIdx).PostMouseWheel(e) = True Then Return True
            End If

            For X As Int32 = mlWindowUB To 0 Step -1
                If myWindowUsed(X) > 0 AndAlso moWindows(X) Is Nothing = False AndAlso moWindows(X).FullScreen = True Then
                    If moWindows(X).PostMouseWheel(e) = True Then Return True
                End If
            Next X
        End If

        Return False
    End Function

	Public Function PostMessage(ByVal lMsgType As UILibMsgCode, ByVal lX As Int32, ByVal lY As Int32, ByVal lButton As System.Windows.Forms.MouseButtons) As Boolean
		Dim bRes As Boolean = False
		Dim X As Int32

		UIControl.mbToolTipSet = False

		'Clear our tool tip
		If RetainTooltip = False Then SetToolTip(False)

		'ok, check our render uis
        If yRenderUI = 0 Then 
            Return False
        ElseIf yRenderUI = 1 Then
            'only post to chat and notification
            For X = 0 To mlWindowUB
                If myWindowUsed(X) > 0 Then
                    If moWindows(X).ControlName = "frmChat" Then
                        If moWindows(X).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
                    ElseIf moWindows(X).ControlName = "frmNotification" Then
                        If moWindows(X).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
                    ElseIf moWindows(X).ControlName = "frmNoteHistory" Then
                        If moWindows(X).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
                    End If
                End If
            Next X
            Return False
        ElseIf yRenderUI = 2 Then
            For X = 0 To mlWindowUB
                If myWindowUsed(X) > 0 Then
                    If moWindows(X).ControlName = "frmChat" Then
                        If moWindows(X).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
                    ElseIf moWindows(X).ControlName = "frmNotification" Then
                        If moWindows(X).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
                    ElseIf moWindows(X).ControlName = "frmNoteHistory" Then
                        If moWindows(X).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
                    ElseIf moWindows(X).ControlName = "frmSketchPad" Then
                        If moWindows(X).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
                    End If
                End If
            Next X
            Return False
        End If


        If moMsgBox Is Nothing = False Then
            bRes = moMsgBox.PostMessage(lMsgType, lX, lY, lButton)
            'If bRes = True Then Return True
            Return True
        End If

        If Me.CurrentComboBoxSelected Is Nothing = False Then
            If Me.CurrentComboBoxSelected.Enabled = True AndAlso Me.CurrentComboBoxSelected.Visible = True AndAlso Me.CurrentComboBoxSelected.TestRegion(lX, lY) = True Then
                If Me.CurrentComboBoxSelected.PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
            ElseIf lMsgType <> UILibMsgCode.eMouseMoveMsgCode AndAlso lMsgType = UILibMsgCode.eMouseDownMsgCode Then
                CType(Me.CurrentComboBoxSelected, UIComboBox).CollapseListBox()
            End If
        End If

        'Ok, check if we are in window move
        If Me.lUISelectState = eSelectState.eScrollBarDrag Then
            If Me.oScrollingBar Is Nothing = False Then
                If lMsgType = UILibMsgCode.eMouseDownMsgCode Then
                    Me.oScrollingBar.btnScroller_OnMouseDown(lX, lY, lButton)
                ElseIf lMsgType = UILibMsgCode.eMouseMoveMsgCode Then
                    Me.oScrollingBar.btnScroller_OnMouseMove(lX, lY, lButton)
                Else
                    Me.oScrollingBar.btnScroller_OnMouseUp(lX, lY, lButton)
                End If
                Return True
            ElseIf Me.oTechProp Is Nothing = False Then
                If lMsgType = UILibMsgCode.eMouseMoveMsgCode Then
                    Me.oTechProp.ctlTechProp_OnMouseMove(lX, lY, lButton)
                Else
                    Me.lUISelectState = eSelectState.eNoSelectState
                    Me.oTechProp = Nothing
                End If
                Return True
            ElseIf Me.oScrollingRel Is Nothing = False Then
                If lMsgType = UILibMsgCode.eMouseMoveMsgCode Then
                    Me.oScrollingRel.ctlDiplomacy_OnMouseMove(lX, lY, lButton)
                Else
                    Me.lUISelectState = eSelectState.eNoSelectState
                    Me.oScrollingRel = Nothing
                End If
                Return True
            End If
        ElseIf Me.lUISelectState = eSelectState.eMoveWindow Then
            If Me.oMovingWindow Is Nothing = False Then
                Me.oMovingWindow.UIWindow_OnMouseMove(lX, lY, lButton)
                Return True
            End If
        ElseIf Me.lUISelectState = eSelectState.eResizeWindow Then
            If Me.oMovingWindow Is Nothing = False Then
                Me.oMovingWindow.HandleResize(lX, lY)
                Return True
            End If
        End If

        If glCurrentEnvirView < CurrentView.eFullScreenInterface Then
            If NewTutorialManager.TutorialOn = True Then
                Dim lTutorialWinIdx As Int32 = -1
                For X = mlWindowUB To 0 Step -1
                    If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmTutorialStep" Then
                        lTutorialWinIdx = X
                        Exit For
                    End If
                Next X
                If lTutorialWinIdx > -1 AndAlso moWindows(lTutorialWinIdx).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
            End If

            For X = mlWindowUB To 0 Step -1
                If myWindowUsed(X) > 0 Then
                    If moWindows(X).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
                End If
            Next X

            If moSelection Is Nothing = False Then bRes = moSelection.PostMessage(lMsgType, lX, lY, lButton)
            If bRes = True Then Return bRes
            If moAdvDisplay Is Nothing = False Then bRes = moAdvDisplay.PostMessage(lMsgType, lX, lY, lButton)
            If bRes = True Then Return bRes
            'If moContents Is Nothing = False Then bRes = moContents.PostMessage(lMsgType, lX, lY, lButton)
            'If bRes = True Then Return bRes
            'If moCommands Is Nothing = False Then bRes = moCommands.PostMessage(lMsgType, lX, lY, lButton)
            'If bRes = True Then Return bRes
            'If moMinMax Is Nothing = False Then bRes = moMinMax.PostMessage(lMsgType, lX, lY, lButton)
            'If bRes = True Then Return bRes
        Else
            If NewTutorialManager.TutorialOn = True Then
                Dim lTutorialWinIdx As Int32 = -1
                For X = mlWindowUB To 0 Step -1
                    If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmTutorialStep" Then
                        lTutorialWinIdx = X
                        Exit For
                    End If
                Next X
                If lTutorialWinIdx > -1 AndAlso moWindows(lTutorialWinIdx).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
            End If

            For X = mlWindowUB To 0 Step -1
                If myWindowUsed(X) > 0 AndAlso moWindows(X) Is Nothing = False AndAlso moWindows(X).FullScreen = True Then
                    If moWindows(X).PostMessage(lMsgType, lX, lY, lButton) = True Then Return True
                End If
            Next X
        End If

        'If we are here, then clear any items with focus
        If lButton <> MouseButtons.None Then
            If FocusedControl Is Nothing = False Then
                FocusedControl.HasFocus = False
                FocusedControl = Nothing
            End If
        End If

        Return False
    End Function

    Private mbPenBegun As Boolean = False
    Public Sub BeginPenSprite(ByVal lFlags As SpriteFlags)
        If mbPenBegun = False Then
            Pen.Begin(lFlags)
            mbPenBegun = True
        End If
    End Sub
    Public Sub EndPenSprite()
        If mbPenBegun = True Then
            Pen.End()
            mbPenBegun = False
        End If
    End Sub

	Public Sub DoAlphaBlendColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point)
		Dim rcSrc As Rectangle

		Dim fX As Single
		Dim fY As Single

        If rcDest.Width = 0 OrElse rcDest.Height = 0 OrElse oInterfaceTexture Is Nothing Then Return

		'rcSrc = System.Drawing.Rectangle.FromLTRB(225, 0, 255, 30)
        rcSrc.Location = New Point(192, 0)
		rcSrc.Width = 62
		rcSrc.Height = 64

		'Now, draw it...
		With Pen
            '.Begin(SpriteFlags.AlphaBlend)
            BeginPenSprite(SpriteFlags.AlphaBlend)
            fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
            fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))
            .Draw2D(oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), System.Drawing.Color.FromArgb(255, 0, 0, 0))
            .Draw2D(oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFill)
            '.End()
            EndPenSprite()
        End With
	End Sub


	'Public Sub DoAlphaBlendColorFill_EX(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point, ByRef oImgSprite As Sprite)
	'    If oImgSprite Is Nothing Then
	'        DoAlphaBlendColorFill(rcDest, clrFill, ptLoc)
	'        Return
	'    End If

	'    Dim rcSrc As Rectangle

	'    Dim fX As Single
	'    Dim fY As Single

	'    If rcDest.Width = 0 OrElse rcDest.Height = 0 Then Exit Sub

	'    'rcSrc = System.Drawing.Rectangle.FromLTRB(225, 0, 255, 30)
	'    rcSrc.Location = New Point(192, 0)
	'    rcSrc.Width = 62
	'    rcSrc.Height = 64

	'    'Now, draw it...
	'    With oImgSprite
	'        'If bPenBegun = False Then .Begin(SpriteFlags.AlphaBlend)
	'        '.Begin(SpriteFlags.AlphaBlend)
	'        fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
	'        fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))
	'        .Draw2D(oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), System.Drawing.Color.FromArgb(255, 0, 0, 0))
	'        .Draw2D(oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFill)
	'        'If bPenBegun = False Then .End()
	'        '.End()
	'    End With

	'End Sub

	Public Sub AddNotification(ByVal sText As String, ByVal cColor As System.Drawing.Color, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lLocX As Int32, ByVal lLocZ As Int32)
        If cColor = Color.Blue Then cColor = Color.Teal
        Dim oTmpWin As frmNotification
		Dim X As Int32

		For X = 0 To mlWindowUB
			If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmNotification" Then
				oTmpWin = CType(moWindows(X), frmNotification)
				oTmpWin.AddNotification(sText, cColor, False, lEnvirID, iEnvirTypeID, lEntityID, iEntityTypeID, lLocX, lLocZ)
				oTmpWin = Nothing
				Exit Sub
			End If
		Next X

		'if we are here, then we need to add it
		oTmpWin = New frmNotification(Me)
		oTmpWin.AddNotification(sText, cColor, False, lEnvirID, iEnvirTypeID, lEntityID, iEntityTypeID, lLocX, lLocZ)
		oTmpWin = Nothing
	End Sub

	Public Sub HandleProductionQueueMsg(ByVal yData() As Byte)
		Dim oTmpWin As frmBuildWindow
		Dim X As Int32

		For X = 0 To mlWindowUB
			If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = "frmBuildWindow" Then
				oTmpWin = CType(moWindows(X), frmBuildWindow)
				oTmpWin.HandleProductionQueueMsg(yData)
				oTmpWin = Nothing
				Exit For
			End If
		Next X

	End Sub

    Private Sub CompactWindowArray()

        Dim oTmp(mlWindowUB) As UIWindow
        Dim yUsed(mlWindowUB) As Byte
        Dim lTmpUB As Int32 = -1

        For X As Int32 = 0 To mlWindowUB
            If myWindowUsed(X) <> 0 Then
                lTmpUB += 1
                oTmp(lTmpUB) = moWindows(X)
                yUsed(lTmpUB) = 255
            End If
        Next X

        mlWindowUB = -1
        moWindows = oTmp
        myWindowUsed = yUsed
        mlWindowUB = lTmpUB
    End Sub

	Public Sub AddWindow(ByVal oWin As UIWindow)
		Dim X As Int32
		Dim lIdx As Int32 = -1

		oWin.Visible = True

		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eWindowOpened, -1, -1, -1, oWin.ControlName)
        End If

        If oWin.lWindowMetricID <> -1 Then BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eViewInterface, oWin.lWindowMetricID, 0)

        'Ok, let's compact our window array
        CompactWindowArray()

        For X = mlWindowUB To 0 Step -1
            If myWindowUsed(X) = 0 Then
                lIdx = X
                Exit For
            Else
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlWindowUB += 1
            ReDim Preserve moWindows(mlWindowUB)
            ReDim Preserve myWindowUsed(mlWindowUB)
            lIdx = mlWindowUB
        End If

        moWindows(lIdx) = oWin
        myWindowUsed(lIdx) = 255
    End Sub

    Public Sub RemoveWindow(ByVal sName As String)
        Dim bForceCollect As Boolean = False
        Dim X As Int32

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eWindowClosed, -1, -1, -1, sName)
        End If

        For X = 0 To mlWindowUB
            If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = sName Then

                If moWindows(X).lWindowMetricID <> -1 Then BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eCloseInterface, moWindows(X).lWindowMetricID, 0)

                If sName.ToLower = "frmresearchmain" Then
                    If moAdvDisplay Is Nothing = False Then moAdvDisplay.ResetBuildButton()
                End If

                If Me.FocusedControl Is Nothing = False Then
                    Dim sTemp As String = sName.ToLower
                    If Me.FocusedControl.ControlName Is Nothing = False AndAlso Me.FocusedControl.ControlName.ToLower = sTemp Then
                        Me.FocusedControl = Nothing
                    Else
                        Dim octrl As UIControl = Me.FocusedControl.ParentControl
                        Dim sTopMostParent As String = ""
                        While octrl Is Nothing = False
                            If octrl.ParentControl Is Nothing Then
                                If octrl.ControlName Is Nothing = False Then sTopMostParent = octrl.ControlName.ToLower
                            End If
                            octrl = octrl.ParentControl
                        End While
                        If sTemp = sTopMostParent Then
                            Me.FocusedControl = Nothing
                        End If
                    End If
                End If
                If moWindows(X) Is Nothing = False Then
                    moWindows(X).Visible = False
                    moWindows(X).RemovedFromUILibList()
                End If
                myWindowUsed(X) = 0
                moWindows(X) = Nothing
                bForceCollect = True
                Exit For
            End If
        Next X

        If bForceCollect = True Then
            'System.GC.Collect()
            Dim oThread As New Threading.Thread(AddressOf _EnsureGCCleanup)
            oThread.Start()
        End If
    End Sub

    Private Sub _EnsureGCCleanup()
        Threading.Thread.Sleep(500)
        System.GC.Collect()
    End Sub

	Public Function GetWindow(ByVal sName As String) As UIWindow
		Dim X As Int32

		For X = 0 To mlWindowUB
			If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = sName Then
				Return moWindows(X)
			End If
		Next X

		If moSelection Is Nothing = False AndAlso moSelection.ControlName = sName Then Return moSelection
		If moAdvDisplay Is Nothing = False AndAlso moAdvDisplay.ControlName = sName Then Return moAdvDisplay
		'If moContents Is Nothing = False AndAlso moContents.ControlName = sName Then Return moContents
		'If moCommands Is Nothing = False AndAlso moCommands.ControlName = sName Then Return moCommands
		'If moMinMax Is Nothing = False AndAlso moMinMax.ControlName = sName Then Return moMinMax

		Return Nothing
	End Function

	Public Sub RemoveAllWindows()
		Dim X As Int32
		For X = 0 To mlWindowUB
			If myWindowUsed(X) > 0 Then
				myWindowUsed(X) = 0
				moWindows(X) = Nothing
				Exit For
			End If
		Next X
		mlWindowUB = -1
		ReDim myWindowUsed(-1)
		ReDim moWindows(-1)
		System.GC.Collect()
	End Sub

	Public Function GetMsgSys() As MsgSystem
		Return moMsgSys
	End Function

	Public Sub HandleColonyStatsMsg(ByRef yData() As Byte)

		If muSettings.ExpandedColonyStatsScreen = True Then
			Dim frmTemp As frmColonyStats = CType(GetWindow("frmColonyStats"), frmColonyStats)
			If frmTemp Is Nothing = False Then frmTemp.HandleColonyMsg(yData)
			frmTemp = Nothing
		Else
			Dim frmTemp As frmColonyStatsSmall = CType(GetWindow("frmColonyStatsSmall"), frmColonyStatsSmall)
			If frmTemp Is Nothing = False Then frmTemp.HandleColonyMsg(yData)
			frmTemp = Nothing
		End If
	End Sub

	Public Sub ReleaseInterfaceTextures()
        Try
            FocusedControl = Nothing
            oScrollingBar = Nothing
            oTechProp = Nothing
            oScrollingRel = Nothing
            oButtonDown = Nothing
            oOptionDown = Nothing
            CurrentComboBoxSelected = Nothing

            For X As Int32 = 0 To mlWindowUB
                Try
                    If moWindows(X) Is Nothing = False Then
                        moWindows(X).KillCanvas()
                    End If
                Catch
                End Try

                Try
                    If moWindows(X) Is Nothing = False Then moWindows(X).KillChildrenDefaultPoolResources()
                Catch
                End Try
            Next X
            If Pen Is Nothing = False Then Pen.Dispose()
            Pen = Nothing
            Try
                If moToolTip Is Nothing = False Then
                    moToolTip.KillCanvas()
                    moToolTip.KillChildrenDefaultPoolResources()
                End If
            Catch
            End Try

            moSelection = Nothing
            If moInterfaceTexture Is Nothing = False Then moInterfaceTexture.Dispose()
            If moSingleSel Is Nothing = False Then moSingleSel.KillCanvas()
            If moMultiSel Is Nothing = False Then moMultiSel.KillCanvas()
            If moAdvDisplay Is Nothing = False Then moAdvDisplay.KillCanvas()
            moMsgBox = Nothing
            BuildGhost = Nothing
            'UIComboBox.ReleaseBorderLine()
            'If lnTemp Is Nothing = False Then lnTemp.Dispose()
            'lnTemp = Nothing
            ctlCalendar.ReleaseDefaultPool()
            'UILine.ReleasePen()
            BPLine.ClearLine()

            'frmSketchPad.ReleaseFont()
            'UIWindow.ReleaseBorderLine()
            frmQuickBar.ReleaseSprite()
        Catch
        End Try
    End Sub

    Public Sub DirtyAllWindows(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color)
        For X As Int32 = 0 To mlWindowUB
            If moWindows(X) Is Nothing = False Then
                moWindows(X).DoResetInterfaceColors(yType, clrPrev)
                moWindows(X).IsDirty = True
            End If
        Next X
    End Sub
	Public Sub DirtyChatWindow()
		For X As Int32 = 0 To mlWindowUB
			If myWindowUsed(X) > 0 AndAlso moWindows(X).FullScreen = True Then
				If moWindows(X).ControlName = "frmChat" Then
					moWindows(X).IsDirty = True
					Return
				End If
			End If
		Next X
	End Sub

    Public Sub BringWindowToForeground(ByVal sName As String)

        Try
            If myWindowUsed(mlWindowUB) > 0 AndAlso moWindows(mlWindowUB).ControlName = sName Then Return

            For X As Int32 = 0 To mlWindowUB
                If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName = sName Then

                    'If moWindows(X) Is Nothing = False Then
                    '    moWindows(X).Visible = False
                    '    moWindows(X).RemovedFromUILibList()
                    'End If
                    'myWindowUsed(X) = 0
                    'moWindows(X) = Nothing

                    Dim oTmpWin As UIWindow = moWindows(X)
                    Dim yUsed As Byte = myWindowUsed(X)

                    For Y As Int32 = X To mlWindowUB - 1
                        moWindows(Y) = moWindows(Y + 1)
                        myWindowUsed(Y) = myWindowUsed(Y + 1)
                    Next Y

                    moWindows(mlWindowUB) = oTmpWin
                    myWindowUsed(mlWindowUB) = yUsed
                    Exit For
                End If
            Next X
        Catch
        End Try
    End Sub

#Region "  Tutorial UI Filter Management  "
	Private Structure uUIControlCmd
		Public sCmdText As String
		Public sParams() As String
		Public bUICommand As Boolean

		Public Function ParamsMatch(ByVal sTestParms() As String, ByVal bPartialMatch As Boolean) As Boolean
			If sParams Is Nothing Then Return True
			If sTestParms Is Nothing Then
				'only way this works is if sparams is all -1s
				For X As Int32 = 0 To sParams.GetUpperBound(0)
					If sParams(X) Is Nothing = False AndAlso sParams(X) <> "-1" AndAlso sParams(X) <> "" Then Return False
				Next X
				Return True
			End If

			'if sparams(x) = -1 then it matches
			'if sparams(x) = sTestParms then it matches

            Dim lUB As Int32 = sParams.GetUpperBound(0)

            Dim bDone As Boolean = False
            While bDone = False
                bDone = True
                For X As Int32 = 0 To sParams.GetUpperBound(0)
                    sParams(X) = sParams(X).Trim
                    If IsNumeric(sParams(X)) = False Then
                        'ok, possible variable name...
                        Dim lID As Int32 = -1
                        Dim iTypeID As Int16 = -1
                        NewTutorialManager.FillStoredGUIDValues(sParams(X), lID, iTypeID)
                        If lID <> -1 OrElse iTypeID <> -1 Then
                            sParams(X) = lID.ToString
                            ReDim Preserve sParams(sParams.GetUpperBound(0) + 1)
                            For Y As Int32 = sParams.GetUpperBound(0) To X + 1 Step -1
                                sParams(Y) = sParams(Y - 1)
                            Next Y
                            sParams(X + 1) = iTypeID.ToString
                            bDone = False
                            Exit For
                        End If
                    End If
                Next X
            End While
            lUB = sParams.GetUpperBound(0)

			Dim lTestUB As Int32 = sTestParms.GetUpperBound(0)

			For X As Int32 = 0 To lUB
				If X > lTestUB Then
					If sParams(X) Is Nothing = False AndAlso sParams(X) <> "-1" AndAlso sParams(X) <> "" Then Return False
                Else
                    If sParams(X) Is Nothing = False AndAlso sParams(X) <> "-1" AndAlso sParams(X) <> "" Then

                        If bPartialMatch = True Then
                            If sParams(X).Trim.ToUpper.StartsWith(sTestParms(X).Trim.ToUpper) = False Then Return False
                        ElseIf sParams(X).Trim.ToUpper <> sTestParms(X).Trim.ToUpper Then
                            Return False
                        End If
                    End If
				End If
			Next X
			Return True
		End Function
	End Structure

	'Private msUIEnabledCmds(-1) As String
	'Private msHotKeyEnabledCmds(-1) As String
	Private muUICmds(-1) As uUIControlCmd
	Private muHotkeyCmds(-1) As uUIControlCmd

	Public Sub HandleDisableCmd(ByVal sCmd As String, ByVal bUI As Boolean)
		sCmd = sCmd.ToUpper.Trim

		If sCmd = "ALL" Then
			If bUI = True Then ReDim muUICmds(-1) Else ReDim muHotkeyCmds(-1)
		Else
			If sCmd.Contains("<") = True Then
				sCmd = sCmd.Substring(0, sCmd.IndexOf("<"))
			End If

			If bUI = True Then
				For X As Int32 = 0 To muUICmds.GetUpperBound(0)
					If muUICmds(X).sCmdText = sCmd Then
						muUICmds(X).sCmdText = ""
					End If
				Next X
			Else
				For X As Int32 = 0 To muHotkeyCmds.GetUpperBound(0)
					If muHotkeyCmds(X).sCmdText = sCmd Then
						muHotkeyCmds(X).sCmdText = ""
					End If
				Next X
			End If
		End If
    End Sub

	Public Sub HandleEnableCmd(ByVal sCmd As String, ByVal bUI As Boolean)
		sCmd = sCmd.ToUpper.Trim

		Dim uNewCmd As uUIControlCmd
		With uNewCmd
			.bUICommand = bUI

			Dim lIdx As Int32 = sCmd.IndexOf("<"c)
			If lIdx > -1 Then
				.sCmdText = sCmd.Substring(0, lIdx)
				sCmd = sCmd.Substring(lIdx).Replace("<", "").Replace(">", "")
				.sParams = Split(sCmd, ",")
				If .sParams Is Nothing = False Then
					Dim bDone As Boolean = False
					While bDone = False
						bDone = True
						For X As Int32 = 0 To .sParams.GetUpperBound(0)
							.sParams(X) = .sParams(X).Trim
							If IsNumeric(.sParams(X)) = False Then
                                'ok, possible variable name...
                                Dim lID As Int32 = -1
								Dim iTypeID As Int16 = -1
								NewTutorialManager.FillStoredGUIDValues(.sParams(X), lID, iTypeID)
                                If lID <> -1 OrElse iTypeID <> -1 Then
                                    .sParams(X) = lID.ToString
                                    ReDim Preserve .sParams(.sParams.GetUpperBound(0) + 1)
                                    For Y As Int32 = .sParams.GetUpperBound(0) To X + 1 Step -1
                                        .sParams(Y) = .sParams(Y - 1)
                                    Next Y
                                    .sParams(X + 1) = iTypeID.ToString
                                    bDone = False
                                    Exit For
                                End If
                            End If
                        Next X

					End While
				End If
			Else
				.sCmdText = sCmd
				.sParams = Nothing
			End If
		End With

		If uNewCmd.sCmdText = "ALL" Then
			If bUI = True Then
				ReDim muUICmds(0)
				muUICmds(0) = uNewCmd
			Else
				ReDim muHotkeyCmds(0)
				muHotkeyCmds(0) = uNewCmd
			End If
		Else
			If bUI = True Then

				If uNewCmd.sCmdText.Trim.ToUpper = "FRMHULLBUILDER.HULMAIN" Then
					frmHullBuilder.sParms = uNewCmd.sParams
                End If

                If uNewCmd.sCmdText.Trim.ToUpper = "ROUTESETLOCATION" Then
                    frmRouteConfig.sParms = uNewCmd.sParams
                End If

                For X As Int32 = 0 To muUICmds.GetUpperBound(0)
                    If muUICmds(X).sCmdText = "" Then
                        muUICmds(X) = uNewCmd
                        Exit For
                    End If
                Next X
                ReDim Preserve muUICmds(muUICmds.GetUpperBound(0) + 1)
                muUICmds(muUICmds.GetUpperBound(0)) = uNewCmd
            Else
                For X As Int32 = 0 To muHotkeyCmds.GetUpperBound(0)
                    If muHotkeyCmds(X).sCmdText = "" Then
                        muHotkeyCmds(X) = uNewCmd
                        Exit For
                    End If
                Next X
                ReDim Preserve muHotkeyCmds(muHotkeyCmds.GetUpperBound(0) + 1)
                muHotkeyCmds(muHotkeyCmds.GetUpperBound(0)) = uNewCmd
            End If
        End If
	End Sub
    Public Function CommandAllowed(ByVal bUI As Boolean, ByVal sCmdText As String) As Boolean
        'Return True
        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then Return True
        If sCmdText.ToUpper.Contains("FRMCREDITS") = True Then Return True
        'If True = True Then Return True
        sCmdText = sCmdText.ToUpper.Trim

        If sCmdText.StartsWith("FRMTUTORIALSTEP") = True Then Return True
        If sCmdText.StartsWith("FRMDEATH") = True Then Return True

        If sCmdText.EndsWith(".") = True Then sCmdText = sCmdText.Substring(0, sCmdText.Length - 1)
        'Return True
        'Ok, now here's the only caveat... we check for the command but we also need to see if the parent is simply enabled...
        Dim sParent As String = ""
        Dim lIdx As Int32 = sCmdText.LastIndexOf("."c)
        If lIdx > -1 Then sParent = sCmdText.Substring(lIdx + 1)

        If bUI = True Then
            If sCmdText.ToUpper.Trim.StartsWith("FRMOPTIONS") = True Then Return True
            If sCmdText.ToUpper.Contains("FRMCHAT") = True Then Return True
            If sCmdText.ToUpper = "FRMTUTORIALSTEP.BTNNEXT" Then Return True
            If sCmdText.ToUpper = "FRMTUTORIALSTEP.BTNPREVIOUS" Then Return True
            If sCmdText.ToUpper = "FRMTUTORIALSTEP.BTNPAUSE" Then Return True
            If sCmdText.ToUpper.Contains("FRMCONFIRMQUIT") Then Return True
            For X As Int32 = 0 To muUICmds.GetUpperBound(0)
                If muUICmds(X).sCmdText = "ALL" Then Return True
                If muUICmds(X).sCmdText = sCmdText Then Return True
                If muUICmds(X).sCmdText = sParent Then Return True
            Next X
        Else
            For X As Int32 = 0 To muHotkeyCmds.GetUpperBound(0)
                If muHotkeyCmds(X).sCmdText = "ALL" Then Return True
                If muHotkeyCmds(X).sCmdText = sCmdText Then Return True
                If muHotkeyCmds(X).sCmdText = sParent Then Return True
            Next X
        End If
        Return False
    End Function
    Public Function CommandAllowedWithParms(ByVal bUI As Boolean, ByVal sCmdText As String, ByVal sParms() As String, ByVal bPartialMatch As Boolean) As Boolean
        'If True = True Then Return True
        'Return True
        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then Return True
        If sCmdText.ToUpper.Contains("FRMCREDITS") = True Then Return True
        sCmdText = sCmdText.ToUpper.Trim
        If sCmdText.EndsWith(".") = True Then sCmdText = sCmdText.Substring(0, sCmdText.Length - 1)
        'Return True
        'Ok, now here's the only caveat... we check for the command but we also need to see if the parent is simply enabled...

        If sCmdText.StartsWith("FRMTUTORIALSTEP") = True Then Return True
        If sCmdText.StartsWith("FRMDEATH") = True Then Return True

        Dim sParent As String = ""
        Dim lIdx As Int32 = sCmdText.LastIndexOf("."c)
        If lIdx > -1 Then sParent = sCmdText.Substring(lIdx + 1)

        If bUI = True Then
            If sCmdText.ToUpper.Contains("FRMCHAT") = True Then Return True
            If sCmdText.ToUpper = "FRMTUTORIALSTEP.BTNNEXT" Then Return True
            If sCmdText.ToUpper = "FRMTUTORIALSTEP.BTNPREVIOUS" Then Return True
            If sCmdText.ToUpper = "FRMTUTORIALSTEP.BTNPAUSE" Then Return True
            If sCmdText.ToUpper.Contains("FRMCONFIRMQUIT") Then Return True
            If sCmdText.ToUpper.Trim.StartsWith("FRMOPTIONS") = True Then Return True
            For X As Int32 = 0 To muUICmds.GetUpperBound(0)
                If muUICmds(X).sCmdText = "ALL" Then
                    If muUICmds(X).ParamsMatch(sParms, bPartialMatch) = True Then Return True
                End If
                If muUICmds(X).sCmdText = sCmdText Then
                    If muUICmds(X).ParamsMatch(sParms, bPartialMatch) = True Then Return True
                End If
                If muUICmds(X).sCmdText = sParent Then
                    If muUICmds(X).ParamsMatch(sParms, bPartialMatch) = True Then Return True
                End If
            Next X
        Else
            For X As Int32 = 0 To muHotkeyCmds.GetUpperBound(0)
                If muHotkeyCmds(X).sCmdText = "ALL" Then
                    If muHotkeyCmds(X).ParamsMatch(sParms, bPartialMatch) = True Then Return True
                End If
                If muHotkeyCmds(X).sCmdText = sCmdText Then
                    If muHotkeyCmds(X).ParamsMatch(sParms, bPartialMatch) = True Then Return True
                End If
                If muHotkeyCmds(X).sCmdText = sParent Then
                    If muHotkeyCmds(X).ParamsMatch(sParms, bPartialMatch) = True Then Return True
                End If
            Next X
        End If
        Return False
    End Function
#End Region

#Region "  Alert Control Border Render  "
    'Private lnTemp As Line = Nothing
	Private mcolAlert() As System.Drawing.Color
	Private mlLastUpdate As Int32
    Private Sub UpdateAlertedControl(ByRef oControl As UIControl)
        Try
            If oControl Is Nothing = False Then
                If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return

                If mcolAlert Is Nothing Then
                    ReDim mcolAlert(2)
                    mcolAlert(0) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
                    mcolAlert(1) = System.Drawing.Color.FromArgb(255, 192, 192, 0)
                    mcolAlert(2) = System.Drawing.Color.FromArgb(255, 128, 128, 0)
                End If

                If glCurrentCycle - mlLastUpdate > 1 Then
                    Dim colTemp As System.Drawing.Color = mcolAlert(0)
                    mcolAlert(0) = mcolAlert(1)
                    mcolAlert(1) = mcolAlert(2)
                    mcolAlert(2) = colTemp
                    mlLastUpdate = glCurrentCycle
                End If

                'If lnTemp Is Nothing OrElse lnTemp.Disposed = True Then
                '    lnTemp = Nothing
                '    Device.IsUsingEventHandlers = False
                '    lnTemp = New Line(oDevice)
                '    Device.IsUsingEventHandlers = True
                'End If

                Dim ptLoc As Point = oControl.GetAbsolutePosition

                Dim lOff1 As Int32 = 6  ' 0
                Dim lOff2 As Int32 = 8  ' 2
                Dim lOff3 As Int32 = 10 ' 4

                Dim uVecs1(4) As Vector2
                uVecs1(0).X = ptLoc.X - lOff1 : uVecs1(0).Y = ptLoc.Y - lOff1
                uVecs1(1).X = ptLoc.X + oControl.Width + lOff1 : uVecs1(1).Y = ptLoc.Y - lOff1
                uVecs1(2).X = ptLoc.X + oControl.Width + lOff1 : uVecs1(2).Y = ptLoc.Y + oControl.Height + lOff1
                uVecs1(3).X = ptLoc.X - lOff1 : uVecs1(3).Y = ptLoc.Y + oControl.Height + lOff1
                uVecs1(4) = uVecs1(0)

                Dim uVecs2(4) As Vector2
                uVecs2(0).X = ptLoc.X - lOff2 : uVecs2(0).Y = ptLoc.Y - lOff2
                uVecs2(1).X = ptLoc.X + oControl.Width + lOff2 : uVecs2(1).Y = ptLoc.Y - lOff2
                uVecs2(2).X = ptLoc.X + oControl.Width + lOff2 : uVecs2(2).Y = ptLoc.Y + oControl.Height + lOff2
                uVecs2(3).X = ptLoc.X - lOff2 : uVecs2(3).Y = ptLoc.Y + oControl.Height + lOff2
                uVecs2(4) = uVecs2(0)

                Dim uVecs3(4) As Vector2
                uVecs3(0).X = ptLoc.X - lOff3 : uVecs3(0).Y = ptLoc.Y - lOff3
                uVecs3(1).X = ptLoc.X + oControl.Width + lOff3 : uVecs3(1).Y = ptLoc.Y - lOff3
                uVecs3(2).X = ptLoc.X + oControl.Width + lOff3 : uVecs3(2).Y = ptLoc.Y + oControl.Height + lOff3
                uVecs3(3).X = ptLoc.X - lOff3 : uVecs3(3).Y = ptLoc.Y + oControl.Height + lOff3
                uVecs3(4) = uVecs3(0)

                'BPLine.DrawLine(3, True, uVecs1, mcolAlert(0))
                'BPLine.DrawLine(3, True, uVecs2, mcolAlert(1))
                'BPLine.DrawLine(3, True, uVecs3, mcolAlert(2))

                BPLine.PrepareMultiDraw(3, True)
                BPLine.MultiDrawLine(uVecs1, mcolAlert(0))
                BPLine.MultiDrawLine(uVecs2, mcolAlert(1))
                BPLine.MultiDrawLine(uVecs3, mcolAlert(2))
                BPLine.EndMultiDraw()
 
            End If
        Catch
        End Try
    End Sub
#End Region

    Public Function GetNextExclamationPointLeft() As Int32
        Dim lBaseLeft As Int32 = 50 'oDevice.PresentationParameters.BackBufferWidth - 50
        For X As Int32 = 0 To mlWindowUB
            If myWindowUsed(X) > 0 AndAlso moWindows(X).ControlName.StartsWith("frmExclPoint") = True Then
                lBaseLeft += 50
            End If
        Next X
        Return lBaseLeft
    End Function
End Class
