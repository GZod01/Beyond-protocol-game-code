Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmMining
    Inherits UIWindow

    Private Enum eyMineralBidStatus As Byte
        eNoBid = 0
        eBiddingFirstPlace = 1
        eBiddingSecondPlace = 2
        eBiddingThirdPlace = 3
        eBiddingFourthPlace = 4
        eBiddingNotInTopFour = 5
        eBiddingNotMetMinimum = 6
    End Enum

    Private Const ml_ITEM_HEIGHT As Int32 = 27
    Private Const ml_ITEM_WIDTH As Int32 = 478

    '"Mining Facility - Calcicese ( 150 / 12345789 ) - Last Pickup 5d 19h 20m 54s ago"
    Private Structure MiningItem
        Public lEntityIndex As Int32        'of the item in the environment
        Public lMineralCacheIndex As Int32  'in the enviornment
        Public lLastPickup As Int32         'cycle that the last pickup occurred... this will be the hardest to manage
        Public lMineralCacheID As Int32

        Public yBidStatus As Byte

        Public bActive As Boolean
        Public bAutoLaunch As Boolean
        Public bRequestedDetails As Boolean

        Public sText As String              'placeholder only
    End Structure

    Private lblTitle As UILabel
    Private WithEvents btnClose As UIButton
    Private WithEvents btnHelp As UIButton
    Private lnDiv1 As UILine
    Private vscrScroll As UIScrollBar

    Private muItems() As MiningItem
    Private mlItemUB As Int32 = -1

    Private lItemsOnScreen As Int32 = 0
    Private moLineItemFont As System.Drawing.Font

    Private mlSelectedItemIdx As Int32 = -1

    Private mlLastBidRequestTime As Int32 = 0

    Private mlForceRefreshCycle As Int32

    Private mlLastScrollPosition As Int32 = -1

    'Private Shared moFont As Font
    'Private Shared moBoldFont As Font
    'Private Shared moSprite As Sprite
    'Public Shared Sub ReleaseSprite()
    '    If moSprite Is Nothing = False Then moSprite.Dispose()
    '    moSprite = Nothing
    '    'If moFont Is Nothing = False Then moFont.Dispose()
    '    'moFont = Nothing
    '    'If moBoldFont Is Nothing = False Then moBoldFont.Dispose()
    '    'moBoldFont = Nothing
    'End Sub

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenMiningWindow)

        'frmMining initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eMining
            .ControlName = "frmMining"
            .Height = 256
            .Width = 512
            '.Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            '.Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            If muSettings.MiningWindowX < 0 OrElse muSettings.MiningWindowY < 0 OrElse muSettings.MiningWindowX > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width OrElse muSettings.MiningWindowY > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height OrElse NewTutorialManager.TutorialOn = True Then
                .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
                .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
                muSettings.SpecialTechLeft = .Left
                muSettings.SpecialTechTop = .Top
            Else
                .Left = muSettings.MiningWindowX
                .Top = muSettings.MiningWindowY
            End If

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .mbAcceptReprocessEvents = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 2
            .Width = 275
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Mining Facilities in Current Environment"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

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

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 2
            .Top = 28
            .Width = 508
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'vscrScroll initial props
        vscrScroll = New UIScrollBar(oUILib, True)
        With vscrScroll
            .ControlName = "vscrScroll"
            .Left = 486
            .Top = 31
            .Width = 24
            .Height = 222
            .Enabled = True
            .Visible = True
            .Value = 0
            .MaxValue = 100
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 4
            .ReverseDirection = True
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(vscrScroll, UIControl))

        moLineItemFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0)

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewMining) = True Then
            RefreshItemList()
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack rights to view the mining interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        mlForceRefreshCycle = glCurrentCycle + 30
    End Sub

    Public Sub FindAndSelectFacility(ByVal lIdx As Int32)
        For X As Int32 = 0 To mlItemUB
            If muItems(X).lEntityIndex = lIdx Then
                mlSelectedItemIdx = X

                'Now, show our form
                Dim oFrm As frmMineralBid = CType(MyBase.moUILib.GetWindow("frmMineralBid"), frmMineralBid)
                If oFrm Is Nothing Then oFrm = New frmMineralBid(MyBase.moUILib)
                oFrm.Visible = True

                If muItems(X).lMineralCacheID = -1 Then
                    If goCurrentEnvir.lCacheUB >= muItems(X).lMineralCacheIndex AndAlso muItems(X).lMineralCacheIndex > -1 Then
                        muItems(X).lMineralCacheID = goCurrentEnvir.lCacheIdx(muItems(X).lMineralCacheIndex)
                    End If
                End If

                oFrm.SetIDs(muItems(X).lMineralCacheID, goCurrentEnvir.oEntity(muItems(X).lEntityIndex).ObjectID)
                oFrm.Left = Me.Left + Me.Width + 2
                oFrm.Top = Me.Top

                Me.IsDirty = True

                Exit For
            End If
        Next X
    End Sub

	Private Sub RefreshItemList()
        Static lCurrentEnvironment As Int32 = -1
        mlLastResortCall = 0

		mlItemUB = -1
        ReDim muItems(mlItemUB)
        mlLastScrollPosition = 0
        If lCurrentEnvironment = goCurrentEnvir.ObjectID Then
            mlLastScrollPosition = vscrScroll.Value
        End If
        lCurrentEnvironment = goCurrentEnvir.ObjectID


        vscrScroll.Value = 0
        vscrScroll.MinValue = 0
        vscrScroll.MaxValue = 0
        vscrScroll.Enabled = False

        If goCurrentEnvir Is Nothing = False Then
            If goCurrentEnvir.HasUnitsHere() = True Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 Then 'AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                        If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eMining Then
                            'Ok add our item
                            mlItemUB += 1
                            ReDim Preserve muItems(mlItemUB)
                            With muItems(mlItemUB)
                                .bActive = (goCurrentEnvir.oEntity(X).CurrentStatus And elUnitStatus.eFacilityPowered) <> 0
                                .bAutoLaunch = goCurrentEnvir.oEntity(X).AutoLaunch
                                .lEntityIndex = X
                                .lLastPickup = Int32.MinValue
                                .lMineralCacheIndex = GetClosestMineralCache(goCurrentEnvir.oEntity(X).LocX, goCurrentEnvir.oEntity(X).LocZ)
                                .yBidStatus = 0
                                .lMineralCacheID = -1
                                .bRequestedDetails = False

                                If .lMineralCacheIndex > -1 AndAlso goCurrentEnvir.lCacheIdx(.lMineralCacheIndex) <> -1 Then
                                    Dim oCache As MineralCache = goCurrentEnvir.oCache(.lMineralCacheIndex)
                                    If oCache Is Nothing = False AndAlso (oCache.Concentration = Int32.MinValue OrElse oCache.Quantity = Int32.MinValue) Then
                                        .lMineralCacheID = oCache.ObjectID
                                        If .bRequestedDetails = False Then
                                            .bRequestedDetails = True
                                            MyBase.moUILib.GetMsgSys.RequestEntityDetails(goCurrentEnvir.oCache(.lMineralCacheIndex), goCurrentEnvir)
                                        End If
                                    End If
                                End If
                            End With

                            With goCurrentEnvir.oEntity(muItems(mlItemUB).lEntityIndex)
                                If .ObjTypeID = ObjectType.eFacility AndAlso .MaxColonists = 0 AndAlso .MaxWorkers = 0 AndAlso .PowerConsumption = 0 Then
                                    Dim yMsg(8) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityPersonnel).CopyTo(yMsg, 0)
                                    .GetGUIDAsString.CopyTo(yMsg, 2)
                                    yMsg(8) = 2     'pass two to indicate a simple request
                                    MyBase.moUILib.SendMsgToPrimary(yMsg)
                                End If
                            End With
                        End If
                    End If
                Next X
            End If
        End If

        lItemsOnScreen = vscrScroll.Height \ ml_ITEM_HEIGHT
        If mlItemUB + 1 > lItemsOnScreen Then
            With vscrScroll
                .MinValue = 0
                .MaxValue = (mlItemUB + 1) - lItemsOnScreen
                .Enabled = True
            End With
        End If
        If mlLastScrollPosition > 0 Then
            vscrScroll.Value = mlLastScrollPosition
        End If
    End Sub

    Private Function GetClosestMineralCache(ByVal fLocX As Single, ByVal fLocZ As Single) As Int32
        Dim lHalfMin As Int32 = 50 'gl_MINIMUM_MINING_FAC_SNAP_TO_DISTANCE \ 2
        Dim rcTemp As Rectangle = Rectangle.FromLTRB(CInt(fLocX - lHalfMin), CInt(fLocZ - lHalfMin), CInt(fLocX + lHalfMin), CInt(fLocZ + lHalfMin))
        Dim fDist As Single
        Dim lCurrIdx As Int32 = -1
        Dim fTemp As Single
        Dim lLowestID As Int32 = -1

        'Now, check the mineral caches
        For X As Int32 = 0 To goCurrentEnvir.lCacheUB
            If goCurrentEnvir.lCacheIdx(X) <> -1 AndAlso goCurrentEnvir.oCache(X).CacheTypeID = MineralCacheType.eMineable Then
                If rcTemp.Contains(goCurrentEnvir.oCache(X).LocX, goCurrentEnvir.oCache(X).LocZ) = True Then
                    'ok, contains...
                    If lCurrIdx = -1 Then
                        'Ok, this one is good
                        lCurrIdx = X
                        lLowestID = goCurrentEnvir.oCache(X).ObjectID
                        fDist = Distance(CInt(fLocX), CInt(fLocZ), goCurrentEnvir.oCache(X).LocX, goCurrentEnvir.oCache(X).LocZ)
                    Else
                        fTemp = Distance(CInt(fLocX), CInt(fLocZ), goCurrentEnvir.oCache(X).LocX, goCurrentEnvir.oCache(X).LocZ)
                        If fTemp <= fDist Then
                            If Math.Abs(fTemp - fDist) < 15 Then
                                If lLowestID < goCurrentEnvir.oCache(X).ObjectID Then Continue For
                            End If
                            lLowestID = goCurrentEnvir.oCache(X).ObjectID
                            lCurrIdx = X
                            fDist = fTemp
                        End If
                    End If
                End If
            End If
        Next X

        Return lCurrIdx
    End Function

    Private mlLastMouseDown As Int32 = 0
    Private Sub frmMining_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        'Adjust MouseX and Y to be within the bounds of this form
        lMouseX -= Me.GetAbsolutePosition.X
        lMouseY -= Me.GetAbsolutePosition.Y

        Dim bDoubleClick As Boolean = Math.Abs(glCurrentCycle - mlLastMouseDown) < 10
        mlLastMouseDown = glCurrentCycle

        'click on an item?
        If lMouseY > vscrScroll.Top Then
            'If we were clicking on an item, which would it be?
            lMouseY -= vscrScroll.Top
            Dim lItemIdx As Int32 = lMouseY \ ml_ITEM_HEIGHT

            lItemIdx += vscrScroll.Value
            If lItemIdx > mlItemUB Then Return

            'Now, check what the user clicked on...
            If lMouseX > 10 AndAlso lMouseX < 26 Then
                'Autolaunch button
                HandleAutoLaunchButtonClick(muItems(lItemIdx).lEntityIndex)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            ElseIf lMouseX > 26 AndAlso lMouseX < 42 Then
                'Active button
                HandleActiveButtonClick(muItems(lItemIdx).lEntityIndex)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            ElseIf lMouseX > 55 AndAlso lMouseX < 480 Then
                'Label
                Try
                    If bDoubleClick = True Then
                        goCurrentEnvir.DeselectAll()
                        If glCurrentEnvirView <> CurrentView.ePlanetView Then
                            'goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem
                            glCurrentEnvirView = CurrentView.ePlanetView
                            goCamera.mlCameraY = 6000
                            goCamera.mlCameraAtY = 0
                        End If
                        With goCurrentEnvir.oEntity(muItems(lItemIdx).lEntityIndex)
                            .bSelected = True
                            Dim lDiffX As Int32 = goCamera.mlCameraX - goCamera.mlCameraAtX
                            Dim lDiffZ As Int32 = goCamera.mlCameraZ - goCamera.mlCameraAtZ

                            Dim lPrevHt As Int32 = 0
                            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                                lPrevHt = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraX, goCamera.mlCameraZ, True))
                            End If

                            goCamera.mlCameraAtX = CInt(.LocX)
                            goCamera.mlCameraAtZ = CInt(.LocZ)

                            goCamera.mlCameraX = goCamera.mlCameraAtX + lDiffX
                            goCamera.mlCameraZ = goCamera.mlCameraAtZ + lDiffZ

                            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                                Dim lNewHt As Int32 = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCamera.mlCameraX, goCamera.mlCameraZ, True))
                                Dim lDiffY As Int32 = goCamera.mlCameraY - lPrevHt
                                goCamera.mlCameraY = lNewHt + lDiffY
                            End If
                        End With
                    Else
                        mlSelectedItemIdx = lItemIdx

                        'Now, show our form
                        Dim oFrm As frmMineralBid = CType(MyBase.moUILib.GetWindow("frmMineralBid"), frmMineralBid)
                        If oFrm Is Nothing Then oFrm = New frmMineralBid(MyBase.moUILib)
                        oFrm.Visible = True

                        If muItems(lItemIdx).lMineralCacheID = -1 Then
                            If goCurrentEnvir.lCacheUB >= muItems(lItemIdx).lMineralCacheIndex AndAlso muItems(lItemIdx).lMineralCacheIndex > -1 Then
                                muItems(lItemIdx).lMineralCacheID = goCurrentEnvir.lCacheIdx(muItems(lItemIdx).lMineralCacheIndex)
                            End If
                        End If

                        oFrm.SetIDs(muItems(lItemIdx).lMineralCacheID, goCurrentEnvir.oEntity(muItems(lItemIdx).lEntityIndex).ObjectID)
                        oFrm.Left = Me.Left + Me.Width + 2
                        oFrm.Top = Me.Top
                        frmMining_WindowMoved()

                        Me.IsDirty = True
                    End If
                Catch
                    'do nothing
                End Try
            End If
        End If
    End Sub

    Private Sub HandleAutoLaunchButtonClick(ByVal lEntityIndex As Int32)
        If goCurrentEnvir Is Nothing Then Return
        If lEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(lEntityIndex) = -1 Then Return
        If HasAliasedRights(AliasingRights.eAlterAutoLaunchPower) = False Then
            MyBase.moUILib.AddNotification("You lack rights to alter auto-launch and power settings.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim yMsg(8) As Byte
        Dim lPos As Int32 = 0

        Dim bCurrentVal As Boolean

        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityPersonnel).CopyTo(yMsg, lPos) : lPos += 2
        goCurrentEnvir.oEntity(lEntityIndex).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        bCurrentVal = goCurrentEnvir.oEntity(lEntityIndex).AutoLaunch

        'Now, toggle our currentval...
        bCurrentVal = Not bCurrentVal

        If bCurrentVal = True Then
            yMsg(lPos) = 1
        Else : yMsg(lPos) = 0
        End If

        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Private Sub HandleActiveButtonClick(ByVal lEntityIndex As Int32)
        If goCurrentEnvir Is Nothing Then Exit Sub
        If lEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(lEntityIndex) = -1 Then Return
        If HasAliasedRights(AliasingRights.eAlterAutoLaunchPower) = False Then
            MyBase.moUILib.AddNotification("You lack rights to alter auto-launch and power settings.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim yMsg(11) As Byte
        Dim lStatus As Int32

        Dim lEntityStatus As Int32 = goCurrentEnvir.oEntity(lEntityIndex).CurrentStatus

        If (lEntityStatus And elUnitStatus.eFacilityPowered) <> 0 Then
            lStatus = -elUnitStatus.eFacilityPowered
            goCurrentEnvir.oEntity(lEntityIndex).CurrentStatus -= elUnitStatus.eFacilityPowered
        Else
            lStatus = elUnitStatus.eFacilityPowered
            goCurrentEnvir.oEntity(lEntityIndex).CurrentStatus = goCurrentEnvir.oEntity(lEntityIndex).CurrentStatus Or lStatus
        End If

        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yMsg, 0)
        goCurrentEnvir.oEntity(lEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(lStatus).CopyTo(yMsg, 8)

        MyBase.moUILib.SendMsgToPrimary(yMsg)

	End Sub

	Private Sub SortList()
		Try
			Dim lSorted() As Int32 = Nothing
			Dim lSortedUB As Int32 = -1
			Dim sSortedName() As String = Nothing

			For X As Int32 = 0 To mlItemUB
				Dim lIdx As Int32 = -1

				'get our new name
				Dim sName As String = ""
				'muitems(x).lMineralCacheIndex
				If muItems(X).lMineralCacheIndex <> -1 Then
					With goCurrentEnvir.oCache(muItems(X).lMineralCacheIndex)
						If .oMineral Is Nothing Then
							sName = "Unknown Mineral"
						Else : sName = .oMineral.MineralName
						End If
					End With
				End If

				For Y As Int32 = 0 To lSortedUB
					If sSortedName(Y) > sName Then
						lIdx = Y
						Exit For
					End If
				Next Y
				lSortedUB += 1
				ReDim Preserve lSorted(lSortedUB)
				ReDim Preserve sSortedName(lSortedUB)
				If lIdx = -1 Then
					lSorted(lSortedUB) = X
					sSortedName(lSortedUB) = sName
				Else
					For Y As Int32 = lSortedUB To lIdx + 1 Step -1
						lSorted(Y) = lSorted(Y - 1)
						sSortedName(Y) = sSortedName(Y - 1)
					Next Y
					lSorted(lIdx) = X
					sSortedName(lIdx) = sName
				End If
			Next X

			If lSortedUB = mlItemUB AndAlso lSorted Is Nothing = False Then
				Dim uTemp(mlItemUB) As MiningItem
				For X As Int32 = 0 To mlItemUB
					uTemp(X) = muItems(lSorted(X))
				Next X
				muItems = uTemp
			End If
		Catch
		End Try
	End Sub

	Private mlLastResortCall As Int32
	Private Sub frmMining_OnNewFrame() Handles Me.OnNewFrame

        If mlForceRefreshCycle <> Int32.MinValue Then
            If mlForceRefreshCycle < glCurrentCycle Then
                mlForceRefreshCycle = glCurrentCycle + 30
                Me.IsDirty = True
            End If
        End If
        'Check for new items in the list
		Dim lCnt As Int32 = 0

		Dim oEnvir As BaseEnvironment = goCurrentEnvir
		If oEnvir Is Nothing = False Then
			Try
				For X As Int32 = 0 To oEnvir.lEntityUB
					If oEnvir.lEntityIdx(X) <> -1 Then 'AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
						Dim oEntity As BaseEntity = oEnvir.oEntity(X)
						If oEntity Is Nothing = False AndAlso oEntity.yProductionType = ProductionType.eMining Then lCnt += 1
					End If
				Next X
				If lCnt - 1 <> mlItemUB Then RefreshItemList()
			Catch
			End Try
		End If

		If glCurrentCycle - mlLastResortCall > 30 Then
			mlLastResortCall = glCurrentCycle
			SortList()
		End If

		'check our entities that will be drawn... have their Active or Autolaunch settings changed? If so, me.isdirty = true
		If Me.IsDirty = False AndAlso oEnvir Is Nothing = False Then
            Try

                Dim bRequestBidStatus As Boolean = (glCurrentCycle - mlLastBidRequestTime > 90) OrElse vscrScroll.Value <> mlLastRequestValue

                Dim yBidStatus() As Byte = Nothing
                Dim lPos As Int32 = 0
                If bRequestBidStatus = True Then
                    ReDim yBidStatus((lItemsOnScreen * 4) + 5)
                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestBidStatus).CopyTo(yBidStatus, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lItemsOnScreen).CopyTo(yBidStatus, lPos) : lPos += 4
                End If

                Dim lMinX As Int32 = vscrScroll.Value
                Dim lMaxX As Int32 = Math.Min(vscrScroll.Value + lItemsOnScreen - 1, mlItemUB)
                For X As Int32 = lMinX To lMaxX
                    If oEnvir.lEntityIdx(muItems(X).lEntityIndex) <> -1 Then
                        With oEnvir.oEntity(muItems(X).lEntityIndex)
                            If .ObjTypeID = ObjectType.eFacility Then
                                If ((.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0) <> muItems(X).bActive Then
                                    Me.IsDirty = True
                                    If bRequestBidStatus = False Then Exit For
                                End If

                                If bRequestBidStatus = True Then
                                    System.BitConverter.GetBytes(.ObjectID).CopyTo(yBidStatus, lPos) : lPos += 4
                                End If
                            End If
                        End With
                    End If
                Next X

                If bRequestBidStatus = True Then
                    MyBase.moUILib.SendMsgToPrimary(yBidStatus)
                    mlLastBidRequestTime = glCurrentCycle
                    mlLastRequestValue = vscrScroll.Value
                End If
            Catch
                'Do nothing
            End Try
		End If
    End Sub
    Private mlLastRequestValue As Int32 = 0

    Private Sub frmMining_OnNewFrameEnd() Handles Me.OnNewFrameEnd
        'Now, we have to render our text seperately because it is rendered each frame...
        If moLineItemFont Is Nothing Then moLineItemFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0)
        If goCurrentEnvir Is Nothing Then Return

        Dim lMinX As Int32 = vscrScroll.Value
        Dim lMaxX As Int32 = Math.Min(vscrScroll.Value + lItemsOnScreen - 1, mlItemUB)

        Try
            For X As Int32 = lMinX To lMaxX
                Dim sTemp As String = ""
                If goCurrentEnvir.lEntityIdx(muItems(X).lEntityIndex) = -1 Then Continue For

                With goCurrentEnvir.oEntity(muItems(X).lEntityIndex)
                    If .EntityName Is Nothing OrElse .EntityName.Contains("Unknown") = True Then
                        .EntityName = GetCacheObjectValue(.ObjectID, .ObjTypeID)
                    End If
                    sTemp = .EntityName & " - "
                End With

                If muItems(X).lMineralCacheIndex <> -1 Then
                    With goCurrentEnvir.oCache(muItems(X).lMineralCacheIndex)
                        If .oMineral Is Nothing = False Then
                            sTemp &= .oMineral.MineralName
                        Else : sTemp &= "Unknown Mineral"
                        End If

                        If .Concentration = Int32.MinValue OrElse .Quantity = Int32.MinValue Then
                            sTemp &= " ( Acquiring... ) "
                            If muItems(X).bRequestedDetails = False Then
                                muItems(X).bRequestedDetails = True
                                MyBase.moUILib.GetMsgSys.RequestEntityDetails(goCurrentEnvir.oCache(muItems(X).lMineralCacheIndex), goCurrentEnvir)
                            End If
                        Else
                            sTemp &= " ( " & .Concentration & " / " & .Quantity.ToString("#,###") & " ) "
                        End If
                    End With
                End If

                With muItems(X)
                    If .sText <> sTemp Then
                        .sText = sTemp
                        Me.IsDirty = True
                    End If
                End With
            Next X
        Catch
            'do nothing
        End Try

    End Sub

    Private Sub frmMining_OnRender() Handles Me.OnRender
        If goCurrentEnvir Is Nothing Then Return
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

        Dim v2Lines()() As Vector2
        ReDim v2Lines(lItemsOnScreen - 1)
        For X As Int32 = 0 To lItemsOnScreen - 1
            ReDim v2Lines(X)(8)
        Next X

        Dim lMinX As Int32 = vscrScroll.Value
        Dim lMaxX As Int32 = Math.Min(vscrScroll.Value + lItemsOnScreen - 1, mlItemUB)
        Dim oLoc As Point

        Using oSprite As New Sprite(MyBase.moUILib.oDevice)
            Dim oSelColor As System.Drawing.Color
            oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)

            oSprite.Begin(SpriteFlags.AlphaBlend)

            For X As Int32 = lMinX To lMaxX
                'Ok, X is our index in muItems... for the OnRender, we render the border, the active button and the autolaunch button

                'Render the border... which means we need to add the lines...
                oLoc = New Point(5, ((X - lMinX) * ml_ITEM_HEIGHT) + vscrScroll.Top)

                With v2Lines(X - lMinX)(0)
                    .X = oLoc.X + 3
                    .Y = oLoc.Y
                End With
                With v2Lines(X - lMinX)(1)
                    .X = oLoc.X + ml_ITEM_WIDTH - 3
                    .Y = oLoc.Y
                End With
                With v2Lines(X - lMinX)(2)
                    .X = oLoc.X + ml_ITEM_WIDTH
                    .Y = oLoc.Y + 3
                End With
                With v2Lines(X - lMinX)(3)
                    .X = oLoc.X + ml_ITEM_WIDTH
                    .Y = oLoc.Y + ml_ITEM_HEIGHT - 3
                End With
                With v2Lines(X - lMinX)(4)
                    .X = oLoc.X + ml_ITEM_WIDTH - 3
                    .Y = oLoc.Y + ml_ITEM_HEIGHT
                End With
                With v2Lines(X - lMinX)(5)
                    .X = oLoc.X + 3
                    .Y = oLoc.Y + ml_ITEM_HEIGHT
                End With
                With v2Lines(X - lMinX)(6)
                    .X = oLoc.X
                    .Y = oLoc.Y + ml_ITEM_HEIGHT - 3
                End With
                With v2Lines(X - lMinX)(7)
                    .X = oLoc.X
                    .Y = oLoc.Y + 3
                End With
                With v2Lines(X - lMinX)(8)
                    .X = oLoc.X + 3
                    .Y = oLoc.Y
                End With


                If X = mlSelectedItemIdx Then
                    Dim clrFill As System.Drawing.Color = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
                    MyBase.moUILib.DoAlphaBlendColorFill(New Rectangle(oLoc, New Size(ml_ITEM_WIDTH, ml_ITEM_HEIGHT)), clrFill, oLoc)
                End If

                Try
                    If goCurrentEnvir.lEntityIdx(muItems(X).lEntityIndex) <> -1 Then
                        With goCurrentEnvir.oEntity(muItems(X).lEntityIndex)
                            If .ObjTypeID = ObjectType.eFacility Then
                                Dim clrVal As System.Drawing.Color

                                'If .AutoLaunch = True Then
                                '    clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                                'Else : clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                                'End If
                                'muItems(X).bAutoLaunch = .AutoLaunch
                                'oSprite.Draw2D(goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.Pak"), _
                                '    New Rectangle(16, 32, 16, 16), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(oLoc.X + 5, oLoc.Y + (ml_ITEM_HEIGHT \ 2) - 8), clrVal)

                                If (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                                    clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                                    muItems(X).bActive = True
                                Else
                                    clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                                    muItems(X).bActive = False
                                End If
                                oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eSphere), New Rectangle(0, 0, 16, 16), _
                                   Point.Empty, 0, New Point(oLoc.X + 21, oLoc.Y + (ml_ITEM_HEIGHT \ 2) - 8), clrVal)
                            End If
                        End With
                    End If
                Catch
                    'Do nothing
                End Try
            Next X

            oSprite.End()
        End Using

        BPLine.PrepareMultiDraw(1, True)
        For X As Int32 = 0 To lItemsOnScreen - 1
            BPLine.MultiDrawLine(v2Lines(X), muSettings.InterfaceBorderColor)
        Next X
        BPLine.EndMultiDraw()
        'ValidateBorderLine()
        'With moBorderLine
        '    .Antialias = True
        '    .Width = 1
        '    .Begin()
        '    For X As Int32 = 0 To lItemsOnScreen - 1
        '        .Draw(v2Lines(X), muSettings.InterfaceBorderColor)
        '    Next X
        '    .End()
        'End With

        'If moFont Is Nothing OrElse moFont.Disposed = True Then
        '    Device.IsUsingEventHandlers = False
        '    moFont = New Font(MyBase.moUILib.oDevice, moLineItemFont)
        '    Device.IsUsingEventHandlers = True
        'End If
        'If moBoldFont Is Nothing OrElse moBoldFont.Disposed = True Then
        '    Device.IsUsingEventHandlers = False
        '    moBoldFont = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font(moLineItemFont, FontStyle.Bold))
        '    Device.IsUsingEventHandlers = True
        'End If
        'If moSprite Is Nothing OrElse moSprite.Disposed = True Then
        '    Device.IsUsingEventHandlers = False
        '    moSprite = New Sprite(MyBase.moUILib.oDevice)
        '    Device.IsUsingEventHandlers = True
        'End If
 
        oLoc = Me.GetAbsolutePosition()
        oLoc.X += 55
        oLoc.Y += vscrScroll.Top

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        Dim bCheckGuild As Boolean = False
        Dim oGuild As Guild = Nothing
        If goCurrentPlayer Is Nothing = False Then
            oGuild = goCurrentPlayer.oGuild
            If oGuild Is Nothing = False Then
                bCheckGuild = (goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.ShareUnitVision) <> 0
            End If
        End If

        Try
            'moSprite.Begin(SpriteFlags.AlphaBlend)
            Using moFont As New Font(MyBase.moUILib.oDevice, moLineItemFont)
                Using moBoldFont As New Font(MyBase.moUILib.oDevice, New System.Drawing.Font(moLineItemFont, FontStyle.Bold))
                    For X As Int32 = lMinX To lMaxX
                        Dim rcDest As New Rectangle(oLoc.X, oLoc.Y + (ml_ITEM_HEIGHT * (X - lMinX)) + 2, 425, 26)
                        If muItems(X).sText <> "" Then
                            Dim clrItem As System.Drawing.Color = muSettings.InterfaceBorderColor
                            If muItems(X).lEntityIndex > -1 Then
                                If oEnvir Is Nothing = False Then
                                    If oEnvir.lEntityUB >= muItems(X).lEntityIndex Then
                                        Dim oEntity As BaseEntity = oEnvir.oEntity(muItems(X).lEntityIndex)
                                        If oEntity Is Nothing = False Then

                                            If oEntity.OwnerID = glPlayerID Then
                                                clrItem = ConvertVector4ToColor(muSettings.MyAssetColor)
                                            ElseIf (bCheckGuild = True AndAlso oGuild.MemberInGuild(oEntity.OwnerID) = True) Then
                                                clrItem = ConvertVector4ToColor(muSettings.GuildAssetColor)
                                            Else
                                                Dim yRel As Byte = goCurrentPlayer.GetPlayerRelScore(oEntity.OwnerID)

                                                If yRel <= elRelTypes.eWar Then
                                                    clrItem = ConvertVector4ToColor(muSettings.EnemyAssetColor)
                                                ElseIf yRel <= elRelTypes.ePeace Then       'elRelTypes.eNeutral
                                                    clrItem = ConvertVector4ToColor(muSettings.NeutralAssetColor)
                                                Else : clrItem = ConvertVector4ToColor(muSettings.AllyAssetColor)
                                                End If
                                            End If

                                        End If
                                    End If
                                End If
                            End If
                            moFont.DrawText(Nothing, muItems(X).sText, rcDest, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, clrItem)  'muSettings.InterfaceBorderColor)

                            Dim sBidStatusText As String = ""
                            Dim clrBidStatus As System.Drawing.Color = muSettings.InterfaceBorderColor
                            Select Case CType(muItems(X).yBidStatus, eyMineralBidStatus)
                                Case eyMineralBidStatus.eBiddingFirstPlace
                                    clrBidStatus = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                                    sBidStatusText = "1st"
                                Case eyMineralBidStatus.eBiddingFourthPlace
                                    clrBidStatus = System.Drawing.Color.FromArgb(255, 255, 128, 64)
                                    sBidStatusText = "4th"
                                Case eyMineralBidStatus.eBiddingNotInTopFour
                                    clrBidStatus = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                                    sBidStatusText = "OUT"
                                Case eyMineralBidStatus.eBiddingNotMetMinimum
                                    clrBidStatus = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                                    sBidStatusText = "<MIN"
                                Case eyMineralBidStatus.eBiddingSecondPlace
                                    clrBidStatus = System.Drawing.Color.FromArgb(255, 0, 255, 255)
                                    sBidStatusText = "2nd"
                                Case eyMineralBidStatus.eBiddingThirdPlace
                                    clrBidStatus = System.Drawing.Color.FromArgb(255, 255, 255, 0)
                                    sBidStatusText = "3rd"
                            End Select
                            If sBidStatusText <> "" Then moBoldFont.DrawText(Nothing, sBidStatusText, rcDest, DrawTextFormat.Right Or DrawTextFormat.VerticalCenter, clrBidStatus)
                        End If
                    Next X
                End Using
            End Using

            'moSprite.End()
        Catch
            'moSprite = Nothing
        End Try


    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Public Sub SetEntityStatusMsgRcvd(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lStatus As Int32)
		If goCurrentEnvir Is Nothing Then Return

		For X As Int32 = 0 To mlItemUB
			If muItems(X).lEntityIndex <> -1 AndAlso goCurrentEnvir.lEntityIdx(muItems(X).lEntityIndex) = lID Then
				If goCurrentEnvir.oEntity(muItems(X).lEntityIndex).ObjTypeID = iTypeID Then
					If lStatus = -elUnitStatus.eFacilityPowered Then
						If muItems(X).bActive = True Then
							'We were expecting it to be powered, so alert the player to low power
							goUILib.AddNotification("Not enough power!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
							If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowPower.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
						End If
					Else
						If muItems(X).bActive = False Then
							goUILib.AddNotification("Facilities require this facility to be active!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
							If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowPower.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
						End If
					End If
				End If
			End If
		Next X

		Me.IsDirty = True
	End Sub

    'Protected Overrides Sub Finalize()
    '	If moSprite Is Nothing = False Then moSprite.Dispose()
    '	moSprite = Nothing
    '       'If moFont Is Nothing = False Then moFont.Dispose()
    '       'moFont = Nothing
    '       'If moBoldFont Is Nothing = False Then moBoldFont.Dispose()
    '       'moBoldFont = Nothing
    '	MyBase.Finalize()
    'End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eMining)
    End Sub

    Public Sub HandleMiningBidListMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        For X As Int32 = 0 To lCnt - 1
            Dim lCacheID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yType As Byte = yData(lPos) : lPos += 1

            For Y As Int32 = 0 To mlItemUB
                If muItems(Y).lMineralCacheID = -1 Then
                    If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.lCacheUB >= muItems(Y).lMineralCacheIndex AndAlso muItems(Y).lMineralCacheIndex > -1 Then
                        muItems(Y).lMineralCacheID = goCurrentEnvir.lCacheIdx(muItems(Y).lMineralCacheIndex)
                    End If
                End If
                If muItems(Y).lMineralCacheID = lCacheID Then
                    muItems(Y).yBidStatus = yType
                    Exit For
                End If
            Next Y
        Next X

        Me.IsDirty = True
    End Sub

    Private Sub frmMining_WindowClosed() Handles Me.WindowClosed
        goUILib.RemoveWindow("frmMineralBid")
    End Sub

    Private Sub frmMining_WindowMoved() Handles Me.WindowMoved
        muSettings.MiningWindowX = Me.Left
        muSettings.MiningWindowY = Me.Top

        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmMineralBid")

        Dim lFinalRight As Int32
        If ofrm Is Nothing = False Then
            lFinalRight = Me.Left + Me.Width + 256
            If lFinalRight > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then
                Me.Left -= (lFinalRight - MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth) - 1
            End If

            ofrm.Left = Me.Left + Me.Width
            ofrm.Top = Me.Top
        Else
            lFinalRight = Me.Left + Me.Width
            If lFinalRight > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then
                Me.Left -= (lFinalRight - MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth) - 1
            End If
        End If

        

    End Sub

    Private Function ConvertVector4ToColor(ByVal vec4 As Vector4) As System.Drawing.Color
        Return System.Drawing.Color.FromArgb(255, CInt(vec4.X * 255), CInt(vec4.Y * 255), CInt(vec4.Z * 255))
    End Function

End Class