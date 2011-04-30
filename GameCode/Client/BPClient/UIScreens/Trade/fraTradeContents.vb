Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class fraTradeContents
    Inherits UIWindow

    Private Enum CurrentPriceSort As Byte
        ePriceAscending = 0
        ePriceDescending = 1
        eItemNameAscending = 2
        eItemNameDescending = 3
        eQuantityAscending = 4
        eQuantityDescending = 5

        eEscrowAscending = 6
        eEscrowDescending = 7
        eDeadlineAscending = 8
        eDeadlineDescending = 9
    End Enum

    Private lblContents As UILabel
	Private lblCurrentPrice As UILabel
	Private lblOrderCap As UILabel
    Private WithEvents lstContents As UIListBox
    Private WithEvents lstCurrent As UIListBox
    Private WithEvents btnAction1 As UIButton
    Private WithEvents btnAction2 As UIButton
    Private WithEvents btnAction3 As UIButton
    Private WithEvents chkShowMineOnly As UICheckBox
    Private WithEvents chkThisEnvironmentOnly As UICheckBox

    Private fraDetail As fraDetailPane

    Private myShowState As Byte
    Private mlTradepostID As Int32 = -1

    Private moList() As TradeOrder
    Private myListUsed() As Byte
    Private mlListUB As Int32 = -1

    Private mlTPCLast As Int32 = -1

    Private mySortBy As CurrentPriceSort

    Private myBO_MineOnly As Byte       '0 = no filter, 1 = mine only

    Private mlLastGTCRefresh As Int32 = -1
    Private mlPreviousContentsIndex As Int32 = -1

	Private mfraDeliver As fraDeliverBuyOrder = Nothing

	Private Shared mlLastContentsListItemData As Int32 = -1
    Private Shared mlLastContentsListItemData2 As Int32 = -1
    Private Shared mbLastThisEnvironmentOnly As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        ' initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTradeContents
            .ControlName = "fraTradeContents"
            .Left = 65
            .Top = 58
            .Width = 790
            .Height = 543
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .BorderLineWidth = 2
            .mbAcceptReprocessEvents = True
        End With

        'lstContents initial props
        lstContents = New UIListBox(oUILib)
        With lstContents
            .ControlName = "lstContents"
            .Left = 5
            .Top = 25
            .Width = 250
            .Height = 510
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(lstContents, UIControl))

        'lblContents initial props
        lblContents = New UILabel(oUILib)
        With lblContents
            .ControlName = "lblContents"
            .Left = 5
            .Top = 5
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Sellables"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblContents, UIControl))

        'lblCurrentPrice initial props
        lblCurrentPrice = New UILabel(oUILib)
        With lblCurrentPrice
            .ControlName = "lblCurrentPrice"
            .Left = 265
            .Top = 5
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Current Prices"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCurrentPrice, UIControl))

        'lstCurrent initial props
        lstCurrent = New UIListBox(oUILib)
        With lstCurrent
            .ControlName = "lstCurrent"
            .Left = 265
            .Top = 25
            .Width = 520
            .Height = 300
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(lstCurrent, UIControl))

        'btnAction1 initial props
        btnAction1 = New UIButton(oUILib)
        With btnAction1
            .ControlName = "btnAction1"
            .Left = 685
            .Top = 510
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Action 1"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAction1, UIControl))

        'btnAction2 initial props
        btnAction2 = New UIButton(oUILib)
        With btnAction2
            .ControlName = "btnAction2"
            .Left = 580
            .Top = 510
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Action 2"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAction2, UIControl))

        'btnAction3 initial props
        btnAction3 = New UIButton(oUILib)
        With btnAction3
            .ControlName = "btnAction3"
            .Left = 475
            .Top = 510
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Action 3"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
		Me.AddChild(CType(btnAction3, UIControl))

		'lblOrderCap initial props
		lblOrderCap = New UILabel(oUILib)
		With lblOrderCap
			.ControlName = "lblOrderCap"
			.Left = 265
			.Top = 5
			.Width = 100
			.Height = 18
			.Enabled = True
			.Visible = False
			.Caption = "0 used / 0 max"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
        Me.AddChild(CType(lblOrderCap, UIControl))

        'chkThisEnvironmentOnly initial props
        chkThisEnvironmentOnly = New UICheckBox(oUILib)
        With chkThisEnvironmentOnly
            .ControlName = "chkThisEnvironmentOnly"
            .Left = 690
            .Top = 4
            .Width = 95
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "This Star Only"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = mbLastThisEnvironmentOnly
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkThisEnvironmentOnly, UIControl))

    End Sub

    Public Sub SetFrameDetails(ByVal yType As Byte, ByVal lTradepostID As Int32)
        mlTradepostID = lTradepostID
        myShowState = yType

        lblOrderCap.Visible = False
        chkThisEnvironmentOnly.Visible = False
        btnAction1.Height = btnAction2.Height
        If myShowState = 0 Then
            'Buy screen - shows other's sell orders
            lblContents.Caption = "Market List"

            btnAction3.Left = btnAction2.Left
            btnAction2.Left = btnAction1.Left

            btnAction1.Top = lstCurrent.Top + lstCurrent.Height + 5
            btnAction1.Caption = "Refresh"
            btnAction2.Caption = "Purchase"
            btnAction2.Visible = False
            btnAction3.Visible = False
            If NewTutorialManager.TutorialOn = False Then chkThisEnvironmentOnly.Visible = True
			FillListWithMarketList(lstContents, myShowState)

			If mlLastContentsListItemData > -1 Then
				For X As Int32 = 0 To lstContents.ListCount - 1
					If lstContents.ItemData(X) = mlLastContentsListItemData Then
						lstContents.ListIndex = X
						Exit For
					End If
				Next X
			End If
        ElseIf myShowState = 1 Then
            'Sell screen - shows my sellables and allows me to create new sell orders
            lblContents.Caption = "Sellables"

            'btnAction3.Left = btnAction2.Left
            'btnAction2.Left = btnAction1.Left

            btnAction1.Top = lstCurrent.Top + lstCurrent.Height + 5
            btnAction1.Caption = "Refresh"

            btnAction2.Visible = False
            btnAction3.Visible = False

            mlTPCLast = Int32.MinValue
        Else
            'BUY ORDER
            lblContents.Caption = "Market List"
            btnAction1.Caption = "Deliver"
            btnAction2.Left = lstCurrent.Left
            btnAction2.Caption = "Create Order"
            btnAction3.Left = lstCurrent.Left + lstCurrent.Width - btnAction3.Width
            btnAction3.Caption = "Accept Order"

            Dim lTemp As Int32 = btnAction3.Left + btnAction3.Width - btnAction2.Left
            lTemp \= 2
            lTemp += btnAction2.Left
            lTemp -= (btnAction1.Width \ 2)
            btnAction1.Left = lTemp
			btnAction1.Visible = False
			lblOrderCap.Visible = True
			lblOrderCap.Left = btnAction2.Left + btnAction2.Width + 5
			lblOrderCap.Top = btnAction2.Top + 2


			FillListWithMarketList(lstContents, myShowState)

			If mlLastContentsListItemData > -1 Then
				For X As Int32 = 0 To lstContents.ListCount - 1
					If lstContents.ItemData(X) = mlLastContentsListItemData AndAlso lstContents.ItemData2(X) = mlLastContentsListItemData2 Then
						lstContents.ListIndex = X
						Exit For
					End If
				Next X
			End If
        End If

        If myShowState <> 4 Then
            lstCurrent.sHeaderRow = "Item Name            Price                Quantity"
        Else
            lstCurrent.sHeaderRow = "Escrow               Payment                Deadline"
            lblCurrentPrice.Caption = "Current Orders"
        End If

    End Sub

    Public Sub ListChangeEvent(ByVal bBoldCurrentItem As Boolean)
        mlPreviousContentsIndex = -1
        lstContents_ItemClick(lstContents.ListIndex)
        If lstContents.ItemBold(lstContents.ListIndex) <> bBoldCurrentItem Then lstContents.ItemBold(lstContents.ListIndex) = bBoldCurrentItem
	End Sub
	Public Sub SetItemBold(ByVal lItemID As Int32, ByVal iItemTypeID As Int16, ByVal bValue As Boolean)
		For X As Int32 = 0 To lstContents.ListCount - 1
			If lstContents.ItemData(X) = lItemID AndAlso lstContents.ItemData2(X) = iItemTypeID Then
				If lstContents.ItemBold(X) <> bValue Then lstContents.ItemBold(X) = bValue
				If bValue = True Then
					Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 0)
					If lstContents.ItemCustomColor(X) <> clrVal Then
						lstContents.ItemCustomColor(X) = clrVal
					End If
				Else
					If lstContents.ItemCustomColor(X) <> lstContents.ForeColor Then
						lstContents.ItemCustomColor(X) = lstContents.ForeColor
					End If
				End If
				Exit For
			End If
		Next X
	End Sub

    Private Sub lstContents_ItemClick(ByVal lIndex As Integer) Handles lstContents.ItemClick
        If lIndex <> -1 AndAlso lstContents.ItemData(lIndex) <> -1 Then

            If lIndex <> mlPreviousContentsIndex OrElse glCurrentCycle - mlLastGTCRefresh > 60 Then
                mlLastGTCRefresh = glCurrentCycle
                mlPreviousContentsIndex = lIndex

                mlListUB = -1
                ReDim moList(-1)
                ReDim myListUsed(-1)
				lstCurrent.Clear()

				mlLastContentsListItemData = lstContents.ItemData(lIndex)
				mlLastContentsListItemData2 = lstContents.ItemData2(lIndex)

                Dim yType As Byte
                If myShowState = 1 Then
                    yType = GetMarketList(lstContents.ItemData(lIndex), lstContents.ItemData2(lIndex))
                Else : yType = CByte(lstContents.ItemData(lIndex))
                End If
                If yType = 0 Then Return
                If myShowState <> 4 Then yType = CByte(yType Or MarketListType.eSellOrderShift)

                Dim yMsg(6) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eGetGTCList).CopyTo(yMsg, 0)
                yMsg(2) = yType
                System.BitConverter.GetBytes(mlTradepostID).CopyTo(yMsg, 3)
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                If myShowState = 1 Then
                    If fraDetail Is Nothing = False Then
                        If fraDetail.ControlName.ToUpper <> "FRACREATESELLORDER" Then
                            For X As Int32 = 0 To Me.ChildrenUB
                                If Me.moChildren(X).ControlName = fraDetail.ControlName Then
                                    Me.RemoveChild(X)
                                    Exit For
                                End If
                            Next
                            fraDetail = Nothing
                            fraDetail = New fraCreateSellOrder(MyBase.moUILib)
                            fraDetail.Top = btnAction1.Top + btnAction1.Height + 5
                            fraDetail.Left = lstCurrent.Left
                            Me.AddChild(CType(fraDetail, UIControl))
                        End If
                    Else
                        fraDetail = New fraCreateSellOrder(MyBase.moUILib)
                        fraDetail.Top = btnAction1.Top + btnAction1.Height + 5
                        fraDetail.Left = lstCurrent.Left
                        Me.AddChild(CType(fraDetail, UIControl))
                    End If

                    'Ok, find our TPC
                    Dim oTPC As TradePostContents = Nothing
                    For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
                        If TradePostContents.lTradePostContentsIdx(X) = mlTradepostID Then
                            oTPC = TradePostContents.oTradePostContents(X)
                            Exit For
                        End If
                    Next X
                    If oTPC Is Nothing = False Then
                        Dim blQty As Int64 = 0
                        Dim blCurrPrice As Int64 = 0
						Dim blCurrQty As Int64 = 0

                        oTPC.GetItemsDetails(lstContents.ItemData(lIndex), CShort(lstContents.ItemData2(lIndex)), blQty, blCurrPrice, blCurrQty, CShort(lstContents.ItemData3(lIndex)))
                        CType(fraDetail, fraCreateSellOrder).SetItemDetails(Me.mlTradepostID, lstContents.ItemData(lIndex), CShort(lstContents.ItemData2(lIndex)), blQty, oTPC.ySellSlotsUsed, blCurrPrice, blCurrQty, CShort(lstContents.ItemData3(lIndex)))
                    End If
                ElseIf myShowState = 0 Then
                    BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eSellOrderCategoryClick, mlLastContentsListItemData, mlLastContentsListItemData2)
                End If
            End If
        End If
    End Sub

    Private Function GetMarketList(ByVal lID As Int32, ByVal lTypeID As Int32) As Byte
        Select Case Math.Abs(lTypeID)
            Case ObjectType.eAgent
                Return MarketListType.eAgent
            Case ObjectType.eAlloyTech
                Return MarketListType.eAlloys
            Case ObjectType.eArmorTech
                If lTypeID < 0 Then Return MarketListType.eArmorComponent Else Return MarketListType.eArmorDesign
            Case ObjectType.eColonists
                Return MarketListType.eColonists
            Case ObjectType.eEngineTech
                If lTypeID < 0 Then Return MarketListType.eEngineComponent Else Return MarketListType.eEngineDesign
            Case ObjectType.eEnlisted
                Return MarketListType.eEnlisted
			Case ObjectType.eMineral, ObjectType.eMineralCache
				If Math.Abs(lTypeID) = ObjectType.eMineralCache Then
					'let's make sure this is not an alloy
					Dim oTPC As TradePostContents = Nothing
					For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
						If TradePostContents.lTradePostContentsIdx(X) = mlTradepostID Then
							oTPC = TradePostContents.oTradePostContents(X)
							Exit For
						End If
					Next X
					If oTPC Is Nothing = False Then
						Dim lTmpID As Int32 = oTPC.GetMineralIDFromCache(lID)
						If lTmpID > -1 Then
							Try
								For X As Int32 = 0 To goCurrentPlayer.mlTechUB
									If goCurrentPlayer.moTechs(X) Is Nothing = False Then
										If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eAlloyTech Then
											Dim oAlloy As AlloyTech = CType(goCurrentPlayer.moTechs(X), AlloyTech)
											If oAlloy Is Nothing = False Then
												If oAlloy.AlloyResultID = lTmpID Then
													Return MarketListType.eAlloys
												End If
											End If
										End If
									End If
								Next X
							Catch
							End Try
						End If
					End If
				End If
				Return MarketListType.eMinerals
            Case ObjectType.eOfficers
                Return MarketListType.eOfficers
            Case ObjectType.eRadarTech
                If lTypeID < 0 Then Return MarketListType.eRadarComponent Else Return MarketListType.eRadarDesign
            Case ObjectType.eShieldTech
                If lTypeID < 0 Then Return MarketListType.eShieldComponent Else Return MarketListType.eShieldDesign
            Case ObjectType.eSpecialTech
                Return MarketListType.eSpecialComponent
            Case ObjectType.eUnit
                Dim oTPC As TradePostContents = Nothing
                For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
                    If TradePostContents.lTradePostContentsIdx(X) = mlTradepostID Then
                        oTPC = TradePostContents.oTradePostContents(X)
                        Exit For
                    End If
                Next X
                If oTPC Is Nothing = False Then
					Dim yChassis As Byte = oTPC.GetUnitItemChassisType(lID, lTypeID)
					If (yChassis And ChassisType.eSpaceBased) <> 0 Then Return MarketListType.eSpaceBasedUnit
					If (yChassis And ChassisType.eAtmospheric) <> 0 Then Return MarketListType.eAerialUnit
                    If (yChassis And ChassisType.eGroundBased) <> 0 Then Return MarketListType.eLandBasedUnit
                    If (yChassis And ChassisType.eNavalBased) <> 0 Then Return MarketListType.eNavalUnit
				End If
				Return MarketListType.eAerialUnit
			Case ObjectType.eWeaponTech
				For X As Int32 = 0 To goCurrentPlayer.mlTechUB
					If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjectID = lID AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = lTypeID Then
						Select Case CType(goCurrentPlayer.moTechs(X), WeaponTech).WeaponClassTypeID
							Case WeaponClassType.eBomb
                                Return MarketListType.eBombComponent
							Case WeaponClassType.eEnergyBeam
								Return MarketListType.eBeamSolidComponent
							Case WeaponClassType.eEnergyPulse
								Return MarketListType.eBeamPulseComponent
							Case WeaponClassType.eMine
                                Return MarketListType.eMineWeaponComponent
							Case WeaponClassType.eMissile
								Return MarketListType.eMissileComponent
							Case WeaponClassType.eProjectile
								Return MarketListType.eProjectileComponent
						End Select
					End If
				Next X
				For X As Int32 = 0 To glPlayerTechKnowledgeUB
					If glPlayerTechKnowledgeIdx(X) <> -1 Then
						If goPlayerTechKnowledge(X).oTech.ObjectID = lID AndAlso goPlayerTechKnowledge(X).oTech.ObjTypeID = lTypeID Then
							Select Case CType(goPlayerTechKnowledge(X).oTech, WeaponTech).WeaponClassTypeID
								Case WeaponClassType.eBomb
                                    Return MarketListType.eBombComponent
								Case WeaponClassType.eEnergyBeam
									Return MarketListType.eBeamSolidComponent
								Case WeaponClassType.eEnergyPulse
									Return MarketListType.eBeamPulseComponent
								Case WeaponClassType.eMine
                                    Return MarketListType.eMineWeaponComponent
								Case WeaponClassType.eMissile
									Return MarketListType.eMissileComponent
								Case WeaponClassType.eProjectile
									Return MarketListType.eProjectileComponent
							End Select
						End If
					End If
				Next X
				Return MarketListType.eProjectileComponent

			Case ObjectType.ePlayerIntel 'TODO: for now these are hard-coded
				Return MarketListType.ePlayerStats
			Case ObjectType.ePlayerItemIntel
				Return MarketListType.eColonyLocation
			Case ObjectType.ePlayerTechKnowledge
				Return MarketListType.ePlayerComponentDesigns
		End Select
        Return 0
    End Function

    Private Sub RefreshCurrentList()
        lstCurrent.Clear()
        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        Try
            For X As Int32 = 0 To mlListUB
                Dim lIdx As Int32 = -1
                If myListUsed(X) <> 0 Then
                    For Y As Int32 = 0 To lSortedUB
                        'Ok, determine what we are sorting by
                        Select Case mySortBy
                            Case CurrentPriceSort.eItemNameAscending
                                If CType(moList(lSorted(Y)), SellOrder).sItemName > CType(moList(X), SellOrder).sItemName Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Case CurrentPriceSort.eItemNameDescending
                                If CType(moList(lSorted(Y)), SellOrder).sItemName < CType(moList(X), SellOrder).sItemName Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Case CurrentPriceSort.ePriceDescending
                                If moList(lSorted(Y)).blPrice < moList(X).blPrice Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Case CurrentPriceSort.eQuantityAscending
                                If moList(lSorted(Y)).blQty > moList(X).blQty Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Case CurrentPriceSort.eQuantityDescending
                                If moList(lSorted(Y)).blQty < moList(X).blQty Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Case CurrentPriceSort.eDeadlineAscending
                                If CType(moList(lSorted(Y)), BuyOrder).lDeadline > CType(moList(X), BuyOrder).lDeadline Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Case CurrentPriceSort.eDeadlineDescending
                                If CType(moList(lSorted(Y)), BuyOrder).lDeadline < CType(moList(X), BuyOrder).lDeadline Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Case CurrentPriceSort.eEscrowAscending
                                If CType(moList(lSorted(Y)), BuyOrder).blEscrow > CType(moList(X), BuyOrder).blEscrow Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Case CurrentPriceSort.eEscrowDescending
                                If CType(moList(lSorted(Y)), BuyOrder).blEscrow < CType(moList(X), BuyOrder).blEscrow Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Case Else               ' CurrentPriceSort.ePriceAscending
                                If moList(lSorted(Y)).blPrice > moList(X).blPrice Then
                                    lIdx = Y
                                    Exit For
                                End If
                        End Select
                    Next Y
                    lSortedUB += 1
                    ReDim Preserve lSorted(lSortedUB)
                    If lIdx = -1 Then
                        lSorted(lSortedUB) = X
                    Else
                        For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                            lSorted(Y) = lSorted(Y - 1)
                        Next Y
                        lSorted(lIdx) = X
                    End If
                End If
            Next X

            'Now, go through and add them to our list
            For X As Int32 = 0 To lSortedUB
                If NewTutorialManager.TutorialOn = False OrElse (NewTutorialManager.TutorialOn = True AndAlso moList(lSorted(X)).yRouteType < 2) Then
                    If chkThisEnvironmentOnly.Value = False OrElse (chkThisEnvironmentOnly.Value = True AndAlso moList(lSorted(X)).yRouteType < 2) Then
                        lstCurrent.AddItem(moList(lSorted(X)).GetCurrentListText)
                        lstCurrent.ItemData(lstCurrent.NewIndex) = lSorted(X)

                        If myShowState = 4 Then
                            If CType(moList(lSorted(X)), BuyOrder).lAcceptedByID = glPlayerID Then
                                lstCurrent.ItemBold(lstCurrent.NewIndex) = True
                            End If
                        End If
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub

    Public Sub HandleGTCList(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lTradePost As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yType As Byte = yData(lPos) : lPos += 1

        If lstContents.ListIndex = -1 Then Return

        If yType = 255 Then
            'Ok, end of the list... refresh it now
            RefreshCurrentList()
        Else
            Dim ySelType As Byte = 0
            If myShowState <> 1 Then ySelType = CByte(lstContents.ItemData(lstContents.ListIndex)) Else ySelType = GetMarketList(lstContents.ItemData(lstContents.ListIndex), lstContents.ItemData2(lstContents.ListIndex))
            If myShowState <> 4 Then ySelType = CByte(ySelType Or MarketListType.eSellOrderShift)
            If lTradePost <> mlTradepostID OrElse ySelType <> yType Then Return

            Dim oNew As TradeOrder
            If (yType And MarketListType.eSellOrderShift) <> 0 Then
                'Sell Order
                oNew = New SellOrder()
            Else
                'Buy Order
				oNew = New BuyOrder()
            End If
            oNew.SetFromMsg(yData, lPos)

            'Now place it
            Dim lIdx As Int32 = -1
            Dim lFirstIdx As Int32 = -1
            For X As Int32 = 0 To mlListUB
                If myListUsed(X) = 0 AndAlso lFirstIdx <> -1 Then
                    lFirstIdx = X
                ElseIf myListUsed(X) <> 0 Then
                    With moList(X)
                        If .lTradePostID = oNew.lTradePostID AndAlso .lObjectID = oNew.lObjectID AndAlso .iObjTypeID = oNew.iObjTypeID Then
                            lIdx = X
                            Exit For
                        End If
                    End With
                End If
            Next X
            If lIdx = -1 Then
                If lFirstIdx = -1 Then
                    mlListUB += 1
                    ReDim Preserve myListUsed(mlListUB)
                    ReDim Preserve moList(mlListUB)
                    lIdx = mlListUB
                Else : lIdx = lFirstIdx
                End If
            End If
            moList(lIdx) = oNew
            myListUsed(lIdx) = 255

        End If

    End Sub

	Private Sub btnAction1_Click(ByVal sName As String) Handles btnAction1.Click
		Select Case myShowState
			Case 0, 1			'BUY and SELL Screen: Refresh GTC 
				lstContents_ItemClick(lstContents.ListIndex)
			Case 4				'BUY ORDER Screen: Deliver
				If HasAliasedRights(AliasingRights.eAlterTrades) = False Then
					MyBase.moUILib.AddNotification("You lack rights to alter/place trades.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
					Return
				End If

				If lstCurrent.ListIndex = -1 Then Return
				Dim lIdx As Int32 = lstCurrent.ItemData(lstCurrent.ListIndex)
				If lIdx > -1 AndAlso lIdx <= mlListUB Then
					If mfraDeliver Is Nothing Then mfraDeliver = New fraDeliverBuyOrder(MyBase.moUILib)
					With mfraDeliver
						.ControlName = "fraDeliverBuyOrder"
						.Left = (btnAction1.Left + (btnAction1.Width \ 2)) - (mfraDeliver.Width \ 2)
						.Top = btnAction1.Top - .Height
						.Enabled = True
						.Visible = True
						.mbAcceptReprocessEvents = True
					End With
					Me.AddChild(CType(mfraDeliver, UIControl))

					fraDetail.Enabled = False
					AddHandler mfraDeliver.CancelClicked, AddressOf mfraDeliver_CancelClick
					AddHandler mfraDeliver.DeliverClicked, AddressOf mfraDeliver_DeliverClick

					mfraDeliver.SetDetails(mlTradepostID, CType(moList(lIdx), BuyOrder))
				End If
		End Select
	End Sub

	Private Sub mfraDeliver_CancelClick()
		If mfraDeliver Is Nothing = False Then
			For X As Int32 = 0 To Me.ChildrenUB
				If mfraDeliver.ControlName = Me.moChildren(X).ControlName Then
					Me.RemoveChild(X)
					Exit For
				End If
			Next X
			RemoveHandler mfraDeliver.CancelClicked, AddressOf mfraDeliver_CancelClick
			RemoveHandler mfraDeliver.DeliverClicked, AddressOf mfraDeliver_DeliverClick
		End If
		fraDetail.Enabled = True
	End Sub
	Private Sub mfraDeliver_DeliverClick(ByVal lCacheID As Int32, ByVal iCacheTypeID As Int16)
		If mfraDeliver Is Nothing = False Then
			For X As Int32 = 0 To Me.ChildrenUB
				If mfraDeliver.ControlName = Me.moChildren(X).ControlName Then
					Me.RemoveChild(X)
					Exit For
				End If
			Next X
			RemoveHandler mfraDeliver.CancelClicked, AddressOf mfraDeliver_CancelClick
			RemoveHandler mfraDeliver.DeliverClicked, AddressOf mfraDeliver_DeliverClick
		End If

		If lstCurrent.ListIndex = -1 Then Return
		Dim lIdx As Int32 = lstCurrent.ItemData(lstCurrent.ListIndex)
		If lIdx > -1 AndAlso lIdx <= mlListUB Then
			Dim yMsg(19) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eDeliverBuyOrder).CopyTo(yMsg, lPos) : lPos += 2
			With moList(lIdx)
				System.BitConverter.GetBytes(.lTradePostID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.lObjectID).CopyTo(yMsg, lPos) : lPos += 4
			End With
			System.BitConverter.GetBytes(mlTradepostID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(lCacheID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(iCacheTypeID).CopyTo(yMsg, lPos) : lPos += 2

			MyBase.moUILib.SendMsgToPrimary(yMsg)
			MyBase.moUILib.AddNotification("Delivery Notification posted to Buy Order...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End If
		fraDetail.Enabled = True
	End Sub

	Private Sub btnAction2_Click(ByVal sName As String) Handles btnAction2.Click
		Select Case myShowState
			Case 0, 1			'BUY and SELL Screen: not used
			Case 4				'BUY ORDER screen: Create Order
				If HasAliasedRights(AliasingRights.eAlterTrades) = False Then
					MyBase.moUILib.AddNotification("You lack rights to alter/place trades.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
					Return
				End If
				CType(Me.ParentControl, frmTradeMain).SwitchToCreateBuyOrder()
		End Select
	End Sub

	Private Sub btnAction3_Click(ByVal sName As String) Handles btnAction3.Click
		Select Case myShowState
			Case 0, 1			'BUY and SELL Screen: not used
			Case 4				'BUY ORDER Screen: Accept Buy Order
				If HasAliasedRights(AliasingRights.eAlterTrades) = False Then
					MyBase.moUILib.AddNotification("You lack rights to alter/place trades.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
					Return
				End If
				If lstCurrent.ListIndex <> -1 Then
					Dim lIdx As Int32 = lstCurrent.ItemData(lstCurrent.ListIndex)
					If lIdx > -1 AndAlso lIdx <= mlListUB AndAlso myListUsed(lIdx) <> 0 Then
						btnAction3.Enabled = False
						Dim yMsg() As Byte = CType(moList(lIdx), BuyOrder).GetAcceptRequest(mlTradepostID)
						If yMsg Is Nothing = False Then
							MyBase.moUILib.SendMsgToPrimary(yMsg)
							'TODO: What do we do here?
						End If
						btnAction3.Enabled = True
					End If
				End If
		End Select
	End Sub

	Private Sub lstCurrent_HeaderRowClick(ByVal lX As Integer) Handles lstCurrent.HeaderRowClick
		'Ok, determine what column was clicked on...

		If lX < 100 Then
			If myShowState = 4 Then
				If mySortBy = CurrentPriceSort.eEscrowAscending Then
					mySortBy = CurrentPriceSort.eEscrowDescending
				Else : mySortBy = CurrentPriceSort.eEscrowAscending
				End If
			Else
				'Item Name
				If mySortBy = CurrentPriceSort.eItemNameAscending Then
					mySortBy = CurrentPriceSort.eItemNameDescending
				Else : mySortBy = CurrentPriceSort.eItemNameAscending
				End If
			End If
		ElseIf lX < 250 Then
			'Price
			If mySortBy = CurrentPriceSort.ePriceAscending Then
				mySortBy = CurrentPriceSort.ePriceDescending
			Else : mySortBy = CurrentPriceSort.ePriceAscending
			End If
		Else
			If myShowState = 4 Then
				'Deadline
				If mySortBy = CurrentPriceSort.eDeadlineAscending Then
					mySortBy = CurrentPriceSort.eDeadlineDescending
				Else : mySortBy = CurrentPriceSort.eDeadlineAscending
				End If
			Else
				'Quantity
				If mySortBy = CurrentPriceSort.eQuantityAscending Then
					mySortBy = CurrentPriceSort.eQuantityDescending
				Else : mySortBy = CurrentPriceSort.eQuantityAscending
				End If
			End If

		End If

		RefreshCurrentList()
	End Sub

    Private Sub lstCurrent_ItemClick(ByVal lIndex As Integer) Handles lstCurrent.ItemClick
        If lIndex <> -1 Then
            If myShowState = 0 Then
                Dim lIdx As Int32 = lstCurrent.ItemData(lIndex)
                If lIdx <> -1 AndAlso lIdx <= mlListUB AndAlso myListUsed(lIdx) <> 0 Then
                    If fraDetail Is Nothing Then
                        fraDetail = New fraSellOrderDetail(MyBase.moUILib)
                        fraDetail.Top = btnAction1.Top + btnAction1.Height + 5
                        fraDetail.Left = lstCurrent.Left
                        Me.AddChild(CType(fraDetail, UIControl))
                    ElseIf fraDetail.ControlName.ToLower <> "frasellorderdetail" Then
                        For X As Int32 = 0 To Me.ChildrenUB
                            If Me.moChildren(X).ControlName = fraDetail.ControlName Then
                                Me.RemoveChild(X)
                                Exit For
                            End If
                        Next X
                        fraDetail = Nothing
                        fraDetail = New fraSellOrderDetail(MyBase.moUILib)
                        fraDetail.Top = btnAction1.Top + btnAction1.Height + 5
                        fraDetail.Left = lstCurrent.Left
                        Me.AddChild(CType(fraDetail, UIControl))
                    Else
                        fraDetail.Visible = True
                    End If
                    CType(fraDetail, fraSellOrderDetail).SetDetails(CType(moList(lIdx), SellOrder), mlTradepostID)
                End If
            ElseIf myShowState = 4 Then
                'Buy Order details
                Dim lIdx As Int32 = lstCurrent.ItemData(lIndex)
                If lIdx <> -1 AndAlso lIdx <= mlListUB AndAlso myListUsed(lIdx) <> 0 Then
                    If fraDetail Is Nothing Then
                        fraDetail = New fraBuyOrderDetails(MyBase.moUILib)
                        fraDetail.Top = lstCurrent.Top + lstCurrent.Height + 5
                        fraDetail.Left = lstCurrent.Left
                        Me.AddChild(CType(fraDetail, UIControl))
                    ElseIf fraDetail.ControlName.ToLower <> "frabuyorderdetails" Then
                        For X As Int32 = 0 To Me.ChildrenUB
                            If Me.moChildren(X).ControlName = fraDetail.ControlName Then
                                Me.RemoveChild(X)
                                Exit For
                            End If
                        Next X
                        fraDetail = Nothing
                        fraDetail = New fraBuyOrderDetails(MyBase.moUILib)
                        fraDetail.Top = lstCurrent.Top + lstCurrent.Height + 5
                        fraDetail.Left = lstCurrent.Left
                        Me.AddChild(CType(fraDetail, UIControl))
                    End If
                    CType(fraDetail, fraBuyOrderDetails).SetDetails(CType(moList(lIdx), BuyOrder), mlTradepostID)

                    If CType(moList(lIdx), BuyOrder).lAcceptedByID = glPlayerID Then
                        btnAction1.Visible = True
                    Else : btnAction1.Visible = False
                    End If
                End If
            End If
        End If
    End Sub

	Public Sub fraTradeContents_OnNewFrame() Handles Me.OnNewFrame
		If fraDetail Is Nothing = False Then
			fraDetail.SubFrame_OnNewFrame()
		End If

		'Ok, are we showing the sell screen?
		If myShowState = 1 Then
			'Yes, ok, find our TradePostContents
			Dim oTPC As TradePostContents = Nothing
			For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
				If TradePostContents.lTradePostContentsIdx(X) = mlTradepostID Then
					oTPC = TradePostContents.oTradePostContents(X)
					Exit For
				End If
			Next X

			If oTPC Is Nothing = False Then
				'Ok... ensure the list has everything
				If mlTPCLast <> oTPC.lLastUpdate Then
					mlTPCLast = oTPC.lLastUpdate
					oTPC.PopulateList(lstContents, 0)
				Else : oTPC.SmartPopulateList(lstContents, 0)
				End If
			Else
				If lstContents.ListCount > 0 Then lstContents.Clear()
			End If
		ElseIf myShowState = 4 Then

			If lblOrderCap.Visible = True Then
				Dim oTPC As TradePostContents = Nothing
				For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
					If TradePostContents.lTradePostContentsIdx(X) = mlTradepostID Then
						oTPC = TradePostContents.oTradePostContents(X)
						Exit For
					End If
				Next X
				If oTPC Is Nothing = False Then
					Dim sText As String = oTPC.yBuySlotsUsed & " used / " & goCurrentPlayer.GetPlayerBuySlots() & " max"
					If lblOrderCap.Caption <> sText Then lblOrderCap.Caption = sText
				End If
			End If

			For X As Int32 = 0 To mlListUB
				Dim bFound As Boolean = False
				For Y As Int32 = 0 To lstCurrent.ListCount - 1
					If lstCurrent.ItemData(Y) = X Then
						Dim sTemp As String = moList(X).GetCurrentListText()
						If lstCurrent.List(Y) <> sTemp Then lstCurrent.List(Y) = sTemp
						bFound = True
						Exit For
					End If
				Next Y
				If bFound = False Then
					RefreshCurrentList()
					Exit For
				End If
			Next X
		End If
	End Sub

	Public Sub HandleGetOrderSpecifics(ByRef yData() As Byte)
		Dim lPos As Int32 = 2
		Dim lTradePostID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim yOrderType As Byte = yData(lPos) : lPos += 1
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		If yOrderType = 0 Then
			'SELL ORDER
			For X As Int32 = 0 To mlListUB
				If myListUsed(X) <> 0 Then
					With moList(X)
						If .lObjectID = lObjID AndAlso .iObjTypeID = iTypeID AndAlso .lTradePostID = lTradePostID Then
							.SetFromDetailMsg(yData, lPos)
							Exit For
						End If
					End With
				End If
			Next X
		Else
			'BUY ORDER
			For X As Int32 = 0 To mlListUB
				If myListUsed(X) <> 0 Then
					With moList(X)
						If .lObjectID = lObjID AndAlso .lTradePostID = lTradePostID Then
							.SetFromDetailMsg(yData, lPos)
							Exit For
						End If
					End With
				End If
			Next X
		End If
	End Sub

#Region "  Detail Panel Base Class  "
	Private MustInherit Class fraDetailPane
		Inherits UIWindow

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)
		End Sub

		Public MustOverride Sub SubFrame_OnNewFrame()
	End Class
#End Region

#Region "  Sell Order Detail Panel (the BUY screen)  "
	'Interface created from Interface Builder
	Private Class fraSellOrderDetail
		Inherits fraDetailPane

		Private txtItemDetails As UITextBox
		Private lblRouteType As UILabel
		Private lblDeliverTime As UILabel
		Private lblSeller As UILabel
		Private lblItemScore As UILabel
		Private lnDiv1 As UILine
		Private lblQuantity As UILabel
		Private WithEvents txtPurchaseQty As UITextBox
		Private lblTotalPrice As UILabel
		Private WithEvents btnPurchase As UIButton

		Private moItem As SellOrder = Nothing
		Private mlTradepostID As Int32 = -1

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

			'fraSellOrderDetail initial props
			With Me
				.ControlName = "fraSellOrderDetail"
				.Left = 71
				.Top = 55
				.Width = 520
				.Height = 180
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
				.BorderLineWidth = 2
				.mbAcceptReprocessEvents = True
			End With

			'txtItemDetails initial props
			txtItemDetails = New UITextBox(oUILib)
			With txtItemDetails
				.ControlName = "txtItemDetails"
				.Left = 5
				.Top = 5
				.Width = 250
				.Height = 170
				.Enabled = True
				.Visible = True
				.Caption = ""
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(0, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
				.MaxLength = 0
				.BorderColor = muSettings.InterfaceBorderColor
				.MultiLine = True
				.Locked = True
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(txtItemDetails, UIControl))

			'lblRouteType initial props
			lblRouteType = New UILabel(oUILib)
			With lblRouteType
				.ControlName = "lblRouteType"
				.Left = 265
				.Top = 5
				.Width = 250
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Route Type: Acquiring..."
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblRouteType, UIControl))

			'lblDeliverTime initial props
			lblDeliverTime = New UILabel(oUILib)
			With lblDeliverTime
				.ControlName = "lblDeliverTime"
				.Left = 265
				.Top = 25
				.Width = 250
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Est. Delivery Time: Calculating..."
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblDeliverTime, UIControl))

			'lblSeller initial props
			lblSeller = New UILabel(oUILib)
			With lblSeller
				.ControlName = "lblSeller"
				.Left = 265
				.Top = 45
				.Width = 250
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Seller: Unknown"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblSeller, UIControl))

			'lblItemScore initial props
			lblItemScore = New UILabel(oUILib)
			With lblItemScore
				.ControlName = "lblItemScore"
				.Left = 265
				.Top = 65
				.Width = 250
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Score: 0 (0%)"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblItemScore, UIControl))

			'lnDiv1 initial props
			lnDiv1 = New UILine(oUILib)
			With lnDiv1
				.ControlName = "lnDiv1"
				.Left = 265
				.Top = 90
				.Width = 245
				.Height = 0
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
			End With
			Me.AddChild(CType(lnDiv1, UIControl))

			'lblQuantity initial props
			lblQuantity = New UILabel(oUILib)
			With lblQuantity
				.ControlName = "lblQuantity"
				.Left = 265
				.Top = 100
				.Width = 114
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Purchase Quantity:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblQuantity, UIControl))

			'txtPurchaseQty initial props
			txtPurchaseQty = New UITextBox(oUILib)
			With txtPurchaseQty
				.ControlName = "txtPurchaseQty"
				.Left = 380
				.Top = 100
				.Width = 130
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "0"
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
				.MaxLength = 0
				.BorderColor = muSettings.InterfaceBorderColor
				.ToolTipText = "Enter the quantity to purchase."
			End With
			Me.AddChild(CType(txtPurchaseQty, UIControl))

			'lblTotalPrice initial props
			lblTotalPrice = New UILabel(oUILib)
			With lblTotalPrice
				.ControlName = "lblTotalPrice"
				.Left = 265
				.Top = 125
				.Width = 250
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Purchase Price: "
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblTotalPrice, UIControl))

			'btnPurchase initial props
			btnPurchase = New UIButton(oUILib)
			With btnPurchase
				.ControlName = "btnPurchase"
				.Left = 335
				.Top = 150
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Purchase"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
				.ToolTipText = "Click to purchase the entered quantity."
			End With
			Me.AddChild(CType(btnPurchase, UIControl))
		End Sub

		Public Sub SetDetails(ByRef oItem As SellOrder, ByVal lTP_ID As Int32)
			moItem = oItem
			mlTradepostID = lTP_ID

            moItem.RequestDetails(lTP_ID)

			If moItem.iObjTypeID = ObjectType.ePlayerIntel OrElse moItem.iObjTypeID = ObjectType.ePlayerItemIntel OrElse moItem.iObjTypeID = ObjectType.ePlayerTechKnowledge Then
                txtItemDetails.Caption = GetNonOwnerIntelItemData(moItem.lTradePostID, moItem.lObjectID, moItem.iObjTypeID, moItem.iExtTypeID)
			Else
				txtItemDetails.Caption = GetNonOwnerItemData(moItem.lObjectID, Math.Abs(moItem.iObjTypeID))
            End If

			''0 = planetary, 1 = intrasystem, 2 = system to system
			Select Case oItem.yRouteType
				Case 0
					lblRouteType.Caption = "Route Type: Planetary"
				Case 1
					lblRouteType.Caption = "Route Type: Intra-System"
				Case Else
					lblRouteType.Caption = "Route Type: Out of System"
			End Select

			lblDeliverTime.Caption = "Est. Delivery Time: Calculating..."
			lblSeller.Caption = "Seller: Unknown"
			lblItemScore.Caption = "Score: 0"
			txtPurchaseQty.Caption = "0"

			btnPurchase.Caption = "Purchase"
		End Sub

        Public Overrides Sub SubFrame_OnNewFrame() Handles Me.OnNewFrame
            Debug.Print("fraTradeContents->SubFrame_OnNewFrame")
            If moItem Is Nothing Then Return
            moItem.RequestDetails(mlTradepostID)
            'Dim iTempTypeID As Int16 = moItem.iObjTypeID
            'If iTempTypeID < 0 Then iTempTypeID = ObjectType.eComponentCache
            Dim sTemp As String = ""
            If moItem.iObjTypeID = ObjectType.ePlayerIntel OrElse moItem.iObjTypeID = ObjectType.ePlayerItemIntel OrElse moItem.iObjTypeID = ObjectType.ePlayerTechKnowledge Then
                sTemp = "Seller: Anonymous"
                If lblSeller.Caption <> sTemp Then lblSeller.Caption = sTemp

                sTemp = GetNonOwnerIntelItemData(mlTradepostID, moItem.lObjectID, moItem.iObjTypeID, moItem.iExtTypeID)
            Else
                If moItem.lSellerID = -1 Then
                    sTemp = "Seller: ...."
                Else : sTemp = "Seller: " & GetCacheObjectValue(moItem.lSellerID, ObjectType.ePlayer)
                End If
                If lblSeller.Caption <> sTemp Then lblSeller.Caption = sTemp

                sTemp = GetNonOwnerItemData(moItem.lObjectID, moItem.iObjTypeID)
            End If
            If txtItemDetails.Caption <> sTemp Then txtItemDetails.Caption = sTemp

            If moItem.lDeliveryTime = -1 Then
                sTemp = "Est. Delivery Time: Calculating..."
            Else : sTemp = "Est. Delivery Time: " & GetTimeStringFromSeconds(moItem.lDeliveryTime)
            End If
            If lblDeliverTime.Caption <> sTemp Then lblDeliverTime.Caption = sTemp



            sTemp = "Score: " & moItem.ItemScore
            If lblItemScore.Caption <> sTemp Then lblItemScore.Caption = sTemp
        End Sub

		Private Function GetTimeStringFromSeconds(ByVal lSeconds As Int32) As String
			Dim lMinutes As Int32 = lSeconds \ 60
			lSeconds -= (lMinutes * 60)
			Dim lHours As Int32 = lMinutes \ 60
			lMinutes -= (lHours * 60)
			Dim lDays As Int32 = lHours \ 24
			lHours -= (lDays * 24)

			Dim sResult As String = ""
			If lDays <> 0 Then
				sResult = lDays.ToString & "d "
			End If
			If lHours <> 0 Then sResult &= lHours.ToString & "h "
			If lMinutes <> 0 Then sResult &= lMinutes.ToString & "m "
			If lSeconds <> 0 OrElse sResult = "" Then sResult &= lSeconds.ToString & "s "
			Return sResult.Trim()
		End Function

		Private Sub txtPurchaseQty_TextChanged() Handles txtPurchaseQty.TextChanged
			If moItem Is Nothing = False Then
				If txtPurchaseQty.Caption <> "" Then
					Dim blTemp As Int64 = CLng(Val(txtPurchaseQty.Caption))

					If blTemp < 0 Then blTemp = 0
					If blTemp > moItem.blQty Then blTemp = moItem.blQty

					If blTemp.ToString <> txtPurchaseQty.Caption Then
						txtPurchaseQty.Caption = blTemp.ToString
					End If

					lblTotalPrice.Caption = "Purchase Price: " & CLng(blTemp * moItem.blPrice * goCurrentPlayer.GetPlayerTradeCost()).ToString("#,##0")
					btnPurchase.Caption = "Purchase"
				End If
			End If
		End Sub

		Private Sub btnPurchase_Click(ByVal sName As String) Handles btnPurchase.Click
			If moItem Is Nothing = False Then
				If btnPurchase.Caption = "Confirm" Then
					'ok, purchase the item
					Dim blQty As Int64 = CLng(Val(txtPurchaseQty.Caption))
					If blQty < 1 Then
						MyBase.moUILib.AddNotification("Please enter a valid quantity to purchase.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Return
					End If

					btnPurchase.Enabled = False
					Dim yMsg() As Byte = moItem.GetPurchaseMessage(blQty, mlTradepostID)
					If yMsg Is Nothing = False Then
						MyBase.moUILib.SendMsgToPrimary(yMsg)
						CType(Me.ParentControl, fraTradeContents).ListChangeEvent(False)
						Me.Visible = False
					Else : btnPurchase.Caption = "Purchase"
					End If
					btnPurchase.Enabled = True
				Else : btnPurchase.Caption = "Confirm"
				End If
			Else : btnPurchase.Caption = "Purchase"
			End If
		End Sub
	End Class
#End Region

#Region "  Create Sell Order  "
	'Interface created from Interface Builder
	Private Class fraCreateSellOrder
		Inherits fraDetailPane

		Private txtItemDetails As UITextBox
		Private lblSellSlots As UILabel
		Private lblQuantity As UILabel
		Private lblPrice As UILabel
		Private txtPrice As UITextBox
		Private WithEvents txtSellQty As UITextBox
		Private WithEvents btnPlace As UIButton
		Private WithEvents btnCancelOrder As UIButton

		Private mlTradePostID As Int32 = -1
		Private mlID As Int32 = -1
		Private miTypeID As Int16 = -1
		Private mblQty As Int64 = 0
		Private mySlotsUsed As Byte = 0
		Private miExtTypeID As Int16 = -1

		Private mblPrice As Int64 = 0
		Private mblForSaleQty As Int64 = 0

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

			'fraCreateSellOrder initial props
			With Me
				.ControlName = "fraCreateSellOrder"
				.Left = 108
				.Top = 283
				.Width = 520
				.Height = 180
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
				.mbAcceptReprocessEvents = True
				.BorderLineWidth = 2
			End With

			'txtItemDetails initial props
			txtItemDetails = New UITextBox(oUILib)
			With txtItemDetails
				.ControlName = "txtItemDetails"
				.Left = 5
				.Top = 5
				.Width = 250
				.Height = 170
				.Enabled = True
				.Visible = True
				.Caption = ""
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(0, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
				.MaxLength = 0
				.BorderColor = muSettings.InterfaceBorderColor
				.MultiLine = True
				.Locked = True
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(txtItemDetails, UIControl))

			'lblSellSlots initial props
			lblSellSlots = New UILabel(oUILib)
			With lblSellSlots
				.ControlName = "lblSellSlots"
				.Left = 265
				.Top = 5
				.Width = 250
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Sell Slots: "
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblSellSlots, UIControl))

			'lblQuantity initial props
			lblQuantity = New UILabel(oUILib)
			With lblQuantity
				.ControlName = "lblQuantity"
				.Left = 265
				.Top = 30
				.Width = 114
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Sell Order Quantity:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblQuantity, UIControl))

			'txtPurchaseQty initial props
			txtSellQty = New UITextBox(oUILib)
			With txtSellQty
				.ControlName = "txtSellQty"
				.Left = 385
				.Top = 30
				.Width = 125
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "0"
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 9
                .bNumericOnly = True
				.BorderColor = muSettings.InterfaceBorderColor
			End With
			Me.AddChild(CType(txtSellQty, UIControl))

			'lblPrice initial props
			lblPrice = New UILabel(oUILib)
			With lblPrice
				.ControlName = "lblPrice"
				.Left = 265
				.Top = 55
				.Width = 100
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Price Per Item:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblPrice, UIControl))

			'btnPlace initial props
			btnPlace = New UIButton(oUILib)
			With btnPlace
				.ControlName = "btnPlace"
				.Left = 335
				.Top = 100
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Place Order"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			Me.AddChild(CType(btnPlace, UIControl))

			'txtPrice initial props
			txtPrice = New UITextBox(oUILib)
			With txtPrice
				.ControlName = "txtPrice"
				.Left = 385
				.Top = 55
				.Width = 125
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "0"
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 13
                .bNumericOnly = True
				.BorderColor = muSettings.InterfaceBorderColor
			End With
			Me.AddChild(CType(txtPrice, UIControl))

			'btnCancelOrder initial props
			btnCancelOrder = New UIButton(oUILib)
			With btnCancelOrder
				.ControlName = "btnCancelOrder"
				.Left = 335
				.Top = 135
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = False
				.Caption = "Cancel Order"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			Me.AddChild(CType(btnCancelOrder, UIControl))
		End Sub

		Public Sub SetItemDetails(ByVal lTP_ID As Int32, ByVal lID As Int32, ByVal iTypeID As Int16, ByVal blQty As Int64, ByVal ySlotsUsed As Byte, ByVal blCurrPrice As Int64, ByVal blCurrQty As Int64, ByVal iExtTypeID As Int16)
			mlTradePostID = lTP_ID
			mlID = lID
			miTypeID = iTypeID
			mblQty = blQty
			mySlotsUsed = ySlotsUsed
			mblPrice = blCurrPrice
			mblForSaleQty = blCurrQty
			miExtTypeID = iExtTypeID

			txtItemDetails.Caption = DetermineDetailsText()

			If mblPrice <> 0 AndAlso mblForSaleQty <> 0 Then
				btnCancelOrder.Visible = True
				txtSellQty.Caption = mblForSaleQty.ToString("###0")
				txtPrice.Caption = mblPrice.ToString("###0")
			Else
				btnCancelOrder.Visible = False
				txtSellQty.Caption = mblQty.ToString("###0")
				txtPrice.Caption = "0"
			End If

			If miTypeID = ObjectType.ePlayerIntel OrElse miTypeID = ObjectType.ePlayerItemIntel OrElse miTypeID = ObjectType.ePlayerTechKnowledge Then
				txtSellQty.Caption = "1"
				mblQty = 1
			End If

			btnPlace.Caption = "Place Order"
			btnCancelOrder.Caption = "Cancel Order"
		End Sub

		Public Overrides Sub SubFrame_OnNewFrame()
			If mlID = -1 OrElse miTypeID = -1 Then Return

			Dim lTempID As Int32 = mlID
			If miTypeID < 0 Then
				Dim oTPC As TradePostContents = Nothing
				For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
					If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
						oTPC = TradePostContents.oTradePostContents(X)
						Exit For
					End If
				Next X
				If oTPC Is Nothing = False Then
					Dim iTypeID As Int16 = -1
					Dim lOwnerID As Int32 = -1
					oTPC.PopulateComponentCacheProperties(mlID, lTempID, iTypeID, lOwnerID)
				End If
			End If
			Dim sTemp As String = DetermineDetailsText()
			If txtItemDetails.Caption <> sTemp Then txtItemDetails.Caption = sTemp

			sTemp = "Sell Slots: " & mySlotsUsed & " Used, " & goCurrentPlayer.GetPlayerSellSlots() & " Available"
			If lblSellSlots.Caption <> sTemp Then lblSellSlots.Caption = sTemp
		End Sub

		Private Function DetermineDetailsText() As String
			Dim sTemp As String = "Requesting Details..."
			Select Case miTypeID
				Case ObjectType.eColonists, ObjectType.eEnlisted, ObjectType.eOfficers, ObjectType.eCredits, ObjectType.eFood, ObjectType.eAmmunition, ObjectType.eColony
					sTemp = ""
				Case ObjectType.ePlayerIntel
					Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(mlID)
					If oRel Is Nothing = False Then
						If oRel.oPlayerIntel Is Nothing = False Then
							With oRel.oPlayerIntel
								sTemp = "Following stats for " & GetCacheObjectValue(mlID, ObjectType.ePlayer) & vbCrLf
								If .lDiplomacyScore <> Int32.MinValue Then
									sTemp &= "Diplomacy Score (" & GetDurationFromSeconds(CInt(Now.ToUniversalTime.Subtract(GetDateFromNumber(.lDiplomacyUpdate)).TotalSeconds), True) & ")" & vbCrLf
								End If
								If .lMilitaryScore <> Int32.MinValue Then
									sTemp &= "Military Score (" & GetDurationFromSeconds(CInt(Now.ToUniversalTime.Subtract(GetDateFromNumber(.lMilitaryUpdate)).TotalSeconds), True) & ")" & vbCrLf
								End If
								If .lPopulationScore <> Int32.MinValue Then
									sTemp &= "Population Score (" & GetDurationFromSeconds(CInt(Now.ToUniversalTime.Subtract(GetDateFromNumber(.lPopulationUpdate)).TotalSeconds), True) & ")" & vbCrLf
								End If
								If .lProductionScore <> Int32.MinValue Then
									sTemp &= "Production Score (" & GetDurationFromSeconds(CInt(Now.ToUniversalTime.Subtract(GetDateFromNumber(.lProductionUpdate)).TotalSeconds), True) & ")" & vbCrLf
								End If
								If .lTechnologyScore <> Int32.MinValue Then
									sTemp &= "Technology Score (" & GetDurationFromSeconds(CInt(Now.ToUniversalTime.Subtract(GetDateFromNumber(.lTechnologyUpdate)).TotalSeconds), True) & ")" & vbCrLf
								End If
								If .lWealthScore <> Int32.MinValue Then
									sTemp &= "Wealth Score (" & GetDurationFromSeconds(CInt(Now.ToUniversalTime.Subtract(GetDateFromNumber(.lWealthUpdate)).TotalSeconds), True) & ")" & vbCrLf
								End If
							End With
						End If
					End If
				Case ObjectType.ePlayerItemIntel
					For X As Int32 = 0 To glItemIntelUB
						If glItemIntelIdx(X) <> -1 Then
							Dim oItem As PlayerItemIntel = goItemIntel(X)
							If oItem Is Nothing = False Then
								If oItem.lItemID = mlID AndAlso oItem.iItemTypeID = miExtTypeID Then
									Select Case oItem.yIntelType
										Case PlayerItemIntel.PlayerItemIntelType.eStatus
											sTemp = "Location and status of " & GetCacheObjectValue(oItem.lItemID, oItem.iItemTypeID)
										Case PlayerItemIntel.PlayerItemIntelType.eFullKnowledge
											sTemp = "Full knowledge of " & GetCacheObjectValue(oItem.lItemID, oItem.iItemTypeID)
										Case Else
											sTemp = "Location of " & GetCacheObjectValue(oItem.lItemID, oItem.iItemTypeID)
									End Select
									Exit For
								End If
							End If
						End If
					Next X
				Case ObjectType.ePlayerTechKnowledge
					sTemp = GetNonOwnerItemData(mlID, miExtTypeID)
				Case Else
					sTemp = GetNonOwnerItemData(mlID, miTypeID)
			End Select
			Return sTemp
		End Function

		Private Sub txtSellQty_TextChanged() Handles txtSellQty.TextChanged
			If mlID = -1 OrElse miTypeID = -1 Then Return

			If txtSellQty.Caption <> "" Then
				Dim blTemp As Int64 = CLng(Val(txtSellQty.Caption.Replace(",", "")))

				If blTemp < 0 Then blTemp = 0
				If blTemp > mblQty Then blTemp = mblQty

				If blTemp.ToString("###0") <> txtSellQty.Caption Then
					txtSellQty.Caption = blTemp.ToString
				End If
				btnPlace.Caption = "Place Order"
			End If

		End Sub

		Private Sub btnPlace_Click(ByVal sName As String) Handles btnPlace.Click
			If mlID <> -1 AndAlso miTypeID <> -1 Then
				If btnPlace.Caption = "Confirm" Then
                    Dim blQty As Int64 = CLng(txtSellQty.Caption)
					If blQty < 1 Then
						MyBase.moUILib.AddNotification("Please enter a valid quantity to sell.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Return
					End If
                    Dim blPrice As Int64 = CLng(txtPrice.Caption)
					If blPrice < 1 Then
						MyBase.moUILib.AddNotification("Please enter a valid sell price.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Return
					End If

					'check if the player has enough slots
					Dim oTPC As TradePostContents = Nothing
					For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
						If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
							oTPC = TradePostContents.oTradePostContents(X)
							Exit For
						End If
					Next X
					If oTPC Is Nothing = False AndAlso oTPC.ySellSlotsUsed >= goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSellOrderSlots) Then
						MyBase.moUILib.AddNotification("You do not have Sell Slots available to place a new sell order.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Return
					End If

					btnPlace.Enabled = False


					For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
						If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
							Dim lTemp As Int32 = TradePostContents.oTradePostContents(X).ySellSlotsUsed
							lTemp += 1
							If lTemp < 0 Then lTemp = 0
							If lTemp > 255 Then lTemp = 255
							TradePostContents.oTradePostContents(X).ySellSlotsUsed = CByte(lTemp)
							Exit For
						End If
                    Next X

                    oTPC.SetItemPrice(blQty, blPrice, mlID, miTypeID)
                    mblPrice = blPrice
                    mblQty = blQty

					Dim yMsg(29) As Byte
					Dim lPos As Int32 = 0
					System.BitConverter.GetBytes(GlobalMessageCode.eSubmitSellOrder).CopyTo(yMsg, lPos) : lPos += 2
					System.BitConverter.GetBytes(mlTradePostID).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(mlID).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(CShort(miTypeID)).CopyTo(yMsg, lPos) : lPos += 2
					System.BitConverter.GetBytes(blQty).CopyTo(yMsg, lPos) : lPos += 8
					System.BitConverter.GetBytes(blPrice).CopyTo(yMsg, lPos) : lPos += 8
					System.BitConverter.GetBytes(miExtTypeID).CopyTo(yMsg, lPos) : lPos += 2
					If yMsg Is Nothing = False Then
						MyBase.moUILib.SendMsgToPrimary(yMsg)
						CType(Me.ParentControl, fraTradeContents).ListChangeEvent(True)
					Else : btnPlace.Caption = "Place Order"
					End If
					btnPlace.Enabled = True
				Else : btnPlace.Caption = "Confirm"
				End If
			Else : btnPlace.Caption = "Place Order"
			End If
		End Sub

		Private Sub btnCancelOrder_Click(ByVal sName As String) Handles btnCancelOrder.Click
			If mlID <> -1 AndAlso miTypeID <> -1 Then
                If btnCancelOrder.Caption = "Confirm" Then
                    btnCancelOrder.Enabled = False

                    If miTypeID = ObjectType.eUnit Then
                        Dim ofrmMsgBox As New frmMsgBox(goUILib, "Cancelling this sell order will result in the unit being recycled. You also incur a 25% sale price penalty to the GTC." & vbCrLf & vbCrLf & "Are you sure?", MsgBoxStyle.YesNo, "Confirm Cancel")
                        AddHandler ofrmMsgBox.DialogClosed, AddressOf CancelUnitMsgBoxResult
                        Return
                    End If

                    For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
                        If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
                            Dim lTemp As Int32 = TradePostContents.oTradePostContents(X).ySellSlotsUsed
                            lTemp -= 1
                            If lTemp < 0 Then lTemp = 0
                            If lTemp > 255 Then lTemp = 255
                            TradePostContents.oTradePostContents(X).ySellSlotsUsed = CByte(lTemp)
                            TradePostContents.oTradePostContents(X).ClearPrice(mlID, miTypeID, miExtTypeID)
                            Exit For
                        End If
                    Next X


                    Dim yMsg(29) As Byte
                    Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSubmitSellOrder).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(mlTradePostID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(mlID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(CShort(miTypeID)).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(-1L).CopyTo(yMsg, lPos) : lPos += 8
                    System.BitConverter.GetBytes(-1L).CopyTo(yMsg, lPos) : lPos += 8
                    System.BitConverter.GetBytes(miExtTypeID).CopyTo(yMsg, lPos) : lPos += 2
                    MyBase.moUILib.SendMsgToPrimary(yMsg)
                    CType(Me.ParentControl, fraTradeContents).SetItemBold(mlID, miTypeID, False)

                    btnCancelOrder.Caption = "Cancel Order"
                    btnCancelOrder.Enabled = True
                    btnCancelOrder.Visible = False
                    CType(Me.ParentControl, fraTradeContents).ListChangeEvent(False)

                Else : btnCancelOrder.Caption = "Confirm"
                End If
			Else : btnCancelOrder.Caption = "Cancel Order"
			End If
        End Sub

        Private Sub CancelUnitMsgBoxResult(ByVal lResult As MsgBoxResult)
            If lResult = MsgBoxResult.Yes Then
                For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
                    If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
                        Dim lTemp As Int32 = TradePostContents.oTradePostContents(X).ySellSlotsUsed
                        lTemp -= 1
                        If lTemp < 0 Then lTemp = 0
                        If lTemp > 255 Then lTemp = 255
                        TradePostContents.oTradePostContents(X).ySellSlotsUsed = CByte(lTemp)
                        TradePostContents.oTradePostContents(X).ClearPrice(mlID, miTypeID, miExtTypeID)
                        Exit For
                    End If
                Next X

                Dim yMsg(29) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSubmitSellOrder).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(mlTradePostID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(mlID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(CShort(miTypeID)).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(-1L).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(-1L).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(miExtTypeID).CopyTo(yMsg, lPos) : lPos += 2
                MyBase.moUILib.SendMsgToPrimary(yMsg)
                CType(Me.ParentControl, fraTradeContents).SetItemBold(mlID, miTypeID, False)
            End If
            btnCancelOrder.Caption = "Cancel Order"
            btnCancelOrder.Enabled = True
            btnCancelOrder.Visible = False
            CType(Me.ParentControl, fraTradeContents).ListChangeEvent(False)
        End Sub
	End Class
#End Region

#Region "  Buy Order Detail Panel (for Buy Orders)  "
    'Interface created from Interface Builder
    Private Class fraBuyOrderDetails
        Inherits fraDetailPane

        Private txtItemDetails As UITextBox
        Private lblEscrow As UILabel
        Private lblPayment As UILabel
        Private lblQuantity As UILabel
        Private lblBuyer As UILabel
        Private lblDeadline As UILabel
        Private lblDelivery As UILabel
        Private lblFinalDeadline As UILabel
        Private lblAcceptedBy As UILabel
        Private lblAcceptedOn As UILabel

        Private moOrder As BuyOrder = Nothing
        Private mlTradePostID As Int32 = -1

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraBuyOrderDetails initial props
            With Me
                .ControlName = "fraBuyOrderDetails"
                .Left = 98
                .Top = 160
                .Width = 520
                .Height = 180
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .mbAcceptReprocessEvents = True
                .BorderLineWidth = 2
            End With

            'txtItemDetails initial props
            txtItemDetails = New UITextBox(oUILib)
            With txtItemDetails
                .ControlName = "txtItemDetails"
                .Left = 5
                .Top = 5
                .Width = 250
                .Height = 170
                .Enabled = True
                .Visible = True
                .Caption = "Requesting Buy Order Requirements..."
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(0, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 0
                .MultiLine = True
                .Locked = True
                .BorderColor = muSettings.InterfaceBorderColor
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(txtItemDetails, UIControl))

            'lblEscrow initial props
            lblEscrow = New UILabel(oUILib)
            With lblEscrow
                .ControlName = "lblEscrow"
                .Left = 260
                .Top = 5
                .Width = 250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Escrow: "
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "The amount of credits deducted from the acceptor's account upon acceptance." & vbCrLf & "On successful completion of the Buy Order, the acceptor receives the credits." & vbCrLf & "Failure to deliver the buy order by the Deadline results in the Buyer receiving the credits."
            End With
            Me.AddChild(CType(lblEscrow, UIControl))

            'lblPayment initial props
            lblPayment = New UILabel(oUILib)
            With lblPayment
                .ControlName = "lblPayment"
                .Left = 260
                .Top = 23
                .Width = 250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Payment:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "The amount of credits received upon timely completion of the buy order." & vbCrLf & "This is in addition to the credits locked in Escrow."
            End With
            Me.AddChild(CType(lblPayment, UIControl))

            'lblQuantity initial props
            lblQuantity = New UILabel(oUILib)
            With lblQuantity
                .ControlName = "lblQuantity"
                .Left = 260
                .Top = 41
                .Width = 250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Quantity:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "The number of items described by the requirements required to fulfill the buy order."
            End With
            Me.AddChild(CType(lblQuantity, UIControl))

            'lblBuyer initial props
            lblBuyer = New UILabel(oUILib)
            With lblBuyer
                .ControlName = "lblBuyer"
                .Left = 260
                .Top = 59
                .Width = 250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Buyer:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Indicates who initiated the Buy Order and will receive the goods upon completion."
            End With
            Me.AddChild(CType(lblBuyer, UIControl))

            'lblDeadline initial props
            lblDeadline = New UILabel(oUILib)
            With lblDeadline
                .ControlName = "lblDeadline"
                .Left = 260
                .Top = 77
                .Width = 250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Deadline:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "The absolute deadline for the goods to be delivered." & vbCrLf & "Failure to deliver the goods in time forfeits the Escrow."
            End With
            Me.AddChild(CType(lblDeadline, UIControl))

            'lblDelivery initial props
            lblDelivery = New UILabel(oUILib)
            With lblDelivery
                .ControlName = "lblDelivery"
                .Left = 260
                .Top = 95
                .Width = 250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Delivery Time:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "The amount of time it will take GTC to deliver" & vbCrLf & "the goods from this tradepost to the buyer."
            End With
            Me.AddChild(CType(lblDelivery, UIControl))

            'lblFinalDeadline initial props
            lblFinalDeadline = New UILabel(oUILib)
            With lblFinalDeadline
                .ControlName = "lblFinalDeadline"
                .Left = 260
                .Top = 113
                .Width = 250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Adjusted Deadline:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "The absolute latest the goods need to be handed over to the GTC for delivery." & vbCrLf & "This considers the Deadline and the time for Delivery."
            End With
            Me.AddChild(CType(lblFinalDeadline, UIControl))

            'lblAcceptedBy initial props
            lblAcceptedBy = New UILabel(oUILib)
            With lblAcceptedBy
                .ControlName = "lblAcceptedBy"
                .Left = 260
                .Top = 131
                .Width = 250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Accepted By:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Who has accepted the buy order and will fulfill it."
            End With
            Me.AddChild(CType(lblAcceptedBy, UIControl))

            'lblAcceptedOn initial props
            lblAcceptedOn = New UILabel(oUILib)
            With lblAcceptedOn
                .ControlName = "lblAcceptedOn"
                .Left = 260
                .Top = 149
                .Width = 250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Accepted On:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "When the buy order was accepted."
            End With
            Me.AddChild(CType(lblAcceptedOn, UIControl))
        End Sub

        Public Sub SetDetails(ByRef oOrder As BuyOrder, ByVal lFromTPID As Int32)
            moOrder = oOrder
            mlTradePostID = lFromTPID
            moOrder.RequestDetails(lFromTPID)
        End Sub

        Public Overrides Sub SubFrame_OnNewFrame()
            If moOrder Is Nothing = False Then
                With moOrder
                    Dim sTemp As String = .GetItemDetailsText()
                    If txtItemDetails.Caption <> sTemp Then txtItemDetails.Caption = sTemp

                    sTemp = "Escrow: " & .blEscrow.ToString("#,##0")
                    If lblEscrow.Caption <> sTemp Then lblEscrow.Caption = sTemp

                    sTemp = "Payment: " & .blPrice.ToString("#,##0")
                    If lblPayment.Caption <> sTemp Then lblPayment.Caption = sTemp

                    sTemp = "Quantity: " & .blQty.ToString("#,##0")
                    If lblQuantity.Caption <> sTemp Then lblQuantity.Caption = sTemp

                    sTemp = "Buyer: " & GetCacheObjectValue(.lSellerID, ObjectType.ePlayer)
                    If lblBuyer.Caption <> sTemp Then lblBuyer.Caption = sTemp

                    sTemp = "Deadline: " & .GetDeadlineText()
                    If lblDeadline.Caption <> sTemp Then lblDeadline.Caption = sTemp

                    sTemp = "Delivery Time: " & GetDurationFromSeconds(.lDeliveryTime, False)
                    If lblDelivery.Caption <> sTemp Then lblDelivery.Caption = sTemp

                    sTemp = "Adjusted Deadline: " & .GetAdjustedDeadline()
                    If lblFinalDeadline.Caption <> sTemp Then lblFinalDeadline.Caption = sTemp

                    If .lAcceptedByID = -1 Then
                        sTemp = "Accepted By: Not Yet Accepted"
                    Else : sTemp = "Accepted By: " & GetCacheObjectValue(.lAcceptedByID, ObjectType.ePlayer)
                    End If
                    If lblAcceptedBy.Caption <> sTemp Then lblAcceptedBy.Caption = sTemp

                    If .lAcceptedByID = -1 Then
                        sTemp = "Accepted On: Not Yet Accepted"
                    Else : sTemp = "Accepted On: " & .GetAcceptedOnText()
                    End If
                    If lblAcceptedOn.Caption <> sTemp Then lblAcceptedOn.Caption = sTemp
                End With
            End If
        End Sub
    End Class
#End Region

	Public Shared Sub FillListWithMarketList(ByRef lstData As UIListBox, ByVal yShowState As Byte)

		Dim lCurrItemData As Int32 = -1
		If lstData.ListIndex > -1 Then lCurrItemData = lstData.ItemData(lstData.ListIndex)

		lstData.Clear()
		If lstData.oIconTexture Is Nothing OrElse lstData.oIconTexture.Disposed = True Then
			lstData.oIconTexture = goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
		End If
		lstData.RenderIcons = True

		lstData.AddItem("COMPONENTS", True) : lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True : lstData.ApplyIconOffset(lstData.NewIndex) = False

		lstData.AddItem("Armor", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eArmorComponent
		lstData.ApplyIconOffset(lstData.NewIndex) = True
		lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(0, 0, 16, 16)
		lstData.IconForeColor(lstData.NewIndex) = Color.White

		lstData.AddItem("Beam (Pulse)", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eBeamPulseComponent
		lstData.ApplyIconOffset(lstData.NewIndex) = True
		lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(32, 32, 16, 16)
		lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 0, 0)

		lstData.AddItem("Beam (Solid)", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eBeamSolidComponent
		lstData.ApplyIconOffset(lstData.NewIndex) = True
		lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(32, 32, 16, 16)
		lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 0, 0)

		lstData.AddItem("Engine", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eEngineComponent
		lstData.ApplyIconOffset(lstData.NewIndex) = True
		lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(0, 16, 16, 16)
		lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 0)

		lstData.AddItem("Missile", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eMissileComponent
		lstData.ApplyIconOffset(lstData.NewIndex) = True
		lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(32, 32, 16, 16)
		lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 0, 0)

		lstData.AddItem("Projectile", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eProjectileComponent
		lstData.ApplyIconOffset(lstData.NewIndex) = True
		lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(32, 32, 16, 16)
		lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 0, 0)

		lstData.AddItem("Radar", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eRadarComponent
		lstData.ApplyIconOffset(lstData.NewIndex) = True
		lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(48, 0, 16, 16)
		lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 0)

		lstData.AddItem("Shield", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eShieldComponent
		lstData.ApplyIconOffset(lstData.NewIndex) = True
		lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 16, 16, 16)
		lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 255)

        'lstData.AddItem("Specials", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eSpecialComponent
        'lstData.ApplyIconOffset(lstData.NewIndex) = True

		'lstData.AddItem("COMPONENT DESIGNS", True) : lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
		'lstData.AddItem("Armor", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eArmorDesign
		'lstData.AddItem("Beam (Pulse)", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eBeamPulseDesign
		'lstData.AddItem("Beam (Solid)", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eBeamSolidDesign
		'lstData.AddItem("Engine", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eEngineDesign
		'lstData.AddItem("Missile", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eMissileDesign
		'lstData.AddItem("Projectile", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eProjectileDesign
		'lstData.AddItem("Radar", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eRadarDesign
		'lstData.AddItem("Shield", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eShieldDesign
		lstData.AddItem("MATERIALS", True) : lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True : lstData.ApplyIconOffset(lstData.NewIndex) = False
		lstData.AddItem("Minerals", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eMinerals
        lstData.ApplyIconOffset(lstData.NewIndex) = True

        lstData.AddItem("Alloys", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eAlloys
		lstData.ApplyIconOffset(lstData.NewIndex) = True

		'lstData.AddItem("PERSONNEL", True) : lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True
		'lstData.AddItem("Agents", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eAgent
		'lstData.AddItem("Colonists", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eColonists
		'lstData.AddItem("Enlisted", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eEnlisted
		'lstData.AddItem("Officers", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eOfficers

		If yShowState <> 4 Then
			lstData.AddItem("PLAYER INTEL", True) : lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True : lstData.ApplyIconOffset(lstData.NewIndex) = False
			lstData.AddItem("Player Stats", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.ePlayerStats : lstData.ApplyIconOffset(lstData.NewIndex) = True
			lstData.AddItem("Component Design Specs", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.ePlayerComponentDesigns : lstData.ApplyIconOffset(lstData.NewIndex) = True
			lstData.AddItem("Colony Locations", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eColonyLocation : lstData.ApplyIconOffset(lstData.NewIndex) = True
			'lstData.AddItem("Foreign Relations List", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eForeignRelationsList : lstData.ApplyIconOffset(lstData.NewIndex) = True
			'lstData.AddItem("Player Senate Vote History", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.ePlayerSenateVoteHistory : lstData.ApplyIconOffset(lstData.NewIndex) = True
			'lstData.AddItem("Budget Details", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eBudgetDetails : lstData.ApplyIconOffset(lstData.NewIndex) = True
			lstData.AddItem("UNITS", True) : lstData.ItemData(lstData.NewIndex) = -1 : lstData.ItemLocked(lstData.NewIndex) = True : lstData.ApplyIconOffset(lstData.NewIndex) = False
			lstData.AddItem("Land-Based", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eLandBasedUnit : lstData.ApplyIconOffset(lstData.NewIndex) = True
			lstData.AddItem("Aerial", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eAerialUnit : lstData.ApplyIconOffset(lstData.NewIndex) = True
            lstData.AddItem("Space-Based", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eSpaceBasedUnit : lstData.ApplyIconOffset(lstData.NewIndex) = True
            'MSC - 06/19/08 - unremark this when naval trading works
            'lstData.AddItem("Naval", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eNavalUnit : lstData.ApplyIconOffset(lstData.NewIndex) = True
        Else
            lstData.AddItem("Specific Minerals", False) : lstData.ItemData(lstData.NewIndex) = MarketListType.eBuyOrderSpecificMineral
            lstData.ApplyIconOffset(lstData.NewIndex) = True
        End If

		If lCurrItemData > -1 Then
			For X As Int32 = 0 To lstData.ListCount - 1
				If lstData.ItemData(X) = lCurrItemData Then
					lstData.ListIndex = X
					Exit For
				End If
			Next X
		End If
	End Sub

    Private Sub chkThisEnvironmentOnly_Click() Handles chkThisEnvironmentOnly.Click
        lstContents_ItemClick(lstContents.ListIndex)
        mbLastThisEnvironmentOnly = chkThisEnvironmentOnly.Value
    End Sub
End Class

Public Enum MarketListType As Byte
    eNotUsed = 0
    eMinerals = 1
    eAlloys = 2
    eArmorComponent = 3
    eBeamPulseComponent = 4
    eBeamSolidComponent = 5
    eEngineComponent = 6
    eMissileComponent = 7
    eProjectileComponent = 8
    eRadarComponent = 9
    eShieldComponent = 10
    eSpecialComponent = 11
    eArmorDesign = 12
    eBeamPulseDesign = 13
    eBeamSolidDesign = 14
    eEngineDesign = 15
    eMissileDesign = 16
    eProjectileDesign = 17
    eRadarDesign = 18
    eShieldDesign = 19
    eLandBasedUnit = 20
    eAerialUnit = 21
    eSpaceBasedUnit = 22
    eNavalUnit = 23
    eSpaceStation = 24
    eAgent = 25
    eColonists = 26
    eEnlisted = 27
    eOfficers = 28
    ePlayerStats = 29 '- scores as of a certain time
    ePlayerComponentDesigns = 30 '- a design schematic
    eColonyLocation = 31 '- bookmarks of colony locations
    eForeignRelationsList = 32 '- the entire diplomacy screen for a player
    ePlayerSenateVoteHistory = 33 '- list of vote history the player took in senate votes of past?
    eBudgetDetails = 34 '- as of a certain time

    eBuyOrderSpecificMineral = 35       'specific mineral

    eBombComponent = 36
    eMineWeaponComponent = 37

    eSellOrderShift = 128           'bitwise item
End Enum