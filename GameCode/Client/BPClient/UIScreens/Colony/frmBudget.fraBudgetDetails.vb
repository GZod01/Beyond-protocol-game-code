Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmBudget
    'Interface created from Interface Builder
    Public Class fraBudgetDetails
        Inherits UIWindow

        Private WithEvents vscrScroll As UIScrollBar
        Private lblH_Type As UILabel
        Private lblH_Envir As UILabel
        Private lblH_Colony As UILabel
        Private lblH_Revenue As UILabel
        Private lblH_Expense As UILabel
        Private lblH_TaxRate As UILabel
        Private lblH_Control As UILabel
        Private lblH_ColonyCnt As UILabel
        Private lnDiv1 As UILine

        Private moItems() As ctlBudgetDetailItem
        Private mlItemUB As Int32 = -1

        Private moList() As Budget.BudgetEnvir
        Private mlListUB As Int32 = -1
        Private mlSelectedItem As Int32 = -1

        Private mlDisplayItemCnt As Int32 = 1
        Private mlLastScrollVal As Int32

        Private mlLastUpdate As Int32

        Private mbHasUnknowns As Boolean = False
        Private moBudget As Budget
        Private Shared mlSortType As Budget.BudgetSort = Budget.BudgetSort.EnvironmentName

        Private Shared mlExpandedSystemIDs() As Int32 = Nothing
        Private Shared mlExpandedSystemUB As Int32 = -1
        Private Shared mlScrollPosition As Int32 = 0
        Private mbScrollTo As Boolean = False

        Public Event ItemSelected(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
        Public Event ListItemDoubleClicked(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraBudgetDetails initial props
            With Me
                .ControlName = "fraBudgetDetails"
                .Left = 101
                .Top = 132
                .Width = 800
                .Height = 200
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 1
                .Moveable = False
                .Caption = "Budget Details By Location"
            End With

            'vscrScroll initial props
            vscrScroll = New UIScrollBar(oUILib, True)
            With vscrScroll
                .ControlName = "vscrScroll"
                .Left = 775
                .Top = 34
                .Width = 24
                .Height = 165
                .Enabled = True
                .Visible = True
                .Value = 0
                .MaxValue = 100
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = True
            End With
            Me.AddChild(CType(vscrScroll, UIControl))

            'lblH_Type initial props
            lblH_Type = New UILabel(oUILib)
            With lblH_Type
                .ControlName = "lblH_Type"
                .Left = 10
                .Top = 10
                .Width = 40
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Type"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(8, DrawTextFormat)
                .bAcceptEvents = True
            End With
            Me.AddChild(CType(lblH_Type, UIControl))

            'lnDiv1 initial props
            lnDiv1 = New UILine(oUILib)
            With lnDiv1
                .ControlName = "lnDiv1"
                .Left = 0
                .Top = 30
                .Width = 800
                .Height = 0
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(lnDiv1, UIControl))

            'lblH_Envir initial props
            lblH_Envir = New UILabel(oUILib)
            With lblH_Envir
                .ControlName = "lblH_Envir"
                .Left = 55
                .Top = 10
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Environment Name"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(8, DrawTextFormat)
                .bAcceptEvents = True
            End With
            Me.AddChild(CType(lblH_Envir, UIControl))

            'lblH_ColonyCnt initial props
            lblH_ColonyCnt = New UILabel(oUILib)
            With lblH_ColonyCnt
                .ControlName = "lblH_ColonyCnt"
                .Left = 310
                .Top = 10
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Colonies"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(8, DrawTextFormat)
                .bAcceptEvents = True
            End With
            Me.AddChild(CType(lblH_ColonyCnt, UIControl))

            'lblH_Colony initial props
            lblH_Colony = New UILabel(oUILib)
            With lblH_Colony
                .ControlName = "lblH_Colony"
                .Left = 200
                .Top = 10
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Colony Name"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(8, DrawTextFormat)
                .bAcceptEvents = True
            End With
            Me.AddChild(CType(lblH_Colony, UIControl))

            'lblH_Revenue initial props
            lblH_Revenue = New UILabel(oUILib)
            With lblH_Revenue
                .ControlName = "lblH_Revenue"
                .Left = 350 ' 360
                .Top = 10
                .Width = 135 '150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Revenue"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.Bottom
                .bAcceptEvents = True
            End With
            Me.AddChild(CType(lblH_Revenue, UIControl))

            'lblH_Expense initial props
            lblH_Expense = New UILabel(oUILib)
            With lblH_Expense
                .ControlName = "lblH_Expense"
                .Left = 495 '520
                .Top = 10
                .Width = 135 '150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Expense"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.Bottom
                .bAcceptEvents = True
            End With
            Me.AddChild(CType(lblH_Expense, UIControl))

            'lblH_Control initial props
            lblH_Control = New UILabel(oUILib)
            With lblH_Control
                .ControlName = "lblH_Control"
                .Left = 630 '640
                .Top = 10
                .Width = 50 '30
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Control"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.Bottom
                .bAcceptEvents = True
            End With
            Me.AddChild(CType(lblH_Control, UIControl))

            'lblH_TaxRate initial props
            lblH_TaxRate = New UILabel(oUILib)
            With lblH_TaxRate
                .ControlName = "lblH_TaxRate"
                .Left = 695
                .Top = 10
                .Width = 90
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Tax Rate"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(8, DrawTextFormat)
                .bAcceptEvents = True
            End With
            Me.AddChild(CType(lblH_TaxRate, UIControl))

            AddHandler lblH_Colony.OnMouseDown, AddressOf lblH_ColonyClick
            AddHandler lblH_Envir.OnMouseDown, AddressOf lblH_EnvirClick
            AddHandler lblH_Expense.OnMouseDown, AddressOf lblH_ExpenseClick
            AddHandler lblH_Revenue.OnMouseDown, AddressOf lblH_RevenueClick
            AddHandler lblH_TaxRate.OnMouseDown, AddressOf lblH_TaxRateClick
            AddHandler lblH_Type.OnMouseDown, AddressOf lblH_TypeClick
            AddHandler lblH_Control.OnMouseDown, AddressOf lblH_ControlClick

            If muSettings.BudgetWindowStaticExpand = False Then
                mlExpandedSystemUB = -1
                ReDim mlExpandedSystemIDs(mlExpandedSystemUB)
            End If
            If mlScrollPosition > 0 Then mbScrollTo = True
        End Sub

        Private Sub ItemClicked(ByVal mlEnvirID As Int32, ByVal miEnvirTypeID As Int16)
            mlSelectedItem = -1
            For X As Int32 = 0 To mlListUB
                If moList(X).lEnvirID = mlEnvirID AndAlso moList(X).iEnvirTypeID = miEnvirTypeID Then
                    mlSelectedItem = X
                    Exit For
                End If
            Next X

            RaiseEvent ItemSelected(mlEnvirID, miEnvirTypeID)

            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBudgetItemSelected, mlEnvirID, miEnvirTypeID, -1, "")
            End If
        End Sub

        Private Sub ItemDoubleClicked(ByVal lEnvirID As Int32, ByVal iTypeID As Int16)
            RaiseEvent ListItemDoubleClicked(lEnvirID, iTypeID)
        End Sub

        Private Sub AlertButtonClicked(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
            Dim yMsg(7) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eClearBudgetAlert).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2

            For X As Int32 = 0 To moBudget.mlItemUB
                If moBudget.muItems(X).lEnvirID = lEnvirID AndAlso moBudget.muItems(X).iEnvirTypeID = iEnvirTypeID Then
                    moBudget.muItems(X).bHasConflict = False
                    Exit For
                End If
            Next X

            MyBase.moUILib.SendMsgToPrimary(yMsg)
            MyBase.moUILib.AddNotification("Sending the all-clear signal.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End Sub

        Private Sub ExpandChanged(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal bNewValue As Boolean)
            If iEnvirTypeID <> ObjectType.eSolarSystem Then Return
            If bNewValue = True Then
                'ok, adding it
                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To mlExpandedSystemUB
                    If mlExpandedSystemIDs(X) = -1 Then
                        If lIdx = -1 Then
                            lIdx = X
                        End If
                    ElseIf mlExpandedSystemIDs(X) = lEnvirID Then
                        Return
                    End If
                Next X

                If lIdx = -1 Then
                    ReDim Preserve mlExpandedSystemIDs(mlExpandedSystemUB + 1)
                    mlExpandedSystemIDs(mlExpandedSystemUB + 1) = lEnvirID
                    mlExpandedSystemUB += 1
                Else
                    mlExpandedSystemIDs(lIdx) = lEnvirID
                End If
            Else
                For X As Int32 = 0 To mlExpandedSystemUB
                    If mlExpandedSystemIDs(X) = lEnvirID Then
                        mlExpandedSystemIDs(X) = -1
                    End If
                Next
            End If
        End Sub

        Public Sub SetList(ByRef oBudget As Budget)
            moBudget = oBudget

            Try
                mlListUB = oBudget.mlItemUB
                ReDim moList(mlListUB)

                'oBudget.SortItems(mlSortType)

                For X As Int32 = 0 To oBudget.mlItemUB
                    moList(X) = oBudget.muItems(X)
                Next X

                Dim lTmp As Int32 = oBudget.mlItemUB + 1
                lTmp -= mlDisplayItemCnt
                If lTmp < 0 Then
                    vscrScroll.Value = 0
                    vscrScroll.MaxValue = 0
                Else : vscrScroll.MaxValue = lTmp
                End If

                ''Ok, determine our viewable item count
                'Dim lVisibleItems As Int32 = 0
                'For lListIdx As Int32 = 0 To mlListUB
                '    With moList(lListIdx)
                '        'Ok, determine if this item is a system
                '        Dim bFound As Boolean = False
                '        For X As Int32 = 0 To mlExpandedSystemUB
                '            If mlExpandedSystemIDs(X) = .lSystemID Then
                '                bFound = True
                '                Exit For
                '            End If
                '        Next X
                '        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255 Then bFound = True
                '        If bFound = True OrElse .iEnvirTypeID = ObjectType.eSolarSystem Then
                '            lVisibleItems += 1
                '        End If
                '    End With
                'Next lListIdx

                'Dim lTmp As Int32 = lVisibleItems
                'lTmp -= mlDisplayItemCnt
                'If lTmp < 0 Then
                '    If vscrScroll.Value <> 0 Then vscrScroll.Value = 0
                '    If vscrScroll.MaxValue <> 0 Then vscrScroll.MaxValue = 0
                'Else
                '    If vscrScroll.MaxValue <> lTmp Then vscrScroll.MaxValue = lTmp
                'End If
            Catch
            End Try
            If mbScrollTo = True Then
                vscrScroll.Value = mlScrollPosition
                mbScrollTo = False
            End If
            Me.IsDirty = True
        End Sub

        Private mlLastFlashCycle As Int32 = 0
        Private mlFlashDelay As Int32 = 30
        Private mbFlashState As Boolean = False

        Private Sub fraBudgetDetails_OnNewFrame() Handles Me.OnNewFrame
            If mbHasUnknowns = True AndAlso glCurrentCycle - mlLastUpdate > 30 Then
                Me.IsDirty = True
            End If
        End Sub

        Public Sub CheckFlashStates()
            If glCurrentCycle - mlLastFlashCycle > mlFlashDelay Then
                mbFlashState = Not mbFlashState
                mlLastFlashCycle = glCurrentCycle
                If mbFlashState = True Then
                    mlFlashDelay = 30
                Else : mlFlashDelay = 7
                End If
                Dim bIsDirty As Boolean = False
                For X As Int32 = 0 To mlItemUB
                    If moItems(X) Is Nothing = False Then
                        bIsDirty = bIsDirty OrElse moItems(X).SetFlashState(mbFlashState)
                    End If
                Next X
                Me.IsDirty = bIsDirty
            End If
        End Sub

        'Private Sub fraBudgetDetails_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles Me.OnRender
        Private Sub fraBudgetDetails_OnRender() Handles Me.OnRender
            'First, check if we need to create the detail item classes...
            If mlItemUB <> mlDisplayItemCnt - 1 Then
                Dim bDone As Boolean = False
                While bDone = False
                    bDone = True
                    For X As Int32 = 0 To Me.ChildrenUB
                        If Me.moChildren(X).ControlName.ToUpper = "CTLBUDGETDETAILITEM" Then
                            bDone = False
                            Dim oTmp As ctlBudgetDetailItem = CType(moChildren(X), ctlBudgetDetailItem)
                            RemoveHandler oTmp.DetailItemClick, AddressOf ItemClicked
                            RemoveHandler oTmp.DetailItemDoubleClick, AddressOf ItemDoubleClicked
                            RemoveHandler oTmp.AlertButtonClick, AddressOf AlertButtonClicked
                            RemoveHandler oTmp.ExpandChanged, AddressOf ExpandChanged
                            oTmp = Nothing
                            Me.RemoveChild(X)
                            Exit For
                        End If
                    Next X
                End While

                mlItemUB = mlDisplayItemCnt - 1
                ReDim moItems(mlItemUB)
                For X As Int32 = 0 To mlItemUB
                    moItems(X) = New ctlBudgetDetailItem(MyBase.moUILib)
                    With moItems(X)
                        .Left = 9
                        .Top = lnDiv1.Top + (X * .Height) '+ 5
                        .Visible = False
                    End With
                    Me.AddChild(CType(moItems(X), UIControl))
                    AddHandler moItems(X).DetailItemClick, AddressOf ItemClicked
                    AddHandler moItems(X).DetailItemDoubleClick, AddressOf ItemDoubleClicked
                    AddHandler moItems(X).AlertButtonClick, AddressOf AlertButtonClicked
                    AddHandler moItems(X).ExpandChanged, AddressOf ExpandChanged
                Next X
            End If

            'Now, refresh our display
            mbHasUnknowns = False

            ''Ok, determine our viewable item count
            'Dim lVisibleItems As Int32 = 0
            'For lListIdx As Int32 = 0 To mlListUB
            '    With moList(lListIdx)
            '        'Ok, determine if this item is a system
            '        Dim bFound As Boolean = False
            '        For X As Int32 = 0 To mlExpandedSystemUB
            '            If mlExpandedSystemIDs(X) = .lSystemID Then
            '                bFound = True
            '                Exit For
            '            End If
            '        Next X
            '        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255 Then bFound = True
            '        If bFound = True OrElse .iEnvirTypeID = ObjectType.eSolarSystem Then
            '            lVisibleItems += 1
            '        End If
            '    End With
            'Next lListIdx

            'Dim lTmp As Int32 = lVisibleItems
            'lTmp -= mlDisplayItemCnt
            'If lTmp < 0 Then
            '    If vscrScroll.Value <> 0 Then vscrScroll.Value = 0
            '    If vscrScroll.MaxValue <> 0 Then vscrScroll.MaxValue = 0
            'Else
            '    If vscrScroll.MaxValue <> lTmp Then vscrScroll.MaxValue = lTmp
            'End If

            Dim lCtrlIdx As Int32 = 0
            Dim lItemIdx As Int32 = vscrScroll.Value
            While lCtrlIdx < mlDisplayItemCnt
                If lItemIdx > mlListUB Then
                    If moItems(lCtrlIdx).Visible = True Then moItems(lCtrlIdx).Visible = False
                    moItems(lCtrlIdx).Selected = False
                Else
                    With moList(lItemIdx)

                        'Ok, determine if this item is a system
                        Dim bFound As Boolean = False
                        For X As Int32 = 0 To mlExpandedSystemUB
                            If mlExpandedSystemIDs(X) = .lSystemID Then
                                bFound = True
                                Exit For
                            End If
                        Next X
                        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255 Then bFound = True
                        If bFound = False AndAlso .iEnvirTypeID <> ObjectType.eSolarSystem Then
                            lItemIdx += 1
                            Continue While
                        End If

                        If lCtrlIdx = 0 Then
                            If lItemIdx <> vscrScroll.Value Then
                                mlPreviousValue = lItemIdx
                                vscrScroll.Value = lItemIdx
                            End If
                        End If

                        If moItems(lCtrlIdx).Visible = False Then moItems(lCtrlIdx).Visible = True
                        moItems(lCtrlIdx).Selected = (lItemIdx = mlSelectedItem)

                        Dim sEnvir As String = GetCacheObjectValue(.lEnvirID, .iEnvirTypeID)
                        Dim sColony As String = GetCacheObjectValue(.lColonyID, ObjectType.eColony)

                        If sEnvir.ToUpper = "UNKNOWN" OrElse (.lColonyID > 0 AndAlso sColony.ToUpper = "UNKNOWN") Then mbHasUnknowns = True

                        Dim blTotalExp As Int64 = .GetTotalExpense
                        'If frmBudget.bShowHWSupply = True Then blTotalExp += .GetHomeworldSupplyLineCost(goCurrentPlayer.oBudget.lIronCurtainPlanet)

                        If .iEnvirTypeID = ObjectType.eSolarSystem Then
                            moItems(lCtrlIdx).SetValues(.lEnvirID, .iEnvirTypeID, sEnvir, sColony, .GetTotalRevenue + .iTotalStarIncome, blTotalExp + .iTotalStarExpense, .yTaxRate, .lColonyID, .bHasCC, .bHasConflict, .yPlanetaryControl, -1, -1)
                        Else
                            moItems(lCtrlIdx).SetValues(.lEnvirID, .iEnvirTypeID, sEnvir, sColony, .GetTotalRevenue, blTotalExp, .yTaxRate, .lColonyID, .bHasCC, .bHasConflict, .yPlanetaryControl, .lCurrentColonyCount, .lPlanetCap)
                        End If


                        If bFound = False AndAlso .iEnvirTypeID = ObjectType.eSolarSystem Then
                            moItems(lCtrlIdx).bExpanded = False
                        Else
                            moItems(lCtrlIdx).bExpanded = True
                        End If
                        If .iEnvirTypeID = ObjectType.eSolarSystem Then
                            If .lChildrenItems = 0 Then moItems(lCtrlIdx).SetExpandButtonVisible(False) Else moItems(lCtrlIdx).SetExpandButtonVisible(True)
                        End If
                    End With
                End If
                lItemIdx += 1
                lCtrlIdx += 1
            End While


            'For X As Int32 = 0 To mlDisplayItemCnt - 1
            '    Dim lIdx As Int32 = vscrScroll.Value + X
            '    If lIdx > mlListUB Then
            '        If moItems(X).Visible = True Then moItems(X).Visible = False
            '        moItems(X).Selected = False
            '    Else
            '        If moItems(X).Visible = False Then moItems(X).Visible = True
            '        moItems(X).Selected = (lIdx = mlSelectedItem)
            '        With moList(lIdx)
            '            Dim sEnvir As String = GetCacheObjectValue(.lEnvirID, .iEnvirTypeID)
            '            Dim sColony As String = GetCacheObjectValue(.lColonyID, ObjectType.eColony)

            '            If sEnvir.ToUpper = "UNKNOWN" OrElse (.lColonyID > 0 AndAlso sColony.ToUpper = "UNKNOWN") Then mbHasUnknowns = True

            '            Dim blTotalExp As Int64 = .GetTotalExpense
            '            'If frmBudget.bShowHWSupply = True Then blTotalExp += .GetHomeworldSupplyLineCost(goCurrentPlayer.oBudget.lIronCurtainPlanet)
            '            moItems(X).SetValues(.lEnvirID, .iEnvirTypeID, sEnvir, sColony, .GetTotalRevenue, blTotalExp, .yTaxRate, .lColonyID, .bHasCC, .bHasConflict, .yPlanetaryControl)
            '        End With
            '    End If
            'Next X

            'and set our last forced update
            mlLastUpdate = glCurrentCycle
        End Sub

        Private Sub fraBudgetDetails_OnRenderEnd() Handles Me.OnRenderEnd
            'ok, determine which label is sorting
            Dim bDescending As Boolean = (mlSortType And Budget.BudgetSort.DescendingOrderShift) <> 0
            Dim lSortType As Budget.BudgetSort = mlSortType
            If bDescending = True Then lSortType = lSortType Xor Budget.BudgetSort.DescendingOrderShift

            'Dim lbl As UILabel = Nothing
            Dim lLeft As Int32 = Int32.MinValue
            Dim lTop As Int32 = -1
            Select Case lSortType
                Case Budget.BudgetSort.ColonyName
                    'lbl = lblH_Colony
                    lLeft = lblH_Colony.Left - 10
                    lTop = lblH_Colony.Top
                Case Budget.BudgetSort.EnvironmentName
                    'lbl = lblH_Envir
                    lLeft = lblH_Envir.Left - 10
                    lTop = lblH_Envir.Top
                Case Budget.BudgetSort.EnvironmentType
                    lLeft = lblH_Type.Left - 8
                    lTop = lblH_Type.Top
                Case Budget.BudgetSort.Expense
                    lLeft = lblH_Expense.Left + 30
                    lTop = lblH_Expense.Top
                Case Budget.BudgetSort.Revenue
                    lLeft = lblH_Revenue.Left + 30
                    lTop = lblH_Revenue.Top
                Case Budget.BudgetSort.TaxRate
                    lLeft = lblH_TaxRate.Left - 10
                    lTop = lblH_TaxRate.Top
                Case Budget.BudgetSort.Control
                    lLeft = lblH_Control.Left - 10
                    lTop = lblH_Control.Top
            End Select
            If lLeft <> Int32.MinValue Then
                Dim rcDest As Rectangle = New Rectangle(lLeft, lTop, 10, 10)
                Dim rcSrc As Rectangle
                If bDescending = True Then
                    rcSrc = grc_UI(elInterfaceRectangle.eDownExpander)
                Else
                    rcSrc = grc_UI(elInterfaceRectangle.eUpExpander)
                End If

                Dim pt As Point = Me.GetAbsolutePosition()
                'rcDest.X += pt.X
                rcDest.Y += pt.Y + 5
                If MyBase.moUILib.oInterfaceTexture Is Nothing = True Then Return
                'MyBase.moUILib.Pen.Begin(SpriteFlags.AlphaBlend)
                MyBase.moUILib.BeginPenSprite(SpriteFlags.AlphaBlend)
                MyBase.moUILib.Pen.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, Point.Empty, 0, rcDest.Location, muSettings.InterfaceBorderColor)
                'MyBase.moUILib.Pen.End()
                MyBase.moUILib.EndPenSprite()
            End If
        End Sub

        Private Sub fraBudgetDetails_OnResize() Handles Me.OnResize
            If lnDiv1 Is Nothing = False Then
                Dim lSize As Int32 = Me.Height - lnDiv1.Top '- 10

                mlDisplayItemCnt = lSize \ ctlBudgetDetailItem.ml_ITEM_HEIGHT
            End If
        End Sub

        Private mlPreviousValue As Int32 = 0
        Private mbIgnoreScrollValueChange As Boolean = False
        Private Sub vscrScroll_ValueChange() Handles vscrScroll.ValueChange
            If mbIgnoreScrollValueChange = True Then Return
            For X As Int32 = 0 To mlItemUB
                If moItems(X) Is Nothing = False Then
                    moItems(X).ClearEditingTaxRate()
                End If
            Next X

            'Ok, determine our change to scroll val
            Dim lItemIdx As Int32 = vscrScroll.Value
            'goUILib.AddNotification("Previous: " & mlPreviousValue.ToString & ", New: " & vscrScroll.Value.ToString, Color.Blue, -1, -1, -1, -1, -1, -1)

            If moList Is Nothing = False AndAlso lItemIdx < mlPreviousValue Then
                'Now, is that item visible
                Dim bDone As Boolean = False
                While bDone = False
                    Dim bFound As Boolean = False

                    With moList(lItemIdx)
                        For X As Int32 = 0 To mlExpandedSystemUB
                            If mlExpandedSystemIDs(X) = .lSystemID Then
                                bFound = True
                                Exit For
                            End If
                        Next X
                        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255 Then bFound = True
                        If bFound = False AndAlso .iEnvirTypeID <> ObjectType.eSolarSystem Then
                            lItemIdx -= 1
                            If lItemIdx < 0 Then
                                mbIgnoreScrollValueChange = True
                                vscrScroll.Value = 0
                                mbIgnoreScrollValueChange = False
                                bDone = True
                            End If
                            Continue While
                        Else
                            mbIgnoreScrollValueChange = True
                            vscrScroll.Value = lItemIdx
                            mbIgnoreScrollValueChange = False
                            bDone = True
                        End If
                    End With

                End While
            End If
            mlPreviousValue = vscrScroll.Value

            Me.IsDirty = True
        End Sub

        Private Sub SortItemClicked(ByVal lNewSortType As Budget.BudgetSort)
            Dim lSortType As Budget.BudgetSort = mlSortType
            If (lSortType And Budget.BudgetSort.DescendingOrderShift) <> 0 Then lSortType = lSortType Xor Budget.BudgetSort.DescendingOrderShift
            If lSortType = lNewSortType Then
                If (mlSortType And Budget.BudgetSort.DescendingOrderShift) <> 0 Then mlSortType = lNewSortType Else mlSortType = lNewSortType Or Budget.BudgetSort.DescendingOrderShift
            Else
                mlSortType = lNewSortType
            End If
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eBudgetWindowSortClick, mlSortType, 0)
            moBudget.ResetLastSortTime()
            moBudget.SortItems(mlSortType)
            SetList(moBudget)
            Me.IsDirty = True
        End Sub

        Private Sub lblH_ControlClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButtons As MouseButtons)
            SortItemClicked(Budget.BudgetSort.Control)
        End Sub
        Private Sub lblH_TypeClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButtons As MouseButtons)
            SortItemClicked(Budget.BudgetSort.EnvironmentType)
        End Sub
        Private Sub lblH_EnvirClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButtons As MouseButtons)
            SortItemClicked(Budget.BudgetSort.EnvironmentName)
        End Sub
        Private Sub lblH_ColonyClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButtons As MouseButtons)
            SortItemClicked(Budget.BudgetSort.ColonyName)
        End Sub
        Private Sub lblH_RevenueClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButtons As MouseButtons)
            SortItemClicked(Budget.BudgetSort.Revenue)
        End Sub
        Private Sub lblH_ExpenseClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButtons As MouseButtons)
            SortItemClicked(Budget.BudgetSort.Expense)
        End Sub
        Private Sub lblH_TaxRateClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButtons As MouseButtons)
            SortItemClicked(Budget.BudgetSort.TaxRate)
        End Sub

        Public Sub SaveScrollPosition()
            mlScrollPosition = vscrScroll.Value
        End Sub
    End Class
End Class
