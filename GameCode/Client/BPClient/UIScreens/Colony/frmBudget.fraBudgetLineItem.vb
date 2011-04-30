Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmBudget

    'Interface created from Interface Builder
    Public Class fraBudgetLineItem
        Inherits UIWindow

        Private lblRevenue As UILabel
        Private txtRevenue As UITextBox
        Private lblExpense As UILabel
		Private txtExpenses As UITextBox 
        Public lblIronCurtain As UILabel
		Private WithEvents btnSetIronCurtain As UIButton

        Private WithEvents btnAbandon As UIButton

		Private mlEnvirID As Int32
        Private miEnvirTypeID As Int16

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraBudgetLineItem initial props
            With Me
                .ControlName = "fraBudgetLineItem"
                .Left = 106
                .Top = 116
                .Width = 535 '780
                .Height = 200
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .BorderLineWidth = 1
                .Caption = "Line Item Specifics"
            End With

            'lblRevenue initial props
            lblRevenue = New UILabel(oUILib)
            With lblRevenue
                .ControlName = "lblRevenue"
                .Left = 10 '20
                .Top = 10
                .Width = 143
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Revenue Sources"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblRevenue, UIControl))

            'txtRevenue initial props
            txtRevenue = New UITextBox(oUILib)
            With txtRevenue
                .ControlName = "txtRevenue"
                .Left = lblRevenue.Left
                .Top = 30
                .Width = 250 '350
                .Height = 130
                .Enabled = True
                .Visible = True
                .Caption = "Source Data Here..."
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(0, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 0
                .BorderColor = muSettings.InterfaceBorderColor
                .MultiLine = True
                .Locked = True
            End With
            Me.AddChild(CType(txtRevenue, UIControl))

            'lblExpense initial props
            lblExpense = New UILabel(oUILib)
            With lblExpense
                .ControlName = "lblExpense"
                .Left = 275 '310
                .Top = 10
                .Width = 143
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Expenses Details"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblExpense, UIControl))

            'txtExpenses initial props
            txtExpenses = New UITextBox(oUILib)
            With txtExpenses
                .ControlName = "txtExpenses"
                .Left = lblExpense.Left
                .Top = 30
                .Width = 250 '350
                .Height = 130
                .Enabled = True
                .Visible = True
                .Caption = "Source Data Here..."
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(0, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 0
                .BorderColor = muSettings.InterfaceBorderColor
                .MultiLine = True
                .Locked = True
            End With
            Me.AddChild(CType(txtExpenses, UIControl))

            'btnSetIronCurtain initial props
            btnSetIronCurtain = New UIButton(oUILib)
            With btnSetIronCurtain
                .ControlName = "btnSetIronCurtain"
                .Left = txtRevenue.Left
                .Top = txtRevenue.Top + txtRevenue.Height + 5
                .Width = 50
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
                .ToolTipText = "Click to set this environment as the off-line invulnerable environment."
            End With
            Me.AddChild(CType(btnSetIronCurtain, UIControl))

            'lblIronCurtain initial props
            lblIronCurtain = New UILabel(oUILib)
            With lblIronCurtain
                .ControlName = "lblIronCurtain"
                .Left = btnSetIronCurtain.Left + btnSetIronCurtain.Width + 5
                .Top = btnSetIronCurtain.Top
                .Width = 143
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Off-Line Invulnerability: Unknown"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Off-line invulnerability begins one hour after logging off." & vbCrLf & _
                 "Only planets can be selected as an off-line invulnerable target." & vbCrLf & _
                 "Units and facilities within this environment will not fire weapons" & vbCrLf & _
                 "and cannot be fired upon while the environment is invulnerable."
            End With
            Me.AddChild(CType(lblIronCurtain, UIControl))

            'btnAbandon initial props
            btnAbandon = New UIButton(oUILib)
            With btnAbandon
                .ControlName = "btnAbandon"
                .Left = txtExpenses.Left + txtExpenses.Width - 140 'txtExpenses.Left + ((txtExpenses.Width \ 2) - 70)
                .Top = btnSetIronCurtain.Top
                .Width = 140
                .Height = 24
                .Enabled = Not gbAliased
                .Visible = True
                .Caption = "Abandon Colony"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
                .ToolTipText = "Click to self-destruct the colony."
            End With
            Me.AddChild(CType(btnAbandon, UIControl))

        End Sub

        Public Sub PopulateData(ByRef oBudget As Budget, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
            Dim oSB As New System.Text.StringBuilder
            Dim sName As String
            Dim sValue As String

            Const lLeftSidePad As Int32 = 21
            Const lRightSidePad As Int32 = 9

            mlEnvirID = lEnvirID
            miEnvirTypeID = iEnvirTypeID

            txtRevenue.Caption = ""
            txtExpenses.Caption = ""

            For X As Int32 = 0 To oBudget.mlItemUB
                If oBudget.muItems(X).lEnvirID = lEnvirID AndAlso oBudget.muItems(X).iEnvirTypeID = iEnvirTypeID Then
                    With oBudget.muItems(X)
                        oSB.AppendLine("COLONY COSTS:")
                        'oSB.AppendLine("-------------")

                        'If frmBudget.bShowHWSupply = True Then
                        '    Dim blHW As Int64 = .GetHomeworldSupplyLineCost(oBudget.lIronCurtainPlanet)
                        '    sName = "Howeworld Supplies:"
                        '    sValue = blHW.ToString("#,##0")
                        '    oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        'End If
                        sName = "Howeworld Supplies:"
                        sValue = .lHWSupplyLineCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))

                        sName = "Population Upkeep:"
                        sValue = .PopUpkeep.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Research Facilities:"
                        sValue = .ResearchCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Factories:"
                        sValue = .FactoryCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Spaceports:"
                        sValue = .SpaceportCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Defenses:"
                        sValue = .TurretCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Excess Storage:"
                        sValue = .ExcessStorage.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Other Facilities:"
                        sValue = .OtherFacCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Unemployment:"
                        sValue = .UnemploymentCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        oSB.AppendLine()
                        sName = "Mining Bids:"
                        sValue = .blMiningBidExpense.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        oSB.AppendLine()

                        oSB.AppendLine("UNIT COSTS:")
                        'oSB.AppendLine("-----------")
                        sName = "Non-Aerial Support:"
                        sValue = .NonAirCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Aerial Support:"
                        sValue = .AirCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Docked Non-Aerial:"
                        sValue = .DockedNonAirCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Docked Aerial:"
                        sValue = .DockedAirCost.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        If .bHasCC = False Then oSB.AppendLine("No Command Center Requires More Support!")
                        If .lCurrentColonyCount > .lPlanetCap Then oSB.AppendLine("Too many colonies on this planet, the population is rioting!")

                        txtExpenses.Caption = oSB.ToString()

                        oSB.Length = 0
                        sName = "Colony Tax Income:"
                        sValue = .TaxIncome.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        sName = "Mining Income:"
                        sValue = .blMiningBidIncome.ToString("#,##0")
                        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))


                        'If .lTradePlayerID Is Nothing = False AndAlso .lTradePlayerID.GetUpperBound(0) > -1 Then
                        '    oSB.AppendLine()
                        '    oSB.AppendLine("TRADE INCOME")
                        '    oSB.AppendLine("------------")

                        '    For Y As Int32 = 0 To .lTradePlayerID.GetUpperBound(0)
                        '        sName = GetCacheObjectValue(.lTradePlayerID(Y), ObjectType.ePlayer)
                        '        sValue = .lTradeValue(Y).ToString("#,##0")
                        '        oSB.AppendLine(sName.PadRight(lLeftSidePad, " "c) & sValue.PadLeft(lRightSidePad, " "c))
                        '    Next Y
                        'End If

                        txtRevenue.Caption = oSB.ToString()

                        If .lColonyID > 0 Then
                            btnAbandon.Enabled = True
                        Else : btnAbandon.Enabled = False
                        End If

                        btnAbandon.Caption = "Abandon Colony"
                    End With

                    Exit For
                End If
            Next X

            If iEnvirTypeID = ObjectType.eSolarSystem OrElse lEnvirID = oBudget.lIronCurtainPlanet OrElse gbAliased = True Then
                btnSetIronCurtain.Enabled = False
            Else
                btnSetIronCurtain.Enabled = True
            End If

        End Sub

        Public Sub fraBudgetLineItem_OnNewFrame() Handles Me.OnNewFrame
            If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oBudget Is Nothing = False Then
                Dim sText As String = "Off-Line Invulnerability: " & GetCacheObjectValue(goCurrentPlayer.oBudget.lIronCurtainPlanet, ObjectType.ePlanet)
                If lblIronCurtain.Caption <> sText Then lblIronCurtain.Caption = sText
            End If
        End Sub

        Private Sub fraBudgetLineItem_OnResize() Handles Me.OnResize
            If txtRevenue Is Nothing = False Then
                txtRevenue.Height = Me.Height - txtRevenue.Top - 35
            End If
            If txtExpenses Is Nothing = False Then
                txtExpenses.Height = Me.Height - txtExpenses.Top - 35
            End If
            If btnSetIronCurtain Is Nothing = False AndAlso txtRevenue Is Nothing = False AndAlso lblIronCurtain Is Nothing = False Then
                btnSetIronCurtain.Top = txtRevenue.Top + txtRevenue.Height + 5
                lblIronCurtain.Top = btnSetIronCurtain.Top + 2

                'lblIronCurtain.Width = txtRevenue.Width - (lblIronCurtain.Left - txtRevenue.Left)
                If btnAbandon Is Nothing = False Then
                    lblIronCurtain.Width = btnAbandon.Left - lblIronCurtain.Left
                End If

            End If
            If btnSetIronCurtain Is Nothing = False AndAlso btnAbandon Is Nothing = False Then
                btnAbandon.Top = btnSetIronCurtain.Top
            End If
        End Sub

        Private Sub btnSetIronCurtain_Click(ByVal sName As String) Handles btnSetIronCurtain.Click
            If miEnvirTypeID <> ObjectType.ePlanet Then Return
            Dim yMsg(10) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetIronCurtain).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlEnvirID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = 1
            MyBase.moUILib.SendMsgToPrimary(yMsg)
        End Sub

        Private Sub btnAbandon_Click(ByVal sName As String) Handles btnAbandon.Click
            If btnAbandon.Caption.ToUpper = "CONFIRM" Then
                Dim yMsg(7) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eAbandonColony).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(mlEnvirID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(miEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            Else
                btnAbandon.Caption = "Confirm"
            End If
        End Sub
    End Class
End Class