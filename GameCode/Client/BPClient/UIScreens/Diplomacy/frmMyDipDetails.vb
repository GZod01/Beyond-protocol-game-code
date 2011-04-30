Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmDiplomacy
    'Interface created from Interface Builder
    Public Class fraMyDipDetails
        Inherits UIWindow

        Private lblTitle As UILabel 
        Private lnDiv1 As UILine
        Private lblScores As UILabel
        Private lblTech As UILabel
        Private fraIcon As UIWindow
        Private lnDiv2 As UILine
        Private lblDiplomacy As UILabel
        Private lblMilitary As UILabel
        Private lblPopulation As UILabel
        Private lblProduction As UILabel
        Private lblWealth As UILabel
        Private lblTotal As UILabel

        Private lblCurrTitle As UILabel
        Private lblNextTitle As UILabel
        Private lblPlanetControl As UILabel
        Private txtPlanetControl As UITextBox

        Private rcBack As Rectangle
        Private rcFore1 As Rectangle
        Private rcFore2 As Rectangle
        Private clrBack As System.Drawing.Color
        Private clrFore1 As System.Drawing.Color
        Private clrFore2 As System.Drawing.Color
        Private mbFirst As Boolean = True

        Private cboCustomTitle As UIComboBox
        Private lblCustomTitle As UILabel
        Private WithEvents btnCustomTitleList As UIButton

        Private lnDiv3 As UILine

        Private WithEvents btnFactions As UIButton
        Private WithEvents btnSetCustomTitle As UIButton
        Private WithEvents btnUpdateIcon As UIButton

        Private mlLastUpdate As Int32

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'frmMyDipDetails initial props
            With Me
                .ControlName = "fraMyDipDetails"
                .Left = 495
                .Top = 1
                .Width = 500 '240
                .Height = 280
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .BorderLineWidth = 1
            End With

            'lblTitle initial props
            lblTitle = New UILabel(oUILib)
            With lblTitle
                .ControlName = "lblTitle"
                .Left = 5
                .Top = 3
                .Width = 196
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Personal Diplomatic Details"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTitle, UIControl))

            'lnDiv1 initial props
            lnDiv1 = New UILine(oUILib)
            With lnDiv1
                .ControlName = "lnDiv1"
                .Left = 1 '0
                .Top = 24
                .Width = 499 '240
                .Height = 0
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(lnDiv1, UIControl))

            'lblScores initial props
            lblScores = New UILabel(oUILib)
            With lblScores
                .ControlName = "lblScores"
                .Left = 80
                .Top = 25
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Scores"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblScores, UIControl))

            'lblTech initial props
            lblTech = New UILabel(oUILib)
            With lblTech
                .ControlName = "lblTech"
                .Left = 85
                .Top = 40
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Technology:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTech, UIControl))

            'fraIcon initial props
            fraIcon = New UIWindow(oUILib)
            With fraIcon
                .ControlName = "fraIcon"
                .Left = 5
                .Top = 30
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
            Me.AddChild(CType(fraIcon, UIControl))

            'lnDiv2 initial props
            lnDiv2 = New UILine(oUILib)
            With lnDiv2
                .ControlName = "lnDiv2"
                .Left = 0
                .Top = 150
                .Width = 240
                .Height = 0
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(lnDiv2, UIControl))

            'lnDiv3 initial props
            lnDiv3 = New UILine(oUILib)
            With lnDiv3
                .ControlName = "lnDiv3"
                .Left = 240
                .Top = 24
                .Width = 0
                .Height = 256
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(lnDiv3, UIControl))

            'lblDiplomacy initial props
            lblDiplomacy = New UILabel(oUILib)
            With lblDiplomacy
                .ControlName = "lblDiplomacy"
                .Left = 85
                .Top = 55
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Diplomacy:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblDiplomacy, UIControl))

            'lblMilitary initial props
            lblMilitary = New UILabel(oUILib)
            With lblMilitary
                .ControlName = "lblMilitary"
                .Left = 85
                .Top = 70
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Military:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMilitary, UIControl))

            'lblPopulation initial props
            lblPopulation = New UILabel(oUILib)
            With lblPopulation
                .ControlName = "lblPopulation"
                .Left = 85
                .Top = 85
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Population:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPopulation, UIControl))

            'lblProduction initial props
            lblProduction = New UILabel(oUILib)
            With lblProduction
                .ControlName = "lblProduction"
                .Left = 85
                .Top = 100
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Production:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblProduction, UIControl))

            'lblWealth initial props
            lblWealth = New UILabel(oUILib)
            With lblWealth
                .ControlName = "lblWealth"
                .Left = 85
                .Top = 115
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Wealth:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblWealth, UIControl))

            'lblTotal initial props
            lblTotal = New UILabel(oUILib)
            With lblTotal
                .ControlName = "lblTotal"
                .Left = 85
                .Top = 130
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Total:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTotal, UIControl))

            'lblCurrTitle initial props
            lblCurrTitle = New UILabel(oUILib)
            With lblCurrTitle
                .ControlName = "lblCurrTitle"
                .Left = 5
                .Top = 155
                .Width = 230
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Title: "
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCurrTitle, UIControl))

            'lblNextTitle initial props
            lblNextTitle = New UILabel(oUILib)
            With lblNextTitle
                .ControlName = "lblNextTitle"
                .Left = 5
                .Top = 173
                .Width = 230
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Next Title: "
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblNextTitle, UIControl))

            'lblPlanetControl initial props
            lblPlanetControl = New UILabel(oUILib)
            With lblPlanetControl
                .ControlName = "lblPlanetControl"
                .Left = 5
                .Top = 190
                .Width = 230
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Planet Control"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPlanetControl, UIControl))

            'txtPlanetControl initial props
            txtPlanetControl = New UITextBox(oUILib)
            With txtPlanetControl
                .ControlName = "txtPlanetControl"
                .Left = 5
                .Top = 210
                .Width = 230
                .Height = 65
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
            End With
			Me.AddChild(CType(txtPlanetControl, UIControl))

            btnUpdateIcon = New UIButton(oUILib)
            With btnUpdateIcon
                .ControlName = "btnUpdateIcon"
                .Left = fraIcon.Left
                .Top = fraIcon.Top + fraIcon.Height + 5
                .Width = fraIcon.Width
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Edit Icon"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnUpdateIcon, UIControl))

            'lblCustomTitle initial props
            lblCustomTitle = New UILabel(oUILib)
            With lblCustomTitle
                .ControlName = "lblCustomTitle"
                .Left = 250
                .Top = 35
                .Width = 80
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Custom Title:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCustomTitle, UIControl))

            btnSetCustomTitle = New UIButton(oUILib)
            With btnSetCustomTitle
                .ControlName = "btnSetCustomTitle"
                .Left = 335
                .Top = lblCustomTitle.Top + 25
                .Width = 75 '50
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetCustomTitle, UIControl))

            'btnCustomTitleList
            btnCustomTitleList = New UIButton(oUILib)
            With btnCustomTitleList
                .ControlName = "btnCustomTitleList"
                .Left = btnSetCustomTitle.Left + btnSetCustomTitle.Width + 5
                .Top = btnSetCustomTitle.Top
                .Width = 75 'fraIcon.Width
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "List"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnCustomTitleList, UIControl))

            btnFactions = New UIButton(oUILib)
            With btnFactions
                .ControlName = "btnFactions"
                .Left = btnSetCustomTitle.Left
                .Top = btnSetCustomTitle.Top + btnSetCustomTitle.Height + 5
                .Width = (btnCustomTitleList.Width + btnCustomTitleList.Left) - .Left
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Manage Factions"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnFactions, UIControl))

            cboCustomTitle = New UIComboBox(oUILib)
            With cboCustomTitle
                .ControlName = "cboCustomTitle"
                .Left = lblCustomTitle.Left + lblCustomTitle.Width + 5
                .Top = lblCustomTitle.Top
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .mbAcceptReprocessEvents = True
                .ToolTipText = "Assigns a separate title that other players will see instead of your senate title."
            End With
            Me.AddChild(CType(cboCustomTitle, UIControl))


        End Sub

		Public Sub frmMyDipDetails_OnNewFrame() Handles Me.OnNewFrame
			If goCurrentEnvir Is Nothing Then Return
			If glCurrentCycle - mlLastUpdate > 15 Then
				With goCurrentPlayer
					Dim sTemp As String

					sTemp = "Technology: " & .lTechnologyScore.ToString("#,##0")
					If lblTech.Caption <> sTemp Then lblTech.Caption = sTemp
					sTemp = "Diplomacy: " & .lDiplomacyScore.ToString("#,##0")
					If lblDiplomacy.Caption <> sTemp Then lblDiplomacy.Caption = sTemp
					sTemp = "Military: " & .lMilitaryScore.ToString("#,##0")
					If lblMilitary.Caption <> sTemp Then lblMilitary.Caption = sTemp
					sTemp = "Population: " & .lPopulationScore.ToString("#,##0")
					If lblPopulation.Caption <> sTemp Then lblPopulation.Caption = sTemp
					sTemp = "Production: " & .lProductionScore.ToString("#,##0")
					If lblProduction.Caption <> sTemp Then lblProduction.Caption = sTemp
					sTemp = "Wealth: " & .lWealthScore.ToString("#,##0")
					If lblWealth.Caption <> sTemp Then lblWealth.Caption = sTemp
					sTemp = "Total: " & .lTotalScore.ToString("#,##0")
					If lblTotal.Caption <> sTemp Then lblTotal.Caption = sTemp
                    Dim yNextTitle As Byte = .yPlayerTitle
                    If (yNextTitle And Player.PlayerRank.ExRankShift) <> 0 Then
                        sTemp = "Title: Ex-" & Player.GetPlayerTitle(.yPlayerTitle, .bIsMale)
                        'yNextTitle = (yNextTitle Xor Player.PlayerRank.ExRankShift)
                        'Tried showing -1 title.  Overall confusing.  Since we now show Ex- in the
                        ' main title, lets juts blank out Next:
                        yNextTitle = Byte.MaxValue
                    Else
                        sTemp = "Title: " & Player.GetPlayerTitle(.yPlayerTitle, .bIsMale)
                        yNextTitle = CByte(yNextTitle + 1)
                    End If
                    If lblCurrTitle.Caption <> sTemp Then lblCurrTitle.Caption = sTemp


                    sTemp = Player.GetPlayerTitle(yNextTitle, .bIsMale)
                    If sTemp <> "" Then sTemp = "Next Title: " & sTemp
                    If lblNextTitle.Caption <> sTemp Then lblNextTitle.Caption = sTemp

                    sTemp = GetNextTitleRequirements(yNextTitle)
                    If lblNextTitle.ToolTipText <> sTemp Then lblNextTitle.ToolTipText = sTemp

                    sTemp = ""
                    If .lControlPlanets Is Nothing = False Then
                        Try
                            'Ok, sort the list by name
                            Dim lSorted() As Int32 = Nothing
                            Dim lSortedUB As Int32 = -1
                            For X As Int32 = 0 To .lControlPlanets.GetUpperBound(0)
                                Dim lIdx As Int32 = -1

                                Dim sName As String = GetCacheObjectValue(.lControlPlanets(X), ObjectType.ePlanet)

                                For Y As Int32 = 0 To lSortedUB
                                    Dim sOtherName As String = GetCacheObjectValue(.lControlPlanets(lSorted(Y)), ObjectType.ePlanet)
                                    If sOtherName > sName Then
                                        lIdx = Y
                                        Exit For
                                    End If
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
                            Next X

                            For X As Int32 = 0 To .lControlPlanets.GetUpperBound(0)
                                If sTemp.Length = 0 Then
                                    sTemp = GetCacheObjectValue(.lControlPlanets(lSorted(X)), ObjectType.ePlanet) & ": " & .yControlPlanetAmt(lSorted(X)) & "%"
                                Else
                                    sTemp &= vbCrLf & GetCacheObjectValue(.lControlPlanets(lSorted(X)), ObjectType.ePlanet) & ": " & .yControlPlanetAmt(lSorted(X)) & "%"
                                End If
                            Next X
                        Catch
                            'do nothing
                        End Try
                    End If
                    If txtPlanetControl.Caption <> sTemp Then txtPlanetControl.Caption = sTemp

                    Dim sItems(31) As String
                    Dim lItemVal(31) As Int32
                    Dim lItemIdx As Int32 = -1
                    If (.lCustomTitlePermission And elCustomRankPermissions.Arbiter) <> 0 Then
                        lItemIdx += 1
                        If .bIsMale = True Then sItems(lItemIdx) = "Arbiter" Else sItems(lItemIdx) = "Arbitress"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Arbiter
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Broker) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Broker"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Broker
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Chancellor) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Chancellor"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Chancellor
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.ChiefBroker) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Chief Broker"
                        lItemVal(lItemIdx) = elCustomRankPermissions.ChiefBroker
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.ChiefScientist) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Chief Scientist"
                        lItemVal(lItemIdx) = elCustomRankPermissions.ChiefScientist
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.CommerceCzar) <> 0 Then
                        lItemIdx += 1
                        If .bIsMale = True Then sItems(lItemIdx) = "Commerce Czar" Else sItems(lItemIdx) = "Commerce Czarina"
                        lItemVal(lItemIdx) = elCustomRankPermissions.CommerceCzar
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Counselor) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Counselor"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Counselor
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Diplomat) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Diplomat"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Diplomat
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Explorer) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Explorer"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Explorer
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.HighSenator) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "High Senator"
                        lItemVal(lItemIdx) = elCustomRankPermissions.HighSenator
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Inquisitor) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Inquisitor"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Inquisitor
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.MasterMerchant) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Master Merchant"
                        lItemVal(lItemIdx) = elCustomRankPermissions.MasterMerchant
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.MasterScientist) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Master Scientist"
                        lItemVal(lItemIdx) = elCustomRankPermissions.MasterScientist
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Merchant) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Merchant"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Merchant
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Preeminence) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Preeminence"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Preeminence
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Scientist) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Scientist"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Scientist
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Senator) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Senator"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Senator
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.SupremeChancellor) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Supreme Chancellor"
                        lItemVal(lItemIdx) = elCustomRankPermissions.SupremeChancellor
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.TradeLord) <> 0 Then
                        lItemIdx += 1
                        If .bIsMale = True Then sItems(lItemIdx) = "Trade Lord" Else sItems(lItemIdx) = "Trade Lady"
                        lItemVal(lItemIdx) = elCustomRankPermissions.TradeLord
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Trader) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Trader"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Trader
                    End If
                    If (.lCustomTitlePermission And elCustomRankPermissions.Transcendent) <> 0 Then
                        lItemIdx += 1
                        sItems(lItemIdx) = "Transcendent"
                        lItemVal(lItemIdx) = elCustomRankPermissions.Transcendent
                    End If

                    If cboCustomTitle.ListCount <> lItemIdx + 2 Then '1 changed to 2.  Var base is 0 so +1, then Rank +1.
                        cboCustomTitle.Clear()
                        cboCustomTitle.AddItem(Player.GetPlayerTitle(.yPlayerTitle, .bIsMale))
                        cboCustomTitle.ItemData(cboCustomTitle.NewIndex) = -1
                        For X As Int32 = 0 To lItemIdx
                            cboCustomTitle.AddItem(sItems(X))
                            cboCustomTitle.ItemData(cboCustomTitle.NewIndex) = lItemVal(X)
                        Next X
                    End If

                End With

                mlLastUpdate = glCurrentCycle
			End If
		End Sub

		Private Sub fraMyDipDetails_OnRenderEnd() Handles Me.OnRenderEnd

			If mbFirst = True Then
				SetIconDetails()
				mbFirst = False
			End If

            If ctlDiplomacy.moSprite Is Nothing OrElse ctlDiplomacy.moSprite.Disposed = True Then
                Device.IsUsingEventHandlers = False
                ctlDiplomacy.moSprite = New Sprite(MyBase.moUILib.oDevice)
                Device.IsUsingEventHandlers = True
            End If
			If ctlDiplomacy.moIconBack Is Nothing OrElse ctlDiplomacy.moIconFore Is Nothing OrElse ctlDiplomacy.moIconBack.Disposed = True OrElse ctlDiplomacy.moIconFore.Disposed = True Then Return
			ctlDiplomacy.moSprite.Begin(SpriteFlags.AlphaBlend)
			Dim rcDest As Rectangle = New Rectangle(0, 0, 64, 64)
			Dim ptDest As Point
			ptDest.X = Me.Left + fraIcon.Left + 3
			ptDest.Y = Me.Top + fraIcon.Top + 3

			ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moIconBack, rcBack, rcDest, ptDest, clrBack)
			ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moIconFore, rcFore1, rcDest, ptDest, clrFore1)
			ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moIconFore, rcFore2, rcDest, ptDest, clrFore2)

			ctlDiplomacy.moSprite.End()
		End Sub

        Private Sub SetIconDetails()
            Dim yBackImg As Byte
            Dim yBackClr As Byte
            Dim yFore1Img As Byte
            Dim yFore1Clr As Byte
            Dim yFore2Img As Byte
            Dim yFore2Clr As Byte

            If goCurrentPlayer Is Nothing Then Return

            PlayerIconManager.FillIconValues(goCurrentPlayer.lPlayerIcon, yBackImg, yBackClr, yFore1Img, yFore1Clr, yFore2Img, yFore2Clr)

            rcBack = PlayerIconManager.ReturnImageRectangle(yBackImg, True)
            rcFore1 = PlayerIconManager.ReturnImageRectangle(yFore1Img, False)
            rcFore2 = PlayerIconManager.ReturnImageRectangle(yFore2Img, False)

            clrBack = PlayerIconManager.GetColorValue(yBackClr)
            clrFore1 = PlayerIconManager.GetColorValue(yFore1Clr)
            clrFore2 = PlayerIconManager.GetColorValue(yFore2Clr)
        End Sub

        Private Function GetNextTitleRequirements(ByVal yNextTitle As Byte) As String
            If (yNextTitle And Player.PlayerRank.ExRankShift) <> 0 Then
                yNextTitle = yNextTitle Xor Player.PlayerRank.ExRankShift
                If yNextTitle > 0 Then yNextTitle = CByte(yNextTitle - 1)
            End If
            Select Case yNextTitle
                Case Player.PlayerRank.Governor     '1
                    Return "At least 200,000 Total Population"
                Case Player.PlayerRank.Overseer     '2
                    Return "At least 500,000 Total Population" '& vbCrLf & "At least 1 space station colony"
                Case Player.PlayerRank.Duke         '3
                    Return "3/4 Population on a single planet (majority owner)" & vbCrLf & "At least 2 planet-based colonies"
                Case Player.PlayerRank.Baron        '4
                    Return "3/4 Population on at least 2 planets (majority owner)" & vbCrLf & "At least 3 planet-based colonies"
                Case Player.PlayerRank.King         '5
                    Return "3/4 Population on more than half" & vbCrLf & "  of the planets in the same system"
                Case Player.PlayerRank.Emperor      '6
                    Return "3/4 Population on all planets in the same system" & vbCrLf & "3/4 Population on a planet outside of the owned system"
                Case Else
                    Return ""
            End Select
        End Function

		Private Sub btnFactions_Click(ByVal sName As String) Handles btnFactions.Click
			Dim ofrm As frmFaction = CType(goUILib.GetWindow("frmFaction"), frmFaction)
			If ofrm Is Nothing Then
				ofrm = New frmFaction(goUILib)
				ofrm.Visible = True
			ElseIf ofrm.Visible = False Then
				ofrm.Visible = True
			Else : goUILib.RemoveWindow("frmFaction")
			End If
		End Sub

        Private Sub btnSetCustomTitle_Click(ByVal sName As String) Handles btnSetCustomTitle.Click
            If cboCustomTitle.ListIndex > -1 Then
                Dim yMsg(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerCustomTitle).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(cboCustomTitle.ItemData(cboCustomTitle.ListIndex)).CopyTo(yMsg, 2)
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            Else
                MyBase.moUILib.AddNotification("Select an item in the list first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End Sub

        Private Sub btnCustomTitleList_Click(ByVal sName As String) Handles btnCustomTitleList.Click
            Dim ofrm As frmCustomTitle = CType(goUILib.GetWindow("frmCustomTitle"), frmCustomTitle)
            If ofrm Is Nothing Then ofrm = New frmCustomTitle(goUILib)
            ofrm.Visible = True
        End Sub

        Private Sub btnUpdateIcon_Click(ByVal sName As String) Handles btnUpdateIcon.Click
            Dim ofrm As frmIconChooser = CType(goUILib.GetWindow("frmIconChooser"), frmIconChooser)
            If ofrm Is Nothing Then ofrm = New frmIconChooser(goUILib, True)
            ofrm.Visible = True
        End Sub
    End Class
End Class