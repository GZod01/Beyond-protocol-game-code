Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmDiplomacy

    'Interface created from Interface Builder
    Public Class fraPlayerDetails
        Inherits UIWindow

        Private lblThreatscores As UILabel
        Private lblTech As UILabel
        Private lblDiplomacy As UILabel
        Private lblMilitary As UILabel
        Private lblPopulace As UILabel
        Private lblProduction As UILabel
        Private lblWealth As UILabel
        Private lblTechScore As UILabel
        Private lblDiplomacyScore As UILabel
        Private lblMilitaryScore As UILabel
        Private lblPopScore As UILabel
        Private lblProdScore As UILabel
        Private lblWealthScore As UILabel
        Private lblTotalScore As UILabel
        Private lblNextChng As UILabel

        Private lnDiv1 As UILine
        Private lblRelationship As UILabel
        Private lblTowardsYou As UILabel
        Private lblTheirScore As UILabel
        Private lblTowardsThem As UILabel
        Private lblOurScore As UILabel
        Private lblTargetScore As UILabel
        Private lblTargetScoreVal As UILabel
        Private WithEvents txtScore As UITextBox
        Private WithEvents hscrRel As UIScrollBar
        Private WithEvents btnSet As UIButton
        Private WithEvents btnReset As UIButton
        Private lnDiv2 As UILine

        Private mclrVals() As System.Drawing.Color

        Private mbLoading As Boolean = True

        Private moRel As PlayerRel = Nothing

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraPlayerDetails initial props
            With Me
                .ControlName = "fraPlayerDetails"
                .Left = 104
                .Top = 182
                .Width = 220 ' 490 '735
                .Height = 280
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 1
                .Caption = "Player Details"
                .Moveable = False
                .mbAcceptReprocessEvents = True
            End With

            'lblThreatscores initial props
            lblThreatscores = New UILabel(oUILib)
            With lblThreatscores
                .ControlName = "lblThreatscores"
                .Left = 5
                .Top = 7
                .Width = 100
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Threat Scores"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblThreatscores, UIControl))

            'lblTech initial props
            lblTech = New UILabel(oUILib)
            With lblTech
                .ControlName = "lblTech"
                .Left = 15
                .Top = 22 ' 25
                .Width = 83
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

            'lblDiplomacy initial props
            lblDiplomacy = New UILabel(oUILib)
            With lblDiplomacy
                .ControlName = "lblDiplomacy"
                .Left = 15
                .Top = 39 ' 45
                .Width = 83
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
                .Left = 15
                .Top = 56 '65
                .Width = 83
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

            'lblPopulace initial props
            lblPopulace = New UILabel(oUILib)
            With lblPopulace
                .ControlName = "lblPopulace"
                .Left = 15
                .Top = 73 '85
                .Width = 83
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Population:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPopulace, UIControl))

            'lblProduction initial props
            lblProduction = New UILabel(oUILib)
            With lblProduction
                .ControlName = "lblProduction"
                .Left = 15
                .Top = 90 '105
                .Width = 83
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
                .Left = 15
                .Top = 107 '125
                .Width = 83
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

            'lblTotalScore initial props
            lblTotalScore = New UILabel(oUILib)
            With lblTotalScore
                .ControlName = "lblTotalScore"
                .Left = 15
                .Top = 124 '145
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Total Score: "
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTotalScore, UIControl))

            'lblTechScore initial props
            lblTechScore = New UILabel(oUILib)
            With lblTechScore
                .ControlName = "lblTechScore"
                .Left = 100
                .Top = lblTech.Top + 2 '27
                .Width = 100
                .Height = 13
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .FontFormat = CType(5, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTechScore, UIControl))

            'lblDiplomacyScore initial props
            lblDiplomacyScore = New UILabel(oUILib)
            With lblDiplomacyScore
                .ControlName = "lblDiplomacyScore"
                .Left = 100
                .Top = lblDiplomacy.Top + 2 '47
                .Width = 100
                .Height = 13
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .FontFormat = CType(5, DrawTextFormat)
            End With
            Me.AddChild(CType(lblDiplomacyScore, UIControl))

            'lblMilitaryScore initial props
            lblMilitaryScore = New UILabel(oUILib)
            With lblMilitaryScore
                .ControlName = "lblMilitaryScore"
                .Left = 100
                .Top = lblMilitary.Top + 2 '67
                .Width = 100
                .Height = 13
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .FontFormat = CType(5, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMilitaryScore, UIControl))

            'lblPopScore initial props
            lblPopScore = New UILabel(oUILib)
            With lblPopScore
                .ControlName = "lblPopScore"
                .Left = 100
                .Top = lblPopulace.Top + 2 '87
                .Width = 100
                .Height = 13
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .FontFormat = CType(5, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPopScore, UIControl))

            'lblProdScore initial props
            lblProdScore = New UILabel(oUILib)
            With lblProdScore
                .ControlName = "lblProdScore"
                .Left = 100
                .Top = lblProduction.Top + 2 '107
                .Width = 100
                .Height = 13
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .FontFormat = CType(5, DrawTextFormat)
            End With
            Me.AddChild(CType(lblProdScore, UIControl))

            'lblWealthScore initial props
            lblWealthScore = New UILabel(oUILib)
            With lblWealthScore
                .ControlName = "lblWealthScore"
                .Left = 100
                .Top = lblWealth.Top + 2 '127
                .Width = 100
                .Height = 13
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .FontFormat = CType(5, DrawTextFormat)
            End With
            Me.AddChild(CType(lblWealthScore, UIControl))

            ''lnDiv1 initial props
            'lnDiv1 = New UILine(oUILib)
            'With lnDiv1
            '    .ControlName = "lnDiv1"
            '    .Left = 220
            '    .Top = 0
            '    .Width = 0
            '    .Height = 280
            '    .Enabled = True
            '    .Visible = True
            '    .BorderColor = muSettings.InterfaceBorderColor
            'End With
            'Me.AddChild(CType(lnDiv1, UIControl))

            'lblRelationship initial props
            lblRelationship = New UILabel(oUILib)
            With lblRelationship
                .ControlName = "lblRelationship"
                .Left = 5 ' 225
                .Top = 147 '170 ' 5
                .Width = 164
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Relationship"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblRelationship, UIControl))

            'lblNextChng initial props
            lblNextChng = New UILabel(oUILib)
            With lblNextChng
                .ControlName = "lblNextChng"
                .Left = 165
                .Top = 147 '170 ' 5
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "00:00"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Right Or DrawTextFormat.VerticalCenter
                .ToolTipText = "Time remaining until you are able to change the relationship down another 5 points."
            End With
            Me.AddChild(CType(lblNextChng, UIControl))

            'lblTowardsYou initial props
            lblTowardsYou = New UILabel(oUILib)
            With lblTowardsYou
                .ControlName = "lblTowardsYou"
                .Left = 7 '235
                .Top = 165 '188 '25
                .Width = 85
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Towards You:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTowardsYou, UIControl))

            'lblTheirScore initial props
            lblTheirScore = New UILabel(oUILib)
            With lblTheirScore
                .ControlName = "lblTheirScore"
                .Left = 105 '325
                .Top = 165 ' 188 '25
                .Width = 160
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTheirScore, UIControl))

            'lblTowardsThem initial props
            lblTowardsThem = New UILabel(oUILib)
            With lblTowardsThem
                .ControlName = "lblTowardsThem"
                .Left = 7 '235
                .Top = 183 '206 '50
                .Width = 99
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Towards Them:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTowardsThem, UIControl))

            'lblOurScore initial props
            lblOurScore = New UILabel(oUILib)
            With lblOurScore
                .ControlName = "lblOurScore"
                .Left = 105 '335
                .Top = 183 '206 '50
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblOurScore, UIControl))

            'lblTargetScore
            lblTargetScore = New UILabel(oUILib)
            With lblTargetScore
                .ControlName = "lblTargetScore"
                .Left = 7 '235
                .Top = 201 '206 '50
                .Width = 99
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Target Score: "
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Relationships can only be dropped by 5 points per 4 hour period." & vbCrLf & _
                               "The exception to this rule is if the other player is already at war with you"
            End With
            Me.AddChild(CType(lblTargetScore, UIControl))

            'lblTargetScoreVal initial props
            lblTargetScoreVal = New UILabel(oUILib)
            With lblTargetScoreVal
                .ControlName = "lblTargetScoreVal"
                .Left = 105
                .Top = 201
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = lblTargetScore.ToolTipText
            End With
            Me.AddChild(CType(lblTargetScoreVal, UIControl))

            'hscrRel initial props
            hscrRel = New UIScrollBar(oUILib, False)
            With hscrRel
                .ControlName = "hscrRel"
                .Left = 5 '235
                .Top = 227 '75
                .Width = 210 '250
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 60
                .MaxValue = 255
                .MinValue = 1
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                '.mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(hscrRel, UIControl))

            'btnSet initial props
            btnSet = New UIButton(oUILib)
            With btnSet
                .ControlName = "btnSet"
                .Left = 5 '265
                .Top = 250 '100
                .Width = 80
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
            Me.AddChild(CType(btnSet, UIControl))

            'btnReset initial props
            btnReset = New UIButton(oUILib)
            With btnReset
                .ControlName = "btnReset"
                .Left = 135 '375
                .Top = 250 '100
                .Width = 80
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Reset"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnReset, UIControl))

            'txtScore initial props
            txtScore = New UITextBox(oUILib)
            With txtScore
                .ControlName = "txtScore"
                .Left = btnSet.Left + btnSet.Width + 5
                .Top = btnSet.Top
                .Width = 40
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = hscrRel.Value.ToString
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 3
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
            End With
            Me.AddChild(CType(txtScore, UIControl))

            'lnDiv2 initial props
            lnDiv2 = New UILine(oUILib)
            With lnDiv2
                .ControlName = "lnDiv2"
                .Left = 1 '220
                .Top = 142 ' 165 ' 130
                .Width = 220 '270
                .Height = 0
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(lnDiv2, UIControl))

            ReDim mclrVals(9)
            mclrVals(0) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            mclrVals(1) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            mclrVals(2) = System.Drawing.Color.FromArgb(255, 255, 128, 0)
            mclrVals(3) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
            mclrVals(4) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            mclrVals(5) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            mclrVals(6) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
            mclrVals(7) = System.Drawing.Color.FromArgb(255, 255, 128, 0)
            mclrVals(8) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            mclrVals(9) = System.Drawing.Color.FromArgb(255, 255, 0, 0)

            mbLoading = False
        End Sub

        'Takes scores in... these scores are the relationship score value, 0 means they are crap compared to you. 9 means they are gods compared to you
        ' 255 means it is unknown
        Public Sub SetData(ByRef oRel As PlayerRel)
            Dim yTech As Byte = 255
            Dim yDiplomacy As Byte = 255
            Dim yMilitary As Byte = 255
            Dim yPopulace As Byte = 255
            Dim yProduction As Byte = 255
            Dim yWealth As Byte = 255
			Dim lTotalScore As Int32 = 0

			Dim lTechUpdate As Int32 = Int32.MinValue
			Dim lDiplomacyUpdate As Int32 = Int32.MinValue
			Dim lMilitaryUpdate As Int32 = Int32.MinValue
			Dim lPopulaceUpdate As Int32 = Int32.MinValue
			Dim lProductionUpdate As Int32 = Int32.MinValue
			Dim lWealthUpdate As Int32 = Int32.MinValue

			Dim lTechScore As Int32 = Int32.MinValue
			Dim lDiplomacyScore As Int32 = Int32.MinValue
			Dim lMilitaryScore As Int32 = Int32.MinValue
			Dim lPopulaceScore As Int32 = Int32.MinValue
			Dim lProductionScore As Int32 = Int32.MinValue
			Dim lWealthScore As Int32 = Int32.MinValue

            moRel = oRel

            If oRel Is Nothing = False Then
                Dim oPlayerIntel As PlayerIntel = Nothing
                If oRel.lPlayerRegards = glPlayerID Then
                    For X As Int32 = 0 To glPlayerIntelUB
                        If glPlayerIntelIdx(X) = oRel.lThisPlayer Then
                            oPlayerIntel = goPlayerIntel(X)
                            Exit For
                        End If
                    Next X
                Else
                    For X As Int32 = 0 To glPlayerIntelUB
                        If glPlayerIntelIdx(X) = oRel.lPlayerRegards Then
                            oPlayerIntel = goPlayerIntel(X)
                            Exit For
                        End If
                    Next X
                End If
                If oPlayerIntel Is Nothing = False Then
					yTech = GetScoreRatioValue(goCurrentPlayer.lTechnologyScore, oPlayerIntel.lTechnologyScore)
					lTechUpdate = oPlayerIntel.lTechnologyUpdate
					lTechScore = oPlayerIntel.lTechnologyScore
					yDiplomacy = GetScoreRatioValue(goCurrentPlayer.lDiplomacyScore, oPlayerIntel.lDiplomacyScore)
					lDiplomacyUpdate = oPlayerIntel.lDiplomacyUpdate
					lDiplomacyScore = oPlayerIntel.lDiplomacyScore
					yMilitary = GetScoreRatioValue(goCurrentPlayer.lMilitaryScore, oPlayerIntel.lMilitaryScore)
					lMilitaryUpdate = oPlayerIntel.lMilitaryUpdate
					lMilitaryScore = oPlayerIntel.lMilitaryScore
					yPopulace = GetScoreRatioValue(goCurrentPlayer.lPopulationScore, oPlayerIntel.lPopulationScore)
					lPopulaceUpdate = oPlayerIntel.lPopulationUpdate
					lPopulaceScore = oPlayerIntel.lPopulationScore
					yProduction = GetScoreRatioValue(goCurrentPlayer.lProductionScore, oPlayerIntel.lProductionScore)
					lProductionUpdate = oPlayerIntel.lProductionUpdate
					lProductionScore = oPlayerIntel.lProductionScore
					yWealth = GetScoreRatioValue(goCurrentPlayer.lWealthScore, oPlayerIntel.lWealthScore)
					lWealthUpdate = oPlayerIntel.lWealthUpdate
					lWealthScore = oPlayerIntel.lWealthScore
                    lTotalScore = oPlayerIntel.lTotalScore

                    lblTheirScore.Caption = GetRelValText(oPlayerIntel.yRegardsCurrentPlayer)
                    lblTheirScore.ForeColor = GetRelValColor(oPlayerIntel.yRegardsCurrentPlayer)
                End If

                lblOurScore.Caption = GetRelValText(oRel.WithThisScore)
                lblOurScore.ForeColor = GetRelValColor(oRel.WithThisScore)

                If oRel.lNextUpdateCycle = -1 Then oRel.TargetScore = oRel.WithThisScore
                hscrRel.Value = oRel.TargetScore
                hscrRel_ValueChange()
            End If

            If yTech = 255 Then
                lblTechScore.DrawBackImage = False
				lblTechScore.Caption = "Unknown"
				lblTechScore.ToolTipText = ""
            Else
                lblTechScore.ControlImageRect = Rectangle.FromLTRB(192, 157 + (yTech * 9), 256, 157 + ((yTech + 1) * 9))
                lblTechScore.BackImageColor = mclrVals(yTech)
				lblTechScore.DrawBackImage = True
				lblTechScore.Caption = ""
                lblTechScore.ToolTipText = "Tech Score: " & lTechScore.ToString("#,##0") & vbCrLf & "Last Updated: " & GetDurationFromSeconds(CInt(Now.Subtract(GetDateFromNumber(lTechUpdate)).TotalSeconds), True)
				lblTechScore.BackImageColor = GetValueColor(yTech)
			End If
            If yDiplomacy = 255 Then
                lblDiplomacyScore.DrawBackImage = False
				lblDiplomacyScore.Caption = "Unknown"
				lblDiplomacyScore.ToolTipText = ""
            Else
				lblDiplomacyScore.ControlImageRect = Rectangle.FromLTRB(192, 157 + (yDiplomacy * 9), 256, 157 + ((yDiplomacy + 1) * 9))
                lblDiplomacyScore.BackImageColor = mclrVals(yDiplomacy)
                lblDiplomacyScore.DrawBackImage = True
				lblDiplomacyScore.Caption = ""
                lblDiplomacyScore.ToolTipText = "Diplomacy Score: " & lDiplomacyScore.ToString("#,##0") & vbCrLf & "Last Updated: " & GetDurationFromSeconds(CInt(Now.Subtract(GetDateFromNumber(lDiplomacyUpdate)).TotalSeconds), True)
				lblDiplomacyScore.BackImageColor = GetValueColor(yDiplomacy)
            End If
            If yMilitary = 255 Then
                lblMilitaryScore.DrawBackImage = False
				lblMilitaryScore.Caption = "Unknown"
				lblMilitaryScore.ToolTipText = ""
            Else
				lblMilitaryScore.ControlImageRect = Rectangle.FromLTRB(192, 157 + (yMilitary * 9), 256, 157 + ((yMilitary + 1) * 9))
                lblMilitaryScore.BackImageColor = mclrVals(yMilitary)
                lblMilitaryScore.DrawBackImage = True
				lblMilitaryScore.Caption = ""
                lblMilitaryScore.ToolTipText = "Military Score: " & lMilitaryScore.ToString("#,##0") & vbCrLf & "Last Updated: " & GetDurationFromSeconds(CInt(Now.Subtract(GetDateFromNumber(lMilitaryUpdate)).TotalSeconds), True)
				lblMilitaryScore.BackImageColor = GetValueColor(yMilitary)
            End If
            If yPopulace = 255 Then
                lblPopScore.DrawBackImage = False
				lblPopScore.Caption = "Unknown"
				lblPopScore.ToolTipText = ""
            Else
				lblPopScore.ControlImageRect = Rectangle.FromLTRB(192, 157 + (yPopulace * 9), 256, 157 + ((yPopulace + 1) * 9))
                lblPopScore.BackImageColor = mclrVals(yPopulace)
                lblPopScore.DrawBackImage = True
				lblPopScore.Caption = ""
                lblPopScore.ToolTipText = "Population Score: " & lPopulaceScore.ToString("#,##0") & vbCrLf & "Last Updated: " & GetDurationFromSeconds(CInt(Now.Subtract(GetDateFromNumber(lPopulaceUpdate)).TotalSeconds), True)
				lblPopScore.BackImageColor = GetValueColor(yPopulace)
            End If
            If yProduction = 255 Then
                lblProdScore.DrawBackImage = False
				lblProdScore.Caption = "Unknown"
				lblProdScore.ToolTipText = ""
            Else
				lblProdScore.ControlImageRect = Rectangle.FromLTRB(192, 157 + (yProduction * 9), 256, 157 + ((yProduction + 1) * 9))
                lblProdScore.BackImageColor = mclrVals(yProduction)
                lblProdScore.DrawBackImage = True
				lblProdScore.Caption = ""
                lblProdScore.ToolTipText = "Production Score: " & lProductionScore.ToString("#,##0") & vbCrLf & "Last Updated: " & GetDurationFromSeconds(CInt(Now.Subtract(GetDateFromNumber(lProductionUpdate)).TotalSeconds), True)
				lblProdScore.BackImageColor = GetValueColor(yProduction)
            End If
            If yWealth = 255 Then
                lblWealthScore.DrawBackImage = False
				lblWealthScore.Caption = "Unknown"
				lblWealthScore.ToolTipText = ""
            Else
				lblWealthScore.ControlImageRect = Rectangle.FromLTRB(192, 157 + (yWealth * 9), 256, 157 + ((yWealth + 1) * 9))
                lblWealthScore.BackImageColor = mclrVals(yWealth)
                lblWealthScore.DrawBackImage = True
				lblWealthScore.Caption = ""
                lblWealthScore.ToolTipText = "Wealth Score: " & lWealthScore.ToString("#,##0") & vbCrLf & "Last Updated: " & GetDurationFromSeconds(CInt(Now.Subtract(GetDateFromNumber(lWealthUpdate)).TotalSeconds), True)
				lblWealthScore.BackImageColor = GetValueColor(yWealth)
            End If

            lblTotalScore.Caption = "Total Score: " & lTotalScore.ToString("#,##0")

            'If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 255 Then
            '    hscrRel.MinValue = Math.Max(1, CInt(moRel.WithThisScore) - 5)
            '    hscrRel.MaxValue = Math.Min(255, CInt(moRel.WithThisScore) + 5)
            'End If
 
		End Sub
        Public Function GetCurrentRelID() As Int32
            If moRel Is Nothing = False Then
                If moRel.oPlayerIntel Is Nothing = False Then Return moRel.oPlayerIntel.ObjectID
            End If
            Return -1
        End Function
		Private Function GetValueColor(ByVal yValue As Byte) As Color
			Select Case yValue
				Case 0
					Return System.Drawing.Color.FromArgb(255, 128, 128, 128)
				Case 1
					Return System.Drawing.Color.FromArgb(255, 0, 128, 0)
				Case 2
					Return System.Drawing.Color.FromArgb(255, 0, 255, 0)
				Case 3
					Return System.Drawing.Color.FromArgb(255, 0, 255, 255)
				Case 4
					Return System.Drawing.Color.FromArgb(255, 0, 0, 255)
				Case 5
					Return System.Drawing.Color.FromArgb(255, 255, 255, 0)
				Case 6
					Return System.Drawing.Color.FromArgb(255, 255, 128, 0)
				Case 7
					Return System.Drawing.Color.FromArgb(255, 128, 0, 0)
				Case 8
					Return System.Drawing.Color.FromArgb(255, 255, 0, 0)
				Case Else
					Return System.Drawing.Color.FromArgb(255, 255, 0, 255)
			End Select
		End Function

		Private Sub btnReset_Click(ByVal sName As String) Handles btnReset.Click
			If moRel Is Nothing Then Return
			hscrRel.Value = moRel.WithThisScore
			hscrRel_ValueChange()
		End Sub

		Private Sub btnSet_Click(ByVal sName As String) Handles btnSet.Click
			Dim yData(10) As Byte
			Dim X As Int32

			If HasAliasedRights(AliasingRights.eAlterDiplomacy) = False Then
				MyBase.moUILib.AddNotification("You lack rights to alter diplomatic relations.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
				Return
			End If

			If moRel Is Nothing Then Return
			If moRel.lPlayerRegards <> glPlayerID Then Return
            If moRel.TargetScore = hscrRel.Value Then Return

			If hscrRel.Value <= elRelTypes.eWar AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
				If (goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.RequirePeaceBetweenMembers) <> 0 Then
					Dim oMember As GuildMember = goCurrentPlayer.oGuild.GetMember(moRel.lThisPlayer)
					If oMember Is Nothing = False Then
						MyBase.moUILib.AddNotification("Unable to go to set relationship to war. That player is in your guild and your guild requires peace.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
						Return
					End If
					oMember = Nothing
				End If
			End If

            If hscrRel.Value <= elRelTypes.eWar Then

                'Ok, get the target's rank... and compare it to my rank. If I am in a guild, get the maximum rank in my guild
                Dim lMyRank As Int32 = goCurrentPlayer.yPlayerTitle
                If (lMyRank And Player.PlayerRank.ExRankShift) <> 0 Then lMyRank = lMyRank Xor Player.PlayerRank.ExRankShift
                If goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.moMembers Is Nothing = False Then
                    For Y As Int32 = 0 To goCurrentPlayer.oGuild.moMembers.GetUpperBound(0)
                        If goCurrentPlayer.oGuild.moMembers(Y) Is Nothing = False Then
                            If (goCurrentPlayer.oGuild.moMembers(Y).yMemberState And GuildMemberState.Approved) <> 0 Then
                                Dim lTemp As Int32 = goCurrentPlayer.oGuild.moMembers(Y).yPlayerTitle
                                If (lTemp And Player.PlayerRank.ExRankShift) <> 0 Then lTemp = lTemp Xor Player.PlayerRank.ExRankShift
                                lMyRank = Math.Max(lMyRank, lTemp)
                            End If
                        End If
                    Next Y
                End If

                Dim lTempPlayerTitle As Int32 = moRel.oPlayerIntel.yPlayerTitle
                If (lTempPlayerTitle And Player.PlayerRank.ExRankShift) <> 0 Then lTempPlayerTitle = lTempPlayerTitle Xor Player.PlayerRank.ExRankShift
                If lTempPlayerTitle < lMyRank Then
                    Dim ofrm As New frmConfirmWar(MyBase.moUILib)
                    ofrm.SetPlayerRel(moRel, CByte(hscrRel.Value), lMyRank - CInt(moRel.oPlayerIntel.yPlayerTitle), lMyRank > goCurrentPlayer.yPlayerTitle)
                    Return
                End If
            End If

            moRel.TargetScore = CByte(hscrRel.Value)
            If moRel.WithThisScore < moRel.TargetScore Then moRel.WithThisScore = moRel.TargetScore
            'moRel.WithThisScore = CByte(hscrRel.Value)
            goCurrentPlayer.SetPlayerRel(moRel.lThisPlayer, moRel.WithThisScore)

			If goCurrentEnvir Is Nothing = False Then
				For X = 0 To goCurrentEnvir.lEntityUB
					If goCurrentEnvir.lEntityIdx(X) <> -1 Then
						If goCurrentEnvir.oEntity(X).OwnerID = moRel.lThisPlayer Then
							goCurrentEnvir.oEntity(X).yRelID = moRel.WithThisScore
						End If
					End If
				Next X
			End If

			System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yData, 0)
			System.BitConverter.GetBytes(goCurrentPlayer.ObjectID).CopyTo(yData, 2)
			System.BitConverter.GetBytes(moRel.lThisPlayer).CopyTo(yData, 6)
            yData(10) = moRel.TargetScore 'moRel.WithThisScore

			'Now, send to primary
            If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 1 AndAlso moRel.lThisPlayer = gl_HARDCODE_PIRATE_PLAYER_ID Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eRelationSet, CByte(hscrRel.Value), -1, -1, "")
                moRel.WithThisScore = 1
            Else
                MyBase.moUILib.SendMsgToPrimary(yData)
            End If
		End Sub

        Private Sub hscrRel_ValueChange() Handles hscrRel.ValueChange
            If mbLoading = True Then Return
            'lblOurScore.Caption = GetRelValText(CByte(hscrRel.Value))
            'lblOurScore.ForeColor = GetRelValColor(CByte(hscrRel.Value))
            lblTargetScoreVal.Caption = GetRelValText(CByte(hscrRel.Value))
            lblTargetScoreVal.ForeColor = GetRelValColor(CByte(hscrRel.Value))

            txtScore.Caption = hscrRel.Value.ToString
        End Sub

        Private Function GetRelValText(ByVal yVal As Byte) As String
            If yVal <= elRelTypes.eBloodWar Then
                Return "BLOOD WAR (" & yVal & ")"
            ElseIf yVal <= elRelTypes.eWar Then
                Return "WAR (" & yVal & ")"
            ElseIf yVal <= elRelTypes.eNeutral Then
                Return "NEUTRAL (" & yVal & ")"
            ElseIf yVal <= elRelTypes.ePeace Then
                Return "PEACE (" & yVal & ")"
            ElseIf yVal <= elRelTypes.eAlly Then
                Return "ALLY (" & yVal & ")"
            Else
                Return "BLOOD ALLY (" & yVal & ")"
            End If
        End Function

        Private Function GetRelValColor(ByVal yVal As Byte) As System.Drawing.Color
            If yVal <= elRelTypes.eWar Then
                Return System.Drawing.Color.FromArgb(255, 255, 0, 0)
            ElseIf yVal <= elRelTypes.eNeutral Then
                Return System.Drawing.Color.FromArgb(255, 255, 255, 255)
            ElseIf yVal <= elRelTypes.ePeace Then
                Return System.Drawing.Color.FromArgb(255, 0, 128, 0)
            Else
                Return System.Drawing.Color.FromArgb(255, 0, 255, 255)
            End If
        End Function

        Private Function GetScoreRatioValue(ByVal lMyScore As Int32, ByVal lTheirScore As Int32) As Byte
            If lTheirScore = Int32.MinValue Then Return 255
			If lMyScore = 0 Then Return 9
            Dim fTemp As Single = CSng(lTheirScore / lMyScore) - 0.5F
            Dim lValue As Int32 = CInt(fTemp * 10)
            If lValue < 0 Then lValue = 0
            If lValue > 9 Then lValue = 9

            Return CByte(lValue)
        End Function

		Private mlLastRelChangeUpdateCycle As Int32 = -1
		Public Sub fraPlayerDetails_OnNewFrame() Handles Me.OnNewFrame
			If glCurrentCycle - mlLastRelChangeUpdateCycle > 30 Then
				mlLastRelChangeUpdateCycle = glCurrentCycle
				If moRel Is Nothing = False Then
					Dim oPlayerIntel As PlayerIntel = Nothing
					If moRel.lPlayerRegards = glPlayerID Then
						For X As Int32 = 0 To glPlayerIntelUB
							If glPlayerIntelIdx(X) = moRel.lThisPlayer Then
								oPlayerIntel = goPlayerIntel(X)
								Exit For
							End If
                        Next X
                    Else
                        For X As Int32 = 0 To glPlayerIntelUB
                            If glPlayerIntelIdx(X) = moRel.lPlayerRegards Then
                                oPlayerIntel = goPlayerIntel(X)
                                Exit For
                            End If
                        Next X
                    End If

                    If moRel.lNextUpdateCycle > glCurrentCycle Then
                        If lblNextChng.ForeColor <> System.Drawing.Color.FromArgb(255, 255, 0, 0) Then lblNextChng.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)

                        Dim lSeconds As Int32 = (moRel.lNextUpdateCycle - glCurrentCycle) \ 30
                        Dim lMinutes As Int32 = lSeconds \ 60
                        Dim lHours As Int32 = lMinutes \ 60
                        lMinutes -= (lHours * 60)

                        Dim sTemp As String = lHours.ToString("0#") & ":" & lMinutes.ToString("0#")
                        If lblNextChng.Caption <> sTemp Then lblNextChng.Caption = sTemp
                    Else
                        If lblNextChng.ForeColor <> muSettings.InterfaceBorderColor Then lblNextChng.ForeColor = muSettings.InterfaceBorderColor
                        If lblNextChng.Caption <> "00:00" Then lblNextChng.Caption = "00:00"
                    End If

                    If oPlayerIntel Is Nothing = False Then
                        Dim sTemp As String = GetRelValText(oPlayerIntel.yRegardsCurrentPlayer)
                        If lblTheirScore.Caption <> sTemp Then lblTheirScore.Caption = sTemp
                        lblTheirScore.Caption = GetRelValText(oPlayerIntel.yRegardsCurrentPlayer)
                        Dim clrVal As System.Drawing.Color = GetRelValColor(oPlayerIntel.yRegardsCurrentPlayer)
                        If lblTheirScore.ForeColor.ToArgb <> clrVal.ToArgb Then lblTheirScore.ForeColor = clrVal
                    End If
                End If

                End If
        End Sub

        Public Sub SetExternalScore(ByVal lVal As Int32)
            hscrRel.Value = lVal
            hscrRel_ValueChange()
        End Sub

        Private Sub txtScore_TextChanged() Handles txtScore.TextChanged
            If mbLoading = True Then Return
            If IsNumeric(txtScore.Caption) = True Then
                Dim lVal As Int32 = CInt(txtScore.Caption)
                If lVal > 255 Then lVal = 255
                If lVal < 1 Then lVal = 1

                mbLoading = True
                hscrRel.Value = lVal
                txtScore.Caption = lVal.ToString
                lblTargetScoreVal.Caption = GetRelValText(CByte(hscrRel.Value))
                lblTargetScoreVal.ForeColor = GetRelValColor(CByte(hscrRel.Value))

                mbLoading = False
            End If
        End Sub
    End Class

End Class
