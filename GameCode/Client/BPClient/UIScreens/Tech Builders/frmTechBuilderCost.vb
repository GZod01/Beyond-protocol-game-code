Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmTechBuilderCost
    Inherits UIWindow

    Private lblResearchFac As UILabel
    Private lblProductionFac As UILabel
    Private WithEvents cboResearchFac As UIComboBox
    Private WithEvents cboProductionFac As UIComboBox

    Private WithEvents tpPowerRequired As ctlTechProp
    Private WithEvents tpHullRequired As ctlTechProp
    Private WithEvents tpProdCredits As ctlTechProp
    Private WithEvents tpProdTime As ctlTechProp
    Private WithEvents tpResCredits As ctlTechProp
    Private WithEvents tpResTime As ctlTechProp
    Private WithEvents tpColonists As ctlTechProp
    Private WithEvents tpEnlisted As ctlTechProp
    Private WithEvents tpOfficers As ctlTechProp

    Private lblCrewHull As UILabel

    Public Event ValueChanged()

    Private mdecBill As Decimal
    Private mdecPayment As Decimal
    Private mrcCostPayment As Rectangle

    Private mbIgnoreEvents As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
 
        'Dim lMaxPowerThrust As Int32 = 10000
        'If goCurrentPlayer Is Nothing = False Then
        '    lMaxPowerThrust = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)
        'End If
        'frmTechBuilderCost initial props
        With Me
            .ControlName = "frmTechBuilderCost"
            .Left = 10
            .Top = 10
            .Width = 490 ' 512
            .Height = 230 '520
            .Enabled = True
            .Visible = True 
            .BorderColor = muSettings.InterfaceBorderColor 
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .mbAcceptReprocessEvents = True
        End With

        tpHullRequired = New ctlTechProp(oUILib)
        With tpHullRequired
            .ControlName = "tpHullRequired"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 10000 '00 ' lMaxPowerThrust
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Hull Required:"
            .PropertyValue = 0
            .bNoMaxValue = True
            .Top = 25
            .Visible = True
            .blAbsoluteMaximum = 100000000
            .SetExtendedProps(False, 100, 120, True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpHullRequired, UIControl))

        tpPowerRequired = New ctlTechProp(oUILib)
        With tpPowerRequired
            .ControlName = "tpPowerRequired"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 10000 '00 'lMaxPowerThrust
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Power Required:"
            .PropertyValue = 0
            .bNoMaxValue = True
            .Top = tpHullRequired.Top + 20
            .Visible = True
            .SetExtendedProps(False, 100, 120, True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpPowerRequired, UIControl))

        tpResCredits = New ctlTechProp(oUILib)
        With tpResCredits
            .ControlName = "tpResCredits"
            .Enabled = True
            .Height = 18
            .Left = tpHullRequired.Left
            .MaxValue = Int32.MaxValue
            .MinValue = 0
            .bNoMaxValue = True
            .PropertyLocked = False
            .PropertyName = "Research Cost:"
            .PropertyValue = 0
            .Top = tpPowerRequired.Top + 20
            .Visible = True
            .SetExtendedProps(False, 100, 120, True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpResCredits, UIControl))

        tpResTime = New ctlTechProp(oUILib)
        With tpResTime
            .ControlName = "tpResTime"
            .Enabled = True
            .Height = 18
            .Left = tpHullRequired.Left
            .MaxValue = Int32.MaxValue
            .MinValue = 0
            .bNoMaxValue = True
            .PropertyLocked = False
            .PropertyName = "Research Time:"
            .PropertyValue = 0
            .Top = tpResCredits.Top + 20
            .Visible = True
            .SetAsTimeIndicator(3000, True)
            .SetExtendedProps(False, 100, 120, True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpResTime, UIControl))

        tpProdCredits = New ctlTechProp(oUILib)
        With tpProdCredits
            .ControlName = "tpProdCredits"
            .Enabled = True
            .Height = 18
            .Left = tpPowerRequired.Left
            .MaxValue = Int32.MaxValue
            .MinValue = 0
            .bNoMaxValue = True
            .PropertyLocked = False
            .PropertyName = "Production Cost:"
            .PropertyValue = 0
            .Top = tpResTime.Top + 20
            .Visible = True
            .SetExtendedProps(False, 100, 120, True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpProdCredits, UIControl))

        tpProdTime = New ctlTechProp(oUILib)
        With tpProdTime
            .ControlName = "tpProdTime"
            .Enabled = True
            .Height = 18
            .Left = tpPowerRequired.Left
            .MaxValue = Int32.MaxValue
            .MinValue = 0
            .bNoMaxValue = True
            .PropertyLocked = False
            .PropertyName = "Production Time:"
            .PropertyValue = 0
            .Top = tpProdCredits.Top + 20
            .Visible = True
            .SetAsTimeIndicator(15000, True)
            .SetExtendedProps(False, 100, 120, True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpProdTime, UIControl))

        tpColonists = New ctlTechProp(oUILib)
        With tpColonists
            .ControlName = "tpColonists"
            .Enabled = True
            .Height = 18
            .Left = tpPowerRequired.Left
            .MaxValue = 100000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Colonists:"
            .PropertyValue = 0
            .Top = tpProdTime.Top + 20
            .Visible = True
            .SetExtendedProps(False, 100, 120, True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpColonists, UIControl))

        tpEnlisted = New ctlTechProp(oUILib)
        With tpEnlisted
            .ControlName = "tpEnlisted"
            .Enabled = True
            .Height = 18
            .Left = tpPowerRequired.Left
            .MaxValue = 100000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Enlisted:"
            .PropertyValue = 0
            .Top = tpColonists.Top + 20
            .Visible = True
            .SetExtendedProps(False, 100, 120, True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpEnlisted, UIControl))

        tpOfficers = New ctlTechProp(oUILib)
        With tpOfficers
            .ControlName = "tpOfficers"
            .Enabled = True
            .Height = 18
            .Left = tpPowerRequired.Left
            .MaxValue = 100000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Officers:"
            .PropertyValue = 0
            .Top = tpEnlisted.Top + 20
            .Visible = True
            .SetExtendedProps(False, 100, 120, True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpOfficers, UIControl))

        'lblCrewHull initial props
        lblCrewHull = New UILabel(oUILib)
        With lblCrewHull
            .ControlName = "lblCrewHull"
            .Left = 15
            .Top = tpOfficers.Top + 20
            .Width = Me.Width - (2 * .Left)
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Crew will take an additional 0 hull"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblCrewHull, UIControl))

        'lblResearchFac initial props
        lblResearchFac = New UILabel(oUILib)
        With lblResearchFac
            .ControlName = "lblResearchFac"
            .Left = 15
            .Top = 5
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Research:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblResearchFac, UIControl))

        'cboResearchFac initial props
        cboResearchFac = New UIComboBox(oUILib)
        With cboResearchFac
            .ControlName = "cboResearchFac"
            .Left = lblResearchFac.Left + lblResearchFac.Width + 5
            .Top = lblResearchFac.Top
            .Width = 150 '150
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboResearchFac, UIControl))

        'lblProductionFac initial props
        lblProductionFac = New UILabel(oUILib)
        With lblProductionFac
            .ControlName = "lblProductionFac"
            .Left = cboResearchFac.Left + cboResearchFac.Width + 10
            .Top = lblResearchFac.Top
            .Width = 70
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Production:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblProductionFac, UIControl))

        'cboProductionFac initial props
        cboProductionFac = New UIComboBox(oUILib)
        With cboProductionFac
            .ControlName = "cboProductionFac"
            .Left = lblProductionFac.Left + lblProductionFac.Width + 5
            .Top = lblProductionFac.Top
            .Width = 150 '150
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboProductionFac, UIControl))
        FillValues()
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub FillValues()
        cboResearchFac.Clear()
        cboProductionFac.Clear()

        For X As Int32 = 0 To glEntityDefUB
            If glEntityDefIdx(X) > -1 Then
                Dim oDef As EntityDef = goEntityDefs(X)
                If oDef Is Nothing = False AndAlso oDef.ObjTypeID = ObjectType.eFacilityDef Then
                    If oDef.ProductionTypeID = ProductionType.eResearch Then
                        cboResearchFac.AddItem(oDef.DefName)
                        cboResearchFac.ItemData(cboResearchFac.NewIndex) = oDef.ObjectID
                    ElseIf (oDef.ProductionTypeID And ProductionType.eProduction) <> 0 Then
                        cboProductionFac.AddItem(oDef.DefName)
                        cboProductionFac.ItemData(cboProductionFac.NewIndex) = oDef.ObjectID
                    End If
                End If
            End If
        Next X

        cboResearchFac.FindComboItemData(52)
        cboProductionFac.FindComboItemData(49)
    End Sub

    Public Sub SetBillPaymentValues(ByVal decBill As Decimal, ByVal decPayment As Decimal)
        mdecBill = decBill
        mdecPayment = decPayment
        Me.IsDirty = True
    End Sub

    Private Function GetPercValueFromDecs(ByVal decTotal As Decimal, ByVal decValue As Decimal) As String
        Dim decTemp As Decimal = (decValue / decTotal) * 100
        If decTemp < 0 Then Return "0%"
        If decTemp > 255 Then Return "255%"
        Return decTemp.ToString("#,##0.#") & "%"
    End Function

    Private mlNextLvlPower As Int32 = 1
    Private mlNextLvlHull As Int32 = 1
    Private mlNextLvlColonist As Int32 = 1
    Private mlNextLvlEnlisted As Int32 = 1
    Private mlNextLvlOfficers As Int32 = 1
    Private mblNextProdCost As Int64 = 1
    Private mblNextProdTime As Int64 = 1
    Private mblNextResCost As Int64 = 1
    Private mblNextResTime As Int64 = 1

    Private Function SetNextPointValue32(ByVal decPay As Decimal, ByVal decTotal As Decimal, ByVal decCoeff As Decimal) As Int32
        Try
            Dim decTemp As Decimal = CInt((decPay / decTotal) * 100)
            Dim lNextPercPoint As Int32
            If decTemp > 100 Then
                lNextPercPoint = 100
            ElseIf decTemp < 0 Then
                lNextPercPoint = 0
            Else
                lNextPercPoint = CInt(Math.Min(100, Math.Floor(decTemp) + 1))
            End If
            If decCoeff = 0 Then Return 0

            Dim decPerc As Decimal = CDec(Math.Pow(((lNextPercPoint / 100D) * decTotal), (1 / decCoeff)))
            If decPerc > Int32.MaxValue Then Return Int32.MaxValue Else Return CInt(Math.Ceiling(decPerc))
        Catch
            Return 0
        End Try
    End Function
    Private Function SetNextPointValue64(ByVal decPay As Decimal, ByVal decTotal As Decimal, ByVal decCoeff As Decimal) As Int64
        Try
            Dim decTemp As Decimal = CLng((decPay / decTotal) * 100)
            Dim lNextPercPoint As Int32
            If decTemp > 100 Then
                lNextPercPoint = 100
            ElseIf decTemp < 0 Then
                lNextPercPoint = 0
            Else
                lNextPercPoint = CInt(Math.Min(100, Math.Floor(decTemp) + 1))
            End If

            Dim decPerc As Decimal = CDec(Math.Pow(((lNextPercPoint / 100D) * decTotal), (1 / decCoeff)))
            If decPerc > Int64.MaxValue Then Return Int64.MaxValue Else Return CLng(Math.Ceiling(decPerc))
        Catch
            Return 0
        End Try
    End Function
    Public Sub SetPaymentPercs(ByVal oTech As TechBuilderComputer)
        If oTech.decTotalBill > 0 Then
            With oTech
                tpPowerRequired.SetPercOfPayment(GetPercValueFromDecs(.decTotalBill, .decPowerPayment))
                tpHullRequired.SetPercOfPayment(GetPercValueFromDecs(.decTotalBill, .decHullPayment))
                tpProdCredits.SetPercOfPayment(GetPercValueFromDecs(.decTotalBill, .decProdCostPayment))
                tpProdTime.SetPercOfPayment(GetPercValueFromDecs(.decTotalBill, .decProdTimePayment))
                tpResCredits.SetPercOfPayment(GetPercValueFromDecs(.decTotalBill, .decResCostPayment))
                tpResTime.SetPercOfPayment(GetPercValueFromDecs(.decTotalBill, .decResTimePayment))
                tpColonists.SetPercOfPayment(GetPercValueFromDecs(.decTotalBill, .decColonistPayment))
                tpEnlisted.SetPercOfPayment(GetPercValueFromDecs(.decTotalBill, .decEnlistedPayment))
                tpOfficers.SetPercOfPayment(GetPercValueFromDecs(.decTotalBill, .decOfficersPayment))

                mlNextLvlColonist = SetNextPointValue32(.decColonistPayment, .decTotalBill, .GetCoeffValue(TechBuilderComputer.elPropCoeffLookup.eColonist))
                mlNextLvlEnlisted = SetNextPointValue32(.decEnlistedPayment, .decTotalBill, .GetCoeffValue(TechBuilderComputer.elPropCoeffLookup.eEnlisted))
                mlNextLvlOfficers = SetNextPointValue32(.decOfficersPayment, .decTotalBill, .GetCoeffValue(TechBuilderComputer.elPropCoeffLookup.eOfficer))
                mlNextLvlHull = SetNextPointValue32(.decHullPayment, .decTotalBill, .GetCoeffValue(TechBuilderComputer.elPropCoeffLookup.eHull))
                mlNextLvlPower = SetNextPointValue32(.decPowerPayment, .decTotalBill, .GetCoeffValue(TechBuilderComputer.elPropCoeffLookup.ePower))
                mblNextProdCost = SetNextPointValue64(.decProdCostPayment, .decTotalBill, .GetCoeffValue(TechBuilderComputer.elPropCoeffLookup.eProdCost))
                mblNextProdTime = SetNextPointValue64(.decProdTimePayment, .decTotalBill, .GetCoeffValue(TechBuilderComputer.elPropCoeffLookup.eProdTime))
                mblNextResCost = SetNextPointValue64(.decResCostPayment, .decTotalBill, .GetCoeffValue(TechBuilderComputer.elPropCoeffLookup.eResCost))
                mblNextResTime = SetNextPointValue64(.decResTimePayment, .decTotalBill, .GetCoeffValue(TechBuilderComputer.elPropCoeffLookup.eResTime))
            End With
        Else
            tpPowerRequired.SetPercOfPayment("0%")
            tpHullRequired.SetPercOfPayment("0%")
            tpProdCredits.SetPercOfPayment("0%")
            tpProdTime.SetPercOfPayment("0%")
            tpResCredits.SetPercOfPayment("0%")
            tpResTime.SetPercOfPayment("0%")
            tpColonists.SetPercOfPayment("0%")
            tpEnlisted.SetPercOfPayment("0%")
            tpOfficers.SetPercOfPayment("0%")
        End If
    End Sub
 
    Public Sub SetFromErrorCode(ByVal lError As TechBuilderComputer.elErrorReasons, ByVal iTypeID As Int16, ByRef tpMin1 As ctlTechProp, ByRef tpMin2 As ctlTechProp, ByRef tpMin3 As ctlTechProp, ByRef tpMin4 As ctlTechProp, ByRef tpMin5 As ctlTechProp, ByRef tpMin6 As ctlTechProp)
        Dim sbHull As New System.Text.StringBuilder()
        Dim sbPower As New System.Text.StringBuilder()
        Dim sbResTime As New System.Text.StringBuilder()
        Dim sbResCost As New System.Text.StringBuilder()
        Dim sbProdTime As New System.Text.StringBuilder()
        Dim sbProdCost As New System.Text.StringBuilder()
        Dim sbCol As New System.Text.StringBuilder()
        Dim sbEnl As New System.Text.StringBuilder()
        Dim sbOff As New System.Text.StringBuilder()
        Dim sbMin1 As New System.Text.StringBuilder()
        Dim sbMin2 As New System.Text.StringBuilder()
        Dim sbMin3 As New System.Text.StringBuilder()
        Dim sbMin4 As New System.Text.StringBuilder()
        Dim sbMin5 As New System.Text.StringBuilder()
        Dim sbMin6 As New System.Text.StringBuilder()

        If (lError And TechBuilderComputer.elErrorReasons.ePaymentOverBill) <> 0 Then
            Dim sText As String = "Total of all percentages cannot exceed 100%"
            sbHull.AppendLine(sText)
            sbPower.AppendLine(sText)
            sbMin1.AppendLine(sText)
            sbMin2.AppendLine(sText)
            sbMin3.AppendLine(sText)
            sbMin4.AppendLine(sText)
            sbMin5.AppendLine(sText)
            sbMin6.AppendLine(sText)
            sbCol.AppendLine(sText)
            sbEnl.AppendLine(sText)
            sbOff.AppendLine(sText)
            sbResTime.AppendLine(sText)
            sbResCost.AppendLine(sText)
            sbProdCost.AppendLine(sText)
            sbProdTime.AppendLine(sText)
        End If

        If (lError And TechBuilderComputer.elErrorReasons.eHullPowerPointsLow) <> 0 Then
            Dim sText As String
            'If iTypeID = ObjectType.eRadarTech Then
            '    sText = "Hull and Power Percentage combined must exceed 2%"
            'Else
            sText = "Hull and Power Percentage combined must exceed 10%"
            'End If
            sbHull.AppendLine(sText)
            sbPower.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eHullLessThan1Perc) <> 0 Then
            Dim sText As String = "Hull percentage must exceed 1%"
            sbHull.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.ePowerLessThan1Perc) <> 0 Then
            Dim sText As String = "Power percentage must exceed 1%"
            sbPower.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eHullMineralPointsLow) <> 0 Then
            Dim sText As String = "Hull and Mineral Percentages combined must exceed 15%"
            sbHull.AppendLine(sText)
            sbMin1.AppendLine(sText)
            sbMin2.AppendLine(sText)
            sbMin3.AppendLine(sText)
            sbMin4.AppendLine(sText)
            sbMin5.AppendLine(sText)
            sbMin6.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eMineralsToHullSize) <> 0 Then
            Dim sText As String = "Mineral quantities combined must exceed Hull Size"
            sbHull.AppendLine(sText)
            sbMin1.AppendLine(sText)
            sbMin2.AppendLine(sText)
            sbMin3.AppendLine(sText)
            sbMin4.AppendLine(sText)
            sbMin5.AppendLine(sText)
            sbMin6.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eHullToCrewRatio) <> 0 Then
            Dim sText As String = "Colonists, Enlisted and Officers combined cannot exceed Hull Size"
            sbHull.AppendLine(sText)
            sbCol.AppendLine(sText)
            sbEnl.AppendLine(sText)
            sbOff.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eHullToCrewHullRatio) <> 0 Then
            Dim sText As String = "Hull taken up by Colonists, Enlisted and Officers cannot exceed Hull Size"
            sbHull.AppendLine(sText)
            sbCol.AppendLine(sText)
            sbEnl.AppendLine(sText)
            sbOff.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eMilitaryToCitizen) <> 0 Then
            Dim sText As String = "Colonists cannot exceed Officers and Enlisted combined"
            sbCol.AppendLine(sText)
            sbEnl.AppendLine(sText)
            sbOff.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eLackOfOfficers) <> 0 Then
            Dim sText As String = "Enlisted cannot exceed 5 times the Officers"
            sbEnl.AppendLine(sText)
            sbOff.AppendLine(sText)
        End If

        If (lError And TechBuilderComputer.elErrorReasons.eProdCostToResCostRatio) <> 0 Then
            Dim sText As String
            If tpProdCredits.PropertyValue > tpResCredits.PropertyValue Then
                sText = "Research Credits must be greater than 1% of Production Credits"
            Else
                sText = "Production Credits must be greater than 1% of Research Credits"
            End If
            sbProdCost.AppendLine(sText)
            sbResCost.AppendLine(sText)
        End If

        If (lError And TechBuilderComputer.elErrorReasons.eProdTimeToResTimeRatio) <> 0 Then
            Dim sText As String
            If tpProdTime.PropertyValue > tpResTime.PropertyValue Then
                sText = "Research Time must be greater than 1/1000th of Production Time"
            Else
                sText = "Production Time must be greater than 1/1000th of Research Time"
            End If
            sbProdTime.AppendLine(sText)
            sbResTime.AppendLine(sText)
        End If

        If (lError And TechBuilderComputer.elErrorReasons.eResTimeProdTimePoints) <> 0 Then
            Dim sText As String = "Research Time percentange and Production Time percentage" & vbCrLf & "  combined must be greater than 10% and less than 90%"
            sbProdTime.AppendLine(sText)
            sbResTime.AppendLine(sText)
        End If

        If (lError And TechBuilderComputer.elErrorReasons.eResCostProdCostPoints) <> 0 Then
            Dim sText As String = "Production Credits percentage and Research Credits percentage" & vbCrLf & "  combined must be greater than 5% and less than 50%"
            sbProdCost.AppendLine(sText)
            sbResCost.AppendLine(sText)
        End If

        If (lError And TechBuilderComputer.elErrorReasons.eMineral1LessThan15) <> 0 Then
            Dim sText As String
            If iTypeID = ObjectType.eEngineTech Then
                sText = "Structural Body"
            Else
                sText = tpMin1.PropertyName
            End If
            sText &= " cannot be less than 15% of all mineral quantities combined."
            sbMin1.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eMineral2LessThan15) <> 0 Then
            Dim sText As String
            If iTypeID = ObjectType.eEngineTech Then
                sText = "Structural Frame"
            Else
                sText = tpMin2.PropertyName
            End If
            sText &= " cannot be less than 15% of all mineral quantities combined."
            sbMin2.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eMineral3LessThan15) <> 0 Then
            Dim sText As String
            If iTypeID = ObjectType.eEngineTech Then
                sText = "Structural Meld"
            Else
                sText = tpMin3.PropertyName
            End If
            sText &= " cannot be less than 15% of all mineral quantities combined."
            sbMin3.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eMineral4LessThan15) <> 0 Then
            Dim sText As String
            If iTypeID = ObjectType.eEngineTech Then
                sText = "Drive Body"
            Else
                sText = tpMin4.PropertyName
            End If
            sText &= " cannot be less than 15% of all mineral quantities combined."
            sbMin4.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eMineral5LessThan15) <> 0 Then
            Dim sText As String
            If iTypeID = ObjectType.eEngineTech Then
                sText = "Drive Frame"
            Else
                sText = tpMin5.PropertyName
            End If
            sText &= " cannot be less than 15% of all mineral quantities combined."
            sbMin5.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eMineral6LessThan15) <> 0 Then
            Dim sText As String
            If iTypeID = ObjectType.eEngineTech Then
                sText = "Drive Meld"
            Else
                sText = tpMin6.PropertyName
            End If
            sText &= " cannot be less than 15% of all mineral quantities combined."
            sbMin6.AppendLine(sText)
        End If

        If (lError And TechBuilderComputer.elErrorReasons.e80PercMarkerExceed) <> 0 Then
            Dim sText As String = "Percentage cannot exceed 80%"
            If Val(tpColonists.GetPercOfPayment()) > 80 Then sbCol.AppendLine(sText)
            If Val(tpEnlisted.GetPercOfPayment()) > 80 Then sbEnl.AppendLine(sText)
            If Val(tpOfficers.GetPercOfPayment()) > 80 Then sbOff.AppendLine(sText)
            If Val(tpProdCredits.GetPercOfPayment()) > 80 Then sbProdCost.AppendLine(sText)
            If Val(tpProdTime.GetPercOfPayment()) > 80 Then sbProdTime.AppendLine(sText)
            If Val(tpResCredits.GetPercOfPayment()) > 80 Then sbResCost.AppendLine(sText)
            If Val(tpResTime.GetPercOfPayment()) > 80 Then sbResTime.AppendLine(sText)
            If Val(tpPowerRequired.GetPercOfPayment()) > 80 Then sbPower.AppendLine(sText)
            If Val(tpHullRequired.GetPercOfPayment()) > 80 Then sbHull.AppendLine(sText)
            If tpMin1 Is Nothing = False AndAlso Val(tpMin1.GetPercOfPayment()) > 80 Then sbMin1.AppendLine(sText)
            If tpMin2 Is Nothing = False AndAlso Val(tpMin2.GetPercOfPayment()) > 80 Then sbMin2.AppendLine(sText)
            If tpMin3 Is Nothing = False AndAlso Val(tpMin3.GetPercOfPayment()) > 80 Then sbMin3.AppendLine(sText)
            If tpMin4 Is Nothing = False AndAlso Val(tpMin4.GetPercOfPayment()) > 80 Then sbMin4.AppendLine(sText)
            If tpMin5 Is Nothing = False AndAlso Val(tpMin5.GetPercOfPayment()) > 80 Then sbMin5.AppendLine(sText)
            If tpMin6 Is Nothing = False AndAlso Val(tpMin6.GetPercOfPayment()) > 80 Then sbMin6.AppendLine(sText)
        End If
        If (lError And TechBuilderComputer.elErrorReasons.eValueIsBelowZero) <> 0 Then
            Dim sText As String = "Cannot go below zero and be locked"
            If tpColonists.PropertyLocked = True AndAlso tpColonists.PropertyValue < 1 Then sbCol.AppendLine(sText)
            If tpEnlisted.PropertyLocked = True AndAlso tpEnlisted.PropertyValue < 1 Then sbEnl.AppendLine(sText)
            If tpOfficers.PropertyLocked = True AndAlso tpOfficers.PropertyValue < 1 Then sbOff.AppendLine(sText)
            If tpProdCredits.PropertyLocked = True AndAlso tpProdCredits.PropertyValue < 1 Then sbProdCost.AppendLine(sText)
            If tpProdTime.PropertyLocked = True AndAlso tpProdTime.PropertyValue < 1 Then sbProdTime.AppendLine(sText)
            If tpResCredits.PropertyLocked = True AndAlso tpResCredits.PropertyValue < 1 Then sbResCost.AppendLine(sText)
            If tpResTime.PropertyLocked = True AndAlso tpResTime.PropertyValue < 1 Then sbResTime.AppendLine(sText)
            If tpPowerRequired.PropertyLocked = True AndAlso tpPowerRequired.PropertyValue < 1 Then sbPower.AppendLine(sText)
            If tpHullRequired.PropertyLocked = True AndAlso tpHullRequired.PropertyValue < 1 Then sbHull.AppendLine(sText)

            If tpMin1 Is Nothing = False AndAlso tpMin1.PropertyLocked = True AndAlso tpMin1.PropertyValue < 1 Then sbMin1.AppendLine(sText)
            If tpMin2 Is Nothing = False AndAlso tpMin2.PropertyLocked = True AndAlso tpMin2.PropertyValue < 1 Then sbMin2.AppendLine(sText)
            If tpMin3 Is Nothing = False AndAlso tpMin3.PropertyLocked = True AndAlso tpMin3.PropertyValue < 1 Then sbMin3.AppendLine(sText)
            If tpMin4 Is Nothing = False AndAlso tpMin4.PropertyLocked = True AndAlso tpMin4.PropertyValue < 1 Then sbMin4.AppendLine(sText)
            If tpMin5 Is Nothing = False AndAlso tpMin5.PropertyLocked = True AndAlso tpMin5.PropertyValue < 1 Then sbMin5.AppendLine(sText)
            If tpMin6 Is Nothing = False AndAlso tpMin6.PropertyLocked = True AndAlso tpMin6.PropertyValue < 1 Then sbMin6.AppendLine(sText)
        End If

        'Now, check our results
        Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        If sbCol.Length > 0 Then
            tpColonists.SetPercError(clrVal, sbCol.ToString)
        Else : tpColonists.SetPercError(muSettings.InterfaceBorderColor, "")
        End If
        If sbEnl.Length > 0 Then
            tpEnlisted.SetPercError(clrVal, sbEnl.ToString)
        Else : tpEnlisted.SetPercError(muSettings.InterfaceBorderColor, "")
        End If
        If sbOff.Length > 0 Then
            tpOfficers.SetPercError(clrVal, sbOff.ToString)
        Else : tpOfficers.SetPercError(muSettings.InterfaceBorderColor, "")
        End If
        If sbProdCost.Length > 0 Then
            tpProdCredits.SetPercError(clrVal, sbProdCost.ToString)
        Else : tpProdCredits.SetPercError(muSettings.InterfaceBorderColor, "")
        End If
        If sbProdTime.Length > 0 Then
            tpProdTime.SetPercError(clrVal, sbProdTime.ToString)
        Else : tpProdTime.SetPercError(muSettings.InterfaceBorderColor, "")
        End If
        If sbResCost.Length > 0 Then
            tpResCredits.SetPercError(clrVal, sbResCost.ToString)
        Else : tpResCredits.SetPercError(muSettings.InterfaceBorderColor, "")
        End If
        If sbResTime.Length > 0 Then
            tpResTime.SetPercError(clrVal, sbResTime.ToString)
        Else : tpResTime.SetPercError(muSettings.InterfaceBorderColor, "")
        End If
        If sbPower.Length > 0 Then
            tpPowerRequired.SetPercError(clrVal, sbPower.ToString)
        Else : tpPowerRequired.SetPercError(muSettings.InterfaceBorderColor, "")
        End If
        If sbHull.Length > 0 Then
            tpHullRequired.SetPercError(clrVal, sbHull.ToString)
        Else : tpHullRequired.SetPercError(muSettings.InterfaceBorderColor, "")
        End If

        If tpMin1 Is Nothing = False Then
            If sbMin1.Length > 0 Then
                tpMin1.SetPercError(clrVal, sbMin1.ToString)
            Else : tpMin1.SetPercError(muSettings.InterfaceBorderColor, "")
            End If
        End If
        If tpMin2 Is Nothing = False Then
            If sbMin2.Length > 0 Then
                tpMin2.SetPercError(clrVal, sbMin2.ToString)
            Else : tpMin2.SetPercError(muSettings.InterfaceBorderColor, "")
            End If
        End If
        If tpMin3 Is Nothing = False Then
            If sbMin3.Length > 0 Then
                tpMin3.SetPercError(clrVal, sbMin3.ToString)
            Else : tpMin3.SetPercError(muSettings.InterfaceBorderColor, "")
            End If
        End If
        If tpMin4 Is Nothing = False Then
            If sbMin4.Length > 0 Then
                tpMin4.SetPercError(clrVal, sbMin4.ToString)
            Else : tpMin4.SetPercError(muSettings.InterfaceBorderColor, "")
            End If
        End If
        If tpMin5 Is Nothing = False Then
            If sbMin5.Length > 0 Then
                tpMin5.SetPercError(clrVal, sbMin5.ToString)
            Else : tpMin5.SetPercError(muSettings.InterfaceBorderColor, "")
            End If
        End If
        If tpMin6 Is Nothing = False Then
            If sbMin6.Length > 0 Then
                tpMin6.SetPercError(clrVal, sbMin6.ToString)
            Else : tpMin6.SetPercError(muSettings.InterfaceBorderColor, "")
            End If
        End If 
    End Sub

    Public Sub SetBaseValues(ByVal lHullReq As Int32, ByVal lPowerReq As Int32, ByVal blResCost As Int64, ByVal blResTime As Int64, ByVal blProdCost As Int64, ByVal blProdTime As Int64, ByVal lColonists As Int32, ByVal lEnlisted As Int32, ByVal lOfficers As Int32, ByVal lHullTypeID As Int32)
        mbIgnoreEvents = True

        Dim decNormalizer As Decimal = 0D
        Dim lMaxGuns As Int32 = 0
        Dim lMaxDPS As Int32 = 0
        Dim lMaxHullSize As Int32 = 0
        Dim lHullAvail As Int32 = 0
        Dim lPower As Int32 = 0
        TechBuilderComputer.GetTypeValues(lHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lPower)

        lPower *= 4

        If blResCost = 0 Then tpResCredits.MaxValue = 100000000 Else tpResCredits.MaxValue = Math.Min(Int64.MaxValue, Math.Max(tpResCredits.MaxValue, CLng(blResCost * 1.5)))
        If blResTime = 0 Then tpResTime.MaxValue = 100000000 Else tpResTime.MaxValue = Math.Min(Int64.MaxValue, Math.Max(tpResTime.MaxValue, CLng(blResTime * 1.5)))
        If blProdCost = 0 Then tpProdCredits.MaxValue = 100000000 Else tpProdCredits.MaxValue = Math.Min(Int64.MaxValue, Math.Max(tpProdCredits.MaxValue, CLng(blProdCost * 1.5)))
        If blProdTime = 0 Then tpProdTime.MaxValue = 100000000 Else tpProdTime.MaxValue = Math.Min(Int64.MaxValue, Math.Max(tpProdTime.MaxValue, CLng(blProdTime * 1.5)))
        If lHullReq > tpHullRequired.MaxValue Then tpHullRequired.MaxValue = lHullReq
        If lHullReq < lMaxHullSize Then tpHullRequired.MaxValue = lMaxHullSize

        If lPowerReq > tpPowerRequired.MaxValue Then tpPowerRequired.MaxValue = lPowerReq
        If lPowerReq < lPower Then tpPowerRequired.MaxValue = lPower

        tpHullRequired.PropertyValue = lHullReq
        tpPowerRequired.PropertyValue = lPowerReq
        tpResCredits.PropertyValue = blResCost
        tpResTime.PropertyValue = blResTime
        tpProdCredits.PropertyValue = blProdCost
        tpProdTime.PropertyValue = blProdTime
        tpColonists.PropertyValue = lColonists
        tpEnlisted.PropertyValue = lEnlisted
        tpOfficers.PropertyValue = lOfficers

        Dim lTotalCrew As Int32 = lColonists + lEnlisted + lOfficers
        lTotalCrew *= goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBetterHullToResidence)
        lblCrewHull.Caption = "Crew will take an additional " & lTotalCrew & " hull"

        mbIgnoreEvents = False
    End Sub
    Public Sub GetLockedValues(ByRef lHullReq As Int32, ByRef lPowerReq As Int32, ByRef blResCost As Int64, ByRef blResTime As Int64, ByRef blProdCost As Int64, ByRef blProdTime As Int64, ByRef lColonists As Int32, ByRef lEnlisted As Int32, ByRef lOfficers As Int32)
        If tpHullRequired.PropertyLocked = True Then lHullReq = CInt(tpHullRequired.PropertyValue) Else lHullReq = -1
        If tpPowerRequired.PropertyLocked = True Then lPowerReq = CInt(tpPowerRequired.PropertyValue) Else lPowerReq = -1
        If tpResCredits.PropertyLocked = True Then blResCost = tpResCredits.PropertyValue Else blResCost = -1
        If tpResTime.PropertyLocked = True Then blResTime = tpResTime.PropertyValue Else blResTime = -1
        If tpProdCredits.PropertyLocked = True Then blProdCost = tpProdCredits.PropertyValue Else blProdCost = -1
        If tpProdTime.PropertyLocked = True Then blProdTime = tpProdTime.PropertyValue Else blProdTime = -1
        If tpColonists.PropertyLocked = True Then lColonists = CInt(tpColonists.PropertyValue) Else lColonists = -1
        If tpEnlisted.PropertyLocked = True Then lEnlisted = CInt(tpEnlisted.PropertyValue) Else lEnlisted = -1
        If tpOfficers.PropertyLocked = True Then lOfficers = CInt(tpOfficers.PropertyValue) Else lOfficers = -1
    End Sub
    Public Sub SetPreConfigValues(ByVal lHullReq As Int32, ByVal lPowerReq As Int32, ByVal blResCost As Int64, ByVal blResTime As Int64, ByVal blProdCost As Int64, ByVal blProdTime As Int64, ByVal lColonists As Int32, ByVal lEnlisted As Int32, ByVal lOfficers As Int32)
        If lHullReq = -1 Then
            tpHullRequired.PropertyLocked = False
        Else
            tpHullRequired.PropertyLocked = True
            tpHullRequired.PropertyValue = lHullReq
        End If
        If lPowerReq = -1 Then
            tpPowerRequired.PropertyLocked = False
        Else
            tpPowerRequired.PropertyLocked = True
            tpPowerRequired.PropertyValue = lPowerReq
        End If
        If blResCost = -1 Then
            tpResCredits.PropertyLocked = False
        Else
            tpResCredits.PropertyLocked = True
            tpResCredits.PropertyValue = blResCost
        End If
        If blResTime = -1 Then
            tpResTime.PropertyLocked = False
        Else
            tpResTime.PropertyLocked = True
            tpResTime.PropertyValue = blResTime
        End If
        If blProdCost = -1 Then
            tpProdCredits.PropertyLocked = False
        Else
            tpProdCredits.PropertyLocked = True
            tpProdCredits.PropertyValue = blProdCost
        End If
        If blProdTime = -1 Then
            tpProdTime.PropertyLocked = False
        Else
            tpProdTime.PropertyLocked = True
            tpProdTime.PropertyValue = blProdTime
        End If
        If lColonists = -1 Then
            tpColonists.PropertyLocked = False
        Else
            tpColonists.PropertyLocked = True
            tpColonists.PropertyValue = lColonists
        End If
        If lEnlisted = -1 Then
            tpEnlisted.PropertyLocked = False
        Else
            tpEnlisted.PropertyLocked = True
            tpEnlisted.PropertyValue = lEnlisted
        End If
        If lOfficers = -1 Then
            tpOfficers.PropertyLocked = False
        Else
            tpOfficers.PropertyLocked = True
            tpOfficers.PropertyValue = lOfficers
        End If
    End Sub

    Public Sub SetMinValsFromCalcs(ByVal lCalcHull As Int32, ByVal lCalcPower As Int32, ByVal blCalcResCost As Int64, ByVal blCalcProdCost As Int64, ByVal lCalcColonists As Int32, ByVal lCalcEnlisted As Int32, ByVal lCalcOfficers As Int32)
        tpHullRequired.MinValue = lCalcHull \ 2
        tpPowerRequired.MinValue = lCalcPower \ 2
        tpResCredits.MinValue = blCalcResCost \ 2
        tpProdCredits.MinValue = blCalcProdCost \ 2
        tpColonists.MinValue = Math.Min(lCalcColonists, 1)
        tpEnlisted.MinValue = Math.Min(lCalcEnlisted, 1)
        tpOfficers.MinValue = Math.Min(lCalcOfficers, 1)
    End Sub

    Private Sub tp_PropValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp) Handles tpHullRequired.PropertyValueChanged, tpPowerRequired.PropertyValueChanged, tpProdCredits.PropertyValueChanged, tpProdTime.PropertyValueChanged, tpResCredits.PropertyValueChanged, tpResTime.PropertyValueChanged, tpColonists.PropertyValueChanged, tpEnlisted.PropertyValueChanged, tpOfficers.PropertyValueChanged
        If mbIgnoreEvents = True Then Return
        If sPropName.ToUpper = "RESEARCH COST:" OrElse sPropName.ToUpper = "RESEARCH TIME:" Then
            If NewTutorialManager.TutorialOn = True Then
                Try
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eTechBuilderCostValueChange, CInt(ctl.PropertyValue), -1, -1, "")
                Catch
                End Try
            End If
        End If
        RaiseEvent ValueChanged()
    End Sub
    Private Sub tp_LockBoxDoubleClick(ByVal sName As String, ByRef oCTL As ctlTechProp) Handles tpHullRequired.LockBoxDoubleClick, tpPowerRequired.LockBoxDoubleClick, tpProdCredits.LockBoxDoubleClick, tpProdTime.LockBoxDoubleClick, tpResCredits.LockBoxDoubleClick, tpResTime.LockBoxDoubleClick, tpColonists.LockBoxDoubleClick, tpEnlisted.LockBoxDoubleClick, tpOfficers.LockBoxDoubleClick
        Select Case oCTL.ControlName
            Case tpPowerRequired.ControlName
                oCTL.PropertyValue = mlNextLvlPower
            Case tpHullRequired.ControlName
                oCTL.PropertyValue = mlNextLvlHull
            Case tpProdCredits.ControlName
                oCTL.PropertyValue = mblNextProdCost
            Case tpProdTime.ControlName
                oCTL.PropertyValue = mblNextProdTime
            Case tpResCredits.ControlName
                oCTL.PropertyValue = mblNextResCost
            Case tpResTime.ControlName
                oCTL.PropertyValue = mblNextResTime
            Case tpColonists.ControlName
                oCTL.PropertyValue = mlNextLvlColonist
            Case tpEnlisted.ControlName
                oCTL.PropertyValue = mlNextLvlEnlisted
            Case tpOfficers.ControlName
                oCTL.PropertyValue = mlNextLvlOfficers
        End Select
        oCTL.PropertyLocked = True
    End Sub


    Public Sub SetAsHullBuilder()
        'power, restime, rescost
        tpHullRequired.Visible = False
        tpProdCredits.Visible = False
        tpProdTime.Visible = False
        tpColonists.Visible = False
        tpEnlisted.Visible = False
        tpOfficers.Visible = False
        lblCrewHull.Visible = False

        cboResearchFac.Width = 135
        cboProductionFac.Width = 135
        lblProductionFac.Left = cboResearchFac.Left + cboResearchFac.Width + 5
        cboProductionFac.Left = lblProductionFac.Left + lblProductionFac.Width

        tpPowerRequired.Top = lblResearchFac.Top + 20
        tpResCredits.Top = tpPowerRequired.Top + 20
        tpResTime.Top = tpResCredits.Top + 20
        Me.Height = tpResTime.Top + tpResTime.Height + 3

        Me.BorderLineWidth = 2
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub
    Public Sub SetAsArmorBuilder()
        tpPowerRequired.Visible = False
        tpHullRequired.Visible = False
        tpColonists.Visible = False
        tpEnlisted.Visible = False
        tpOfficers.Visible = False
        lblCrewHull.Visible = False

        'tpProdTime.SetAsArmorTimeIndicator(500)
        tpProdTime.SetAsTimeIndicator(500, True)

        tpPowerRequired.MaxValue = 0 : tpPowerRequired.MinValue = 0 : tpHullRequired.MaxValue = 0 : tpHullRequired.MinValue = 0
        tpColonists.MaxValue = 0 : tpColonists.MinValue = 0 : tpEnlisted.MaxValue = 0 : tpEnlisted.MinValue = 0
        tpOfficers.MaxValue = 0 : tpOfficers.MinValue = 0
    End Sub

    Private Sub frmTechBuilderCost_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim pt As Point = Me.GetAbsolutePosition()
        Dim lTmpX As Int32 = lMouseX - pt.X
        Dim lTmpY As Int32 = lMouseY - pt.Y

        If mrcCostPayment.Contains(lTmpX, lTmpY) = True Then
            MyBase.moUILib.SetToolTip("Indicates the overall value of this component design. Being on the right of the scale indicates a harder design to accomplish.", lMouseX, lMouseY)
        End If
    End Sub

    
    Private Sub frmTechBuilderCost_OnNewFrameEnd() Handles Me.OnNewFrameEnd
        If goCurrentPlayer Is Nothing = False Then
            If goCurrentPlayer.yPlayerPhase <> 255 Then
                Dim oSysFont As New System.Drawing.Font("Microsoft Sans Serif", 20.0F, FontStyle.Bold, GraphicsUnit.Point, 0)
                Dim sText As String = "Research Speed Bonus: 100x"

                Dim rcTemp As Rectangle = BPFont.MeasureString(oSysFont, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter) 'moCPFlashFont.MeasureString(Nothing, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.Red)
                rcTemp.X = Me.GetAbsolutePosition.X
                rcTemp.Y = Me.GetAbsolutePosition.Y + Me.Height

                Dim clrVal As System.Drawing.Color = muSettings.InterfaceBorderColor

                BPFont.DrawText(oSysFont, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Right, clrVal)
            End If
        End If
    End Sub

    Private Sub frmTechBuilderCost_OnRenderEnd() Handles Me.OnRenderEnd
        If lblCrewHull Is Nothing = False Then
            mrcCostPayment.X = Me.Width - 185
            mrcCostPayment.Y = lblCrewHull.Top
            mrcCostPayment.Height = lblCrewHull.Height
            mrcCostPayment.Width = 180
        End If

        'ok, draw our CostPayment bar... fill the box with black
        MyBase.moUILib.DoAlphaBlendColorFill(mrcCostPayment, Color.Black, mrcCostPayment.Location)

        'draw the box border
        RenderBox(mrcCostPayment, 2, muSettings.InterfaceBorderColor)
        'now draw a line in the center... 2 if even number, 3 if odd
        Dim vPts(1) As Vector2
        If mrcCostPayment.Width Mod 2 = 0 Then
            'even - 2 lines
            Dim lMiddle As Int32 = mrcCostPayment.X + (mrcCostPayment.Width \ 2)
            vPts(0) = New Vector2(lMiddle, mrcCostPayment.Y + 1)
            vPts(1) = New Vector2(lMiddle, mrcCostPayment.Bottom)
            BPLine.DrawLine(1.5, True, vPts, muSettings.InterfaceBorderColor)
            vPts(0).X += 1 : vPts(1).X += 1
            BPLine.DrawLine(1.5, True, vPts, muSettings.InterfaceBorderColor)
        Else
            'odd
            Dim lMiddle As Int32 = mrcCostPayment.X + (mrcCostPayment.Width \ 2)
            vPts(0) = New Vector2(lMiddle, mrcCostPayment.Y + 1)
            vPts(1) = New Vector2(lMiddle, mrcCostPayment.Bottom)
            BPLine.DrawLine(1.5, True, vPts, muSettings.InterfaceBorderColor)
            vPts(0).X -= 1 : vPts(1).X -= 1
            BPLine.DrawLine(1.5, True, vPts, muSettings.InterfaceBorderColor)
            vPts(0).X += 2 : vPts(1).X += 2
            BPLine.DrawLine(1.5, True, vPts, muSettings.InterfaceBorderColor)
        End If

        'Draw a circle at the point along the bar where the bill/payment exist
        '(((blMorePay + decPayment) / decBill) * 100).ToString("##0.##") & "%")
        Dim decPerc As Decimal = mdecBill / 1350000000D
        'Dim fPerc As Single = 1.0F
        'If mdecBill <> 0 Then
        '    fPerc = CSng(mdecPayment / mdecBill)
        'End If
        Dim lLeft As Int32
        If decPerc < 0 Then
            lLeft = mrcCostPayment.X
        ElseIf decPerc > 2 Then
            lLeft = mrcCostPayment.Right
        Else
            lLeft = CInt(mrcCostPayment.X + (decPerc * (mrcCostPayment.Width \ 2)))  'CInt(((fPerc - 0.9F) / 0.2F) * mrcCostPayment.Width)
        End If
        Dim rcDest As Rectangle = New Rectangle(lLeft - 6, mrcCostPayment.Y + 4, 12, 12)
        MyBase.moUILib.oDevice.RenderState.Lighting = False
        BPSprite.Draw2DOnce(MyBase.moUILib.oDevice, MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eSphere), rcDest, System.Drawing.Color.FromArgb(192, 255, 255, 255), 256, 256)
        MyBase.moUILib.oDevice.RenderState.Lighting = True
    End Sub

    Private Sub frmTechBuilderCost_OnResize() Handles Me.OnResize
        Dim lNewWidth As Int32 = Me.Width - 30
        If tpPowerRequired Is Nothing = False Then tpPowerRequired.Width = lNewWidth
        If tpHullRequired Is Nothing = False Then tpHullRequired.Width = lNewWidth
        If tpProdCredits Is Nothing = False Then tpProdCredits.Width = lNewWidth
        If tpProdTime Is Nothing = False Then tpProdTime.Width = lNewWidth
        If tpResCredits Is Nothing = False Then tpResCredits.Width = lNewWidth
        If tpResTime Is Nothing = False Then tpResTime.Width = lNewWidth
        If tpColonists Is Nothing = False Then tpColonists.Width = lNewWidth
        If tpEnlisted Is Nothing = False Then tpEnlisted.Width = lNewWidth
        If tpOfficers Is Nothing = False Then tpOfficers.Width = lNewWidth



    End Sub

    Private Sub cboProductionFac_ItemSelected(ByVal lItemIndex As Integer) Handles cboProductionFac.ItemSelected
        If cboProductionFac.ListIndex > -1 Then
            Dim lID As Int32 = cboProductionFac.ItemData(cboProductionFac.ListIndex)
            For X As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(X) = lID Then
                    Dim oDef As EntityDef = goEntityDefs(X)
                    If oDef Is Nothing = False AndAlso oDef.ObjTypeID = ObjectType.eFacilityDef Then
                        'ok, let's determine prod factor...
                        Dim lProdFactor As Int32 = oDef.ProdFactor
                        'that's per cycle... multiply by 30 for per second
                        lProdFactor *= 30
                        tpProdTime.SetAsTimeIndicator(lProdFactor, True)
                        tp_PropValueChanged("", Nothing)
                        Exit For
                    End If
                End If
            Next X
        End If
    End Sub

    Private Sub cboResearchFac_ItemSelected(ByVal lItemIndex As Integer) Handles cboResearchFac.ItemSelected
        If cboResearchFac.ListIndex > -1 Then
            Dim lID As Int32 = cboResearchFac.ItemData(cboResearchFac.ListIndex)
            For X As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(X) = lID Then
                    Dim oDef As EntityDef = goEntityDefs(X)
                    If oDef Is Nothing = False AndAlso oDef.ObjTypeID = ObjectType.eFacilityDef Then
                        'ok, let's determine prod factor...
                        Dim lProdFactor As Int32 = oDef.ProdFactor
                        'that's per cycle... multiply by 30 for per second
                        lProdFactor *= 30
                        tpResTime.SetAsTimeIndicator(lProdFactor, True)
                        tp_PropValueChanged("", Nothing)
                        Exit For
                    End If
                End If
            Next X
        End If
    End Sub

    Public Sub SetAndLockValues(ByVal oTech As Base_Tech)
        With oTech
            If .lSpecifiedColonists <> -1 Then
                tpColonists.PropertyValue = .lSpecifiedColonists
                tpColonists.PropertyLocked = True
            End If
            If .lSpecifiedEnlisted <> -1 Then
                tpEnlisted.PropertyValue = .lSpecifiedEnlisted
                tpEnlisted.PropertyLocked = True
            End If
            If .lSpecifiedHull <> -1 Then
                tpHullRequired.PropertyValue = .lSpecifiedHull
                tpHullRequired.PropertyLocked = True
            End If
            If .lSpecifiedOfficers <> -1 Then
                tpOfficers.PropertyValue = .lSpecifiedOfficers
                tpOfficers.PropertyLocked = True
            End If
            If .lSpecifiedPower <> -1 Then
                tpPowerRequired.PropertyValue = .lSpecifiedPower
                tpPowerRequired.PropertyLocked = True
            End If
            If .blSpecifiedProdCost <> -1 Then
                tpProdCredits.PropertyValue = .blSpecifiedProdCost
                tpProdCredits.PropertyLocked = True
            End If
            If .blSpecifiedProdTime <> -1 Then
                tpProdTime.PropertyValue = .blSpecifiedProdTime
                tpProdTime.PropertyLocked = True
            End If
            If .blSpecifiedResCost <> -1 Then
                tpResCredits.PropertyValue = .blSpecifiedResCost
                tpResCredits.PropertyLocked = True
            End If
            If .blSpecifiedResTime <> -1 Then
                tpResTime.PropertyValue = .blSpecifiedResTime
                tpResTime.PropertyLocked = True
            End If
        End With
    End Sub

    Public Sub ResetAll()
        tpPowerRequired.PropertyLocked = False
        tpHullRequired.PropertyLocked = False
        tpProdCredits.PropertyLocked = False
        tpProdTime.PropertyLocked = False
        tpResCredits.PropertyLocked = False
        tpResTime.PropertyLocked = False
        tpColonists.PropertyLocked = False
        tpEnlisted.PropertyLocked = False
        tpOfficers.PropertyLocked = False
    End Sub
End Class
