Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmProdCost
    Inherits UIWindow

    Private WithEvents lblPCType As UILabel
    Private WithEvents txtProdCosts As UITextBox

    Private moProdCost As ProductionCost
    Private mlProdFactor As Int32
    Private mbResearch As Boolean
    Private mlHullUsed As Int32
    Private mlPowerUsed As Int32

    Private mblActualPoints As Int64 

    Public Sub New(ByRef oUILib As UILib, ByVal sProdCostWindowName As String, Optional ByVal bSubForm As Boolean = False)
        MyBase.New(oUILib)

        If sProdCostWindowName.Length = 0 Then sProdCostWindowName = "frmProdCost"

        'frmProdCost initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eProdCost
            .ControlName = sProdCostWindowName
            .Left = 158
            .Top = 115

            If MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth < 1280 Then
                .Width = 200
            Else
                .Width = 250
            End If
            If MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight < 1024 Then
                .Height = 180
            Else
                .Height = 230
            End If

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
        End With

        'lblPCType initial props
        lblPCType = New UILabel(oUILib)
        With lblPCType
            .ControlName = "lblPCType"
            .Left = 5
            .Top = 5
            .Width = 190
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Production Costs:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPCType, UIControl))

        'txtProdCosts initial props
        txtProdCosts = New UITextBox(oUILib)
        With txtProdCosts
            .ControlName = "txtProdCosts"
            .Left = 5
            .Top = 25
            .Width = Me.Width - .Left - 5
            .Height = Me.Height - .Top - 5
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
        Me.AddChild(CType(txtProdCosts, UIControl))

        If bSubForm = False Then
            MyBase.moUILib.RemoveWindow(Me.ControlName)
            MyBase.moUILib.AddWindow(Me)
        End If

    End Sub

    Public Sub SetFromProdCost(ByRef oProdCost As ProductionCost, ByVal lProdFactor As Int32, ByVal bResearch As Boolean, ByVal lHullUsed As Int32, ByVal lPowerUsed As Int32)
        moProdCost = oProdCost
        mlProdFactor = lProdFactor
        mbResearch = bResearch
        mlHullUsed = lHullUsed
        mlPowerUsed = lPowerUsed
        If moProdCost Is Nothing = False Then mblActualPoints = moProdCost.PointsRequired
        RequestTimeToDoHere()
        RefreshDisplay()
    End Sub

    Private Function GetMaxFactionBonus() As Single
        Dim fValue As Single = 1.0F

        Dim fMult As Single = 1.0F
        Select Case goCurrentPlayer.yPlayerTitle
            Case Player.PlayerRank.Emperor
                fMult = 0.85F
            Case Player.PlayerRank.King
                fMult = 0.76F
            Case Player.PlayerRank.Baron
                fMult = 0.81F
            Case Player.PlayerRank.Duke
                fMult = 0.96F
            Case Player.PlayerRank.Overseer
                fMult = 0.99F
            Case Player.PlayerRank.Governor
                fMult = 0.99F
        End Select
        For X As Int32 = 0 To 4
            fValue *= fMult
        Next X
        Return fValue
    End Function
    Private Function GetMaxOtherFactionBonus() As Single
        Dim fValue As Single = 1.0F
        Dim lMaxCnt As Int32 = 0
        If goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Magistrate Then
            lMaxCnt = 3
        ElseIf goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Governor Then
            lMaxCnt = 2
        ElseIf goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Overseer OrElse goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Duke Then
            lMaxCnt = 1
        End If
        For X As Int32 = 0 To lMaxCnt - 1
            fValue *= 0.5F
        Next X
        Return fValue
    End Function

    Public Sub SetFromFailureCode(ByVal yErrorCode As Byte)
        Me.lblPCType.Caption = "Design Flaw:"

        Dim sFinal As String = "The component design is impossible to comprehend."

        Select Case CType(yErrorCode, TechBuilderErrorReason)
            Case TechBuilderErrorReason.eBeamWeapon_CatalystCompositeKnowFail
                sFinal = "Insufficient knowledge of the Catalyst and Composite materials."
            Case TechBuilderErrorReason.eBeamWeapon_CompCompressZero
                'If goCurrentPlayer.PlayerKnowsProperty("COMPRESSIBILITY") = True Then 
                sFinal = "Composite Compressibility property is too low."
            Case TechBuilderErrorReason.eBeamWeapon_CompQuantumZero
                'If goCurrentPlayer.PlayerKnowsProperty("QUANTUM") = True Then 
                sFinal = "Composite Quantum property is too low."
            Case TechBuilderErrorReason.eBeamWeapon_CompReflectionZero
                'If goCurrentPlayer.PlayerKnowsProperty("REFLECTION") = True Then 
                sFinal = "Composite Reflection property is too low."
            Case TechBuilderErrorReason.eBeamWeapon_ConcReflectionKnowZero
                'If goCurrentPlayer.PlayerKnowsProperty("REFLECTION") = True Then 
                sFinal = "Concentrator Reflection property is not understood."
            Case TechBuilderErrorReason.eBeamWeapon_ConcReflectionZero
                'If goCurrentPlayer.PlayerKnowsProperty("REFLECTION") = True Then 
                sFinal = "Concentrator Reflection property is too low."
            Case TechBuilderErrorReason.eBeamWeapon_HousingDensityZero
                'If goCurrentPlayer.PlayerKnowsProperty("DENSITY") = True Then 
                sFinal = "Housing Density property is too low."
            Case TechBuilderErrorReason.eBeamWeapon_HousingThermCondZero
                'If goCurrentPlayer.PlayerKnowsProperty("THERMAL CONDUCTION") = True Then 
                sFinal = "Housing Thermal Conduction property is too low."
            Case TechBuilderErrorReason.eEngine_DriveBodyMagReactZero
                'If goCurrentPlayer.PlayerKnowsProperty("MAGNETIC REACTANCE") = True Then 
                sFinal = "Drive Body Magnetic Reaction property is too low."
            Case TechBuilderErrorReason.eEngine_DriveFrameHardnessZero
                'If goCurrentPlayer.PlayerKnowsProperty("HARDNESS") = True Then 
                sFinal = "Drive Frame Hardness property is too low."
            Case TechBuilderErrorReason.eEngine_DriveFrameSuperCZero
                'If goCurrentPlayer.PlayerKnowsProperty("SUPERCONDUCTIVE POINT") = True Then 
                sFinal = "Drive Frame Superconductive Point is too low."
            Case TechBuilderErrorReason.eEngine_DriveMeldChemReactZero
                'If goCurrentPlayer.PlayerKnowsProperty("CHEMICAL REACTANCE") = True Then 
                sFinal = "Drive Meld Chemical Reactance property is too low."
            Case TechBuilderErrorReason.eEngine_StructureBodyDensityZero
                'If goCurrentPlayer.PlayerKnowsProperty("DENSITY") = True Then 
                sFinal = "Structure Body Density property is too low."
            Case TechBuilderErrorReason.eEngine_StructureBodyMeltingPtZero
                'If goCurrentPlayer.PlayerKnowsProperty("MELTING POINT") = True Then 
                sFinal = "Structure Body Melting Point is too low."
            Case TechBuilderErrorReason.eEngine_StructureFrameHardnessZero
                'If goCurrentPlayer.PlayerKnowsProperty("HARDNESS") = True Then 
                sFinal = "Structure Frame Hardness property is too low."
            Case TechBuilderErrorReason.eEngine_StructureFrameMalleableZero
                'If goCurrentPlayer.PlayerKnowsProperty("MALLEABLE") = True Then 
                sFinal = "Structure Frame Malleable property is too low."
            Case TechBuilderErrorReason.eEngine_StructureMeldMeltingPtZero
                'If goCurrentPlayer.PlayerKnowsProperty("MELTING POINT") = True Then 
                sFinal = "Structure Meld Melting Point is too low."
            Case TechBuilderErrorReason.eHull_HardnessZero
                'If goCurrentPlayer.PlayerKnowsProperty("HARDNESS") = True Then 
                sFinal = "Hull material's Hardness property is too low."
            Case TechBuilderErrorReason.eShield_AccMagReactZero
                'If goCurrentPlayer.PlayerKnowsProperty("MAGNETIC REACTANCE") = True Then 
                sFinal = "Accelerator's Magnetic Reactance property is too low."
            Case TechBuilderErrorReason.eShield_AccQuantumZero
                'If goCurrentPlayer.PlayerKnowsProperty("QUANTUM") = True Then 
                sFinal = "Accelerator's Quantum property is too low."
            Case TechBuilderErrorReason.eShield_CasingThermCondZero
                'If goCurrentPlayer.PlayerKnowsProperty("THERMAL CONDUCTION") = True Then 
                sFinal = "Casing material's Thermal Conductance property is too low."
            Case TechBuilderErrorReason.eShield_CoilDensityZero
                'If goCurrentPlayer.PlayerKnowsProperty("DENSITY") = True Then 
                sFinal = "Coil material is not dense enough."
            Case TechBuilderErrorReason.eShield_CoilMagProdZero
                'If goCurrentPlayer.PlayerKnowsProperty("MAGNETIC PRODUCTION") = True Then 
                sFinal = "Coil material's Magnetic Production property is too low."
            Case TechBuilderErrorReason.ePulse_AccMagReactZero
                sFinal = "Accelerator material's Magnetic Reactance property is too low."
            Case TechBuilderErrorReason.ePulse_AccQuantumZero
                sFinal = "Accelerator material's Quantum property is too low."
            Case TechBuilderErrorReason.ePulse_AccSuperCZero
                sFinal = "Accelerator's Superconductive Point is too low."
            Case TechBuilderErrorReason.ePulse_CasingDensityZero
                sFinal = "Casing Material's density is too low."
            Case TechBuilderErrorReason.ePulse_CasingThermCondZero
                sFinal = "Casing material's Thermal Conductance is too low."
            Case TechBuilderErrorReason.ePulse_CoilDensityZero
                sFinal = "Coil material's Density is too low."
            Case TechBuilderErrorReason.ePulse_CoilMagReactZero
                sFinal = "Coil material's Magnetic Reactance is too low."
            Case TechBuilderErrorReason.ePulse_CoilSuperCZero
                sFinal = "Coil's Superconductive Point is too low."
            Case TechBuilderErrorReason.ePulse_CompressChamberDensityZero
                sFinal = "Compression Chamber's Density is too low."
            Case TechBuilderErrorReason.ePulse_CompressChmaberCompressExceedDensity
                sFinal = "Compression Chamber's Compression exceeds Density."
            Case TechBuilderErrorReason.ePulse_CoilMagProdExceedCoilMagReact
                sFinal = "Coil's Magnetic Production Exceeds Coil's Magnetic Reaction"
            Case TechBuilderErrorReason.eInvalidMaterialsSelected
                sFinal = "Invalid Materials selected."
            Case TechBuilderErrorReason.eInvalidValuesEntered
                sFinal = "Invalid values for properties were entered."
        End Select

        Me.txtProdCosts.Caption = sFinal
    End Sub

    Private Sub frmProdCost_OnResize() Handles Me.OnResize
        If txtProdCosts Is Nothing = False Then
            txtProdCosts.Height = Me.Height - txtProdCosts.Top - 5
            'txtProdCosts.Width = Me.Width - 10
            txtProdCosts.Width = Me.Width - txtProdCosts.Left - 5
        End If
    End Sub

    Public Sub SetFromWeaponDefResult(ByRef oWpnDef As WeaponDef, ByVal yType As WeaponClassType)

        Dim lWidth As Int32 = 250
        Dim lTxtWidth As Int32 = lWidth - 10

        If MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth < 1280 Then
            lWidth = 140
            lTxtWidth = 130
            Me.Height = 180
        Else
            Me.Height = 230
        End If

        Me.Width = lWidth

        txtProdCosts.Width = lTxtWidth
        txtProdCosts.Height = Me.Height - txtProdCosts.Top - 5
        lblPCType.Caption = "Weapon Statistics"

        Dim oSB As New System.Text.StringBuilder()
        If oWpnDef Is Nothing = False Then
            With oWpnDef
                Dim fROF As Single = Math.Max(1, .ROF_Delay) / 30.0F
                oSB.AppendLine("Rate of Fire: " & fROF.ToString("##0.##"))

                If yType = WeaponClassType.eMissile Then
                    oSB.AppendLine("Flight Time: " & (.Range / 30.0F).ToString("##0.##"))
                Else
                    oSB.AppendLine("Range: " & .Range)
                End If

                oSB.AppendLine("Accuracy: " & .Accuracy)
                oSB.AppendLine("Area Effect: " & .AOERange)
                oSB.AppendLine("Firepower Rating: " & .lFirePowerRating)

                Dim blTotalDmgs As Int64 = .PiercingMaxDmg
                blTotalDmgs += .ECMMaxDmg
                blTotalDmgs += .ImpactMaxDmg
                blTotalDmgs += .FlameMaxDmg
                blTotalDmgs += .ChemicalMaxDmg
                If (.yWeaponType < WeaponType.eMissile_Color_1 OrElse .yWeaponType > WeaponType.eMissile_Color_9) Then blTotalDmgs += .BeamMaxDmg

                Dim fDPS As Single = (blTotalDmgs / fROF)
                oSB.AppendLine()
                oSB.AppendLine("Damage Per Second: " & fDPS.ToString("#,##0.###"))
                oSB.AppendLine()

                oSB.AppendLine("DAMAGE")
                If (.yWeaponType < WeaponType.eMissile_Color_1 OrElse .yWeaponType > WeaponType.eMissile_Color_9) AndAlso (.BeamMaxDmg <> 0 OrElse .BeamMinDmg <> 0) Then oSB.AppendLine("  Beam: " & .BeamMinDmg & " - " & .BeamMaxDmg)
                If .ChemicalMinDmg <> 0 OrElse .ChemicalMaxDmg <> 0 Then oSB.AppendLine("  Chemical: " & .ChemicalMinDmg & " - " & .ChemicalMaxDmg)
                If .FlameMaxDmg <> 0 OrElse .FlameMinDmg <> 0 Then oSB.AppendLine("  Flame: " & .FlameMinDmg & " - " & .FlameMaxDmg)
                If .ImpactMaxDmg <> 0 OrElse .ImpactMinDmg <> 0 Then oSB.AppendLine("  Impact: " & .ImpactMinDmg & " - " & .ImpactMaxDmg)
                If .ECMMaxDmg <> 0 OrElse .ECMMinDmg <> 0 Then oSB.AppendLine("  Magnetic: " & .ECMMinDmg & " - " & .ECMMaxDmg)
                If .PiercingMaxDmg <> 0 OrElse .PiercingMinDmg <> 0 Then oSB.AppendLine("  Pierce: " & .PiercingMinDmg & " - " & .PiercingMaxDmg)
            End With
        End If

        txtProdCosts.Caption = oSB.ToString
    End Sub

    Public Sub SetTextDirect(ByVal sText As String, ByVal sTitle As String)
        Me.Width = 256
        Me.Height = 230

        txtProdCosts.Width = 246
        txtProdCosts.Height = Me.Height - txtProdCosts.Top - 5
        lblPCType.Caption = sTitle

        txtProdCosts.Caption = sText
    End Sub

    Private Function GetBestResearchTimeString() As String
        'full faction, best research 766,
        Dim lProdFactor As Int32 = 766
        Dim blPoints As Int64 = mblActualPoints
        'Dim fMult As Single = (GetMaxFactionBonus() * GetMaxOtherFactionBonus())
        'blPoints = CLng(blPoints * fMult)
        Dim blSecondsToFinish As Int64 = CLng(blPoints / (lProdFactor * 30))

        'Ok, format seconds to finish to minutes, hours, whatever
        Dim blHours As Int64 = blSecondsToFinish \ 3600L
        blSecondsToFinish -= (blHours * 3600L)
        Dim blMinutes As Int64 = blSecondsToFinish \ 60L
        blSecondsToFinish -= (blMinutes * 60L)
        If blSecondsToFinish <> 0 Then blMinutes += 1

        If blMinutes = 60 Then
            blMinutes = 0
            blHours += 1
        End If

        Dim sTimeString As String = ""
        If blHours > 24 Then
            Dim blDays As Int64 = blHours \ 24L
            sTimeString = blDays.ToString & " days"
            blHours -= (blDays * 24L)
        End If
        If blHours > 0 Then
            If sTimeString.Length > 0 Then sTimeString &= ", "
            sTimeString &= blHours.ToString & " hours"
        End If
        If blMinutes > 0 Then
            If sTimeString.Length > 0 Then sTimeString &= ", "
            sTimeString &= blMinutes.ToString & " minutes"
        End If
        If sTimeString = "" Then sTimeString = "< 1 minute"

        Return "Best Time to Research (at 100 morale, efficiency, and full faction bonus for your title and best research facility): " & vbCrLf & "  " & sTimeString
    End Function

    Private Sub RefreshDisplay()
        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder()

        If moProdCost Is Nothing = False Then
            With moProdCost
                oSB.AppendLine("Credits: " & .CreditCost.ToString("#,##0"))
                If .ColonistCost > 0 Then oSB.AppendLine("Colonists: " & .ColonistCost.ToString("#,##0"))
                If .EnlistedCost > 0 Then oSB.AppendLine("Enlisted: " & .EnlistedCost.ToString("#,##0"))
                If .OfficerCost > 0 Then oSB.AppendLine("Officers: " & .OfficerCost.ToString("#,##0"))

                If mbResearch = True Then

                    If mlProdFactor > 0 Then
                        'full faction, best research 766,
                        Dim blPoints As Int64 = mblActualPoints
                        'Dim fMult As Single = (GetMaxFactionBonus() * GetMaxOtherFactionBonus())
                        'blPoints = CLng(blPoints * fMult)
                        Dim blSecondsToFinish As Int64 = CLng(blPoints / (mlProdFactor * 30))

                        'Ok, format seconds to finish to minutes, hours, whatever
                        Dim blHours As Int64 = blSecondsToFinish \ 3600L
                        blSecondsToFinish -= (blHours * 3600L)
                        Dim blMinutes As Int64 = blSecondsToFinish \ 60L
                        blSecondsToFinish -= (blMinutes * 60L)
                        If blSecondsToFinish <> 0 Then blMinutes += 1

                        If blMinutes = 60 Then
                            blMinutes = 0
                            blHours += 1
                        End If

                        Dim sTimeString As String = ""
                        If blHours > 24 Then
                            Dim blDays As Int64 = blHours \ 24L
                            sTimeString = blDays.ToString & " days"
                            blHours -= (blDays * 24L)
                        End If
                        If blHours > 0 Then
                            If sTimeString.Length > 0 Then sTimeString &= ", "
                            sTimeString &= blHours.ToString & " hours"
                        End If
                        If blMinutes > 0 Then
                            If sTimeString.Length > 0 Then sTimeString &= ", "
                            sTimeString &= blMinutes.ToString & " minutes"
                        End If
                        If sTimeString = "" Then sTimeString = "< 1 minute"

                        oSB.AppendLine("Current time to research: " & vbCrLf & "  " & sTimeString)
                        oSB.AppendLine()
                        oSB.AppendLine(GetBestResearchTimeString())
                    Else
                        oSB.AppendLine("Time to Research Here: Unknown")
                    End If
                Else
                    'oSB.AppendLine("Production Complexity: " & (.PointsRequired \ 1000).ToString("#,##0"))
                    mlProdFactor = Math.Max(mlProdFactor, 1)
                    Dim blPoints As Int64 = mblActualPoints
                    Dim blSecondsToFinish As Int64 = CLng(blPoints / (mlProdFactor * 30))

                    'Ok, format seconds to finish to minutes, hours, whatever
                    Dim blHours As Int64 = blSecondsToFinish \ 3600L
                    blSecondsToFinish -= (blHours * 3600L)
                    Dim blMinutes As Int64 = blSecondsToFinish \ 60L
                    blSecondsToFinish -= (blMinutes * 60L)
                    If blSecondsToFinish <> 0 Then blMinutes += 1

                    If blMinutes = 60 Then
                        blMinutes = 0
                        blHours += 1
                    End If

                    Dim sTimeString As String = ""
                    If blHours > 24 Then
                        Dim blDays As Int64 = blHours \ 24L
                        sTimeString = blDays.ToString & " days"
                        blHours -= (blDays * 24L)
                    End If
                    If blHours > 0 Then
                        If sTimeString.Length > 0 Then sTimeString &= ", "
                        sTimeString &= blHours.ToString & " hours"
                    End If
                    If blMinutes > 0 Then
                        If sTimeString.Length > 0 Then sTimeString &= ", "
                        sTimeString &= blMinutes.ToString & " minutes"
                    End If
                    If sTimeString = "" Then sTimeString = "< 1 minute"

                    oSB.AppendLine("Time to produce here: " & vbCrLf & "  " & sTimeString)
                End If

                If mlHullUsed <> 0 Then
                    oSB.AppendLine("Hull Consumption: " & mlHullUsed.ToString("#,##0"))
                End If
                If mlPowerUsed <> 0 Then
                    oSB.AppendLine("Power Consumption: " & mlPowerUsed.ToString("#,##0"))
                End If

                'Minerals.
                If .ItemCostUB > -1 Then oSB.AppendLine(vbCrLf & "MATERIALS")
                For X As Int32 = 0 To .ItemCostUB
                    oSB.AppendLine(.ItemCosts(X).GetItemName & ": " & .ItemCosts(X).QuantityNeeded.ToString("#,##0"))
                Next X
            End With
        End If

        If mbResearch = True Then
            Me.lblPCType.Caption = "Research Costs:"
        End If

        Me.txtProdCosts.Caption = oSB.ToString()
        oSB = Nothing
    End Sub

    Public Sub HandleRequestTimeToDoHere(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        lPos += 7   'guid and type

        If mbResearch = True Then
            mblActualPoints = System.BitConverter.ToInt64(yData, lPos)
        End If
        lPos += 8

        mlProdFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        RefreshDisplay()
    End Sub

    Private Sub RequestTimeToDoHere()
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            Dim yMsg(16) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yMsg, 0)

            For X As Int32 = 0 To oEnvir.lEntityUB
                Try
                    If oEnvir.lEntityIdx(X) > -1 Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False Then
                            If oEntity.OwnerID = glPlayerID AndAlso oEntity.bSelected = True Then
                                oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
                                System.BitConverter.GetBytes(mblActualPoints).CopyTo(yMsg, 8)
                                If mbResearch = True Then yMsg(16) = 0 Else yMsg(16) = 1
                                MyBase.moUILib.SendMsgToPrimary(yMsg)
                                Exit For
                            End If
                        End If
                    End If
                Catch
                End Try
            Next X


        End If
    End Sub
End Class
