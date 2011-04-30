Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmGuildBillboardBid
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblPickLoc As UILabel
    Private txtDetails As UITextBox
    Private lblLocDetails As UILabel
    Private fraBBSel As UIWindow
    Private lblBillboardSelection As UILabel
    Private fraPlacement As UIWindow
    Private lblBidDetails As UILabel
    Private txtBidAmt As UITextBox
    Private lblDuration As UILabel
    Private txtDuration As UITextBox
    Private lblCurrentOwner As UILabel
    Private lblCurrentBid As UILabel

    Private WithEvents btnPlaceBid As UIButton
    Private WithEvents btnClose As UIButton
    Private WithEvents btnWithdraw As UIButton
    Private WithEvents lstLocations As UIListBox
    Private WithEvents optTL As UIOption
    Private WithEvents optBL As UIOption
    Private WithEvents optTR As UIOption
    Private WithEvents optRTM As UIOption
    Private WithEvents optRBM As UIOption
    Private WithEvents optBR As UIOption

    Private mbIgnoreOptClick As Boolean = False

    Private moAurelium As SolarSystem = Nothing
    Private mbRequested As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmGuildBillboardBid initial props
        With Me
            .ControlName = "frmGuildBillboardBid"
            .Left = 225
            .Top = 165
            .Width = 512
            .Height = 316
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 0
            .Width = 211
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Guild Billboard Placement"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 25
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblPickLoc initial props
        lblPickLoc = New UILabel(oUILib)
        With lblPickLoc
            .ControlName = "lblPickLoc"
            .Left = 5
            .Top = 25
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Pick a Location:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPickLoc, UIControl))

        'lstLocations initial props
        lstLocations = New UIListBox(oUILib)
        With lstLocations
            .ControlName = "lstLocations"
            .Left = 5
            .Top = 45
            .Width = 170
            .Height = 148
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstLocations, UIControl))

        'txtDetails initial props
        txtDetails = New UITextBox(oUILib)
        With txtDetails
            .ControlName = "txtDetails"
            .Left = 5
            .Top = 220
            .Width = 170
            .Height = 90
            .Enabled = True
            .Visible = True
            .Caption = "Info about the planet goes here"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .Locked = True
            .MultiLine = True
        End With
        Me.AddChild(CType(txtDetails, UIControl))

        'lblLocDetails initial props
        lblLocDetails = New UILabel(oUILib)
        With lblLocDetails
            .ControlName = "lblLocDetails"
            .Left = 5
            .Top = 200
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Location Details"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblLocDetails, UIControl))

        'fraBBSel initial props
        fraBBSel = New UIWindow(oUILib)
        With fraBBSel
            .ControlName = "fraBBSel"
            .Left = 185
            .Top = 45
            .Width = 200
            .Height = 150
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
        End With
        Me.AddChild(CType(fraBBSel, UIControl))

        'lblBillboardSelection initial props
        lblBillboardSelection = New UILabel(oUILib)
        With lblBillboardSelection
            .ControlName = "lblBillboardSelection"
            .Left = 185
            .Top = 25
            .Width = 121
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Billboard Selection:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblBillboardSelection, UIControl))

        'optTL initial props
        optTL = New UIOption(oUILib)
        With optTL
            .ControlName = "optTL"
            .Left = 400
            .Top = 40
            .Width = 65
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Top Left"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optTL, UIControl))

        'optBL initial props
        optBL = New UIOption(oUILib)
        With optBL
            .ControlName = "optBL"
            .Left = 400
            .Top = 60
            .Width = 82
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Bottom Left"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optBL, UIControl))

        'optTR initial props
        optTR = New UIOption(oUILib)
        With optTR
            .ControlName = "optTR"
            .Left = 400
            .Top = 80
            .Width = 75
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Top Right"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optTR, UIControl))

        'optRTM initial props
        optRTM = New UIOption(oUILib)
        With optRTM
            .ControlName = "optRTM"
            .Left = 400
            .Top = 100
            .Width = 88
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Right Upper"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optRTM, UIControl))

        'optRBM initial props
        optRBM = New UIOption(oUILib)
        With optRBM
            .ControlName = "optRBM"
            .Left = 400
            .Top = 120
            .Width = 86
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Right Lower"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optRBM, UIControl))

        'optBR initial props
        optBR = New UIOption(oUILib)
        With optBR
            .ControlName = "optBR"
            .Left = 400
            .Top = 140
            .Width = 92
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Bottom Right"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optBR, UIControl))

        'fraPlacement initial props
        fraPlacement = New UIWindow(oUILib)
        With fraPlacement
            .ControlName = "fraPlacement"
            .Left = 189
            .Top = 70
            .Width = 40
            .Height = 20
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceBorderColor
            .FullScreen = False
            .Moveable = False
        End With
        Me.AddChild(CType(fraPlacement, UIControl))

        'lblBidDetails initial props
        lblBidDetails = New UILabel(oUILib)
        With lblBidDetails
            .ControlName = "lblBidDetails"
            .Left = 185
            .Top = 255
            .Width = 130
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Bid Amount / minute:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblBidDetails, UIControl))

        'txtBidAmt initial props
        txtBidAmt = New UITextBox(oUILib)
        With txtBidAmt
            .ControlName = "txtBidAmt"
            .Left = 305
            .Top = 255
            .Width = 84
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 9
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
        End With
        Me.AddChild(CType(txtBidAmt, UIControl))

        'lblDuration initial props
        lblDuration = New UILabel(oUILib)
        With lblDuration
            .ControlName = "lblDuration"
            .Left = 185
            .Top = 285
            .Width = 109
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Duration (minutes):"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDuration, UIControl))

        'txtDuration initial props
        txtDuration = New UITextBox(oUILib)
        With txtDuration
            .ControlName = "txtDuration"
            .Left = 305
            .Top = 285
            .Width = 84
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 9
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
        End With
        Me.AddChild(CType(txtDuration, UIControl))

        'lblCurrentOwner initial props
        lblCurrentOwner = New UILabel(oUILib)
        With lblCurrentOwner
            .ControlName = "lblCurrentOwner"
            .Left = 185
            .Top = 200
            .Width = 314
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Current Owner: Unknown"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCurrentOwner, UIControl))

        'lblCurrentBid initial props
        lblCurrentBid = New UILabel(oUILib)
        With lblCurrentBid
            .ControlName = "lblCurrentBid"
            .Left = 185
            .Top = 225
            .Width = 203
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Current Bid: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCurrentBid, UIControl))

        'btnPlaceBid initial props
        btnPlaceBid = New UIButton(oUILib)
        With btnPlaceBid
            .ControlName = "btnPlaceBid"
            .Left = 400
            .Top = 230
            .Width = 100
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Place Bid"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnPlaceBid, UIControl))

        'btnWithdraw initial props
        btnWithdraw = New UIButton(oUILib)
        With btnWithdraw
            .ControlName = "btnWithdraw"
            .Left = 400
            .Top = 275
            .Width = 100
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Withdraw"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnWithdraw, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 25
            .Top = 2
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)


        optTL.Value = True
        FillLocationsList()
    End Sub

    Private Sub FillLocationsList()
        lstLocations.Clear()
        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X) Is Nothing = False AndAlso goGalaxy.moSystems(X).ObjectID = 36 Then
                moAurelium = goGalaxy.moSystems(X)
                If goGalaxy.moSystems(X).PlanetUB = -1 Then
                    If mbRequested = False Then
                        Dim yMsg(7) As Byte
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestSystemDetails).CopyTo(yMsg, lPos) : lPos += 2
                        goGalaxy.moSystems(X).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                        MyBase.moUILib.SendMsgToPrimary(yMsg)
                        mbRequested = True
                    End If
                Else

                    For Y As Int32 = 0 To goGalaxy.moSystems(X).PlanetUB
                        If goGalaxy.moSystems(X).moPlanets(Y) Is Nothing = False Then
                            lstLocations.AddItem(goGalaxy.moSystems(X).moPlanets(Y).PlanetName, False)
                            lstLocations.ItemData(lstLocations.NewIndex) = goGalaxy.moSystems(X).moPlanets(Y).ObjectID
                            lstLocations.ItemData2(lstLocations.NewIndex) = ObjectType.ePlanet
                        End If
                    Next Y

                End If
                Exit For
            End If
        Next X
    End Sub

    Private Sub btnPlaceBid_Click(ByVal sName As String) Handles btnPlaceBid.Click
        If lstLocations.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("Select a location to bid for first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        Dim lSlotID As Int32 = GetSelectedSlotID()
        If lSlotID = -1 Then Return

        Dim lID As Int32 = lstLocations.ItemData(lstLocations.ListIndex)
        If IsNumeric(txtBidAmt.Caption) = False OrElse txtBidAmt.Caption.Contains("-") OrElse txtBidAmt.Caption.Contains(".") OrElse txtBidAmt.Caption = "" Then
            MyBase.moUILib.AddNotification("Enter a valid bid amount.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        If IsNumeric(txtDuration.Caption) = False OrElse txtDuration.Caption.Contains("-") OrElse txtDuration.Caption.Contains(".") OrElse txtDuration.Caption = "" Then
            MyBase.moUILib.AddNotification("Enter a valid duration.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lAmt As Int32 = CInt(txtBidAmt.Caption)
        Dim lDuration As Int32 = CInt(txtDuration.Caption)

        Dim yMsg(17) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.ePlaceGuildBillboardBid).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSlotID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lAmt).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDuration).CopyTo(yMsg, lPos) : lPos += 4

        MyBase.moUILib.AddNotification("Bid Placed. If you win, you will be updated within 1 minute.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Private Sub btnWithdraw_Click(ByVal sName As String) Handles btnWithdraw.Click
        If lstLocations.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("Select a location to withdraw from first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        Dim lSlotID As Int32 = GetSelectedSlotID()
        If lSlotID = -1 Then Return

        Dim lID As Int32 = lstLocations.ItemData(lstLocations.ListIndex)

        Dim yMsg(17) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.ePlaceGuildBillboardBid).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSlotID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4

        MyBase.moUILib.AddNotification("Bid Withdrawn.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Private Sub lstLocations_ItemClick(ByVal lIndex As Integer) Handles lstLocations.ItemClick
        If moAurelium Is Nothing = False AndAlso lIndex > -1 Then
            Dim lID As Int32 = lstLocations.ItemData(lIndex)

            Dim yMsg(7) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildBillboards).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(ObjectType.ePlanet).CopyTo(yMsg, lPos) : lPos += 2
            MyBase.moUILib.SendMsgToPrimary(yMsg)

            For X As Int32 = 0 To moAurelium.PlanetUB
                If moAurelium.moPlanets(X) Is Nothing = False AndAlso moAurelium.moPlanets(X).ObjectID = lID Then
                    Dim oSB As New System.Text.StringBuilder

                    With moAurelium.moPlanets(X)
                        Select Case .PlanetSizeID
                            Case 0
                                oSB.AppendLine("Size Class: Tiny")
                            Case 1
                                oSB.AppendLine("Size Class: Small")
                            Case 2
                                oSB.AppendLine("Size Class: Medium")
                            Case 3
                                oSB.AppendLine("Size Class: Large")
                            Case 4
                                oSB.AppendLine("Size Class: Huge")
                        End Select

                        Select Case .MapTypeID
                            Case PlanetType.eAcidic
                                oSB.AppendLine("Classification: Acidic")
                            Case PlanetType.eAdaptable
                                oSB.AppendLine("Classification: Adaptable")
                            Case PlanetType.eBarren
                                oSB.AppendLine("Classification: Barren")
                            Case PlanetType.eDesert
                                oSB.AppendLine("Classification: Desert")
                            Case PlanetType.eGasGiant
                                oSB.AppendLine("Classification: Gas Giant")
                            Case PlanetType.eGeoPlastic
                                oSB.AppendLine("Classification: Lava")
                            Case PlanetType.eSuperGiant
                                oSB.AppendLine("Classification: Super Giant")
                            Case PlanetType.eTerran
                                oSB.AppendLine("Classification: Terran")
                            Case PlanetType.eTundra
                                oSB.AppendLine("Classification: Frozen")
                            Case PlanetType.eWaterWorld
                                oSB.AppendLine("Classification: Oceanic")
                        End Select
 
                    End With
                    txtDetails.Caption = oSB.ToString
                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub PlacePlacementBox()
        Dim lLeft As Int32
        Dim lTop As Int32

        Dim lHalf As Int32 = fraBBSel.Height \ 2

        If optTL.Value = True Then
            lLeft = fraBBSel.Left + 3
            lTop = fraBBSel.Top + lHalf - (fraPlacement.Height * 3)
        ElseIf optBL.Value = True Then
            lLeft = fraBBSel.Left + 3
            lTop = fraBBSel.Top + lHalf + (fraPlacement.Height * 2)
        ElseIf optTR.Value = True Then
            lLeft = fraBBSel.Left + fraBBSel.Width - fraPlacement.Width - 3
            lTop = fraBBSel.Top + lHalf - (fraPlacement.Height * 3)
        ElseIf optRTM.Value = True Then
            lLeft = fraBBSel.Left + fraBBSel.Width - fraPlacement.Width - 3
            lTop = fraBBSel.Top + lHalf - CInt(fraPlacement.Height * 1.5)
        ElseIf optRBM.Value = True Then
            lLeft = fraBBSel.Left + fraBBSel.Width - fraPlacement.Width - 3
            lTop = fraBBSel.Top + lHalf
        Else
            'bottom right
            lLeft = fraBBSel.Left + fraBBSel.Width - fraPlacement.Width - 3
            lTop = fraBBSel.Top + lHalf + (fraPlacement.Height * 2)
        End If

        fraPlacement.Left = lLeft
        fraPlacement.Top = lTop
    End Sub

    Private Function GetSelectedSlotID() As Int32
        Dim lSlotID As Int32 = -1
        If optTL.Value = True Then
            lSlotID = 0
        ElseIf optBL.Value = True Then
            lSlotID = 1
        ElseIf optTR.Value = True Then
            lSlotID = 2
        ElseIf optRTM.Value = True Then
            lSlotID = 3
        ElseIf optRBM.Value = True Then
            lSlotID = 4
        Else
            lSlotID = 5
        End If
        Return lSlotID
    End Function

#Region "  Option Button Clicks  "
    Private Sub optBL_Click() Handles optBL.Click
        'bottom left
        If mbIgnoreOptClick = True Then Return
        mbIgnoreOptClick = True
        optBL.Value = True
        optBR.Value = False
        optRBM.Value = False
        optRTM.Value = False
        optTL.Value = False
        optTR.Value = False
        mbIgnoreOptClick = False
        PlacePlacementBox()
    End Sub

    Private Sub optBR_Click() Handles optBR.Click
        If mbIgnoreOptClick = True Then Return
        mbIgnoreOptClick = True
        optBL.Value = False
        optBR.Value = True
        optRBM.Value = False
        optRTM.Value = False
        optTL.Value = False
        optTR.Value = False
        mbIgnoreOptClick = False
        PlacePlacementBox()
    End Sub

    Private Sub optRBM_Click() Handles optRBM.Click
        If mbIgnoreOptClick = True Then Return
        mbIgnoreOptClick = True
        optBL.Value = False
        optBR.Value = False
        optRBM.Value = True
        optRTM.Value = False
        optTL.Value = False
        optTR.Value = False
        mbIgnoreOptClick = False
        PlacePlacementBox()
    End Sub

    Private Sub optRTM_Click() Handles optRTM.Click
        If mbIgnoreOptClick = True Then Return
        mbIgnoreOptClick = True
        optBL.Value = False
        optBR.Value = False
        optRBM.Value = False
        optRTM.Value = True
        optTL.Value = False
        optTR.Value = False
        mbIgnoreOptClick = False
        PlacePlacementBox()
    End Sub

    Private Sub optTL_Click() Handles optTL.Click
        If mbIgnoreOptClick = True Then Return
        mbIgnoreOptClick = True
        optBL.Value = False
        optBR.Value = False
        optRBM.Value = False
        optRTM.Value = False
        optTL.Value = True
        optTR.Value = False
        mbIgnoreOptClick = False
        PlacePlacementBox()
    End Sub

    Private Sub optTR_Click() Handles optTR.Click
        If mbIgnoreOptClick = True Then Return
        mbIgnoreOptClick = True
        optBL.Value = False
        optBR.Value = False
        optRBM.Value = False
        optRTM.Value = False
        optTL.Value = False
        optTR.Value = True
        mbIgnoreOptClick = False
        PlacePlacementBox()
    End Sub
#End Region

    Private Sub frmGuildBillboardBid_OnNewFrame() Handles Me.OnNewFrame
        If lstLocations.ListCount <> moAurelium.PlanetUB + 1 Then
            FillLocationsList()
        End If

        If lstLocations.ListIndex > -1 AndAlso moAurelium Is Nothing = False Then
            Dim lID As Int32 = lstLocations.ItemData(lstLocations.ListIndex)
            For X As Int32 = 0 To moAurelium.PlanetUB
                If moAurelium.moPlanets(X) Is Nothing = False AndAlso moAurelium.moPlanets(X).ObjectID = lID Then

                    Dim lSlotID As Int32 = GetSelectedSlotID()

                    With moAurelium.moPlanets(X).uBillboards(lSlotID)
                        Dim sText As String = "Current Owner: " & GetCacheObjectValue(.lGuildID, ObjectType.eGuild)
                        If lblCurrentOwner.Caption <> sText Then lblCurrentOwner.Caption = sText
                        sText = "Current Bid: " & .BidAmount.ToString("#,##0")
                        If lblCurrentBid.Caption <> sText Then lblCurrentBid.Caption = sText
                    End With

                    Exit For
                End If
            Next X
        End If

    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub
End Class