Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmMineralBid
    Inherits UIWindow


    Private Structure PlayerBid
        Public lPlayerID As Int32
        Public lBidAmt As Int32
        Public lPrevQtyRcvd As Int32
    End Structure

    Private muBids(3) As PlayerBid

    Private lblMineral As UILabel
    Private lblCurrentHigh As UILabel
    Private lblMinimum As UILabel
    Private lblLastResults As UILabel
    Private lblFirstAmt As UILabel
    Private lblFirstPlayer As UILabel
    Private lblSecondAmt As UILabel
    Private lblSecondPlayer As UILabel
    Private lblThirdAmt As UILabel
    Private lblFourthAmt As UILabel
    Private lblThirdPerson As UILabel
    Private lblFourthPlayer As UILabel
    Private lblYourBid As UILabel
    Private lblProdRate As UILabel
    Private WithEvents txtBid As UITextBox
    Private lblQuantity As UILabel
    Private WithEvents txtQuantity As UITextBox

    Private btnSubmit As UIButton
    Private btnWithdraw As UIButton
    Private btnClose As UIButton

    Private mlCacheID As Int32 = -1
    Private mlFacilityID As Int32 = -1
    Private mlLastRequestTime As Int32 = -1
    Private mlMinBidAmt As Int32 = 0
    Private mlProductionRate As Int32 = 0

    Private mbEdittingQuantity As Boolean = False
    Private mbEdittingBid As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmMineralBid initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eMiningBid
            .ControlName = "frmMineralBid"
            .Left = 160
            .Top = 111
            .Width = 256
            .Height = 256
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
        End With

        'lblMineral initial props
        lblMineral = New UILabel(oUILib)
        With lblMineral
            .ControlName = "lblMineral"
            .Left = 5
            .Top = 5
            .Width = 220
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Product: Nitnesese"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMineral, UIControl))

        'lblCurrentHigh initial props
        lblCurrentHigh = New UILabel(oUILib)
        With lblCurrentHigh
            .ControlName = "lblCurrentHigh"
            .Left = 5
            .Top = 25
            .Width = 250
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Current High Bid: 5000000"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCurrentHigh, UIControl))

        'lblMinimum initial props
        lblMinimum = New UILabel(oUILib)
        With lblMinimum
            .ControlName = "lblMinimum"
            .Left = 5
            .Top = 170
            .Width = 250
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Minimum Bid: 25"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMinimum, UIControl))

        'lblProdRate initial props
        lblProdRate = New UILabel(oUILib)
        With lblProdRate
            .ControlName = "lblProdRate"
            .Left = 5
            .Top = 45
            .Width = 250
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Production: "
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "This is the rate at which minerals are being pulled from the ground." & vbCrLf & _
                           "This number is based on the mining facility's production rate and" & vbCrLf & _
                           "will fluctuate based on the owning player's colonial morale as the" & vbCrLf & _
                           "miners who work at the mine live in the owning player's colony."
        End With
        Me.AddChild(CType(lblProdRate, UIControl))

        'lblLastResults initial props
        lblLastResults = New UILabel(oUILib)
        With lblLastResults
            .ControlName = "lblLastResults"
            .Left = 5
            .Top = 65
            .Width = 150
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Last Bid Results:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblLastResults, UIControl))

        'lblFirstAmt initial props
        lblFirstAmt = New UILabel(oUILib)
        With lblFirstAmt
            .ControlName = "lblFirstAmt"
            .Left = 20
            .Top = 85
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "240 for 25"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblFirstAmt, UIControl))

        'lblFirstPlayer initial props
        lblFirstPlayer = New UILabel(oUILib)
        With lblFirstPlayer
            .ControlName = "lblFirstPlayer"
            .Left = 120
            .Top = 85
            .Width = 130
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Enoch Dagor"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblFirstPlayer, UIControl))

        'lblSecondAmt initial props
        lblSecondAmt = New UILabel(oUILib)
        With lblSecondAmt
            .ControlName = "lblSecondAmt"
            .Left = 20
            .Top = 105
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "240 for 19"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSecondAmt, UIControl))

        'lblSecondPlayer initial props
        lblSecondPlayer = New UILabel(oUILib)
        With lblSecondPlayer
            .ControlName = "lblSecondPlayer"
            .Left = 120
            .Top = 105
            .Width = 130
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Csaj"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSecondPlayer, UIControl))

        'lblThirdAmt initial props
        lblThirdAmt = New UILabel(oUILib)
        With lblThirdAmt
            .ControlName = "lblThirdAmt"
            .Left = 20
            .Top = 125
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "240 for 18"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblThirdAmt, UIControl))

        'lblFourthAmt initial props
        lblFourthAmt = New UILabel(oUILib)
        With lblFourthAmt
            .ControlName = "lblFourthAmt"
            .Left = 20
            .Top = 145
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "240 for 14"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblFourthAmt, UIControl))

        'lblThirdPerson initial props
        lblThirdPerson = New UILabel(oUILib)
        With lblThirdPerson
            .ControlName = "lblThirdPerson"
            .Left = 120
            .Top = 125
            .Width = 130
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Aurelius"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblThirdPerson, UIControl))

        'lblFourthPlayer initial props
        lblFourthPlayer = New UILabel(oUILib)
        With lblFourthPlayer
            .ControlName = "lblFourthPlayer"
            .Left = 120
            .Top = 145
            .Width = 130
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Norma Cenva"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblFourthPlayer, UIControl))

        'lblYourBid initial props
        lblYourBid = New UILabel(oUILib)
        With lblYourBid
            .ControlName = "lblYourBid"
            .Left = 5
            .Top = 195
            .Width = 55
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Your Bid:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblYourBid, UIControl))

        'txtBid initial props
        txtBid = New UITextBox(oUILib)
        With txtBid
            .ControlName = "txtBid"
            .Left = 65
            .Top = 195
            .Width = 60
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "18"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 8
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
        End With
        Me.AddChild(CType(txtBid, UIControl))

        'lblQuantity initial props
        lblQuantity = New UILabel(oUILib)
        With lblQuantity
            .ControlName = "lblQuantity"
            .Left = 130
            .Top = 195
            .Width = 55
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Quantity:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblQuantity, UIControl))

        'txtQuantity initial props
        txtQuantity = New UITextBox(oUILib)
        With txtQuantity
            .ControlName = "txtQuantity"
            .Left = 185
            .Top = 195
            .Width = 63
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "24760120"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 8
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
        End With
        Me.AddChild(CType(txtQuantity, UIControl))

        'btnSubmit initial props
        btnSubmit = New UIButton(oUILib)
        With btnSubmit
            .ControlName = "btnSubmit"
            .Left = 15
            .Top = 225
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Submit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSubmit, UIControl))

        'btnWithdraw initial props
        btnWithdraw = New UIButton(oUILib)
        With btnWithdraw
            .ControlName = "btnWithdraw"
            .Left = 141
            .Top = 225
            .Width = 100
            .Height = 24
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
            .Left = 231
            .Top = 2
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

        AddHandler btnClose.Click, AddressOf btnClose_Click
        AddHandler btnSubmit.Click, AddressOf btnSubmit_Click
        AddHandler btnWithdraw.Click, AddressOf btnWithdraw_Click

        Dim sRemain As String = "gets 25% of the production of the mine." & vbCrLf & "Any remainder is sold to the first ranked bidder."
        Dim sTemp As String = "First ranked bid " & sRemain
        lblFirstAmt.ToolTipText = sTemp : lblFirstPlayer.ToolTipText = sTemp
        sRemain = "gets 25% of the production of the mine, if available." & vbCrLf & "Any remainder is sold to the first ranked bidder."
        sTemp = "Second ranked bid " & sRemain
        lblSecondAmt.ToolTipText = sTemp : lblSecondPlayer.ToolTipText = sTemp
        sTemp = "Third ranked bid " & sRemain
        lblThirdAmt.ToolTipText = sTemp : lblThirdPerson.ToolTipText = sTemp
        sTemp = "Fourth ranked bid " & sRemain
        lblFourthAmt.ToolTipText = sTemp : lblFourthPlayer.ToolTipText = sTemp

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Public Sub SetIDs(ByVal lCacheID As Int32, ByVal lFacID As Int32)
        mlCacheID = lCacheID
        mlFacilityID = lFacID

        lblMineral.Caption = ""
        lblCurrentHigh.Caption = "Requesting Data..."
        lblMinimum.Caption = "Requesting Data..."
        lblFirstAmt.Caption = "Requesting..."
        lblFirstPlayer.Caption = ""
        lblSecondAmt.Caption = "Requesting..."
        lblSecondPlayer.Caption = ""
        lblThirdAmt.Caption = "Requesting..."
        lblThirdPerson.Caption = ""
        lblFourthAmt.Caption = "Requesting..."
        lblFourthPlayer.Caption = ""
        txtBid.Caption = ""
        txtQuantity.Caption = ""
        mbEdittingQuantity = False
        mbEdittingBid = False
        muBids = Nothing
        Me.IsDirty = True
        SendRequestForData()
    End Sub

    Private Sub SendRequestForData()
        Dim yMsg(9) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eMineralBid).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(mlFacilityID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(mlCacheID).CopyTo(yMsg, lPos) : lPos += 4
        MyBase.moUILib.SendMsgToPrimary(yMsg)
        mlLastRequestTime = glCurrentCycle
    End Sub

    Private Sub btnClose_Click(ByVal sName As String)
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnSubmit_Click(ByVal sName As String)

        If txtBid.Caption.Trim = "" OrElse txtBid.Caption.Contains(".") = True OrElse txtBid.Caption.Contains("-") OrElse IsNumeric(txtBid.Caption) = False Then
            MyBase.moUILib.AddNotification("You must enter a numeric bid amount that is greater than the minimum bid.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        If txtQuantity.Caption.Trim = "" OrElse txtQuantity.Caption.Contains(".") = True OrElse txtQuantity.Caption.Contains("-") OrElse IsNumeric(txtQuantity.Caption) = False Then
            MyBase.moUILib.AddNotification("You must enter a numeric quantity that is greater than zero.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lBidAmt As Int32 = CInt(txtBid.Caption)
        Dim lMaxQty As Int32 = CInt(txtQuantity.Caption)

        If lBidAmt < mlMinBidAmt Then
            MyBase.moUILib.AddNotification("You must enter a numeric bid amount that is greater than the minimum bid.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        SubmitBidSetting(lBidAmt, lMaxQty)

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eMiningBidSubmitButton, -1, -1, -1, "")
        End If
    End Sub

    Private Sub btnWithdraw_Click(ByVal sName As String)
        SubmitBidSetting(-1, -1)
    End Sub

    Private Sub SubmitBidSetting(ByVal lBidAmt As Int32, ByVal lMaxQty As Int32)
        Dim yMsg(17) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eSetMineralBid).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(mlCacheID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(mlFacilityID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lBidAmt).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMaxQty).CopyTo(yMsg, lPos) : lPos += 4

        MyBase.moUILib.SendMsgToPrimary(yMsg)
        MyBase.moUILib.AddNotification("Bid Settings Updated.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub frmMineralBid_OnNewFrame() Handles Me.OnNewFrame
        If glCurrentCycle - mlLastRequestTime > 90 Then
            SendRequestForData()
        End If

        If muBids Is Nothing = False Then
            For X As Int32 = 0 To 3
                Dim uAmt As UILabel = Nothing
                Dim uPlr As UILabel = Nothing

                Select Case X
                    Case 0
                        uAmt = lblFirstAmt
                        uPlr = lblFirstPlayer
                    Case 1
                        uAmt = lblSecondAmt
                        uPlr = lblSecondPlayer
                    Case 2
                        uAmt = lblThirdAmt
                        uPlr = lblThirdPerson
                    Case 3
                        uAmt = lblFourthAmt
                        uPlr = lblFourthPlayer
                End Select

                If muBids(X).lPlayerID > 0 Then
                    Dim sTemp As String = muBids(X).lPrevQtyRcvd & " for " & muBids(X).lBidAmt
                    If uAmt.Caption <> sTemp Then uAmt.Caption = sTemp
                    sTemp = GetCacheObjectValue(muBids(X).lPlayerID, ObjectType.ePlayer)
                    If uPlr.Caption <> sTemp Then uPlr.Caption = sTemp

                    If muBids(X).lPlayerID = glPlayerID Then
                        If uPlr.GetFont.Bold = False Then
                            uPlr.SetFont(New System.Drawing.Font(uPlr.GetFont, FontStyle.Bold))
                        End If
                    ElseIf uPlr.GetFont.Bold = True Then
                        uPlr.SetFont(New System.Drawing.Font(uPlr.GetFont, FontStyle.Regular))
                    End If
                Else
                    If uAmt.Caption <> "" Then uAmt.Caption = ""
                    If uPlr.Caption <> "" Then uPlr.Caption = ""
                End If
            Next X
        End If
        'mlProductionRate = 1000
        Dim sProdLine As String = "Production: " & mlProductionRate.ToString & " affected by morale"
        If lblProdRate.Caption <> sProdLine Then lblProdRate.Caption = sProdLine



        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return

        Try
            For X As Int32 = 0 To oEnvir.lCacheUB
                If oEnvir.lCacheIdx(X) = mlCacheID Then
                    Dim oCache As MineralCache = oEnvir.oCache(X)
                    If oCache Is Nothing = False Then
                        With oCache
                            Dim sTemp As String = ""
                            If .oMineral Is Nothing = False Then
                                sTemp &= .oMineral.MineralName
                            Else : sTemp = "Unknown Mineral"
                            End If

                            If .Concentration = Int32.MinValue OrElse .Quantity = Int32.MinValue Then
                                sTemp &= " ( Acquiring... ) "
                            Else
                                sTemp &= " ( " & .Concentration & " / " & .Quantity.ToString("#,###") & " ) "
                            End If

                            If lblMineral.Caption <> sTemp Then lblMineral.Caption = sTemp

                        End With
                    End If
                    Exit For
                End If
            Next X

        Catch
            'do nothing
        End Try

    End Sub

    Public Sub HandleMineralBidData(ByVal yData() As Byte)
        Dim lPos As Int32 = 2

        Dim lFacID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lCacheID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If lFacID <> mlFacilityID Then Return 'OrElse mlCacheID <> lCacheID Then Return

        mlMinBidAmt = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        mlProductionRate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return
        Dim lCacheQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Try
            For X As Int32 = 0 To oEnvir.lCacheUB
                If oEnvir.lCacheIdx(X) = mlCacheID Then
                    Dim oCache As MineralCache = oEnvir.oCache(X)
                    If oCache Is Nothing = False Then
                        With oCache
                            .Quantity = lCacheQty
                        End With
                    End If
                    Exit For
                End If
            Next X

        Catch
            'do nothing
        End Try


        Dim lCurrentHigh As Int32 = 0

        ReDim muBids(3)
        For X As Int32 = 0 To 3
            With muBids(X)
                .lBidAmt = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lCurrentHigh = Math.Max(lCurrentHigh, .lBidAmt)
                .lPrevQtyRcvd = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End With
        Next X

        Dim lPlayerBid As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4


        'Ok, we have all of our details, now we need to go ahead and populate our form
        'Private lblMineral As UILabel  - get this from the mineral cache

        lblCurrentHigh.Caption = "Current High Bid: " & lCurrentHigh.ToString("#,##0")
        lblMinimum.Caption = "Minimum Bid: " & mlMinBidAmt.ToString("#,##0")

        For X As Int32 = 0 To 3
            Dim uAmt As UILabel = Nothing
            Dim uPlr As UILabel = Nothing

            Select Case X
                Case 0
                    uAmt = lblFirstAmt
                    uPlr = lblFirstPlayer
                Case 1
                    uAmt = lblSecondAmt
                    uPlr = lblSecondPlayer
                Case 2
                    uAmt = lblThirdAmt
                    uPlr = lblThirdPerson
                Case 3
                    uAmt = lblFourthAmt
                    uPlr = lblFourthPlayer
            End Select

            If muBids(X).lPlayerID > 0 Then
                uAmt.Caption = muBids(X).lPrevQtyRcvd & " for " & muBids(X).lBidAmt
                uPlr.Caption = GetCacheObjectValue(muBids(X).lPlayerID, ObjectType.ePlayer)
            Else
                uAmt.Caption = ""
                uPlr.Caption = ""
            End If
        Next X

        If lPlayerBid <= 1 Then
            lPlayerBid = mlMinBidAmt
        End If
        If lPlayerQty <= 1 Then
            lPlayerQty = lCacheQty
        End If

        If mbEdittingBid = False Then
            txtBid.Caption = lPlayerBid.ToString
            mbEdittingBid = False
        End If
        If mbEdittingQuantity = False Then
            txtQuantity.Caption = lPlayerQty.ToString
            mbEdittingQuantity = False
        End If
        
    End Sub

    Private Sub txtBid_TextChanged() Handles txtBid.TextChanged
        mbEdittingBid = True
        If NewTutorialManager.TutorialOn = True Then
            Dim lVal As Int32 = CInt(Val(txtBid.Caption))
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eMiningBidAmountChange, lVal, -1, -1, "")
        End If
    End Sub

    Private Sub txtQuantity_TextChanged() Handles txtQuantity.TextChanged
        mbEdittingQuantity = True
        If NewTutorialManager.TutorialOn = True Then
            Dim lVal As Int32 = CInt(Val(txtQuantity.Caption))
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eMiningBidQuantityChange, lVal, -1, -1, "")
        End If
    End Sub

    Private Sub frmMineralBid_OnRender() Handles Me.OnRender
        Dim rcDest As Rectangle

        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return

        Dim lWidth As Int32 = lblFirstPlayer.Left + lblFirstPlayer.Width - lblFirstAmt.Left + 4
        Dim lHeight As Int32 = lblFirstPlayer.Height + 2

        rcDest = New Rectangle(lblFirstAmt.Left - 2, lblFirstAmt.Top - 1, lWidth, lHeight)
        MyBase.RenderRoundedBorder(rcDest, 2, muSettings.InterfaceBorderColor)
        rcDest.Y = lblSecondAmt.Top - 1
        MyBase.RenderRoundedBorder(rcDest, 2, muSettings.InterfaceBorderColor)
        rcDest.Y = lblThirdAmt.Top - 1
        MyBase.RenderRoundedBorder(rcDest, 2, muSettings.InterfaceBorderColor)
        rcDest.Y = lblFourthAmt.Top - 1
        MyBase.RenderRoundedBorder(rcDest, 2, muSettings.InterfaceBorderColor)
    End Sub

    Private Sub frmMineralBid_WindowMoved() Handles Me.WindowMoved
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmMining")
        If ofrm Is Nothing = False Then
            ofrm.Left = Me.Left - ofrm.Width
            ofrm.Top = Me.Top
        End If
    End Sub
End Class