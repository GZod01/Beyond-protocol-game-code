Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmColonyStatsSmall
    Inherits UIWindow

    Private Const ml_REQUEST_DELAY As Int32 = 5000

    Private WithEvents lblExpand As UILabel
    Private lblPowerInd As UILabel
    Private lblPopInd As UILabel
    Private lblMoraleInd As UILabel
    Private shpPowerBack As UIWindow
    Private shpPowerFore As UIWindow
    Private shpPopHomeless As UIWindow
    Private shpMoraleBack As UIWindow
    Private shpPopPowerless As UIWindow
    Private shpPopGood As UIWindow
    Private shpMoraleFore As UIWindow
    Private lblPopGrowth As UILabel

    Private msw_Delay As Stopwatch

    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        Dim lTempTop As Int32
        Dim oWin As UIWindow = oUILib.GetWindow("frmEnvirDisplay")

        If oWin Is Nothing = False Then
            lTempTop = oWin.Top + oWin.Height + 1
        Else : lTempTop = CInt(oUILib.oDevice.PresentationParameters.BackBufferHeight / 2) - 85
        End If
        oWin = Nothing

        If lTempTop + 64 > oUILib.oDevice.PresentationParameters.BackBufferHeight Then lTempTop = oUILib.oDevice.PresentationParameters.BackBufferHeight - 64

        'frmColonyStats_Small initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eColonyStatsSmall
            .ControlName = "frmColonyStatsSmall"
            '.Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - 165
            '.Top = lTempTop

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.ColonyStatsX
                lTop = muSettings.ColonyStatsY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 165
            If lTop < 0 Then lTop = lTempTop
            If lLeft + 163 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 163
            If lTop + 64 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 64

            .Left = lLeft
            .Top = lTop

            .Width = 163
            .Height = 64 '82
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With

        'lblExpand initial props
        lblExpand = New UILabel(oUILib)
        With lblExpand
            .ControlName = "lblExpand"
            .Left = 2
            .Top = 2
            .Width = 160
            .Height = 10
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eColonyStatsMinimizer)
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblExpand, UIControl))

        'lblPowerInd initial props
        lblPowerInd = New UILabel(oUILib)
        With lblPowerInd
            .ControlName = "lblPowerInd"
            .Left = 6
            .Top = 13 '18
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eLightning)
        End With
        Me.AddChild(CType(lblPowerInd, UIControl))

        'lblPopInd initial props
        lblPopInd = New UILabel(oUILib)
        With lblPopInd
            .ControlName = "lblPopInd"
            .Left = 5
            .Top = 29 '37
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eSingleDude)
        End With
        Me.AddChild(CType(lblPopInd, UIControl))

        'lblMoraleInd initial props
        lblMoraleInd = New UILabel(oUILib)
        With lblMoraleInd
            .ControlName = "lblMoraleInd"
            .Left = 6
            .Top = 46 '57
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eHappySadFace)
        End With
        Me.AddChild(CType(lblMoraleInd, UIControl))

        'shpPowerBack initial props
        shpPowerBack = New UIWindow(oUILib)
        With shpPowerBack
            .ControlName = "shpPowerBack"
            .Left = 25
            .Top = 13 '18
            .Width = 120
            .Height = 14
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = Color.Red
            .FullScreen = False
            .BorderLineWidth = 1
			.bRoundedBorder = False
			.Moveable = False
        End With
        Me.AddChild(CType(shpPowerBack, UIControl))

        'shpPowerFore initial props
        shpPowerFore = New UIWindow(oUILib)
        With shpPowerFore
            .ControlName = "shpPowerFore"
            .Left = 26
            .Top = 14 '19
            .Width = 0 '119
            .Height = 13
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            .FullScreen = False
            .DrawBorder = False
			.bRoundedBorder = False
			.Moveable = False
        End With
        Me.AddChild(CType(shpPowerFore, UIControl))

        'shpPopHomeless initial props
        shpPopHomeless = New UIWindow(oUILib)
        With shpPopHomeless
            .ControlName = "shpPopHomeless"
            .Left = 25
            .Top = 30 '38
            .Width = 120
            .Height = 14
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = Color.Red
            .FullScreen = False
            .BorderLineWidth = 1
			.bRoundedBorder = False
			.Moveable = False
        End With
        Me.AddChild(CType(shpPopHomeless, UIControl))

        'shpMoraleNeg initial props
        shpMoraleBack = New UIWindow(oUILib)
        With shpMoraleBack
            .ControlName = "shpMoraleBack"
            .Left = 25
            .Top = 47 '58
            .Width = 120
            .Height = 14
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 1
			.bRoundedBorder = False
			.Moveable = False
        End With
        Me.AddChild(CType(shpMoraleBack, UIControl))

        'shpPopPowerless initial props
        shpPopPowerless = New UIWindow(oUILib)
        With shpPopPowerless
            .ControlName = "shpPopPowerless"
            .Left = 26
            .Top = 31 '39
            .Width = 0 '80
            .Height = 13
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 0)
            .FullScreen = False
            .DrawBorder = False
			.bRoundedBorder = False
			.Moveable = False
        End With
        Me.AddChild(CType(shpPopPowerless, UIControl))

        'shpPopGood initial props
        shpPopGood = New UIWindow(oUILib)
        With shpPopGood
            .ControlName = "shpPopGood"
            .Left = 26
            .Top = 31 '39
            .Width = 0 ' 61
            .Height = 13
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            .DrawBorder = False
            .FullScreen = False
			.bRoundedBorder = False
			.Moveable = False
        End With
        Me.AddChild(CType(shpPopGood, UIControl))

        'shpMoraleGood initial props
        shpMoraleFore = New UIWindow(oUILib)
        With shpMoraleFore
            .ControlName = "shpMoraleFore"
            .Left = 26
            .Top = 48 '59
            .Width = 0 '80
            .Height = 13
            .Enabled = True
            .Visible = False
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            .DrawBorder = False
            .FullScreen = False
			.bRoundedBorder = False
			.Moveable = False
        End With
        Me.AddChild(CType(shpMoraleFore, UIControl))

        'lblPopGrowth initial props
        lblPopGrowth = New UILabel(oUILib)
        With lblPopGrowth
            .ControlName = "lblPopGrowth"
            .Left = 148
            .Top = 31 '39
            .Width = 14
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 15.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPopGrowth, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewColonyStats) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack the rights to view colony statistics.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        'msw_Delay = Stopwatch.StartNew

        AddHandler shpPopGood.OnMouseDown, AddressOf HousingItemClick
        AddHandler shpPopHomeless.OnMouseDown, AddressOf HousingItemClick
        AddHandler shpPopPowerless.OnMouseDown, AddressOf HousingItemClick

        muSettings.ExpandedColonyStatsScreen = False

        mbLoading = False
        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.ColonyDetailsSmallShown)
    End Sub

    Private Sub RefreshDetails()
        Dim yMsg(7) As Byte
        Try
            System.BitConverter.GetBytes(GlobalMessageCode.eGetColonyDetails).CopyTo(yMsg, 0)
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, 2)
            Else
                Dim bFound As Boolean = False
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                        If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eSpaceStationSpecial Then
                            goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yMsg, 2)
                            bFound = True
                        End If
                        Exit For
                    End If
                Next X
                If bFound = False Then
                    Me.Visible = False
                    MyBase.moUILib.RemoveWindow(Me.ControlName)
                    Exit Sub
                End If
            End If
            MyBase.moUILib.SendMsgToPrimary(yMsg)
        Catch
        Finally
            If msw_Delay Is Nothing Then
                msw_Delay = Stopwatch.StartNew
            Else
                msw_Delay.Reset()
                msw_Delay.Start()
            End If
        End Try
    End Sub

    Public Sub HandleColonyMsg(ByVal yData() As Byte)
        Dim lTemp As Int32
        Dim yTemp(19) As Byte

        Dim lPopulation As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim lJobs As Int32 = System.BitConverter.ToInt32(yData, 12)
        Dim lEnlisted As Int32 = System.BitConverter.ToInt32(yData, 16)
        Dim lOfficers As Int32 = System.BitConverter.ToInt32(yData, 20)
        Dim lPowerGen As Int32 = System.BitConverter.ToInt32(yData, 24)
        Dim lPowerNeed As Int32 = System.BitConverter.ToInt32(yData, 28)
        Dim yUnemployment As Byte = yData(52)
        Dim lPoweredHomes As Int32 = System.BitConverter.ToInt32(yData, 53)
        Dim lUnpoweredHomes As Int32 = System.BitConverter.ToInt32(yData, 57)
		Dim lMorale As Int32 = System.BitConverter.ToInt32(yData, 61)
		Dim lGrowthRate As Int32 = System.BitConverter.ToInt16(yData, 65)

        'Power...
        lTemp = lPowerGen + lPowerNeed
        If lTemp <> 0 Then
            lTemp = CInt((lPowerGen / lTemp) * (shpPowerBack.Width - 1))
        Else : lTemp = shpPowerBack.Width - 1
        End If
        If shpPowerFore.Width <> lTemp Then shpPowerFore.Width = lTemp 'change it only if we need to
        Try
            'Growth Rate Indicator
            If lGrowthRate > 0 Then
                If lblPopGrowth.Caption <> "+" Then
                    lblPopGrowth.Caption = "+"
                    lblPopGrowth.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                End If
                lblPopGrowth.ToolTipText = "+" & lGrowthRate
            ElseIf lGrowthRate = 0 Then
                If lblPopGrowth.Caption <> "" Then
                    lblPopGrowth.Caption = ""
                End If
                lblPopGrowth.ToolTipText = "No Growth"
            ElseIf lblPopGrowth.Caption <> "-" Then
                lblPopGrowth.Caption = "-"
                lblPopGrowth.ForeColor = Color.Red
                lblPopGrowth.ToolTipText = lGrowthRate.ToString
            End If

            'Ok, calculate our numbers
            Dim lHomeless As Int32 = Math.Max(0, lPopulation - (lPoweredHomes + lUnpoweredHomes))
            Dim lPowerless As Int32 = Math.Max(0, Math.Min((lPopulation - lPoweredHomes), lUnpoweredHomes))

            'Now... powerless
            If lPowerless <> 0 AndAlso shpPopPowerless.Visible = False Then
                shpPopPowerless.Visible = True
            ElseIf lPowerless = 0 AndAlso shpPopPowerless.Visible = True Then
                shpPopPowerless.Visible = False
            End If

            If lPowerless <> 0 AndAlso lHomeless = 0 AndAlso shpPopPowerless.Width <> shpPopHomeless.Width - 1 Then
                shpPopPowerless.Width = shpPopHomeless.Width - 1
            Else
                lTemp = lPoweredHomes + lPowerless
                lTemp = CInt((lTemp / lPopulation) * (shpPopHomeless.Width - 1))
                If shpPopPowerless.Width <> lTemp Then shpPopPowerless.Width = lTemp
            End If

            'Remainder... (powered)
            lTemp = lPopulation - (lHomeless + lPowerless)
            lTemp = CInt((lTemp / lPopulation) * (shpPopHomeless.Width - 1))
            If shpPopGood.Width <> lTemp Then shpPopGood.Width = lTemp

            'Now, morale... morale is on a scale from -150 to 150
            Dim lLeft As Int32
            Dim lWidth As Int32
            If shpMoraleFore.Visible = False Then shpMoraleFore.Visible = True

            lTemp = shpMoraleBack.Width \ 2
            If lMorale <= -150 Then
                lLeft = shpMoraleBack.Left + 1
                lWidth = lTemp
                If shpMoraleFore.FillColor <> Color.Red Then shpMoraleFore.FillColor = Color.Red
            ElseIf lMorale >= 150 Then
                lLeft = shpMoraleBack.Left + 1 + lTemp
                lWidth = lTemp - 1
                If shpMoraleFore.FillColor <> System.Drawing.Color.FromArgb(255, 0, 255, 0) Then shpMoraleFore.FillColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            ElseIf lMorale < 0 Then
                lWidth = CInt(Math.Abs(lTemp * (lMorale / 150)))
                lLeft = shpMoraleBack.Left + 1 + (lTemp - lWidth)
                If shpMoraleFore.FillColor <> Color.Red Then shpMoraleFore.FillColor = Color.Red
            Else
                lLeft = shpMoraleBack.Left + 1 + lTemp
                lWidth = CInt(Math.Abs(lTemp * (lMorale / 150)))
                If shpMoraleFore.FillColor <> System.Drawing.Color.FromArgb(255, 0, 255, 0) Then shpMoraleFore.FillColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            End If
            If shpMoraleFore.Left <> lLeft Then shpMoraleFore.Left = lLeft
            If shpMoraleFore.Width <> lWidth Then shpMoraleFore.Width = lWidth


            Dim sPowerTip As String = ""
            Dim sPopTip As String = ""
            Dim sMoraleTip As String = ""

            sPowerTip = "Generating: " & lPowerGen & vbCrLf & "Using: " & lPowerNeed
            sPopTip = "Powered Homes: " & (lPopulation - (lHomeless + lPowerless)) & vbCrLf & "Unpowered: " & lPowerless & vbCrLf & "Homeless: " & lHomeless
            sMoraleTip = "Morale: " & lMorale

            'Now, set our tooltips
            lblPowerInd.ToolTipText = sPowerTip : shpPowerBack.ToolTipText = sPowerTip : shpPowerFore.ToolTipText = sPowerTip
            lblPopInd.ToolTipText = sPopTip : shpPopHomeless.ToolTipText = sPopTip : shpPopPowerless.ToolTipText = sPopTip : shpPopGood.ToolTipText = sPopTip
            lblMoraleInd.ToolTipText = sMoraleTip : shpMoraleBack.ToolTipText = sMoraleTip : shpMoraleFore.ToolTipText = sPopTip
        Catch
        End Try
    End Sub

    Private Sub frmColonyStatsSmall_OnNewFrame() Handles Me.OnNewFrame
        If msw_Delay Is Nothing OrElse msw_Delay.ElapsedMilliseconds > ml_REQUEST_DELAY Then
            RefreshDetails()
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If msw_Delay Is Nothing = False Then msw_Delay.Stop()
        msw_Delay = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub lblExpand_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles lblExpand.OnMouseDown
        Dim frmTemp As frmColonyStats = CType(MyBase.moUILib.GetWindow("frmColonyStats"), frmColonyStats)
        If frmTemp Is Nothing Then
            frmTemp = New frmColonyStats(MyBase.moUILib)
        Else : frmTemp.Visible = True
        End If
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        muSettings.ExpandedColonyStatsScreen = True
        frmTemp = Nothing
    End Sub

    Private Sub HousingItemClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As System.Windows.Forms.MouseButtons)
        frmMain.SelectAndGotoNextUnit(frmMain.eSelectNextType.eUnpoweredResidence)
    End Sub

    Private Sub frmColonyStatsSmall_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.ColonyStatsX = Me.Left
            muSettings.ColonyStatsY = Me.Top
        End If
    End Sub
End Class