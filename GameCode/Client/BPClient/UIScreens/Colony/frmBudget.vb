Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmBudget
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private WithEvents btnClose As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents btnExport As UIButton
    Public fraSummary As fraBudgetSummary
    Public fraDeath As fraDeathBudget
    Public fraDetails As fraBudgetDetails
    Public fraLineItem As fraBudgetLineItem
    Private fraAgent As fraAgentBudget

    'Public Shared bShowHWSupply As Boolean = False
    'Private WithEvents chkShowHWSupply As UICheckBox

    Private lblTradeIncome As UILabel
    Private txtTradeIncome As UITextBox

    Private mlLastUpdate As Int32 = Int32.MinValue

    Private mlEnvirID As Int32 = -1
    Private miEnvirTypeID As Int16 = -1

    Private mbInitialized As Boolean = False

    Private mlPrevDeathBudget As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenBudgetWindow)

        'frmBudget initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eBudgetWindow
            .ControlName = "frmBudget"
            '.Left = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 400
            '.Top = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 300
            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.BudgetX
                lTop = muSettings.BudgetY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 300
            If lLeft + 800 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 800
            If lTop + 600 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 600

            .Left = lLeft
            .Top = lTop

            .Width = 800
            .Height = 600
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 1
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 400
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Budget Report for Enoch Dagor"
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
            .Left = 0
            .Top = 30
            .Width = 800
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
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

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = btnClose.Left - 25
            .Top = btnClose.Top
            .Width = btnClose.Width
            .Height = btnClose.Height
            .Enabled = True
            .Visible = True
            .Caption = "?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to begin the tutorial for this window"
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'btnExport
        btnExport = New UIButton(oUILib)
        With btnExport
            .ControlName = "btnExport"
            .Left = btnHelp.Left - 25
            .Top = btnClose.Top
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "E"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            '.ToolTipText = "You can change the format of exported data from CSV to XML " & vbCrLf & "in the general settings window." & vbCrLf & "Exported data will go into your GameDir\ExportedData"
            .ToolTipText = "Click to export data." & vbCrLf & "Exported data will go into your GameDir\ExportedData"
        End With
        Me.AddChild(CType(btnExport, UIControl))

        ''chkShowHWSupply initial props
        'chkShowHWSupply = New UICheckBox(oUILib)
        'With chkShowHWSupply
        '    .ControlName = "chkShowHWSupply"
        '    .Left = Me.Width - 300
        '    .Top = lblTitle.Top
        '    .Width = 240
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True

        '    .Caption = "Show Homeworld Supply Costs"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = bShowHWSupply
        '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        '    .ToolTipText = "Shows the effects of the upcoming Homeworld Supply changes." & vbCrLf & _
        '                   "These changes are still in progress and are subject to change." & vbCrLf & _
        '                   "The expenses apply to colonies in systems away from your homeworld." & vbCrLf & _
        '                   "Your homeworld is determined by where the invulnerability is set." & vbCrLf & _
        '                   "To mitigate the cost of the colony, it must be made self-sustaining." & vbCrLf & _
        '                   "Growing the colony's population will fix the issue until a point" & vbCrLf & _
        '                   "where the colony can generate enough supplies for its population."
        'End With
        'Me.AddChild(CType(chkShowHWSupply, UIControl))

        'fraSummary initial props
        fraSummary = New fraBudgetSummary(oUILib)
        With fraSummary
            .ControlName = "fraSummary"
            .Left = 10
            .Top = 40
            .Width = 310
            .Height = 80
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With
        Me.AddChild(CType(fraSummary, UIControl))

        'fraDeath initial props
        fraDeath = New fraDeathBudget(oUILib)
        With fraDeath
            .ControlName = "fraDeath"
			.Left = 325
            .Top = 40
            .Enabled = True
            .Visible = True
            .FullScreen = False
            .BorderLineWidth = 1
            .Moveable = False
            .SetDeathBalance(0)
        End With
		Me.AddChild(CType(fraDeath, UIControl))

		'fraAgent initial props
		fraAgent = New fraAgentBudget(oUILib)
		With fraAgent
			.ControlName = "fraAgent"
			.Left = fraDeath.Left + fraDeath.Width + 5
			.Top = fraDeath.Top
			.Enabled = True
			.Visible = True
			.FullScreen = False
			.BorderLineWidth = 1
			.Moveable = False
		End With
		Me.AddChild(CType(fraAgent, UIControl))

        'fraDetails initial props
        fraDetails = New fraBudgetDetails(oUILib)
        With fraDetails
            .ControlName = "fraDetails"
            .Left = 0
            .Top = 130
            .Width = 800
            .Height = 200
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With
        Me.AddChild(CType(fraDetails, UIControl))

        'fraLineItem initial props
        fraLineItem = New fraBudgetLineItem(oUILib)
        With fraLineItem
            .ControlName = "fraLineItem"
            .Left = 10
            .Top = 340
            .Height = Me.Height - .Top - 10
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With
        Me.AddChild(CType(fraLineItem, UIControl))

        'lblTradeIncome initial props
        lblTradeIncome = New UILabel(oUILib)
        With lblTradeIncome
            .ControlName = "lblTradeIncome"
            .Left = fraLineItem.Left + fraLineItem.Width + 10
            .Top = fraLineItem.Top
            .Width = Me.Width - .Left - 5
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Trade Income"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTradeIncome, UIControl))

        'txtTradeIncome initial props
        txtTradeIncome = New UITextBox(oUILib)
        With txtTradeIncome
            .ControlName = "txtTradeIncome"
            .Left = fraLineItem.Left + fraLineItem.Width + 10
            .Top = lblTradeIncome.Top + lblTradeIncome.Height
            .Width = lblTradeIncome.Width
            .Height = fraLineItem.Height - lblTradeIncome.Height
            .Enabled = True
            .Visible = True
            .Caption = ""
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
        Me.AddChild(CType(txtTradeIncome, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewBudget) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack rights to view the budget interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        AddHandler fraDetails.ItemSelected, AddressOf ItemClicked
        AddHandler fraDetails.ListItemDoubleClicked, AddressOf ItemDoubleClicked

        Dim yMsg(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerBudget).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
        MyBase.moUILib.SendMsgToPrimary(yMsg)

        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
            goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.IAmLostF7Key)
        End If

        mbInitialized = True
    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub frmBudget_OnNewFrame() Handles Me.OnNewFrame
        If mbInitialized = False Then Return
		If goCurrentPlayer Is Nothing Then Return
		If goCurrentPlayer.oBudget Is Nothing Then Return

		If mlPrevDeathBudget <> goCurrentPlayer.DeathBudgetBalance Then
			mlPrevDeathBudget = goCurrentPlayer.DeathBudgetBalance
			fraDeath.SetDeathBalance(goCurrentPlayer.DeathBudgetBalance)
			If NewTutorialManager.TutorialOn = True Then
				NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eDeathBudgetDeposit, -1, -1, -1, "")
			End If
		End If

		If mlLastUpdate <> goCurrentPlayer.oBudget.lLastUpdateCycle Then
			mlLastUpdate = goCurrentPlayer.oBudget.lLastUpdateCycle

			'Now... fill our data.
            With goCurrentPlayer.oBudget

                'Dim blMoreExp As Int64 = 0
                'If frmBudget.bShowHWSupply = True Then
                '    For X As Int32 = 0 To .mlItemUB
                '        blMoreExp += .muItems(X).GetHomeworldSupplyLineCost(.lIronCurtainPlanet)
                '    Next X
                'End If
                'fraSummary.SetValues(.TotalRevenue, .TotalExpense + blMoreExp, .lAgentMaintCost)
                fraSummary.SetValues(.TotalRevenue, .TotalExpense, .lAgentMaintCost)
                fraDetails.SetList(goCurrentPlayer.oBudget)
                fraAgent.SetText()

                fraDeath.SetMaximum(.lMaxDeathBudget)

                If mlEnvirID <> -1 AndAlso miEnvirTypeID <> -1 Then
                    fraLineItem.Visible = True
                    fraLineItem.PopulateData(goCurrentPlayer.oBudget, mlEnvirID, miEnvirTypeID)
                Else : fraLineItem.Visible = False
                End If

                lblTitle.Caption = "Budget Report for " & goCurrentPlayer.PlayerName
            End With
		End If
        fraLineItem.fraBudgetLineItem_OnNewFrame()
        fraDetails.CheckFlashStates()

        Dim sTemp As String = goCurrentPlayer.oBudget.GetTradeText
        If txtTradeIncome.Caption <> sTemp Then
            txtTradeIncome.Caption = sTemp
            Dim sTotalTradeValue As String = goCurrentPlayer.oBudget.TotalTradeIncome.ToString("#,##0")
            lblTradeIncome.Caption = "Trade Income (" & sTotalTradeValue & ")"
        End If


        'Render Budget Percent
        Using oSprite As New Sprite(MyBase.moUILib.oDevice)
            oSprite.Begin(SpriteFlags.AlphaBlend)
            Dim lTemp As Int32 = (goCurrentPlayer.oBudget.mlItemUB * 100) \ 348 \ 10



            Dim rcSrc As Rectangle = grc_UI(elInterfaceRectangle.eMinBar_0 + lTemp) 'Rectangle.FromLTRB(193, 157 + (lTemp * 9), 256, 157 + ((lTemp + 1) * 9))
            Dim clrVal As System.Drawing.Color
            If lTemp < 7 Then
                clrVal = Color.Green
            ElseIf lTemp < 8 Then
                clrVal = Color.Yellow
            Else : clrVal = Color.Red
            End If
            Dim rcProgress As Rectangle = New Rectangle(390, 7, 64, 9)
            rcProgress.X = Me.Left + Me.Width - 64
            rcProgress.Y = Me.Top - 9

            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcProgress, Point.Empty, 0, rcProgress.Location, clrVal)
            oSprite.End()
            oSprite.Dispose()
        End Using

    End Sub

	Private Sub ItemClicked(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
		mlEnvirID = lEnvirID
		miEnvirTypeID = iEnvirTypeID
		mlLastUpdate = Int32.MinValue
	End Sub

	Private Sub ItemDoubleClicked(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
		goCamera.TrackingIndex = -1

		If lEnvirID < 1 OrElse iEnvirTypeID < 1 Then Return

		Dim lID As Int32 = lEnvirID
		Dim iTypeID As Int16 = iEnvirTypeID

		'If it is a facility (spacestation), then convert to the true parent
		If iTypeID = ObjectType.eFacility Then
			For X As Int32 = 0 To goCurrentPlayer.oBudget.mlItemUB
				If goCurrentPlayer.oBudget.muItems(X).lEnvirID = lEnvirID AndAlso goCurrentPlayer.oBudget.muItems(X).iEnvirTypeID = iEnvirTypeID Then
					If iEnvirTypeID = ObjectType.eFacility Then
						lID = goCurrentPlayer.oBudget.muItems(X).lJumpToID
						iTypeID = ObjectType.eSolarSystem
					End If
					Exit For
				End If
			Next X
        End If
        If iTypeID <> ObjectType.ePlanet AndAlso iTypeID <> ObjectType.eSolarSystem Then
            Return
        End If

		'Ok, lID and TypeID have the correct values
        If goCurrentEnvir Is Nothing = False Then
            Dim bAlreadyHere As Boolean = False
            If goCurrentEnvir.ObjectID = lID AndAlso goCurrentEnvir.ObjTypeID = iTypeID Then
                'we're already here... nothing to do...
                bAlreadyHere = True
            Else : frmMain.ForceChangeEnvironment(lID, iTypeID)
            End If
            If iTypeID = ObjectType.ePlanet Then
                glCurrentEnvirView = CurrentView.ePlanetMapView
            ElseIf iEnvirTypeID = ObjectType.eFacility Then
                glCurrentEnvirView = CurrentView.eSystemView
                
                If bAlreadyHere = True Then
                    goCurrentEnvir.JumpToEntity(goUILib.lJumpToExtended1, goUILib.lJumpToExtended2)
                Else
                    goUILib.lUISelectState = UILib.eSelectState.eJumpToEntity
                    goUILib.lJumpToExtended1 = lEnvirID : goUILib.lJumpToExtended2 = iEnvirTypeID
                End If
            Else
                glCurrentEnvirView = CurrentView.eSystemMapView1
            End If
        End If
		MyBase.moUILib.RemoveWindow(Me.ControlName)

		If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
			goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.FirstBudgetScreenJumpTo)
		End If
	End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eBudget)
    End Sub

    Private Sub frmBudget_WindowClosed() Handles Me.WindowClosed
        fraDetails.SaveScrollPosition()
    End Sub

    Private Sub frmBudget_WindowMoved() Handles Me.WindowMoved
        If mbInitialized = True Then
            muSettings.BudgetX = Me.Left
            muSettings.BudgetY = Me.Top
        End If
    End Sub

    'Private Sub chkShowHWSupply_Click() Handles chkShowHWSupply.Click
    '    bShowHWSupply = chkShowHWSupply.Value
    'End Sub

    Private Sub btnExport_Click(ByVal sName As String) Handles btnExport.Click
        Export_BudgetInfo()
    End Sub

    Private Sub Export_BudgetInfo()
        If goCurrentPlayer Is Nothing Then Return
        Export_BudgetInfo_Csv()
    End Sub

    Private Sub Export_BudgetInfo_Csv()
        Dim sExportData As String = ""
        Dim iTotalIncome As Int64
        Dim iTotalExpense As Int64
        sExportData &= "Name,Total Income,Total Expense,Colony Tax Income,Howeworld Supplies,Population Upkeep,Research Facilities,Factories,Spaceports,Defenses,Excess Storage,Other Facilities,Unemployment,Mining Bids,Non-Aerial Support,Aerial Support,Docked Non-Aerial,Docked Aerial,Mining Income,TaxRate,Control" & vbCrLf
        sExportData &= "TradeIncome," & goCurrentPlayer.oBudget.TotalTradeIncome.ToString & ",,,,,,,,,,,,,,,,," & vbCrLf

        For X As Int32 = 0 To goCurrentPlayer.oBudget.mlItemUB
            With goCurrentPlayer.oBudget.muItems(X)
                If .iEnvirTypeID = ObjectType.eFacility Then
                    sExportData &= GetCacheObjectValue(.lSystemID, ObjectType.eSolarSystem) & " (" & GetCacheObjectValue(.lEnvirID, .iEnvirTypeID) & ")"
                Else
                    sExportData &= GetCacheObjectValue(.lEnvirID, .iEnvirTypeID)
                End If
                iTotalIncome = .TaxIncome + .blMiningBidIncome
                iTotalExpense = .lHWSupplyLineCost + .PopUpkeep + .ResearchCost + .FactoryCost + .SpaceportCost + .TurretCost + .ExcessStorage + .OtherFacCost + .UnemploymentCost + .blMiningBidExpense + .NonAirCost + .AirCost + .DockedNonAirCost + .DockedAirCost
                sExportData &= ","
                sExportData &= iTotalIncome.ToString
                sExportData &= ","
                sExportData &= iTotalExpense.ToString

                sExportData &= ","
                sExportData &= .TaxIncome.ToString
                sExportData &= ","
                sExportData &= .lHWSupplyLineCost.ToString
                sExportData &= ","
                sExportData &= .PopUpkeep.ToString
                sExportData &= ","
                sExportData &= .ResearchCost.ToString
                sExportData &= ","
                sExportData &= .FactoryCost.ToString
                sExportData &= ","
                sExportData &= .SpaceportCost.ToString
                sExportData &= ","
                sExportData &= .TurretCost.ToString
                sExportData &= ","
                sExportData &= .ExcessStorage.ToString
                sExportData &= ","
                sExportData &= .OtherFacCost.ToString
                sExportData &= ","
                sExportData &= .UnemploymentCost.ToString
                sExportData &= ","
                sExportData &= .blMiningBidExpense.ToString
                sExportData &= ","
                sExportData &= .NonAirCost.ToString
                sExportData &= ","
                sExportData &= .AirCost.ToString
                sExportData &= ","
                sExportData &= .DockedNonAirCost.ToString
                sExportData &= ","
                sExportData &= .DockedAirCost.ToString
                sExportData &= ","
                sExportData &= .blMiningBidIncome.ToString
                sExportData &= ","
                sExportData &= .yTaxRate.ToString
                sExportData &= ","
                sExportData &= .yPlanetaryControl.ToString
                sExportData &= vbCrLf
            End With
        Next X
        If sExportData = "" Then Return
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile = sFile & "\"
        sFile &= "ExportedData"
        If Exists(sFile) = False Then MkDir(sFile)
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "Budget_" & goCurrentPlayer.PlayerName & "_" & Now.ToString("MM_dd_yyyy_HHmmss") & ".csv"

        Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Create)

        Dim info As Byte() = New System.Text.UTF8Encoding(True).GetBytes(sExportData)
        fsFile.Write(info, 0, info.Length)
        fsFile.Close()
        fsFile.Dispose()
        If goUILib Is Nothing = False Then
            goUILib.AddNotification("Budget Info Exported.", Color.White, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub
End Class