Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Current Environment details display (envir name, command pop, credits)
Public Class frmEnvirDisplay
    Inherits UIWindow

    Private lblEnvirName As UILabel
    Private lblEnvirNameGoto As UILabel
    Private lblCommandCap As UILabel
    Private lblFacCPCap As UILabel
    Private lblCredits As UILabel
    'Private lblWarpoints As UILabel
    Private ln1 As UILine
    Private ln2 As UILine
    Private WithEvents btnBombardment As UIButton
    Public WithEvents btnGoLeft As UIButton
    Public WithEvents btnGoRight As UIButton
    Private WithEvents btnEnvList As UIButton
    Private WithEvents btnEnvShortcuts As UIButton

	Public Shared dtEndTimer As Date = Date.MinValue
	Public Shared mlPirateWave As Int32 = Int32.MinValue
	Public Shared clrPirateWave As System.Drawing.Color

	Private lWaveFullAlphaTime As Int32 = Int32.MinValue


    Private mdtCPPenaltyEnds As Date = Date.MinValue
    Private mlCPPenalty As Int32 = 0
    Private mlDecPlayer() As Int32
    Private myPenalty() As Byte
    Private mlCPPenaltyUB As Int32 = -1
    Private mbCPMsgSent As Boolean = False
    Private mlLastCPJunkCheck As Int32 = 0
    Private msCPJunk As String = ""
    Private mlLastCPMsgSent As Int32 = 0

    Private miWantedObjTypeID As Short = 0
    Private miWantedObjectID As Int32 = 0

    Private msImportantMsg As String = ""
    Private mclrImportantMsg As System.Drawing.Color
    Private mfImportantMsgAlpha As Single = 0.0F
    Private mlImportantMsgIndex As Int32 = 0
    Private Structure uImportantMsg
        Public sMsg As String
        Public clr As System.Drawing.Color
        Public lIndex As Int32
        Public lBaseAlpha As Int32
    End Structure
    Private mlLastImportantMsgCycle As Int32 = 0
    Private mcolImportantMsgs As Collection


    'Private moCPFlashFont As Font
    'Private moWaveFont As Font

    Public sNextWPUpkeepTime As String
    Public sNextGuildShareUpkeep As String
    Public dtSessionNextUpkeep As DateTime
    Private mlLastWPTimerRequest As Int32 = 0
    
    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
        Dim sGotoTooltip As String
        'sGotoTooltip = "The current environment you are located in. If you are viewing another environment, the active environment will be in parentheses."
        sGotoTooltip = "Click the arrows to cycle through the planets" & vbCrLf
        sGotoTooltip &= "Left-Click the label to go there." & vbCrLf
        sGotoTooltip &= "Or Shift-Left Click the arrows to go directly to the next planet." & vbCrLf
        sGotoTooltip &= "Right-Click the label to reset."

        'frmEnvirDisplay initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eEnvirDisplay
            .ControlName = "frmEnvirDisplay"
            'Initially had this up against the Minimap... but instead, I'm going to throw it on the right side...
            '.Left = muSettings.MiniMapWidthHeight
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 203 '281
            .Top = 1

            If muSettings.EnvirDisplayLocX <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Left = muSettings.EnvirDisplayLocX
            If muSettings.EnvirDisplayLocY <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Top = muSettings.EnvirDisplayLocY

            .Width = 211 '279
            .Height = 63 '25
            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            '.FullScreen = True  'always display the environment display
            .FullScreen = False
            .BorderLineWidth = 2
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'lblEnvirName initial props
        lblEnvirName = New UILabel(oUILib)
        With lblEnvirName
            .ControlName = "lblEnvirName"
            .Left = 0
            .Top = 2
            .Width = 139
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = ""
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
            '.ToolTipText = "The current environment you are located in. If you are viewing another environment, the active environment will be in parentheses."
            .ToolTipText = sGotoTooltip
        End With
        Me.AddChild(CType(lblEnvirName, UIControl))

        'lblEnvirNameGoto initial props
        lblEnvirNameGoto = New UILabel(oUILib)
        With lblEnvirNameGoto
            .ControlName = "lblEnvirNameGoto"
            .Left = 0
            .Top = lblEnvirName.Top + lblEnvirName.Height
            .Width = 139
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = "line2"
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
            .ToolTipText = sGotoTooltip
        End With
        Me.AddChild(CType(lblEnvirNameGoto, UIControl))

        'lblCommandCap initial props
        lblCommandCap = New UILabel(oUILib)
        With lblCommandCap
            .ControlName = "lblCommandCap"
            .Left = 139
            .Top = 2
            .Width = 61
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = ""
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center 'DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
        End With
        Me.AddChild(CType(lblCommandCap, UIControl))

        'lblFacCPCap initial props
        lblFacCPCap = New UILabel(oUILib)
        With lblFacCPCap
            .ControlName = "lblFacCPCap"
            .Left = 139
            .Top = 15
            .Width = 61
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = ""
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center  'DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
        End With
        Me.AddChild(CType(lblFacCPCap, UIControl))

        'lblCredits initial props
        lblCredits = New UILabel(oUILib)
        With lblCredits
            .ControlName = "lblCredits"
            .Left = 2 '203
            .Top = 27 '0
            .Width = 200 '76
            .Height = 27
            .Enabled = True
            .Visible = True
            .Caption = ""
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
            .ToolTipText = "The amount of credits in your treasury."
        End With
        Me.AddChild(CType(lblCredits, UIControl))

        ''lblWarpoints initial props
        'lblWarpoints = New UILabel(oUILib)
        'With lblWarpoints
        '    .ControlName = "lblWarpoints"
        '    .Left = 2 '203
        '    .Top = 40 '0
        '    .Width = 200 '76
        '    .Height = 27
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Arial", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        '    .ToolTipText = "The amount of warpoints you have available." & vbCrLf & "You gain warpoints by killing enemy units."
        'End With
        'Me.AddChild(CType(lblWarpoints, UIControl))

        'ln1 initial props
        ln1 = New UILine(oUILib)
        With ln1
            .ControlName = "ln1"
            .Left = 139
            .Top = 1
            .Width = 0
            .Height = 27
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(ln1, UIControl))

        'ln2 initial props
        ln2 = New UILine(oUILib)
        With ln2
            .ControlName = "ln2"
            .Left = Me.BorderLineWidth  '203
            .Top = 28 '0
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0 '26
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(ln2, UIControl))

        'btnBombardment initial props
        btnBombardment = New UIButton(oUILib)
        With btnBombardment
            .ControlName = "btnBombardment"
            .Left = 2
            .Top = 48
            .Width = Me.Width - 3
            .Height = 14
            .Enabled = True
            .Visible = False
            .Caption = "Orbital Bombardment"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Accesses the Orbital Bombardment options for this planet."
        End With
        Me.AddChild(CType(btnBombardment, UIControl))

        'btnEnvList initial props
        btnEnvList = New UIButton(oUILib)
        With btnEnvList
            .ControlName = "btnEnvList"
            .Width = 10
            .Height = ln2.Top - 4
            .Left = lblEnvirName.Left + lblEnvirName.Width - 10
            .Top = 2
            .Enabled = True
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            '.ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = sGotoTooltip
            .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal)
        End With
        Me.AddChild(CType(btnEnvList, UIControl))

        'btnEnvShortcuts initial props
        btnEnvShortcuts = New UIButton(oUILib)
        With btnEnvShortcuts
            .ControlName = "btnEnvShortcuts"
            .Left = 2
            .Top = 2
            .Width = 10
            .Height = ln2.Top - 4
            .Enabled = True
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = "S"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = sGotoTooltip
        End With
        Me.AddChild(CType(btnEnvShortcuts, UIControl))

        'btnGoLeft initial props
        btnGoLeft = New UIButton(oUILib)
        With btnGoLeft
            .ControlName = "btnGoLeft"
            .Left = btnEnvShortcuts.Left + btnEnvShortcuts.Width + 1
            .Top = 2
            .Width = 10
            .Height = ln2.Top - 4
            .Enabled = True
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = "<"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = sGotoTooltip
        End With
        Me.AddChild(CType(btnGoLeft, UIControl))

        'btnGoRight initial props
        btnGoRight = New UIButton(oUILib)
        With btnGoRight
            .ControlName = "btnGoRight"
            .Top = 2
            .Width = 10
            .Left = btnEnvList.Left - .Width - 1
            .Height = ln2.Top - 4
            .Enabled = True
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = ">"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = sGotoTooltip
        End With
        Me.AddChild(CType(btnGoRight, UIControl))


        oUILib.RemoveWindow(Me.ControlName)

        'Now, add me...
        oUILib.AddWindow(Me)
    End Sub

	Protected Overrides Sub Finalize()
		moSW = Nothing
		Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
		MyBase.Finalize()
	End Sub

    'Public Sub ClearFonts()
    '    If moCPFlashFont Is Nothing = False Then moCPFlashFont.Dispose()
    '    moCPFlashFont = Nothing
    '    If moWaveFont Is Nothing = False Then moWaveFont.Dispose()
    '    moWaveFont = Nothing
    'End Sub

    Private Sub SendCPmsg()
        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetCPPenaltyList).CopyTo(yMsg, 0)
        MyBase.moUILib.SendMsgToPrimary(yMsg)
        mlLastCPMsgSent = glCurrentCycle
    End Sub

    Private Sub SetCPJunk()

        Try
            If mlCPPenalty = 0 Then
                msCPJunk = ""
                Return
            End If

            Dim oSB As New System.Text.StringBuilder()
            oSB.Append(vbCrLf & vbCrLf)
            oSB.AppendLine("Current Penalties:" & mlCPPenalty & " CP per unit")
            oSB.AppendLine("Penalty Ends: " & mdtCPPenaltyEnds.ToString("MM/dd HH:mm"))
            oSB.AppendLine("Penalties are for declaring on: ")
            For X As Int32 = 0 To mlCPPenaltyUB
                Dim sTemp As String = "  " & myPenalty(X) & " CP for " & GetCacheObjectValue(mlDecPlayer(X), ObjectType.ePlayer)
                oSB.AppendLine(sTemp)
            Next X

            msCPJunk = oSB.ToString
            mlLastCPJunkCheck = glCurrentCycle
        Catch
            msCPJunk = ""
            mlLastCPJunkCheck = 0
        End Try

    End Sub

    Private Sub RequestWPTimers()
        Dim yMsg(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eWPTimers).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
        goUILib.SendMsgToPrimary(yMsg)

        mlLastWPTimerRequest = glCurrentCycle
    End Sub

    Private Sub frmEnvirDisplay_OnNewFrame() Handles Me.OnNewFrame
        Dim sName As String
        Try

            If mbCPMsgSent = False Then
                If goCurrentEnvir Is Nothing = False Then
                    mbCPMsgSent = True
                    SendCPmsg()
                End If
            ElseIf glCurrentCycle - mlLastCPMsgSent > 1800 Then
                mbCPMsgSent = False
            End If

            If glCurrentCycle - mlLastCPJunkCheck > 90 AndAlso mbCPMsgSent = True Then
                SetCPJunk()
            End If

            'Now, figure out what we are currently viewing...
            Dim oViewingObject As Base_GUID = Nothing
            If goGalaxy Is Nothing = False Then
                If glCurrentEnvirView = CurrentView.eGalaxyMapView Then
                    sName = goGalaxy.GalaxyName
                Else
                    If goGalaxy.CurrentSystemIdx <> -1 Then
                        With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                            If glCurrentEnvirView = CurrentView.eSystemMapView1 OrElse glCurrentEnvirView = CurrentView.eSystemMapView2 _
                              OrElse glCurrentEnvirView = CurrentView.eSystemView Then
                                sName = .SystemName
                                oViewingObject = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                            Else
                                'Ok, check the planet level
                                If .CurrentPlanetIdx <> -1 Then
                                    If glCurrentEnvirView = CurrentView.ePlanetMapView OrElse glCurrentEnvirView = CurrentView.ePlanetView Then
                                        sName = .moPlanets(.CurrentPlanetIdx).PlanetName
                                        oViewingObject = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(.CurrentPlanetIdx)
                                    Else
                                        sName = .SystemName
                                        oViewingObject = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                                    End If
                                Else
                                    sName = .SystemName
                                    oViewingObject = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                                End If
                            End If
                        End With
                    Else : sName = goGalaxy.GalaxyName
                    End If
                End If
            Else : sName = "Universe"
            End If

            If glCurrentEnvirView = CurrentView.ePlanetView Then
                If Me.Height <> 70 Then
                    Me.Height = 70
                    btnBombardment.Visible = True

                    btnBombardment.Top = Me.Height - btnBombardment.Height - 2
                    'dim ofrm as UIWindow = 
                    If muSettings.ExpandedColonyStatsScreen = True Then
                        Dim ofrm As UIWindow = goUILib.GetWindow("frmColonyStats")
                        If ofrm Is Nothing = False Then ofrm.Top = Me.Height + Me.Top
                    Else
                        Dim ofrm As UIWindow = goUILib.GetWindow("frmColonyStatsSmall")
                        If ofrm Is Nothing = False Then ofrm.Top = Me.Height + Me.Top
                    End If
                End If
            Else
                'If Me.Height <> 50 Then
                '    Me.Height = 50
                '    btnBombardment.Visible = False
                '    MyBase.moUILib.RemoveWindow("frmBombard")
                'End If
                If Me.Height <> 55 Then
                    Me.Height = 55
                    btnBombardment.Visible = False
                    MyBase.moUILib.RemoveWindow("frmBombard")
                End If
            End If

            If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then

                If oViewingObject Is Nothing = False Then
                    Dim bNeedToSet As Boolean = False
                    With CType(goCurrentEnvir.oGeoObject, Base_GUID)
                        If .ObjectID <> goCurrentEnvir.ObjectID OrElse .ObjTypeID <> goCurrentEnvir.ObjTypeID Then
                            bNeedToSet = True
                        End If
                    End With
                    If bNeedToSet = True AndAlso oViewingObject.ObjectID = goCurrentEnvir.ObjectID AndAlso oViewingObject.ObjTypeID = goCurrentEnvir.ObjTypeID Then goCurrentEnvir.oGeoObject = oViewingObject
                End If


                Dim sTemp As String = ""
                If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                    sTemp = CType(goCurrentEnvir.oGeoObject, Planet).PlanetName
                ElseIf CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                    sTemp = CType(goCurrentEnvir.oGeoObject, SolarSystem).SystemName
                End If
                If sTemp <> "" AndAlso sName <> sTemp Then
                    '                    sName = sTemp
                    If miWantedObjTypeID = 0 AndAlso miWantedObjectID = 0 Then
                        sName &= vbCrLf & "(" & sTemp & ")"
                    End If
                End If
            End If
            If miWantedObjTypeID <> 0 AndAlso miWantedObjectID <> 0 Then
                Dim sGoto As String
                If miWantedObjTypeID = ObjectType.eSolarSystem Then
                    sGoto = goGalaxy.GetSystemName(miWantedObjectID)
                Else
                    sGoto = goGalaxy.GetPlanetName(miWantedObjectID)
                End If
                If lblEnvirNameGoto.ForeColor <> Color.Green Then lblEnvirNameGoto.ForeColor = Color.Green
                If lblEnvirNameGoto.Caption <> sGoto Then lblEnvirNameGoto.Caption = sGoto
                If lblEnvirName.Height <> 14 Then lblEnvirName.Height = 14
            Else
                If lblEnvirNameGoto.Caption <> "" Then lblEnvirNameGoto.Caption = ""
                If lblEnvirName.Height <> 26 Then lblEnvirName.Height = 26
            End If
            If lblEnvirName.Caption <> sName Then lblEnvirName.Caption = sName

            If goCurrentPlayer Is Nothing = False Then

                If glCurrentCycle - mlLastWPTimerRequest > 18000 OrElse mlLastWPTimerRequest = 0 Then       '10 minutes
                    mlLastWPTimerRequest = glCurrentCycle
                    RequestWPTimers()
                End If

                'Dim sWP As String = ""
                Dim clrCredits As Color = muSettings.InterfaceBorderColor
                'Dim clrWarpoints As Color = muSettings.InterfaceBorderColor
                If HasAliasedRights(AliasingRights.eViewTreasury) = True Then
                    Dim sCF As String = goCurrentPlayer.blCashFlow.ToString("#,##0")
                    If goCurrentPlayer.blCashFlow > 0 Then sCF = "+" & sCF
                    sName = "Credits: " & goCurrentPlayer.blCredits.ToString("#,##0") & " / " & sCF
                    'sWP = "Warpoints: " & goCurrentPlayer.blWarpoints.ToString("#,##0") & " / " & goCurrentPlayer.lCurrentWarpointUpkeepCost.ToString("#,##0")
                    If goCurrentPlayer.blCashFlow < 0 Then
                        clrCredits = Color.FromArgb(255, 0, 0)
                    Else
                        clrCredits = Color.FromArgb(0, 255, 0)
                    End If
                    'If goCurrentPlayer.lCurrentWarpointUpkeepCost <= 250 Then
                    '    clrWarpoints = Color.FromArgb(0, 255, 0)
                    'ElseIf goCurrentPlayer.lCurrentWarpointUpkeepCost > goCurrentPlayer.blWarpoints Then
                    '    clrWarpoints = Color.FromArgb(255, 0, 0)
                    'Else : clrWarpoints = Color.FromArgb(255, 255, 0)
                    'End If
                Else
                    sName = "Credits: Lack Alias Rights"
                    'sWP = "Warpoints: Lack Alias Rights"
                End If
                If lblCredits.Caption <> sName Then lblCredits.Caption = sName
                If lblCredits.ForeColor <> clrCredits Then lblCredits.ForeColor = clrCredits
                'If lblWarpoints.Caption <> sWP Then lblWarpoints.Caption = sWP
                'If lblWarpoints.ForeColor <> clrWarpoints Then lblWarpoints.ForeColor = clrWarpoints

                'Dim sToolTip As String = "The amount of warpoints you have available and " & vbCrLf & "the current upkeep of your empire." & vbCrLf & "You gain warpoints by killing enemy units." & vbCrLf
                'sToolTip &= "This Session Change : " & (goCurrentPlayer.blWarpoints - goCurrentPlayer.blWarpointsSession).ToString("#,##0")

                'sToolTip &= vbCrLf & vbCrLf
                'sToolTip &= "Units have warpoint upkeep that gets deducted" & vbCrLf & _
                '            "from your warpoints every 24 hours of play time" & vbCrLf & _
                '            "or 30 days of real time whichever is first. This" & vbCrLf & _
                '            "is only deducted if your upkeep is higher than 250." & vbCrLf & _
                '            "Guild share assets upkeep every 24 hours of real" & vbCrLf & _
                '            "time and are always deducted." & vbCrLf

                'Dim lSecondsRemain As Int32 = CInt(dtSessionNextUpkeep.Subtract(Now).TotalSeconds)
                'Dim sDur As String = GetDurationFromSeconds(lSecondsRemain, False)
                'If sDur.StartsWith("00:") = True Then sDur = sDur.Substring(3)
                'sToolTip &= "  Play Time remaining: " & sDur & vbCrLf
                'sToolTip &= "  30 Day Cutoff: " & sNextWPUpkeepTime & vbCrLf
                'sToolTip &= "  Next Guild Upkeep: " & sNextGuildShareUpkeep

                'If lblWarpoints.ToolTipText <> sToolTip Then lblWarpoints.ToolTipText = sToolTip
            End If

                'Now, the command cap
                Dim lCPLimit As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eCPLimit)

                Dim lFacCPLimit As Int32 = lCPLimit \ 5
                sName = lMyFacilities.ToString & "/" & lFacCPLimit.ToString
                If lblFacCPCap.Caption <> sName Then lblFacCPCap.Caption = sName
                If lMyFacilities > lFacCPLimit Then
                    sName = "Command Point Overage Penalties: " & ((lMyFacilities - lFacCPLimit) * 2) & "% maintenance cost increase"
                    If lblFacCPCap.ToolTipText <> sName Then lblFacCPCap.ToolTipText = sName
                    Dim clrTemp As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    If lblFacCPCap.ForeColor <> clrTemp Then lblFacCPCap.ForeColor = clrTemp
                Else
                    sName = "Indicates Facility Command Points for this environment." & vbCrLf & "If the first number exceeds the second number" & vbCrLf & "then a large maintenance cost penalty is applied."
                    If lblFacCPCap.ToolTipText <> sName Then lblFacCPCap.ToolTipText = sName
                    If lblFacCPCap.ForeColor <> muSettings.InterfaceBorderColor Then lblFacCPCap.ForeColor = muSettings.InterfaceBorderColor
                End If

                If goCurrentEnvir Is Nothing = False Then
                    sName = goCurrentEnvir.lCPUsage.ToString & "/" & lCPLimit.ToString()
                    If lblCommandCap.Caption <> sName Then lblCommandCap.Caption = sName
                    sName = "Indicates Command Points for this environment." & vbCrLf & "If the first number exceeds the second number" & vbCrLf & "then penalties begin to affect your troops." & vbCrLf & "Penalties may include increased costs, docking delays," & vbCrLf & "commands becoming unavailable and chaos."
                    If goCurrentEnvir.lCPUsage > lCPLimit Then

                        If goTutorial Is Nothing = False AndAlso goTutorial.EventTriggered(TutorialManager.TutorialTriggerType.ExceedsCommandPoints) = False Then  'AndAlso goTutorial.TutorialOn = True Then
                            goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eCPExceeding)
                            goTutorial.ForcefulSetTrigger(TutorialManager.TutorialTriggerType.ExceedsCommandPoints)
                        End If

                        Dim fOverage As Single = CSng(goCurrentEnvir.lCPUsage / lCPLimit)

                        sName = "Command Point Overage Penalties:" & vbCrLf & _
                          "Maintenance Cost Multiplier (" & fOverage.ToString("#.#0") & ")" & vbCrLf & _
                          "Docking/Undocking Delays" & vbCrLf & "Planet Bombardment Unavailable"

                        Dim lTemp As Int32 = goCurrentEnvir.lCPUsage - lCPLimit
                        Dim lG As Int32 = 255
                        Dim lB As Int32 = 128
                        If lTemp > lB Then
                            lB = 0
                            lTemp -= lB
                        Else
                            lB -= lTemp
                            lTemp = 0
                        End If
                        If lTemp <> 0 Then
                            If lTemp > lG Then
                                lG = 0
                            Else
                                lG -= lTemp
                            End If
                        End If
                        If lblCommandCap.ForeColor.G <> lG OrElse lblCommandCap.ForeColor.B <> lB Then
                            lblCommandCap.ForeColor = System.Drawing.Color.FromArgb(255, 255, lG, lB)
                        End If
                    ElseIf lblCommandCap.ForeColor.ToArgb <> System.Drawing.Color.FromArgb(255, 255, 255, 255).ToArgb Then
                        lblCommandCap.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
                        Me.IsDirty = True
                    End If
                    sName &= msCPJunk
                    If lblCommandCap.ToolTipText <> sName Then lblCommandCap.ToolTipText = sName
                End If
        Catch
            'do nothing
        End Try
    End Sub

	Private mfFlashVal As Single = 1.0F
	Private moSW As Stopwatch
	Private Sub frmEnvirDisplay_OnNewFrameEnd() Handles Me.OnNewFrameEnd
		Try
			If dtEndTimer <> Date.MinValue Then
                'If moCPFlashFont Is Nothing = True OrElse moCPFlashFont.Disposed = True Then
                '	Device.IsUsingEventHandlers = False
                '	moCPFlashFont = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 20.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
                '	Device.IsUsingEventHandlers = True
                'End If

				'ok, got a timer to render
				Dim lSeconds As Int32 = CInt(dtEndTimer.Subtract(Now).TotalSeconds)

				If lSeconds < 0 Then
					dtEndTimer = Date.MinValue
                Else
                    Dim oSysFont As New System.Drawing.Font("Microsoft Sans Serif", 20.0F, FontStyle.Bold, GraphicsUnit.Point, 0)
                    Dim sText As String = GetDurationFromSeconds(lSeconds, False)
                    If sText.StartsWith("00:") = True Then sText = sText.Substring(3)
                    Dim rcTemp As Rectangle = BPFont.MeasureString(oSysFont, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter) 'moCPFlashFont.MeasureString(Nothing, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.Red)
                    rcTemp.X = Me.GetAbsolutePosition.X - rcTemp.Width
                    rcTemp.Y = Me.GetAbsolutePosition.Y + (Me.Height \ 2) - (rcTemp.Height \ 2)
                    Dim clrVal As System.Drawing.Color = muSettings.InterfaceBorderColor
                    If lSeconds < 600 Then
                        clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    End If
                    'moCPFlashFont.DrawText(Nothing, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Right, clrVal)
                    BPFont.DrawText(oSysFont, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Right, clrVal)
				End If

            End If

            If (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255) Then
                'goCurrentPlayer.lAccountStatus <> AccountStatusType.eActiveAccount
                Dim sText As String
                If goCurrentPlayer.lAccountStatus <> AccountStatusType.eActiveAccount AndAlso goCurrentPlayer.lAccountStatus <> AccountStatusType.eMondelisActive Then
                    sText = "Trial System. Press F12 to subscribe and play the full version."
                ElseIf goCurrentPlayer.lAccountStatus = AccountStatusType.eMondelisActive Then
                    sText = "After completing the tutorial, press Escape and Self Destruct to go to live."
                ElseIf NewTutorialManager.TutorialOn = False OrElse NewTutorialManager.GetTutorialStepID() > 314 Then
                    sText = "Trial System. Press Escape and Click Self Destruct to spawn in the galaxy."
                Else : sText = ""
                End If
                If sText <> "" Then
                    Dim oSysFont As New System.Drawing.Font("Microsoft Sans Serif", 14.0F, FontStyle.Bold, GraphicsUnit.Point, 0)
                    Dim rcTemp As Rectangle = BPFont.MeasureString(oSysFont, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter)
                    rcTemp.X = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (rcTemp.Width \ 2)
                    rcTemp.Y = Me.GetAbsolutePosition.Y + (Me.Height * 2) + 50 'added 10 to be below alerts.
                    Dim clrVal As System.Drawing.Color = muSettings.InterfaceBorderColor
                    BPFont.DrawText(oSysFont, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Right, clrVal)
                End If
            ElseIf (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.lAccountStatus <> AccountStatusType.eActiveAccount) Then
                Dim sText As String = "Trial Account. Press F12 to subscribe and play the full version."
                Dim oSysFont As New System.Drawing.Font("Microsoft Sans Serif", 14.0F, FontStyle.Bold, GraphicsUnit.Point, 0)
                Dim rcTemp As Rectangle = BPFont.MeasureString(oSysFont, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter)
                rcTemp.X = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (rcTemp.Width \ 2)
                rcTemp.Y = Me.GetAbsolutePosition.Y + (Me.Height * 2) + 50 'added 10 to be below alerts.
                Dim clrVal As System.Drawing.Color = muSettings.InterfaceBorderColor
                BPFont.DrawText(oSysFont, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Right, clrVal)
            End If

            If goCurrentEnvir Is Nothing = False Then
                Dim sText As String = ""
                Dim oSysFont As New System.Drawing.Font("Microsoft Sans Serif", 20.0F, FontStyle.Bold, GraphicsUnit.Point, 0)
                Dim rcTemp As Rectangle
                Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                If goCurrentEnvir.IsPlayerIronCurtain(glPlayerID) = True Then
                    'If moCPFlashFont Is Nothing = True OrElse moCPFlashFont.Disposed = True Then
                    '    Device.IsUsingEventHandlers = False
                    '    moCPFlashFont = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 20.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
                    '    Device.IsUsingEventHandlers = True
                    'End If
                    sText = "INVULNERABILITY ACTIVE"
                    rcTemp = BPFont.MeasureString(oSysFont, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter) 'moCPFlashFont.MeasureString(Nothing, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.Red)
                    rcTemp.X = Me.GetAbsolutePosition.X - rcTemp.Width
                    rcTemp.Y = Me.GetAbsolutePosition.Y + (Me.Height \ 2) - (rcTemp.Height \ 2)

                    'moCPFlashFont.DrawText(Nothing, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Right, clrVal)
                    BPFont.DrawText(oSysFont, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Right, clrVal)
                End If

                'Alert player that their home is a Secure Spawn
                'if GoCurrentEnvir.Ishomeworld=true then
                '    Dim sStarName As String = ""
                '    Dim iObjectID As Int32
                '    If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                '        For x As Int32 = 0 To goGalaxy.mlSystemUB
                '            For y As Int32 = 0 To goGalaxy.moSystems(x).PlanetUB
                '                If goGalaxy.moSystems(x).moPlanets(y).ObjectID = goCurrentEnvir.ObjectID Then
                '                    iObjectID = goGalaxy.moSystems(x).ObjectID
                '                    Exit For
                '                End If
                '            Next y
                '        Next x

                '    Else
                '        iObjectID = goCurrentEnvir.ObjectID
                '    End If
                '    For x As Int32 = 0 To goGalaxy.mlSystemUB
                '        If goGalaxy.moSystems(x).ObjectID = iObjectID Then
                '            sStarName = goGalaxy.moSystems(x).SystemName
                '            Exit For
                '        End If
                '    Next x
                '    If sStarName.EndsWith(" (S)") Then
                '        sText = "Secure Spawn Star"
                '        clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                '        rcTemp = (BPFont.MeasureString(oSysFont, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter)) 'moCPFlashFont.MeasureString(Nothing, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.Red)
                '        rcTemp.X = Me.GetAbsolutePosition.X - rcTemp.Width
                '        rcTemp.Y = Me.GetAbsolutePosition.Y + (Me.Height \ 2) - (rcTemp.Height \ 2) + rcTemp.Height + 5

                '        BPFont.DrawText(oSysFont, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Right, clrVal)
                '    End If
            End If

            If mlPirateWave <> Int32.MinValue Then
                'If moWaveFont Is Nothing = True OrElse moWaveFont.Disposed = True Then
                '	Device.IsUsingEventHandlers = False
                '	moWaveFont = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 30.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
                '	Device.IsUsingEventHandlers = True
                'End If
                Dim oSysFont As New System.Drawing.Font("Microsoft Sans Serif", 30.0F, FontStyle.Bold, GraphicsUnit.Point, 0)
                Dim lA As Int32 = clrPirateWave.A
                If lA = 255 Then
                    If lWaveFullAlphaTime = Int32.MinValue Then
                        lWaveFullAlphaTime = glCurrentCycle
                    Else
                        If glCurrentCycle - lWaveFullAlphaTime > 90 Then
                            lA -= 10
                        End If
                    End If
                ElseIf lWaveFullAlphaTime = Int32.MinValue Then
                    lA += 10
                Else
                    lA -= 10
                End If
                lA = Math.Max(Math.Min(lA, 255), 0)
                If lA = 0 Then
                    mlPirateWave = Int32.MinValue
                    lWaveFullAlphaTime = Int32.MinValue
                Else
                    clrPirateWave = System.Drawing.Color.FromArgb(lA, clrPirateWave.R, clrPirateWave.G, clrPirateWave.B)
                    Dim sText As String = "WAVE " & mlPirateWave
                    Dim rcTemp As Rectangle = BPFont.MeasureString(oSysFont, sText, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter)

                    rcTemp.X = ((MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (rcTemp.Width \ 2)) + 1
                    rcTemp.Y = ((MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (rcTemp.Height \ 2)) + 1
                    BPFont.DrawText(oSysFont, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Center, System.Drawing.Color.FromArgb(lA, 0, 0, 0))
                    rcTemp.X -= 2
                    rcTemp.Y -= 2
                    BPFont.DrawText(oSysFont, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Center, System.Drawing.Color.FromArgb(lA, 255, 255, 255))
                    rcTemp.X += 1
                    rcTemp.Y += 1
                    BPFont.DrawText(oSysFont, sText, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Center, clrPirateWave)
                End If
            End If

            If msImportantMsg Is Nothing = False AndAlso msImportantMsg <> "" Then
                If glCurrentCycle <> mlLastImportantMsgCycle Then
                    mfImportantMsgAlpha -= 9.0F '(9.0F * (glCurrentCycle - mlLastImportantMsgCycle))
                    mlLastImportantMsgCycle = glCurrentCycle
                End If
                If mfImportantMsgAlpha < 0 Then
                    msImportantMsg = ""
                    mfImportantMsgAlpha = 0
                Else
                    Dim oSysFont As New System.Drawing.Font("Microsoft Sans Serif", 30.0F, FontStyle.Bold, GraphicsUnit.Point, 0)
                    Dim rcTemp As Rectangle = BPFont.MeasureString(oSysFont, msImportantMsg, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter)
                    rcTemp.X = ((MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (rcTemp.Width \ 2)) + 1
                    rcTemp.Y = 78
                    BPFont.DrawText(oSysFont, msImportantMsg, rcTemp, DrawTextFormat.VerticalCenter Or DrawTextFormat.Center, System.Drawing.Color.FromArgb(Math.Max(0, Math.Min(255, CInt(mfImportantMsgAlpha))), mclrImportantMsg.R, mclrImportantMsg.G, mclrImportantMsg.B))
                End If
            ElseIf mcolImportantMsgs Is Nothing = False Then
                If mcolImportantMsgs.Count > 0 Then
                    Try
                        Dim lItem As Int32 = -1
                        For Each uTemp As uImportantMsg In mcolImportantMsgs
                            With uTemp
                                lItem = .lIndex
                                msImportantMsg = .sMsg
                                mclrImportantMsg = .clr
                                mfImportantMsgAlpha = .lBaseAlpha
                                mlLastImportantMsgCycle = glCurrentCycle
                            End With
                        Next
                        If lItem <> -1 Then
                            mcolImportantMsgs.Remove("I" & lItem)
                        End If
                    Catch
                        'Stop
                    End Try
                End If
            End If

            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False AndAlso goCurrentPlayer Is Nothing = False Then
                Dim lCPLimit As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eCPLimit)
                If oEnvir.lCPUsage > lCPLimit Then
                    'If moCPFlashFont Is Nothing = True OrElse moCPFlashFont.Disposed = True Then
                    '    Device.IsUsingEventHandlers = False
                    '    moCPFlashFont = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 20.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
                    '    Device.IsUsingEventHandlers = True
                    'End If
                    Dim oSysFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.0F, FontStyle.Bold, GraphicsUnit.Point, 0)

                    Dim rcTemp As Rectangle = BPFont.MeasureString(oSysFont, "Exceeding Command Points", DrawTextFormat.Center Or DrawTextFormat.VerticalCenter)

                    If moSW Is Nothing Then moSW = Stopwatch.StartNew()
                    Dim fElapsed As Single = moSW.ElapsedMilliseconds / 30.0F
                    moSW.Reset()
                    moSW.Start()
                    mfFlashVal -= (fElapsed * 0.1F)
                    If mfFlashVal < 0.0F Then mfFlashVal = 1.0F

                    Dim lTemp As Int32 = oEnvir.lCPUsage - lCPLimit

                    If lTemp < 1 Then Return
                    Dim lG As Int32 = 255
                    Dim lB As Int32 = 128
                    If lTemp > lB Then
                        lB = 0
                        lTemp -= lB
                    Else
                        lB -= lTemp
                        lTemp = 0
                    End If
                    If lTemp <> 0 Then
                        If lTemp > lG Then
                            lG = 0
                        Else
                            lG -= lTemp
                        End If
                    End If

                    If lTemp >= lCPLimit \ 2 Then
                        mfFlashVal = 1.0F
                    End If

                    Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(CInt(255 * mfFlashVal), 255, lG, lB)

                    Dim ofrm As UIWindow = goUILib.GetWindow("frmNotification")
                    Dim lTop As Int32 = 0
                    If ofrm Is Nothing = False Then
                        If ofrm.Top < 20 Then lTop += ofrm.Top + ofrm.Height
                    End If
                    rcTemp.X = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (rcTemp.Width \ 2)
                    rcTemp.Y = lTop
                    BPFont.DrawText(oSysFont, "Exceeding Command Points", rcTemp, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrVal)
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub frmEnvirDisplay_WindowMoved() Handles Me.WindowMoved
        muSettings.EnvirDisplayLocX = Me.Left
        muSettings.EnvirDisplayLocY = Me.Top
    End Sub

	Private Sub btnBombardment_Click(ByVal sName As String) Handles btnBombardment.Click
        If HasAliasedRights(AliasingRights.eChangeBehavior) = False Then
            MyBase.moUILib.AddNotification("You lack rights to change unit behavior.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Exit Sub
        End If
        Dim ofrm As frmBombard = CType(MyBase.moUILib.GetWindow("frmBombard"), frmBombard)
        If ofrm Is Nothing = False Then
            ofrm.Visible = False
            MyBase.moUILib.RemoveWindow(ofrm.ControlName)
            ofrm = Nothing
        Else : ofrm = New frmBombard(goUILib)
        End If
    End Sub

    Public Shared lMyFacilities As Int32 = 0

    Public Sub HandleGetCPPenaltyMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        mlCPPenalty = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lCyclesRemain As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If lCyclesRemain > 0 Then
            mdtCPPenaltyEnds = Now.AddSeconds(lCyclesRemain / 30)

            Dim lNewUB As Int32 = lCnt - 1
            Dim lIDs(lNewUB) As Int32
            Dim yVals(lNewUB) As Byte
            For X As Int32 = 0 To lNewUB
                lIDs(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                yVals(X) = yData(lPos) : lPos += 1
            Next X

            mlCPPenaltyUB = -1
            myPenalty = yVals
            mlDecPlayer = lIDs
            mlCPPenaltyUB = lNewUB
        Else
            mdtCPPenaltyEnds = Date.MinValue
            mlCPPenaltyUB = -1
            ReDim myPenalty(-1)
            ReDim mlDecPlayer(-1)
        End If
    End Sub

    Public Sub btnGoLeft_Click(ByVal sName As String) Handles btnGoLeft.Click
        Try
            If goCurrentEnvir Is Nothing Then Return
            If goCurrentEnvir.oGeoObject Is Nothing Then Return

            Dim oSystem As SolarSystem = Nothing
            Dim oPlanet As Planet = Nothing

            If miWantedObjTypeID = 0 OrElse miWantedObjectID = 0 Then

                'oObj = goCurrentEnvir.oGeoObject
                If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                    oSystem = CType(goCurrentEnvir.oGeoObject, SolarSystem)
                Else
                    oPlanet = CType(goCurrentEnvir.oGeoObject, Planet)
                    If oPlanet Is Nothing Then Return
                    oSystem = oPlanet.ParentSystem
                End If
                miWantedObjTypeID = goCurrentEnvir.ObjTypeID
                miWantedObjectID = goCurrentEnvir.ObjectID
            Else
                If miWantedObjTypeID = ObjectType.eSolarSystem Then
                    'oObj = miWantedObj
                    For x As Int32 = 0 To goGalaxy.mlSystemUB
                        If goGalaxy.moSystems(x).ObjectID = miWantedObjectID Then
                            oSystem = goGalaxy.moSystems(x)
                            Exit For
                        End If
                    Next
                Else

                    For x As Int32 = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB
                        If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(x).ObjectID = miWantedObjectID Then
                            oPlanet = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(x)
                            Exit For
                        End If
                    Next

                    If oPlanet Is Nothing Then Return
                    oSystem = oPlanet.ParentSystem
                    For x As Int32 = 0 To oSystem.PlanetUB
                        If oSystem.moPlanets(x).ObjectID = miWantedObjectID Then
                            oPlanet = oSystem.moPlanets(x)
                            Exit For
                        End If
                    Next
                End If
            End If
            If oSystem Is Nothing Then Return

            If miWantedObjTypeID = ObjectType.eSolarSystem Then
                'goto last planet
                For x As Int32 = 0 To oSystem.PlanetUB
                    Dim sThere As Int32 = GetPlanetNameValue(oSystem.moPlanets(x).PlanetName)
                    If sThere = oSystem.PlanetUB + 1 Then
                        If goCurrentEnvir.ObjTypeID <> ObjectType.ePlanet OrElse (goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.ObjectID <> oSystem.moPlanets(x).ObjectID) Then
                            If frmMain.mbShiftKeyDown = True Then
                                GotoEnvironmentWrapper_local(oSystem.moPlanets(x).ObjectID, ObjectType.ePlanet)
                            Else
                                miWantedObjTypeID = ObjectType.ePlanet
                                miWantedObjectID = oSystem.moPlanets(x).ObjectID
                            End If
                        Else
                            miWantedObjTypeID = 0
                            miWantedObjectID = 0
                        End If
                        Return
                    End If
                Next
            Else
                oSystem = oPlanet.ParentSystem

                Dim sHere As Int32 = GetPlanetNameValue(oPlanet.PlanetName)
                For x As Int32 = 0 To oSystem.PlanetUB
                    Dim sThere As Int32 = GetPlanetNameValue(oSystem.moPlanets(x).PlanetName)
                    If sThere = sHere - 1 Then
                        If goCurrentEnvir.ObjTypeID <> ObjectType.ePlanet OrElse (goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.ObjectID <> oSystem.moPlanets(x).ObjectID) Then
                            If frmMain.mbShiftKeyDown = True Then
                                GotoEnvironmentWrapper_local(oSystem.moPlanets(x).ObjectID, ObjectType.ePlanet)
                            Else
                                miWantedObjTypeID = ObjectType.ePlanet
                                miWantedObjectID = oSystem.moPlanets(x).ObjectID
                            End If
                        Else
                            miWantedObjTypeID = 0
                            miWantedObjectID = 0
                        End If
                        Return
                    End If
                Next
                'Must be the last planet, lets goto space.
                If goCurrentEnvir.ObjTypeID <> ObjectType.eSolarSystem OrElse (goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso goCurrentEnvir.ObjectID <> oSystem.ObjectID) Then
                    If frmMain.mbShiftKeyDown = True Then
                        GotoEnvironmentWrapper_local(oSystem.ObjectID, ObjectType.eSolarSystem)
                    Else
                        miWantedObjTypeID = ObjectType.eSolarSystem
                        miWantedObjectID = oSystem.ObjectID
                    End If
                Else
                    miWantedObjTypeID = 0
                    miWantedObjectID = 0
                End If
            End If
        Catch
        End Try
    End Sub

    Public Sub btnGoRight_Click(ByVal sName As String) Handles btnGoRight.Click
        Try
            If goCurrentEnvir Is Nothing Then Return
            If goCurrentEnvir.oGeoObject Is Nothing Then Return

            Dim oSystem As SolarSystem = Nothing
            Dim oPlanet As Planet = Nothing

            If miWantedObjTypeID = 0 OrElse miWantedObjectID = 0 Then

                'oObj = goCurrentEnvir.oGeoObject
                If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                    oSystem = CType(goCurrentEnvir.oGeoObject, SolarSystem)
                Else
                    oPlanet = CType(goCurrentEnvir.oGeoObject, Planet)
                    If oPlanet Is Nothing Then Return
                    oSystem = oPlanet.ParentSystem
                End If
                miWantedObjTypeID = goCurrentEnvir.ObjTypeID
                miWantedObjectID = goCurrentEnvir.ObjectID
            Else
                If miWantedObjTypeID = ObjectType.eSolarSystem Then
                    'oObj = miWantedObj
                    For x As Int32 = 0 To goGalaxy.mlSystemUB
                        If goGalaxy.moSystems(x).ObjectID = miWantedObjectID Then
                            oSystem = goGalaxy.moSystems(x)
                            Exit For
                        End If
                    Next
                Else

                    For x As Int32 = 0 To goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).PlanetUB
                        If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(x).ObjectID = miWantedObjectID Then
                            oPlanet = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(x)
                            Exit For
                        End If
                    Next

                    If oPlanet Is Nothing Then Return
                    oSystem = oPlanet.ParentSystem
                    For x As Int32 = 0 To oSystem.PlanetUB
                        If oSystem.moPlanets(x).ObjectID = miWantedObjectID Then
                            oPlanet = oSystem.moPlanets(x)
                            Exit For
                        End If
                    Next
                End If
            End If
            If oSystem Is Nothing Then Return

            If miWantedObjTypeID = ObjectType.eSolarSystem Then
                'osystem = CType(oObj, SolarSystem)
                'If oSystem Is Nothing Then Return

                'goto first planet
                For x As Int32 = 0 To oSystem.PlanetUB
                    If Right(oSystem.moPlanets(x).PlanetName, 2) = " I" OrElse Right(oSystem.moPlanets(x).PlanetName, 6) = " Prime" Or oSystem.moPlanets(x).PlanetName = "Gliese 581b" Then
                        If goCurrentEnvir.ObjTypeID <> ObjectType.ePlanet OrElse (goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.ObjectID <> oSystem.moPlanets(x).ObjectID) Then
                            If frmMain.mbShiftKeyDown = True Then
                                GotoEnvironmentWrapper_local(oSystem.moPlanets(x).ObjectID, ObjectType.ePlanet)
                            Else
                                miWantedObjTypeID = ObjectType.ePlanet
                                miWantedObjectID = oSystem.moPlanets(x).ObjectID
                            End If
                        Else
                        End If
                    End If
                Next
            Else
                'Dim oPlanet As Planet = CType(oObj, Planet)
                'If oPlanet Is Nothing Then Return
                oSystem = oPlanet.ParentSystem


                Dim sHere As Int32 = GetPlanetNameValue(oPlanet.PlanetName)
                For x As Int32 = 0 To oSystem.PlanetUB
                    Dim sThere As Int32 = GetPlanetNameValue(oSystem.moPlanets(x).PlanetName)
                    If sThere = sHere + 1 Then
                        If goCurrentEnvir.ObjTypeID <> ObjectType.ePlanet OrElse (goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.ObjectID <> oSystem.moPlanets(x).ObjectID) Then
                            If frmMain.mbShiftKeyDown = True Then
                                GotoEnvironmentWrapper_local(oSystem.moPlanets(x).ObjectID, ObjectType.ePlanet)
                            Else
                                miWantedObjTypeID = ObjectType.ePlanet
                                miWantedObjectID = oSystem.moPlanets(x).ObjectID
                            End If
                        Else
                            miWantedObjTypeID = 0
                            miWantedObjectID = 0
                        End If
                        Return
                    End If

                Next
                'Must be the last planet, lets goto space.
                If goCurrentEnvir.ObjTypeID <> ObjectType.eSolarSystem OrElse (goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso goCurrentEnvir.ObjectID <> oSystem.ObjectID) Then
                    If frmMain.mbShiftKeyDown = True Then
                        GotoEnvironmentWrapper_local(oSystem.ObjectID, ObjectType.eSolarSystem)
                    Else
                        miWantedObjTypeID = ObjectType.eSolarSystem
                        miWantedObjectID = oSystem.ObjectID
                    End If
                Else
                    miWantedObjTypeID = 0
                    miWantedObjectID = 0
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub frmEnvirDisplay_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        Dim ptLoc As Point = Me.GetAbsolutePosition()
        lMouseX -= ptLoc.X
        lMouseY -= ptLoc.Y
        If lMouseX >= lblEnvirName.Left AndAlso lMouseX <= lblEnvirName.Left + lblEnvirName.Width AndAlso lMouseY >= lblEnvirName.Top AndAlso lMouseY <= lblEnvirName.Top + 26 Then
            If lButton = MouseButtons.Left Then
                If miWantedObjTypeID <> 0 AndAlso miWantedObjectID <> 0 Then
                    GotoEnvironmentWrapper_local(miWantedObjectID, miWantedObjTypeID)
                End If
            ElseIf lButton = MouseButtons.Right Then
                miWantedObjTypeID = 0
                miWantedObjectID = 0
            End If
        End If
    End Sub
    Public Sub ChangeEnviroments()
        miWantedObjTypeID = 0
        miWantedObjectID = 0
    End Sub

    Public Sub GotoEnvironmentWrapper_local()
        If miWantedObjTypeID > 0 AndAlso miWantedObjectID > 0 Then
            GotoEnvironmentWrapper_local(miWantedObjectID, miWantedObjTypeID)
        End If
    End Sub
    Private Sub GotoEnvironmentWrapper_local(ByVal ObjectID As Int32, ByVal ObjTypeID As Short)
        If GotoEnvironmentWrapper(ObjectID, ObjTypeID) = False Then
            miWantedObjTypeID = 0
            miWantedObjectID = 0
        End If
    End Sub

    Private Sub btnEnvList_Click(ByVal sName As String) Handles btnEnvList.Click
        'Close Environment Shortcuts if it's open
        Dim ofrmES As frmEnvirShortcuts = CType(goUILib.GetWindow("frmEnvirShortcuts"), frmEnvirShortcuts)
        If Not ofrmES Is Nothing Then
            goUILib.RemoveWindow(ofrmES.ControlName)
            ofrmES.Visible = False
        End If
        ofrmES = Nothing

        'Toggle Open/Close of Environment List
        Dim ofrmED As frmEnvirList = CType(goUILib.GetWindow("frmEnvirList"), frmEnvirList)
        If ofrmED Is Nothing Then
            ofrmED = New frmEnvirList(goUILib)
            ofrmED.Visible = True
        Else
            goUILib.RemoveWindow(ofrmED.ControlName)
            ofrmED.Visible = False
        End If
        ofrmED = Nothing
    End Sub

    Private Sub btnEnvShortcuts_Click(ByVal sName As String) Handles btnEnvShortcuts.Click
        'Close Environment List if it's open
        Dim ofrmEL As frmEnvirList = CType(goUILib.GetWindow("frmEnvirList"), frmEnvirList)
        If Not ofrmEL Is Nothing Then
            goUILib.RemoveWindow(ofrmEL.ControlName)
            ofrmEL.Visible = False
        End If
        ofrmEL = Nothing

        'Toggle Open/Close of Environment Shortcuts
        Dim ofrmES As frmEnvirShortcuts = CType(goUILib.GetWindow("frmEnvirShortcuts"), frmEnvirShortcuts)
        If ofrmES Is Nothing Then
            ofrmES = New frmEnvirShortcuts(goUILib)
            ofrmES.Visible = True
        Else
            goUILib.RemoveWindow(ofrmES.ControlName)
            ofrmES.Visible = False
        End If
        ofrmES = Nothing
    End Sub

    Public Sub AddImportantMsg(ByVal sMsg As String, ByVal clr As System.Drawing.Color, ByVal lBaseAlpha As Int32)
        Try
            If mcolImportantMsgs Is Nothing Then mcolImportantMsgs = New Collection

            Dim uNew As uImportantMsg
            With uNew
                .clr = clr
                .sMsg = sMsg
                Dim lVal As Int32 = mlImportantMsgIndex
                mlImportantMsgIndex += 1
                .lIndex = lVal
                .lBaseAlpha = lBaseAlpha
            End With
            mcolImportantMsgs.Add(uNew, "I" & uNew.lIndex)
        Catch
            '            Stop
        End Try

    End Sub
End Class

