Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmQuickBar
    Inherits UIWindow

    Private mbFlashes() As Boolean
    Private mbFlashState As Boolean = False
    Private mlLastFlashStateUpdate As Int32 = 0
    Private rcSrc() As Rectangle
    Private msToolTips() As String

    Private WithEvents btnPageLeft As UIButton
    Private lblPageNum As UILabel
    Private WithEvents btnPageRight As UIButton
    Private miPageNumber As Int32 = 1
    Private miPages As Int32 = 2

    Private mbPrevHasFlashes As Boolean = False

    Private Const lButtonCount As Int32 = 31  '15

    Private Shared moSprite As Sprite
    Public Shared Sub ReleaseSprite()
        If moSprite Is Nothing = False Then moSprite.Dispose()
        moSprite = Nothing
    End Sub

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
        If isAdmin() = False Then miPages = 1
        'frmQuickBar initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eQuickBar
            .ControlName = "frmQuickBar"
            .Left = 1
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            .Width = 64
            .Height = 255
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .DrawBorder = False
            .FillColor = Color.Black
            .FullScreen = False
            .BorderLineWidth = 1
            .bRoundedBorder = False
        End With

        ReDim mbFlashes(lButtonCount)
        For X As Int32 = 0 To mbFlashes.GetUpperBound(0)
            mbFlashes(X) = False
        Next X

        ReDim rcSrc(lButtonCount)
        rcSrc(0) = grc_UI(elInterfaceRectangle.eQuickbar_Help)
        rcSrc(1) = grc_UI(elInterfaceRectangle.eQuickbar_Email)
        rcSrc(2) = grc_UI(elInterfaceRectangle.eQuickbar_Battlegroup)
        rcSrc(3) = grc_UI(elInterfaceRectangle.eQuickbar_Trade)
        rcSrc(4) = grc_UI(elInterfaceRectangle.eQuickbar_Diplomacy)
        rcSrc(5) = grc_UI(elInterfaceRectangle.eQuickbar_ColonyStats)
        rcSrc(6) = grc_UI(elInterfaceRectangle.eQuickbar_Budget)
        rcSrc(7) = grc_UI(elInterfaceRectangle.eQuickbar_Mining)
        rcSrc(8) = grc_UI(elInterfaceRectangle.eQuickbar_Agent)
        rcSrc(9) = grc_UI(elInterfaceRectangle.eQuickbar_Formations)
        rcSrc(10) = grc_UI(elInterfaceRectangle.eQuickbar_ColonyResearch)
        rcSrc(11) = grc_UI(elInterfaceRectangle.eQuickbar_AvailResources)
        rcSrc(12) = grc_UI(elInterfaceRectangle.eQuickbar_Command)
        rcSrc(13) = grc_UI(elInterfaceRectangle.eQuickBar_Senate)
        rcSrc(14) = grc_UI(elInterfaceRectangle.eQuickbar_ChatConfig)
        rcSrc(15) = grc_UI(elInterfaceRectangle.eQuickBar_Transports)
        rcSrc(16) = grc_UI(elInterfaceRectangle.eQuickBar_Guild) 'senate
        rcSrc(17) = grc_UI(elInterfaceRectangle.eQuickbar_RouteTemplate) 'format
        rcSrc(18) = grc_UI(elInterfaceRectangle.eQuickbar_EnvironmentRelations)

        ReDim msToolTips(lButtonCount)
        msToolTips(0) = "Open Tutorial Window (F1)"
        msToolTips(1) = "Open Email Window (F2)"
        msToolTips(2) = "Open Battlegroup Window (F3)"
        msToolTips(3) = "Open Trade Window (F4)"
        msToolTips(4) = "Open Diplomacy Window (F5)"
        msToolTips(5) = "Open Colony Window (F6)"
        msToolTips(6) = "Open Budget Window (F7)"
        msToolTips(7) = "Open Mining Window (F8)"
        msToolTips(8) = "Open Agent Window (F9)"
        msToolTips(9) = "Open Formations Window (F10)"
        msToolTips(10) = "Open Colony Research Queue Window"
        msToolTips(11) = "Open Available Resources Window"
        msToolTips(12) = "Open Command Window"
        msToolTips(13) = "Open Senate Window"
        msToolTips(14) = "Open Chat Channel Configuration Window"
        msToolTips(15) = "Open Transport Management Window"
        msToolTips(16) = "Open Guild Mgmt Window"
        msToolTips(17) = "Open Route Templates Window"
        msToolTips(18) = "Open Environment Relations Manager"

        'btnPageLeft initial props
        btnPageLeft = New UIButton(oUILib)
        With btnPageLeft
            .ControlName = "btnPageLeft"
            .Left = 0
            .Width = 25
            .Height = 12
            .Top = Me.Height - 13
            .Enabled = True
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = "<"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnPageLeft, UIControl))

        lblPageNum = New UIButton(oUILib)
        With lblPageNum
            .ControlName = "lblPageNum"
            .Width = 10
            .Height = 16
            .Left = 26
            .Top = Me.Height - .Height
            .Enabled = True
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = "1"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(lblPageNum, UIControl))

        'btnPageRight initial props
        btnPageRight = New UIButton(oUILib)
        With btnPageRight
            .ControlName = "btnPageRight"
            .Width = 25
            .Height = 12
            .Left = 64 - 26
            .Top = Me.Height - 13
            .Enabled = True
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = ">"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnPageRight, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub frmQuickBar_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        Dim lTmpY As Int32 = lMouseY - Me.GetAbsolutePosition.Y
        Dim lTmpX As Int32 = lMouseX - Me.GetAbsolutePosition.X
        If lTmpY > 31 * 8 Then
            Exit Sub
        End If
        lTmpY \= 31
        lTmpX \= 32

        frmMain.mbIgnoreMouseMove = True
        Select Case miPageNumber
            Case 1
                Select Case lTmpY
                    Case 0
                        If lTmpX = 0 Then
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Help"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmTutorialTOC = CType(goUILib.GetWindow("frmTutorialTOC"), frmTutorialTOC)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmTutorialTOC(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing
                            'mbFlashes(0) = False
                        Else

                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Email"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmEmailMain = CType(goUILib.GetWindow("frmEmailMain"), frmEmailMain)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmEmailMain(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing
                            'mbFlashes(1) = False
                        End If
                    Case 1
                        If lTmpX = 0 Then
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Fleet"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmFleet = CType(goUILib.GetWindow("frmFleet"), frmFleet)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmFleet(goUILib, -1)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing
                            'mbFlashes(2) = False
                        Else
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Trade"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim oTmpWin As UIWindow = Nothing
                            If goUILib Is Nothing = False Then
                                oTmpWin = goUILib.GetWindow("frmTradeMain")
                                If oTmpWin Is Nothing = False Then
                                    goUILib.RemoveWindow(oTmpWin.ControlName)
                                Else : oTmpWin = New frmTradeMain(goUILib, -1)
                                End If
                                oTmpWin = Nothing
                            End If
                            If mbFlashes(3) = True AndAlso oTmpWin Is Nothing = False Then
                                'ok, highlight one that has the reason for it...
                                Try
                                    For X As Int32 = 0 To goCurrentPlayer.mlTradeUB
                                        If (goCurrentPlayer.moTrades(X).yTradeState And (Trade.eTradeStateValues.TradeCompleted Or Trade.eTradeStateValues.TradeRejected)) = 0 Then
                                            If goCurrentPlayer.moTrades(X).lPlayer1ID = glPlayerID Then
                                                CType(oTmpWin, frmTradeMain).SetTradePostID(goCurrentPlayer.moTrades(X).lP1TradepostID)
                                            Else
                                                CType(oTmpWin, frmTradeMain).SetTradePostID(goCurrentPlayer.moTrades(X).lP2TradepostID)
                                            End If

                                        End If
                                    Next X
                                Catch
                                End Try
                            End If
                            'mbFlashes(3) = False
                            'Clear any pending TradePost 'busy' indicators so followup new trades this session can alert
                            If (MyBase.moUILib.GetMsgSys.mlBusyStatus And elBusyStatusType.TradesNotification) <> 0 Then
                                MyBase.moUILib.GetMsgSys.mlBusyStatus = MyBase.moUILib.GetMsgSys.mlBusyStatus Xor elBusyStatusType.TradesNotification
                            End If
                        End If
                    Case 2
                        If lTmpX = 0 Then
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Diplomacy"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim oTmpWin As UIWindow = goUILib.GetWindow("frmDiplomacy")
                            If oTmpWin Is Nothing = False Then
                                goUILib.RemoveWindow(oTmpWin.ControlName)
                                If glCurrentEnvirView = CurrentView.eDiplomacyScreen Then ReturnToPreviousView()
                            Else
                                oTmpWin = New frmDiplomacy(goUILib)
                                CType(oTmpWin, frmDiplomacy).SetFromCurrentPlayer()
                            End If
                            oTmpWin = Nothing
                            'mbFlashes(4) = False
                        Else
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"ColonyStats"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim oCStatWin As UIWindow
                            Dim sTemp As String = "frmColonyStatsSmall"
                            If goUILib Is Nothing = False Then
                                If muSettings.ExpandedColonyStatsScreen = True Then sTemp = "frmColonyStats"
                                oCStatWin = goUILib.GetWindow(sTemp)
                                If oCStatWin Is Nothing = False Then
                                    If oCStatWin.Visible = True Then
                                        oCStatWin.Visible = False
                                        goUILib.RemoveWindow(oCStatWin.ControlName)
                                    Else : oCStatWin.Visible = True
                                    End If
                                ElseIf muSettings.ExpandedColonyStatsScreen = True Then
                                    oCStatWin = New frmColonyStats(goUILib)
                                Else
                                    oCStatWin = New frmColonyStatsSmall(goUILib)
                                End If
                                oCStatWin = Nothing
                            End If
                            'mbFlashes(5) = False
                        End If
                    Case 3
                        If lTmpX = 0 Then
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Budget"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim oTmpWin As UIWindow
                            If goUILib Is Nothing = False Then
                                oTmpWin = goUILib.GetWindow("frmBudget")
                                If oTmpWin Is Nothing = False Then
                                    goUILib.RemoveWindow(oTmpWin.ControlName)
                                Else : oTmpWin = New frmBudget(goUILib)
                                End If
                                oTmpWin = Nothing
                            End If
                            'mbFlashes(6) = False
                        Else
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Mining"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmMining = CType(goUILib.GetWindow("frmMining"), frmMining)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmMining(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing
                            'mbFlashes(7) = False
                        End If
                    Case 4
                        If lTmpX = 0 Then
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Agent"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmAgentMain = CType(goUILib.GetWindow("frmAgentMain"), frmAgentMain)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmAgentMain(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing
                            'mbFlashes(8) = False
                        Else
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Formations"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmFormations = CType(goUILib.GetWindow("frmFormations"), frmFormations)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmFormations(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing
                            'mbFlashes(9) = False
                        End If
                    Case 5
                        If lTmpX = 0 Then
                            'Colony Research Manager
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"ColonyResearch"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmColonyResearch = CType(goUILib.GetWindow("frmColonyResearch"), frmColonyResearch)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmColonyResearch(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing
                            'mbFlashes(10) = False
                        Else
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"AvailableResources"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmAvailRes = CType(goUILib.GetWindow("frmAvailRes"), frmAvailRes)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                'New frmAvailRes(MyBase.moUILib, False, CType(goCurrentEnvir.oEntity(mlEntityIndex), Base_GUID))
                                Dim bFound As Boolean = False
                                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                                If oEnvir Is Nothing = False Then
                                    If oEnvir.ObjTypeID = ObjectType.ePlanet Then
                                        bFound = True
                                    ElseIf oEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                        For X As Int32 = 0 To oEnvir.lEntityUB
                                            If oEnvir.lEntityIdx(X) > -1 Then
                                                Dim oTmp As BaseEntity = oEnvir.oEntity(X)
                                                If oTmp Is Nothing = False Then
                                                    If oTmp.OwnerID = glPlayerID AndAlso oTmp.bSelected = True Then
                                                        If oTmp.yProductionType = ProductionType.eSpaceStationSpecial OrElse oTmp.yProductionType = ProductionType.eTradePost Then
                                                            bFound = True
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next X
                                    End If
                                End If
                                If bFound = False Then Return
                                ofrm = New frmAvailRes(goUILib, False, Nothing)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing
                            'mbFlashes(11) = False
                        End If
                    Case 6
                        If lTmpX = 0 Then
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Command"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            'Facility and Unit List Management
                            Dim ofrm As frmCommand = CType(goUILib.GetWindow("frmCommand"), frmCommand)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmCommand(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing

                            'mbFlashes(12) = False
                        Else
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"Senate"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmSenate = CType(goUILib.GetWindow("frmSenate"), frmSenate)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmSenate(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing

                            'mbFlashes(13) = False
                        End If
                    Case 7
                        If lTmpX = 0 Then
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"CHATCHANNELS"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            'Facility and Unit List Management
                            Dim ofrm As frmChannels = CType(goUILib.GetWindow("frmChannels"), frmChannels)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmChannels(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing

                            'mbFlashes(14) = False
                        Else
                            If NewTutorialManager.TutorialOn = True Then
                                Dim sParms() As String = {"TRANSPORTMGMT"}
                                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                            End If

                            Dim ofrm As frmTransportManagement = CType(goUILib.GetWindow("frmTransportManagement"), frmTransportManagement)
                            If ofrm Is Nothing = False Then
                                If ofrm.Visible = True Then
                                    goUILib.RemoveWindow(ofrm.ControlName)
                                Else : ofrm.Visible = True
                                End If
                            Else
                                ofrm = New frmTransportManagement(goUILib)
                                ofrm.Visible = True
                            End If
                            ofrm = Nothing
                            'mbFlashes(15) = False
                        End If
                    Case Else
                        Return
                End Select
            Case 2
                Select Case lTmpY
                    Case 0
                        If lTmpX = 0 Then 'Guild Window
                            OpenWindow_Guild()
                        Else ' Route Templates
                            Dim oFrm As frmRouteTemplate = CType(MyBase.moUILib.GetWindow("frmRouteTemplate"), frmRouteTemplate)
                            If oFrm Is Nothing = False Then
                                If oFrm.Visible = True Then
                                    goUILib.RemoveWindow(oFrm.ControlName)
                                Else : oFrm.Visible = True
                                End If
                            Else
                                oFrm = New frmRouteTemplate(goUILib)
                                oFrm.Visible = True
                            End If
                        End If
                    Case 1
                        If lTmpX = 0 Then 'Environment Relations
                            Dim oFrm As frmEnvirRelations = CType(MyBase.moUILib.GetWindow("frmEnvirRelations"), frmEnvirRelations)
                            If oFrm Is Nothing = False Then
                                If oFrm.Visible = True Then
                                    goUILib.RemoveWindow(oFrm.ControlName)
                                Else : oFrm.Visible = True
                                End If
                            Else
                                oFrm = New frmEnvirRelations(goUILib)
                                oFrm.Visible = True
                            End If
                        Else
                        End If
                    Case 2
                        If lTmpX = 0 Then
                        Else
                        End If
                    Case 3
                        If lTmpX = 0 Then
                        Else
                        End If
                    Case 4
                        If lTmpX = 0 Then
                        Else
                        End If
                    Case 5
                        If lTmpX = 0 Then
                        Else
                        End If
                    Case 6
                        If lTmpX = 0 Then
                            If isAdmin() = False Then Return
                            GFXEngine.gbDeviceLostHard = True
                            GFXEngine.gbDeviceLost = True
                        Else
                        End If
                    Case 7
                        If lTmpX = 0 Then
                            If isAdmin() = False Then Return
                            GFXEngine.gbDeviceLost = True
                        Else
                            If isAdmin() = False Then Return
                            goUILib.RemoveWindow(Me.ControlName)
                            Dim ofrmQB As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                            If ofrmQB Is Nothing Then ofrmQB = New frmQuickBar(goUILib)
                            ofrmQB = Nothing
                            Exit Sub
                        End If
                    Case Else
                        Return
                End Select
            Case Else
                Return
        End Select
        mbFlashes(((miPageNumber - 1) * 16) + (lTmpY * 2) + lTmpX) = False

        If goSound Is Nothing = False Then
            goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
        End If

        Me.IsDirty = True
    End Sub

    Private Sub frmQuickBar_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim lTmpY As Int32 = lMouseY - Me.GetAbsolutePosition.Y
        If lTmpY > 31 * 8 Then
            Exit Sub
        End If
        lTmpY \= 31

        Dim lTmpX As Int32 = lMouseX - Me.GetAbsolutePosition.X
        lTmpX \= 32

        Dim lIdx As Int32 = (lTmpY * 2 + lTmpX) + ((miPageNumber - 1) * 16)

        If lIdx < 0 OrElse lIdx > msToolTips.GetUpperBound(0) OrElse msToolTips(lIdx) Is Nothing = True Then Return

        msToolTips(1) = "Open Email Window (F2)" & vbCrLf & "Unread Messages: " & frmEmailMain.lCurrentUnreadMessages

        MyBase.moUILib.SetToolTip(msToolTips(lIdx), lMouseX, lMouseY)
    End Sub

    Private Sub frmQuickBar_OnNewFrame() Handles Me.OnNewFrame
        'for flash effects
        If glCurrentCycle - mlLastFlashStateUpdate > 15 Then

            Dim bHasFlash As Boolean = False
            For X As Int32 = 0 To mbFlashes.GetUpperBound(0)
                If mbFlashes(X) = True Then
                    'Debug.Print("set bHasFlash [" & X.ToString & "]= True")
                    bHasFlash = True
                    Exit For
                End If
            Next X
            'Debug.Print("read  bHasFlash = " & bHasFlash.ToString)

            mlLastFlashStateUpdate = glCurrentCycle
            mbFlashState = Not mbFlashState
            If bHasFlash = True OrElse bHasFlash <> mbPrevHasFlashes Then Me.IsDirty = True
            mbPrevHasFlashes = bHasFlash
        End If
    End Sub

    Private Sub frmQuickBar_OnRenderEnd() Handles Me.OnRenderEnd
        If moSprite Is Nothing = True OrElse moSprite.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moSprite = New Sprite(MyBase.moUILib.oDevice)
            Device.IsUsingEventHandlers = True
        End If

        moSprite.Begin(SpriteFlags.AlphaBlend)

        Dim oTmpTex As Texture = goResMgr.GetTexture("Flare2.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "xpl.pak")
        If oTmpTex Is Nothing = False Then moSprite.Draw2D(oTmpTex, System.Drawing.Rectangle.Empty, New Rectangle(0, 0, Me.Width, Me.Height), Point.Empty, 0, New Point(0, 0), System.Drawing.Color.FromArgb(255, Math.Min(255, muSettings.InterfaceFillColor.R * 2), Math.Min(255, muSettings.InterfaceFillColor.G * 2), Math.Min(255, muSettings.InterfaceFillColor.B * 2)))

        If mbFlashes Is Nothing = False Then
            Dim rcDest As Rectangle = New Rectangle(0, 0, 32, 31)
            For X As Int32 = 0 + ((miPageNumber - 1) * 16) To 15 + ((miPageNumber - 1) * 16)
                If rcSrc(X).IsEmpty = False AndAlso (mbFlashes(X) = False OrElse mbFlashState = False) Then
                    Dim lLeftVal As Int32 = 0
                    If X Mod 2 = 1 Then lLeftVal = 32
                    Dim lTopVal As Int32 = (X \ 2) * 31 - ((8 * 31) * (miPageNumber - 1))
                    Dim ptLoc As Point = New Point(lLeftVal, lTopVal)
                    rcDest.X = lLeftVal
                    If MyBase.moUILib.oInterfaceTexture Is Nothing = False Then moSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc(X), rcDest, Point.Empty, 0, ptLoc, System.Drawing.Color.White)
                End If
            Next X
        End If
        moSprite.End()
    End Sub

    Public Sub NewEmailHasArrived()
        If goUILib Is Nothing = False Then
            Dim ofrm As frmEmailMain = CType(goUILib.GetWindow("frmEmailMain"), frmEmailMain)
            If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                mbFlashes(1) = False
                Return
            End If
            ofrm = Nothing
        End If


        mbFlashes(1) = True
        If goSound Is Nothing = False Then
            goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
        End If
    End Sub

    Public Sub TradeAgreementUpdate()
        If goUILib Is Nothing = False Then
            Dim ofrm As frmTradeMain = CType(goUILib.GetWindow("frmTradeMain"), frmTradeMain)
            If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                mbFlashes(3) = False
                Return
            End If
            ofrm = Nothing
        End If
        mbFlashes(3) = True
    End Sub

    Public Sub SenateUpdate()
        If goUILib Is Nothing = False Then
            Dim ofrm As frmSenate = CType(goUILib.GetWindow("frmSenate"), frmSenate)
            If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                mbFlashes(13) = False
                Return
            End If
            ofrm = Nothing
        End If
        mbFlashes(13) = True
    End Sub

    Public Sub DiplomacyUpdate()
        If goUILib Is Nothing = False Then
            Dim ofrm As frmDiplomacy = CType(goUILib.GetWindow("frmDiplomacy"), frmDiplomacy)
            If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                mbFlashes(4) = False
                Return
            End If
            ofrm = Nothing
        End If
        mbFlashes(4) = True
    End Sub

    Public Sub AgentEventUpdate()
        If goUILib Is Nothing = False Then
            Dim ofrm As frmAgentMain = CType(goUILib.GetWindow("frmAgentMain"), frmAgentMain)
            If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                mbFlashes(8) = False
                Return
            End If
            ofrm = Nothing
        End If
        mbFlashes(8) = True
    End Sub

    Public Sub ClearFlashState(ByVal yValue As Byte)
        mbFlashes(yValue) = False
        'Clear any pending TradePost 'busy' indicators so followup new trades this session can alert
        If yValue = 3 AndAlso (MyBase.moUILib.GetMsgSys.mlBusyStatus And elBusyStatusType.TradesNotification) <> 0 Then
            MyBase.moUILib.GetMsgSys.mlBusyStatus = MyBase.moUILib.GetMsgSys.mlBusyStatus Xor elBusyStatusType.TradesNotification
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If moSprite Is Nothing = False Then moSprite.Dispose()
        moSprite = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub btnPageLeft_Click(ByVal sName As String) Handles btnPageLeft.Click
        miPageNumber -= 1
        If miPageNumber < 1 Then miPageNumber = 1
        If lblPageNum.Caption <> miPageNumber.ToString Then lblPageNum.Caption = miPageNumber.ToString
    End Sub

    Private Sub btnPageRight_Click(ByVal sName As String) Handles btnPageRight.Click
        miPageNumber += 1
        If miPageNumber > miPages Then miPageNumber = miPages
        If lblPageNum.Caption <> miPageNumber.ToString Then lblPageNum.Caption = miPageNumber.ToString
    End Sub
End Class
