Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAvailRes
    Inherits UIWindow

    Private lblTitle As UILabel
    Private txtData As UITextBox
    Private WithEvents btnRefresh As UIButton
    Private WithEvents btnClose As UIButton

    Private mlLastUpdate As Int32 = -1
    Private mlDelayedUpdate As Int32 = -1

    Private moGUID As Base_GUID = Nothing
    Private mlRefreshReenable As Int32 = -1

    Private mbLoading As Boolean = True
    Private mbUnloading As Boolean = False

    Public Sub New(ByRef oUILib As UILib, ByVal bSubForm As Boolean, ByRef oGUID As Base_GUID)
        MyBase.New(oUILib)

        moGUID = oGUID
        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenAvailableResourcesWindow)

        'frmAvailRes initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAvailableResources
            .ControlName = "frmAvailRes"
            '.Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            '.Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.AvailResX
                lTop = muSettings.AvailResY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            If lLeft + 256 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 256
            If lTop + 256 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 256

            .Left = lLeft
            .Top = lTop

            .Width = 256
            .Height = 256
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True  'so it shows up in full screen interfaces
            .Moveable = True
            .bRoundedBorder = True
            .BorderLineWidth = 2
            .mbAcceptReprocessEvents = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 2
            .Width = 147
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Available Resources"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'txtData initial props
        txtData = New UITextBox(oUILib)
        With txtData
            .ControlName = "txtData"
            .Left = 5
            .Top = 26
            .Width = 246
            .Height = 200
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            '.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .SetFont(New System.Drawing.Font("Courier New", 8.5F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .MultiLine = True
            .Locked = True
        End With
        Me.AddChild(CType(txtData, UIControl))

        'btnRefresh initial props
        btnRefresh = New UIButton(oUILib)
        With btnRefresh
            .ControlName = "btnRefresh"
            .Left = 78
            .Top = 230
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Refresh"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRefresh, UIControl))

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
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 14.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
            .ToolTipText = "Click to close this window."
        End With
		Me.AddChild(CType(btnClose, UIControl))
		MyBase.moUILib.RemoveWindow(Me.ControlName)

		If HasAliasedRights(AliasingRights.eViewBudget Or AliasingRights.eAddProduction) = True Then
			MyBase.moUILib.AddWindow(Me)
		Else
			Me.Visible = False
			goUILib.AddNotification("You lack the rights to view the available resources interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If

		If bSubForm = False Then
			MyBase.moUILib.RemoveWindow(Me.ControlName)
			MyBase.moUILib.AddWindow(Me)
        End If

        mbLoading = False

        'Go ahead and get a refresh here
        'goAvailableResources = Nothing
        'btnRefresh_Click(btnRefresh.ControlName)
    End Sub

    Private Sub frmAvailRes_OnNewFrame() Handles Me.OnNewFrame
        If goAvailableResources Is Nothing = False AndAlso goAvailableResources.LastUpdate <> mlLastUpdate Then
            mlLastUpdate = goAvailableResources.LastUpdate
            txtData.Caption = goAvailableResources.GetTextBoxText()
        ElseIf glCurrentCycle - mlDelayedUpdate > 30 Then
            mlDelayedUpdate = glCurrentCycle
            If goAvailableResources Is Nothing = False Then
                'Automatic Refresh Disabled (04/03/2010)  Too much needless load, and packed planets were chaos
                'Dim sTemp As String = goAvailableResources.GetTextBoxText()
                'If txtData.Caption <> sTemp Then txtData.Caption = sTemp
            Else
                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                If oEnvir Is Nothing = False Then
                    Dim bFound As Boolean = False
                    If oEnvir.ObjTypeID = ObjectType.ePlanet Then
                        bFound = True
                        btnRefresh_Click("btnRefresh")

                    ElseIf oEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                        Try
                            For X As Int32 = 0 To oEnvir.lEntityUB
                                If oEnvir.lEntityIdx(X) > -1 Then
                                    Dim oTmp As BaseEntity = oEnvir.oEntity(X)
                                    If oTmp Is Nothing = False Then
                                        If oTmp.OwnerID = glPlayerID AndAlso oTmp.bSelected = True Then
                                            If oTmp.yProductionType = ProductionType.eSpaceStationSpecial OrElse oTmp.yProductionType = ProductionType.eTradePost Then
                                                btnRefresh_Click("btnRefresh")
                                                bFound = True
                                            End If
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next X
                        Catch
                        End Try
                    End If
                    If bFound = False Then
                        MyBase.moUILib.RemoveWindow(Me.ControlName)
                    End If
                End If
            End If
        End If

        If goAvailableResources Is Nothing = False AndAlso goCurrentEnvir Is Nothing = False Then
            If goAvailableResources.lColonyParentID <> goCurrentEnvir.ObjectID OrElse goAvailableResources.iColonyParentTypeID <> goCurrentEnvir.ObjTypeID Then
                Dim bFound As Boolean = False
                Dim bRefreshAvailRes As Boolean = True
                'If goAvailableResources.iColonyParentTypeID = ObjectType.eFacility AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                'ok, check if the selected facility is a space station
                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                If oEnvir Is Nothing = False Then
                    Try
                        For X As Int32 = 0 To oEnvir.lEntityUB
                            If oEnvir.lEntityIdx(X) > -1 Then
                                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                                If oEntity Is Nothing = False AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.bSelected = True Then
                                    If oEntity.yProductionType = ProductionType.eSpaceStationSpecial OrElse oEntity.yProductionType = ProductionType.eTradePost Then
                                        bFound = True
                                        If goAvailableResources.lColonyParentID = oEntity.ObjectID Then
                                            bRefreshAvailRes = False
                                        End If
                                    End If
                                    Exit For
                                End If
                            End If
                        Next X
                    Catch
                    End Try
                    'end If
                End If
                If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso bFound = False AndAlso bRefreshAvailRes = True Then
                    MyBase.moUILib.RemoveWindow(Me.ControlName)
                    Return
                End If

                If bRefreshAvailRes = True AndAlso mbUnloading = False Then
                    goAvailableResources.ClearResources()
                    goAvailableResources = Nothing
                End If
            End If
        End If

        If glCurrentCycle > mlRefreshReenable AndAlso btnRefresh.Enabled = False Then
            If btnRefresh.Caption <> "Refresh" Then btnRefresh.Caption = "Refresh"
            btnRefresh.Enabled = True
        End If
    End Sub

	Private Sub btnRefresh_Click(ByVal sName As String) Handles btnRefresh.Click
		Dim yMsg(7) As Byte
		mlRefreshReenable = glCurrentCycle + 60
        btnRefresh.Enabled = False
        If btnRefresh.Caption <> "Refreshing..." Then btnRefresh.Caption = "Refreshing..."
		System.BitConverter.GetBytes(GlobalMessageCode.eGetAvailableResources).CopyTo(yMsg, 0)
        If moGUID Is Nothing = False Then
            'Verify our moGUID
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                If moGUID.ObjTypeID <> oEnvir.ObjTypeID Then
                    If moGUID.ObjTypeID = ObjectType.eFacility Then
                        If oEnvir.ObjTypeID <> ObjectType.eSolarSystem Then
                            moGUID = Nothing
                            btnRefresh_Click(sName)
                            Return
                        Else
                            Try
                                For X As Int32 = 0 To oEnvir.lEntityUB
                                    If oEnvir.lEntityIdx(X) = moGUID.ObjectID Then
                                        Dim oTmp As BaseEntity = oEnvir.oEntity(X)
                                        If oTmp Is Nothing = False AndAlso oTmp.ObjectID = moGUID.ObjectID AndAlso oTmp.ObjTypeID = moGUID.ObjTypeID Then
                                            If oTmp.bSelected = False Then
                                                moGUID = Nothing
                                                btnRefresh_Click(sName)
                                                Return
                                            End If
                                            Exit For
                                        End If
                                    End If
                                Next X
                            Catch
                            End Try
                        End If
                    End If
                End If
            End If
            'MyBase.moUILib.AddNotification("Using current moGUID", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            moGUID.GetGUIDAsString.CopyTo(yMsg, 2)
        Else
            If goAvailableResources Is Nothing = False Then
                System.BitConverter.GetBytes(goAvailableResources.lColonyID).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(ObjectType.eColony).CopyTo(yMsg, 6)
                'MyBase.moUILib.AddNotification("Using current AvailRes", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Else
                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                Dim oTmpGuid As Base_GUID = oEnvir
                If oEnvir Is Nothing = False Then
                    'MyBase.moUILib.AddNotification("Using current Envir: " & oTmpGuid.ObjectID & ", " & oTmpGuid.ObjTypeID, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If oEnvir.ObjTypeID = ObjectType.eSolarSystem Then

                        Try
                            For X As Int32 = 0 To oEnvir.lEntityUB
                                If oEnvir.lEntityIdx(X) > -1 Then
                                    Dim oTmp As BaseEntity = oEnvir.oEntity(X)
                                    If oTmp Is Nothing = False Then
                                        If oTmp.OwnerID = glPlayerID AndAlso oTmp.bSelected = True Then
                                            If oTmp.yProductionType = ProductionType.eSpaceStationSpecial OrElse oTmp.yProductionType = ProductionType.eTradePost Then
                                                oTmpGuid = oTmp
                                                moGUID = oTmpGuid
                                            End If
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next X
                        Catch
                        End Try
                    End If
                    oTmpGuid.GetGUIDAsString.CopyTo(yMsg, 2)
                Else : Return
                End If

            End If
        End If
		MyBase.moUILib.SendMsgToPrimary(yMsg)

		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eAvailableResourcesRefreshClick, -1, -1, -1, "")
		End If
	End Sub

    Public Sub SetGUID(ByRef oGUID As Base_GUID)
        moGUID = oGUID
    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub frmAvailRes_WindowClosed() Handles Me.WindowClosed
        If Not mbLoading Then mbUnloading = True
    End Sub

    Private Sub frmAvailRes_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.AvailResX = Me.Left
            muSettings.AvailResY = Me.Top
        End If
    End Sub
End Class