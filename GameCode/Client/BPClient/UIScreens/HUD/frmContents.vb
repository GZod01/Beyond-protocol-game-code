Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Hangar or Cargo Bay contents form
Public Class frmContents
    Inherits UIWindow

    Private Const ml_CONTENTS_REQUEST_DELAY As Int32 = 5000

    'Private WithEvents lstData As UIListBox
    Private WithEvents tvwData As UITreeView
    Private lblTitle As UILabel
    Private lblCapacity As UILabel
    Private WithEvents btnLaunch As UIButton
    Private WithEvents btnLaunchAll As UIButton
    Private WithEvents btnSwitch As UIButton
    Private WithEvents btnLaunchToReinforce As UIButton
    Private lblQuantity As UILabel
    Private txtQuantity As UITextBox
    Private WithEvents btnCancelLaunch As UIButton

    Private mlEntityIndex As Int32 = -1
    Private miEntityTypeID As Int16

    Private myCurrView As Byte = 0      'Hangar
    Private mlPrevCnt As Int32
    Private myPrevView As Byte = 0

    Private mlLastContentUpdate As Int32 = -1

    Private msw_Delay As Stopwatch
    Private mbForceRefresh As Boolean

    Private mlLastStationRqst As Int32 = 0
    Private mlLastSecondUpdate As Int32 = -1
    Private mbNeedSecondUpdates As Boolean = False

    Private mbTransferDisplayed As Boolean = False
    Private mbRepairDisplayed As Boolean = False
    'Private mbAmmoDisplayed As Boolean = False
    Private mbHasNoHangar As Boolean = False
    Private mbHasNoCargo As Boolean = False
    Private mbHasUnknowns As Boolean = False
    Private mlTimeOpened As Int32 = -1


    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenContentsWindow)

        'frmContents initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eContents
            .ControlName = "frmContents"

            Dim lLeft As Int32 = 128
            Dim lTop As Int32 = oUILib.oDevice.PresentationParameters.BackBufferHeight - 512

            Dim oFrm As UIWindow = oUILib.GetWindow("frmSingleSelect")
            If oFrm Is Nothing = False Then
                lLeft = oFrm.Left + oFrm.Width
            End If
            oFrm = Nothing
            oFrm = oUILib.GetWindow("frmAdvanceDisplay")
            If oFrm Is Nothing = False Then
                lTop = oFrm.Top - 512
            End If
            oFrm = Nothing

            If muSettings.ContentsLocX <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Left = muSettings.ContentsLocX
            If muSettings.ContentsLocY <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Top = muSettings.ContentsLocY

            .Left = lLeft
            .Top = lTop
            .Width = 384 '256
            .Height = 512 '256
            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
            .mbAcceptReprocessEvents = True
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 30 '10
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Hangar Contents"
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lblCapacity initial props
        lblCapacity = New UILabel(oUILib)
        With lblCapacity
            .ControlName = "lblCapacity"
            .Left = Me.Width - 150
            .Top = 30 '10
            .Width = 145 '87
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "?/?"
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
        End With
        Me.AddChild(CType(lblCapacity, UIControl))

        'tvwData initial props
        tvwData = New UITreeView(oUILib) 'New UIListBox(oUILib)
        With tvwData
            .ControlName = "tvwData"
            .Left = 5
            .Top = 50
            .Width = Me.Width - 10
            .Height = Me.Height - 125 '105
            .Enabled = True
            .Visible = True
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))

            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(tvwData, UIControl))

        'btnLaunchAll initial props
        btnLaunchAll = New UIButton(oUILib)
        With btnLaunchAll
            .ControlName = "btnLaunchAll"
            .Left = 5
            .Top = tvwData.Top + tvwData.Height + 5
            .Width = 95
            .Height = 19
            .Enabled = True
            .Visible = True
            .Caption = "Launch All"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(btnLaunchAll, UIControl))

        'btnLaunch initial props
        btnLaunch = New UIButton(oUILib)
        With btnLaunch
            .ControlName = "btnLaunch"
            .Left = btnLaunchAll.Left + btnLaunchAll.Width + 5
            .Top = btnLaunchAll.Top
            .Width = 95
            .Height = 19
            .Enabled = False
            .Visible = True
            .Caption = "Launch"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(btnLaunch, UIControl))

        'lblQuantity initial props
        lblQuantity = New UILabel(oUILib)
        With lblQuantity
            .ControlName = "lblQuantity"
            .Left = btnLaunch.Left + btnLaunch.Width + 5
            .Top = btnLaunch.Top
            .Width = 25
            .Height = btnLaunch.Height
            .Enabled = True
            .Visible = True
            .Caption = "Qty:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
        End With
        Me.AddChild(CType(lblQuantity, UIControl))

        'txtQuantity initial props
        txtQuantity = New UITextBox(oUILib)
        With txtQuantity
            .ControlName = "txtQuantity"
            .Left = lblQuantity.Left + lblQuantity.Width + 5
            .Top = btnLaunch.Top
            .Width = 30
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "1"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 4
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtQuantity, UIControl))



        'btnCancelLaunch initial props
        btnCancelLaunch = New UIButton(oUILib)
        With btnCancelLaunch
            .ControlName = "btnCancelLaunch"
            .Left = btnLaunchAll.Left
            .Top = btnLaunchAll.Top + btnLaunchAll.Height + 5
            .Width = 149
            .Height = 19
            .Enabled = True
            .Visible = True
            .Caption = "Cancel Launch"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(btnCancelLaunch, UIControl))


        'btnLaunchToReinforce initial props
        btnLaunchToReinforce = New UIButton(oUILib)
        With btnLaunchToReinforce
            .ControlName = "btnLaunchToReinforce"
            .Left = Me.Width - 155
            .Top = btnCancelLaunch.Top
            .Width = 149
            .Height = 19
            .Enabled = True
            .Visible = True
            .Caption = "Launch to Reinforce"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(btnLaunchToReinforce, UIControl))

        'btnSwitch initial props
        btnSwitch = New UIButton(oUILib)
        With btnSwitch
            .ControlName = "btnSwitch"
            .Left = (Me.Width \ 2) - 78 ' txtQuantity.Left + txtQuantity.Width + 5
            .Top = btnCancelLaunch.Top + btnCancelLaunch.Height + 5
            .Width = 149
            .Height = 19
            .Enabled = True
            .Visible = True
            .Caption = "Switch to Cargo"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(btnSwitch, UIControl))

        'msw_Delay = Stopwatch.StartNew
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        If muSettings.yCurrentContentsView = 1 Then
			btnSwitch_Click(btnSwitch.ControlName)
        End If
        mlTimeOpened = glCurrentCycle
    End Sub

    Private Sub UpdateEntity()
        If goCurrentEnvir Is Nothing Then Return
        If msw_Delay Is Nothing = False AndAlso msw_Delay.ElapsedMilliseconds < ml_CONTENTS_REQUEST_DELAY Then Return
        If mlEntityIndex = -1 Then Return
        If goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return
        If goCurrentEnvir.oEntity(mlEntityIndex) Is Nothing Then Return

        Debug.WriteLine("UpdateEntity")

        If msw_Delay Is Nothing Then
            msw_Delay = Stopwatch.StartNew
        Else
            msw_Delay.Reset()
            msw_Delay.Start()
        End If

        mlPrevCnt = -1  'force a redraw...

        Me.Visible = True

        miEntityTypeID = goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID

        'Because we are requesting new, clear our list
        'goCurrentEnvir.oEntity(mlEntityIndex).colCargo = Nothing
        'goCurrentEnvir.oEntity(mlEntityIndex).colHangar = Nothing

        'With goCurrentEnvir.oEntity(mlEntityIndex)
        '    If .colHangar Is Nothing = False Then
        '        For Each oTmp As EntityContents In .colHangar
        '            If glCurrentCycle - oTmp.ContentsLastUpdate > 120 Then
        '                .colHangar.Remove("ID" & oTmp.ObjectID)
        '                Exit For
        '            End If
        '        Next
        '    End If
        '    If .colCargo Is Nothing = False Then
        '        For Each oTmp As EntityContents In .colCargo
        '            If glCurrentCycle - oTmp.ContentsLastUpdate > 120 Then
        '                .colCargo.Remove("ID" & oTmp.ObjectID)
        '                Exit For
        '            End If
        '        Next
        '    End If
        'End With

        'Ok, send request to get contents
        Dim yData(7) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityContents).CopyTo(yData, 0)
        System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID).CopyTo(yData, 2)
        System.BitConverter.GetBytes(goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID).CopyTo(yData, 6)
        MyBase.moUILib.SendMsgToPrimary(yData)
    End Sub

    Private Sub frmContents_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
        Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y
        'Now, check for what the player is hovering over
        If New Rectangle(4, 5, 77, 22).Contains(lX, lY) = True Then
            If mbHasNoHangar = True Then Return
            'Transfer button
            'mbTransferDisplayed = Not mbTransferDisplayed
            'If mbTransferDisplayed = True Then
            '    Dim oTransfer As frmTransfer = CType(MyBase.moUILib.GetWindow("frmTransfer"), frmTransfer)
            '    If oTransfer Is Nothing Then oTransfer = New frmTransfer(MyBase.moUILib)
            '    oTransfer.Visible = True
            '    oTransfer.SetEntityIndex(mlEntityIndex)
            'Else
            '    MyBase.moUILib.RemoveWindow("frmTransfer")
            'End If
            mbTransferDisplayed = Not mbTransferDisplayed
            If mbTransferDisplayed = True Then
                If goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eUnit Then mbTransferDisplayed = False
                If goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eMining Then mbTransferDisplayed = False
            End If

            If mbTransferDisplayed = True Then
                Dim oTransfer As frmTransfer = CType(MyBase.moUILib.GetWindow("frmTransfer"), frmTransfer)
                If oTransfer Is Nothing Then oTransfer = New frmTransfer(MyBase.moUILib)
                oTransfer.Visible = True
                oTransfer.SetFromEntity(mlEntityIndex, -1, -1)
                'oTransfer.SetEntityIndex(mlEntityIndex)
            Else
                MyBase.moUILib.RemoveWindow("frmTransfer")
            End If
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            Me.IsDirty = True
        ElseIf New Rectangle(82, 5, 77, 22).Contains(lX, lY) = True Then
            'Repair button
            If myCurrView = 0 AndAlso tvwData.oSelectedNode Is Nothing = False AndAlso tvwData.oSelectedNode.lItemData2 = ObjectType.eUnit Then
                HandleRepairButtonClick()
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                Me.IsDirty = True
            End If
        ElseIf New Rectangle(160, 5, 90, 22).Contains(lX, lY) = True Then
            'Repair All
            If mbHasNoHangar = True Then Return
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            HandleRepairAll()
        ElseIf New Rectangle(251, 5, 70, 22).Contains(lX, lY) = True Then
            'Recycle
            If mbHasNoHangar = True Then Return
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            HandleRecycleButtonClick()
        End If
    End Sub

    Private Sub frmContents_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        MyBase.moUILib.SetToolTip(False)

        Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
        Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y
        'Now, check for what the player is hovering over
        If New Rectangle(4, 5, 77, 22).Contains(lX, lY) = True Then
            'Transfer button
            MyBase.moUILib.SetToolTip("Click to show/hide the Transfer window.", lMouseX, lMouseY)
        ElseIf New Rectangle(82, 5, 77, 22).Contains(lX, lY) = True Then
            'Repair button
            MyBase.moUILib.SetToolTip("Click to show/hide the Repair window." & vbCrLf & "You must select a hangar item first.", lMouseX, lMouseY)
        ElseIf New Rectangle(160, 5, 90, 22).Contains(lX, lY) = True Then
            'Repair All
            MyBase.moUILib.SetToolTip("Click to repair all selected units." & vbCrLf & "You must select a hangar item first.", lMouseX, lMouseY)
        ElseIf New Rectangle(251, 5, 70, 22).Contains(lX, lY) = True Then
            'Recycle
            MyBase.moUILib.SetToolTip("Click to dismantle and recycle the selected unit." & vbCrLf & "You must select a hangar item first.", lMouseX, lMouseY)
        End If
    End Sub

    Private Function SortCol(ByVal col As Collection) As Collection

        Dim oTmp As EntityContents

        Dim oResults(-1) As EntityContents
        Dim lResultUB As Int32 = -1
        Dim sResultName(-1) As String
        mbHasUnknowns = False
        For Each oTmp In col
            Dim sName As String = ""
            Select Case oTmp.ObjTypeID
                Case ObjectType.eUnit
                    sName = GetCacheObjectValueCheckUnknowns(oTmp.ObjectID, oTmp.ObjTypeID, mbHasUnknowns)
                    If oTmp.oDetailItem Is Nothing = False Then
                        'sName = CType(oTmp.oDetailItem, EntityDef).DefName

                        If oTmp.lQuantity > -1 Then
                            Dim lSeconds As Int32 = oTmp.lQuantity \ 30
                            Dim lMinutes As Int32 = lSeconds \ 60

                            lSeconds -= (lMinutes * 60)

                            mbNeedSecondUpdates = True
                            mlLastSecondUpdate = glCurrentCycle

                            sName &= " (" & lMinutes.ToString("0#") & ":" & lSeconds.ToString("0#") & ")"

                            'subtract a second
                            oTmp.lQuantity -= 30
                        End If
                    End If
                Case ObjectType.eMineralCache
                    If oTmp.oDetailItem Is Nothing = False Then
                        sName = CType(oTmp.oDetailItem, Mineral).MineralName
                    Else : sName = "Unknown Mineral"
                    End If
                Case ObjectType.eComponentCache
                    If oTmp.oDetailItem Is Nothing = False Then
                        sName = CType(oTmp.oDetailItem, Base_Tech).GetComponentName()
                    Else : sName = GetCacheObjectValueCheckUnknowns(oTmp.ObjectID, oTmp.ObjTypeID, mbHasUnknowns)
                    End If
                Case ObjectType.eColonists
                    sName = oTmp.lQuantity & " Colonists"
                Case ObjectType.eEnlisted
                    sName = oTmp.lQuantity & " Enlisted"
                Case ObjectType.eOfficers
                    sName = oTmp.lQuantity & " Officers"
                Case ObjectType.eAmmunition
                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(oTmp.ObjectID, ObjectType.eWeaponTech)
                    Dim sTemp As String = ""
                    If oTech Is Nothing = False Then
                        sTemp = oTech.GetComponentName
                    Else : sTemp = GetCacheObjectValueCheckUnknowns(oTmp.ObjectID, ObjectType.eWeaponTech, mbHasUnknowns)
                    End If
                    sName = "Ammunition (" & sTemp & ")"
                Case Else
                    'TODO: What else is there?
                    sName = "Define Me!"
            End Select

            sName = sName.ToUpper
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To lResultUB
                If sResultName(X) > sName Then
                    lIdx = X
                    Exit For
                End If
            Next X
            lResultUB += 1
            ReDim Preserve oResults(lResultUB)
            ReDim Preserve sResultName(lResultUB)

            If lIdx = -1 Then
                oResults(lResultUB) = oTmp
                sResultName(lResultUB) = sName
            Else
                For Y As Int32 = lResultUB To lIdx + 1 Step -1
                    oResults(Y) = oResults(Y - 1)
                    sResultName(Y) = sResultName(Y - 1)
                Next Y
                oResults(lIdx) = oTmp
                sResultName(lIdx) = sName
            End If
        Next

        Dim colResult As New Collection
        For X As Int32 = 0 To lResultUB
            colResult.Add(oResults(X))
        Next X
        Return colResult
    End Function

    Private Sub frmContents_OnNewFrame() Handles MyBase.OnNewFrame
        Try
            If mlEntityIndex <> -1 AndAlso goCurrentEnvir Is Nothing = False Then
                Dim oTmp As EntityContents
                Dim yRefresh As Byte = 0
                Dim sName As String
                Dim colTemp As Collection

                Dim lUsed As Int32

                'See if we need to requery the server...
                UpdateEntity()

                If goCurrentEnvir.lEntityUB >= mlEntityIndex AndAlso goCurrentEnvir.lEntityIdx(mlEntityIndex) <> -1 Then
                    With goCurrentEnvir.oEntity(mlEntityIndex)
                        If .ObjTypeID = ObjectType.eFacilityDef AndAlso btnLaunchToReinforce.Enabled = False Then btnLaunchToReinforce.Enabled = True

                        'first, check if we need to update...
                        If myCurrView <> myPrevView Then
                            yRefresh = 1
                        Else
                            If myCurrView = 0 Then
                                If .colHangar Is Nothing = False Then
                                    If mlPrevCnt <> .colHangar.Count Then yRefresh = 1

                                    'OR...
                                    If mbNeedSecondUpdates = True AndAlso (glCurrentCycle - mlLastSecondUpdate) > 30 Then
                                        yRefresh = 1
                                    End If
                                End If
                            ElseIf .colCargo Is Nothing = False Then
                                If mlPrevCnt <> .colCargo.Count Then yRefresh = 1
                            End If
                        End If

                        If .ContentsLastUpdate > mlLastContentUpdate Then
                            mlLastContentUpdate = .ContentsLastUpdate
                            yRefresh = 1
                        End If
                        If mbHasUnknowns AndAlso mlTimeOpened + 30 > glCurrentCycle Then mbForceRefresh = True

                        'Now, if we need to refresh...
                        If yRefresh > 0 Or mbForceRefresh = True Then
                            Dim oGrandPa As UITreeView.UITreeViewItem = Nothing
                            If myCurrView <> myPrevView Then
                                tvwData.Clear()
                            End If

                            mbNeedSecondUpdates = False

                            Debug.WriteLine("frmContents_OnNewFrame")

                            If myCurrView = 0 Then
                                colTemp = .colHangar
                            Else : colTemp = .colCargo
                            End If

                            lUsed = 0

                            If colTemp Is Nothing = False Then

                                If colTemp.Count <> 0 Then
                                    If myCurrView = 0 Then
                                        .colHangar = SortCol(.colHangar)
                                        colTemp = .colHangar
                                    Else
                                        .colCargo = SortCol(.colCargo)
                                        colTemp = .colCargo
                                    End If

                                    If myCurrView = 0 Then
                                        oGrandPa = tvwData.oRootNode
                                        If oGrandPa Is Nothing OrElse oGrandPa.lItemData <> -1 OrElse oGrandPa.lItemData2 <> -1 Then
                                            tvwData.Clear()
                                            oGrandPa = tvwData.AddNode("All Units", -1, -1, -1, Nothing, Nothing)
                                            oGrandPa.bExpanded = True
                                        End If
                                    End If
                                End If

                                Dim iUC(Int16.MaxValue) As Int16
                                mlPrevCnt = colTemp.Count
                                For Each oTmp In colTemp
                                    sName = ""
                                    Select Case oTmp.ObjTypeID
                                        Case ObjectType.eUnit
                                            sName = GetCacheObjectValue(oTmp.ObjectID, oTmp.ObjTypeID)
                                            If oTmp.oDetailItem Is Nothing = False Then
                                                'sName = CType(oTmp.oDetailItem, EntityDef).DefName
                                                lUsed += CType(oTmp.oDetailItem, EntityDef).HullSize

                                                If oTmp.lQuantity > -1 Then
                                                    Dim lSeconds As Int32 = oTmp.lQuantity \ 30
                                                    Dim lMinutes As Int32 = lSeconds \ 60

                                                    lSeconds -= (lMinutes * 60)

                                                    mbNeedSecondUpdates = True
                                                    mlLastSecondUpdate = glCurrentCycle

                                                    sName &= " (" & lMinutes.ToString("0#") & ":" & lSeconds.ToString("0#") & ")"

                                                    'subtract a second
                                                    oTmp.lQuantity -= 30
                                                End If
                                            End If
                                        Case ObjectType.eMineralCache
                                            If oTmp.oDetailItem Is Nothing = False Then
                                                sName = CType(oTmp.oDetailItem, Mineral).MineralName
                                            Else : sName = "Unknown Mineral"
                                            End If
                                            lUsed += oTmp.lQuantity
                                        Case ObjectType.eComponentCache
                                            If oTmp.oDetailItem Is Nothing = False Then
                                                sName = CType(oTmp.oDetailItem, Base_Tech).GetComponentName()
                                            Else : sName = GetCacheObjectValue(oTmp.ObjectID, oTmp.ObjTypeID)
                                            End If
                                            lUsed += oTmp.lQuantity
                                        Case ObjectType.eColonists
                                            'sName = oTmp.lQuantity & " Colonists"
                                            sName = "Colonists"
                                            lUsed += oTmp.lQuantity
                                        Case ObjectType.eEnlisted
                                            'sName = oTmp.lQuantity & " Enlisted"
                                            sName = "Enlisted"
                                            lUsed += oTmp.lQuantity
                                        Case ObjectType.eOfficers
                                            'sName = oTmp.lQuantity & " Officers"
                                            sName = "Officers"
                                            lUsed += oTmp.lQuantity
                                        Case ObjectType.eAmmunition
                                            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(oTmp.ObjectID, ObjectType.eWeaponTech)
                                            Dim sTemp As String = ""
                                            If oTech Is Nothing = False Then
                                                sTemp = oTech.GetComponentName
                                            Else : sTemp = GetCacheObjectValue(oTmp.ObjectID, ObjectType.eWeaponTech)
                                            End If
                                            sName = "Ammunition (" & sTemp & ")"
                                            lUsed += oTmp.lQuantity
                                        Case Else
                                            'TODO: What else is there?
                                            sName = "Define Me!"
                                    End Select

                                    If sName <> "" Then
                                        If oTmp.ObjTypeID = ObjectType.eComponentCache OrElse oTmp.ObjTypeID = ObjectType.eMineralCache OrElse oTmp.ObjTypeID = ObjectType.eAmmunition OrElse oTmp.ObjTypeID = ObjectType.eColonists OrElse oTmp.ObjTypeID = ObjectType.eEnlisted OrElse oTmp.ObjTypeID = ObjectType.eOfficers Then sName &= " x" & oTmp.lQuantity

                                        Dim lIdx As Int32 = -1
                                        Dim oNode As UITreeView.UITreeViewItem = tvwData.GetNodeByItemData2(oTmp.ObjectID, oTmp.ObjTypeID)
                                        Dim oParent As UITreeView.UITreeViewItem = Nothing

                                        If oTmp.ObjTypeID = ObjectType.eComponentCache OrElse oTmp.ObjTypeID = ObjectType.eMineralCache OrElse oTmp.ObjTypeID = ObjectType.eAmmunition OrElse oTmp.ObjTypeID = ObjectType.eColonists OrElse oTmp.ObjTypeID = ObjectType.eEnlisted OrElse oTmp.ObjTypeID = ObjectType.eOfficers Then
                                            If oNode Is Nothing Then
                                                oNode = tvwData.AddNode(sName, oTmp.ObjectID, oTmp.ObjTypeID, -1, oParent, Nothing)
                                                Me.IsDirty = True
                                            Else
                                                If oNode.sItem <> sName Then
                                                    oNode.sItem = sName
                                                    Me.IsDirty = True
                                                End If
                                            End If
                                        Else
                                            oParent = tvwData.GetNodeByItemData2(oTmp.lDetailItemID, oTmp.iDetailItemTypeID)
                                            Dim sParent As String = GetCacheObjectValue(oTmp.lDetailItemID, oTmp.iDetailItemTypeID)
                                            If oParent Is Nothing Then
                                                oParent = tvwData.AddNode(sParent, oTmp.lDetailItemID, oTmp.iDetailItemTypeID, -1, oGrandPa, Nothing)
                                                Me.IsDirty = True
                                            Else
                                                If oParent.sItem.EndsWith(sParent) = False Then 'Reversed with StarsWith since the recent QTY will throw off the startswith.
                                                    oParent.sItem = sParent
                                                    Me.IsDirty = True
                                                End If
                                            End If
                                            iUC(0) = iUC(0) + 1S
                                            iUC(oTmp.lDetailItemID) = iUC(oTmp.lDetailItemID) + 1S
                                            If myCurrView = 0 Then
                                                oParent.sItem = "(" & iUC(oTmp.lDetailItemID).ToString("#,##0") & ") " & sParent
                                                oGrandPa.sItem = " (" & iUC(0).ToString("#,##0") & ") All Units"
                                            End If
                                            If oNode Is Nothing Then
                                                oNode = tvwData.AddNode(sName, oTmp.ObjectID, oTmp.ObjTypeID, -1, oParent, Nothing)
                                                Me.IsDirty = True
                                            Else
                                                If oNode.sItem <> sName Then
                                                    oNode.sItem = sName
                                                    Me.IsDirty = True
                                                End If
                                            End If
                                        End If

                                    End If
                                Next

                                'Now, go back through and remove any items that are no longer there
                                Dim lCurrID1 As Int32 = -1
                                Dim lCurrID2 As Int32 = -1
                                If tvwData.oSelectedNode Is Nothing = False Then  'lstData.ListIndex <> -1 Then
                                    'lCurrID1 = lstData.ItemData(lstData.ListIndex) : lCurrID2 = lstData.ItemData2(lstData.ListIndex)
                                    lCurrID1 = tvwData.oSelectedNode.lItemData : lCurrID2 = tvwData.oSelectedNode.lItemData2
                                End If

                                Dim bDone As Boolean = False
                                While bDone = False
                                    bDone = True


                                    Dim oNode As UITreeView.UITreeViewItem = tvwData.oRootNode
                                    While oNode Is Nothing = False
                                        If oNode.lItemData2 = ObjectType.eUnit Then
                                            Dim bFound As Boolean = False
                                            For Each oTmp In colTemp
                                                If oTmp.ObjectID = oNode.lItemData AndAlso oTmp.ObjTypeID = oNode.lItemData2 Then
                                                    bFound = True
                                                    Exit For
                                                End If
                                            Next oTmp
                                            If bFound = False Then
                                                tvwData.RemoveNode(oNode)
                                                bDone = False
                                                Exit While
                                            End If
                                        ElseIf oNode.lItemData2 = ObjectType.eUnitDef Then
                                            Dim bFound As Boolean = False
                                            For Each oTmp In colTemp
                                                If oTmp.lDetailItemID = oNode.lItemData AndAlso oTmp.iDetailItemTypeID = oNode.lItemData2 Then
                                                    bFound = True
                                                    Exit For
                                                End If
                                            Next oTmp
                                            If bFound = False Then
                                                Dim oGrandPaCheck As UITreeView.UITreeViewItem = oNode.oParentNode
                                                tvwData.RemoveNode(oNode)
                                                If Not oGrandPaCheck Is Nothing Then
                                                    If oGrandPaCheck.oFirstChild Is Nothing Then
                                                        tvwData.RemoveNode(oGrandPaCheck)
                                                    End If
                                                End If
                                            End If
                                        End If
                                        oNode = UITreeView.TraverseNextNode(oNode)
                                    End While

                                    'For X As Int32 = 0 To lstData.ListCount - 1
                                    '    Dim bFound As Boolean = False
                                    '    For Each oTmp In colTemp
                                    '        If oTmp.ObjectID = lstData.ItemData(X) AndAlso oTmp.ObjTypeID = lstData.ItemData2(X) Then
                                    '            bFound = True
                                    '            Exit For
                                    '        End If
                                    '    Next oTmp
                                    '    If bFound = False Then
                                    '        lstData.RemoveItem(X)
                                    '        bDone = False
                                    '        Exit For
                                    '    End If
                                    'Next X
                                End While

                                'bDone = False
                                'For X As Int32 = 0 To lstData.ListCount - 1
                                '    If lstData.ItemData(X) = lCurrID1 AndAlso lstData.ItemData2(X) = lCurrID2 Then
                                '        If lstData.ListIndex <> X Then lstData.ListIndex = X
                                '        bDone = True
                                '        Exit For
                                '    End If
                                'Next X
                                'If bDone = False Then
                                If tvwData.oSelectedNode Is Nothing = True Then MyBase.moUILib.RemoveWindow("frmRepair")
                                'End If

                            Else : mlPrevCnt = 0
                            End If

                            myPrevView = myCurrView

                            If myCurrView = 0 Then
                                lblTitle.Caption = "Hangar Contents"
                                If .oUnitDef Is Nothing Then
                                    lblCapacity.Caption = lUsed.ToString("#,##0") & "/???"
                                Else
                                    lblCapacity.Caption = lUsed.ToString("#,##0") & "/" & .oUnitDef.Hangar_Cap.ToString("#,##0")
                                End If
                            Else
                                lblTitle.Caption = "Cargo Contents"
                                If .oUnitDef Is Nothing Then
                                    lblCapacity.Caption = lUsed.ToString("#,##0") & "/???"
                                Else
                                    lblCapacity.Caption = lUsed.ToString("#,##0") & "/" & .oUnitDef.Cargo_Cap.ToString("#,##0")
                                End If
                            End If

                        End If
                        mbForceRefresh = False

                    End With
                Else
                    mlEntityIndex = -1
                    Me.Visible = False
                End If
            End If
        Catch
            'do nothing, next frame will handle it
        End Try

        Try
            If goCurrentEnvir Is Nothing OrElse mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 OrElse goCurrentEnvir.oEntity(mlEntityIndex).bSelected = False Then
                MyBase.moUILib.RemoveSelection(mlEntityIndex)
                Me.Visible = False
                MyBase.moUILib.RemoveWindow(Me.ControlName)
                MyBase.moUILib.RemoveWindow("frmResearchMain")
                MyBase.moUILib.RemoveWindow("frmBuildWindow")
                Return
            End If
        Catch
            'Do Nothing
        End Try
        'Verify all of the data...
        If mbTransferDisplayed = True Then
            Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmTransfer")
            If ofrm Is Nothing OrElse ofrm.Visible = False Then
                mbTransferDisplayed = False
                Me.IsDirty = True
            End If
        End If
        If mbRepairDisplayed = True Then
            Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmRepair")
            If ofrm Is Nothing OrElse ofrm.Visible = False Then
                mbRepairDisplayed = False
                Me.IsDirty = True
            End If
        End If
        'If mbAmmoDisplayed = True Then
        '    Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmAmmo")
        '    If ofrm Is Nothing OrElse ofrm.Visible = False Then
        '        mbAmmoDisplayed = False
        '        Me.IsDirty = True
        '    End If
        'End If

    End Sub

    Private Sub frmContents_OnRenderEnd() Handles Me.OnRenderEnd
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

        If mbTransferDisplayed = True OrElse mbRepairDisplayed = True Then 'OrElse mbAmmoDisplayed = True Then
            Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                Dim oSelColor As System.Drawing.Color
                oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)

                oSprite.Begin(SpriteFlags.AlphaBlend)
                Dim rcTemp As Rectangle
                If mbTransferDisplayed = True Then
                    rcTemp = New Rectangle(4, 5, 77, 22)
                    DoMultiColorFill(rcTemp, oSelColor, rcTemp.Location, oSprite)
                End If
                If mbRepairDisplayed = True Then
                    rcTemp = New Rectangle(82, 5, 77, 22)
                    DoMultiColorFill(rcTemp, oSelColor, rcTemp.Location, oSprite)
                End If
                'If mbAmmoDisplayed = True Then
                '    rcTemp = New Rectangle(160, 5, 77, 22)
                '    DoMultiColorFill(rcTemp, oSelColor, rcTemp.Location, oSprite)
                'End If
                oSprite.End()
            End Using
        End If

        'Now, draw our borders around the buttons
        Dim rcRect As Rectangle = New Rectangle(4, 5, 77, 22)
        MyBase.RenderRoundedBorder(rcRect, 1, muSettings.InterfaceBorderColor)
        rcRect = New Rectangle(82, 5, 77, 22)
        MyBase.RenderRoundedBorder(rcRect, 1, muSettings.InterfaceBorderColor)
        rcRect = New Rectangle(160, 5, 90, 22)
        MyBase.RenderRoundedBorder(rcRect, 1, muSettings.InterfaceBorderColor)
        rcRect = New Rectangle(251, 5, 70, 22)
        MyBase.RenderRoundedBorder(rcRect, 1, muSettings.InterfaceBorderColor)

        'Now, render our text...
        Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                oTextSpr.Begin(SpriteFlags.AlphaBlend)
                Dim clrTemp As System.Drawing.Color

                rcRect = New Rectangle(4, 5, 77, 22)
                clrTemp = muSettings.InterfaceBorderColor
                If mbHasNoHangar = True Then
                    clrTemp = System.Drawing.Color.FromArgb(255, clrTemp.R \ 2, clrTemp.G \ 2, clrTemp.B \ 2)
                End If
                oFont.DrawText(oTextSpr, "Transfer", rcRect, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrTemp)

                Dim oNode As UITreeView.UITreeViewItem = tvwData.oSelectedNode
                rcRect = New Rectangle(82, 5, 77, 22)
                clrTemp = muSettings.InterfaceBorderColor
                If oNode Is Nothing OrElse oNode.lItemData2 <> ObjectType.eUnit Then
                    clrTemp = System.Drawing.Color.FromArgb(255, clrTemp.R \ 2, clrTemp.G \ 2, clrTemp.B \ 2)
                End If
                oFont.DrawText(oTextSpr, "Repair", rcRect, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrtemp)

                
                rcRect = New Rectangle(160, 5, 90, 22)
                clrTemp = muSettings.InterfaceBorderColor
                If mbHasNoHangar = True Then
                    clrTemp = System.Drawing.Color.FromArgb(255, clrTemp.R \ 2, clrTemp.G \ 2, clrTemp.B \ 2)
                End If
                oFont.DrawText(oTextSpr, "Repair All", rcRect, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrTemp)


                rcRect = New Rectangle(251, 5, 70, 22)
                clrTemp = muSettings.InterfaceBorderColor
                If oNode Is Nothing OrElse oNode.lItemData2 <> ObjectType.eUnit Then
                    'OrElse HasAliasedRights(AliasingRights.eDismantle) = False Then
                    clrTemp = System.Drawing.Color.FromArgb(255, clrTemp.R \ 2, clrTemp.G \ 2, clrTemp.B \ 2)
                End If
                oFont.DrawText(oTextSpr, "Recycle", rcRect, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrTemp)

                oTextSpr.End()
                oTextSpr.Dispose()
            End Using
            oFont.Dispose()
        End Using
    End Sub

	Private Sub btnSwitch_Click(ByVal sName As String) Handles btnSwitch.Click
		If myCurrView = 0 Then
			myCurrView = 1
			btnSwitch.Caption = "Switch to Hangar"
			btnLaunch.Visible = False
			btnLaunchAll.Visible = False
            btnLaunchToReinforce.Visible = False
            btnCancelLaunch.Visible = False
            lblQuantity.Visible = False
            txtQuantity.Visible = False
		Else
			myCurrView = 0
			btnSwitch.Caption = "Switch to Cargo"
			btnLaunch.Visible = True
			btnLaunchAll.Visible = True
            btnLaunchToReinforce.Visible = True
            btnCancelLaunch.Visible = True
            lblQuantity.Visible = True
            txtQuantity.Visible = True
		End If
		muSettings.yCurrentContentsView = myCurrView
	End Sub

    Private Sub tvwData_ItemClick(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwData.NodeSelected
        If myCurrView = 0 Then
            If oNode Is Nothing = False Then
                btnLaunch.Enabled = True

                Try
                    Dim lID As Int32 = oNode.lItemData
                    With goCurrentEnvir.oEntity(mlEntityIndex)
                        For Each oTmp As EntityContents In .colHangar
                            If oTmp.ObjectID = lID Then

                                'If oTmp.lQuantity > 0 AndAlso .yProductionType <> ProductionType.eMining Then
                                '    btnLaunch.Caption = "Cancel"
                                'Else : 
                                btnLaunch.Caption = "Launch"
                                'End If

                                Dim ofrmRepair As frmRepair = CType(MyBase.moUILib.GetWindow("frmRepair"), frmRepair)
                                If ofrmRepair Is Nothing = False AndAlso ofrmRepair.Visible = True Then
                                    ofrmRepair.SetFromEntity(.ObjectID, .ObjTypeID, oTmp.ObjectID, oTmp.ObjTypeID)
                                End If
                                Dim ofrmAmmo As frmAmmo = CType(MyBase.moUILib.GetWindow("frmAmmo"), frmAmmo)
                                If ofrmAmmo Is Nothing = False AndAlso ofrmAmmo.Visible = True Then
                                    ofrmAmmo.SetFromEntity(.ObjectID, .ObjTypeID, oTmp.ObjectID, oTmp.ObjTypeID, oNode.sItem)
                                End If

                                ofrmRepair = Nothing
                                ofrmAmmo = Nothing

                                Exit For
                            End If
                        Next
                    End With

                Catch ex As Exception
                    MyBase.moUILib.RemoveWindow("frmRepair")
                    MyBase.moUILib.RemoveWindow("frmAmmo")
                End Try

            Else
                btnLaunch.Enabled = False
            End If
        End If
    End Sub

    Private Sub tvwData_ItemDblClick() Handles tvwData.NodeDoubleClicked
        btnLaunch_Click(btnLaunch.ControlName)
    End Sub

    Private Sub btnLaunch_Click(ByVal sName As String) Handles btnLaunch.Click
        If myCurrView = 1 Then Exit Sub

        If HasAliasedRights(AliasingRights.eDockUndockUnits) = False Then
            goUILib.AddNotification("You lack rights to dock/undock units.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        If tvwData.oSelectedNode Is Nothing = False Then
            Dim oTmp As EntityContents
            Dim lID As Int32 = tvwData.oSelectedNode.lItemData
            Dim iTypeID As Int16 = CShort(tvwData.oSelectedNode.lItemData2)

            Dim yData(14) As Byte

            If goCurrentEnvir Is Nothing = False Then
                If goCurrentEnvir.lEntityIdx(mlEntityIndex) <> -1 Then
                    With goCurrentEnvir.oEntity(mlEntityIndex)
                        If iTypeID = ObjectType.eUnit Then
                            For Each oTmp In .colHangar
                                If oTmp.ObjectID = lID Then
                                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestUndock).CopyTo(yData, 0)
                                    'Send our object's GUID
                                    oTmp.GetGUIDAsString.CopyTo(yData, 2)
                                    'Send our parent object's GUID
                                    .GetGUIDAsString.CopyTo(yData, 8)

                                    'If btnLaunch.Caption = "Cancel" Then
                                    '    yData(14) = 1
                                    '    'If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, e
                                    'Else
                                    yData(14) = 0
                                    If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, SoundMgr.SpeechType.eUndockRequest Or SoundMgr.SpeechType.eSingleSelect)
                                    'End If

                                    MyBase.moUILib.SendMsgToRegion(yData)

                                    Exit For
                                End If
                            Next
                        ElseIf iTypeID = ObjectType.eUnitDef Then
                            Dim lCnt As Int32 = CInt(Val(txtQuantity.Caption))
                            If lCnt < 1 Then lCnt = 1
                            If lCnt > 255 Then lCnt = 255

                            'ReDim yData(lCnt * 16)
                            'ReDim yData(lCnt * 17)
                            'Dim lPos As Int32 = 0
                            ''Ok, go through our contents
                            'For Each oTmp In .colHangar
                            '    If oTmp.lDetailItemID = lID AndAlso oTmp.iDetailItemTypeID = iTypeID Then

                            ReDim yData(15)
                            System.BitConverter.GetBytes(GlobalMessageCode.eRequestUndock).CopyTo(yData, 0)
                            'oTmp.GetGUIDAsString.CopyTo(yData, 2)
                            System.BitConverter.GetBytes(lID).CopyTo(yData, 2)
                            System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 6)
                            .GetGUIDAsString.CopyTo(yData, 8)
                            yData(14) = 0
                            yData(15) = CByte(lCnt)
                            MyBase.moUILib.SendMsgToRegion(yData)
                            '        lCnt -= 1
                            '        If lCnt < 1 Then Exit For
                            '    End If
                            'Next
                            'MyBase.moUILib.SendLenAppendedMsgToRegion(yData)

                            If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, SoundMgr.SpeechType.eUndockRequest Or SoundMgr.SpeechType.eSingleSelect)
                        Else
                            System.BitConverter.GetBytes(GlobalMessageCode.eRequestUndock).CopyTo(yData, 0)
                            'Send our object's GUID
                            System.BitConverter.GetBytes(lID).CopyTo(yData, 2)
                            System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 6)
                            'Send our parent object's GUID
                            .GetGUIDAsString.CopyTo(yData, 8)

                            yData(14) = 0
                            If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, SoundMgr.SpeechType.eUndockRequest Or SoundMgr.SpeechType.eSingleSelect)

                            MyBase.moUILib.SendMsgToRegion(yData)
                        End If
                    End With
                End If
            End If
        Else
            btnLaunch.Enabled = False
        End If
    End Sub

	Private Sub btnLaunchAll_Click(ByVal sName As String) Handles btnLaunchAll.Click
		If myCurrView = 1 Then Exit Sub

		Dim yData(13) As Byte

		If HasAliasedRights(AliasingRights.eDockUndockUnits) = False Then
			goUILib.AddNotification("You lack rights to dock/undock units.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If

		If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eLaunchAll, 1))

		If goCurrentEnvir Is Nothing = False Then
			If goCurrentEnvir.lEntityIdx(mlEntityIndex) <> -1 Then
                With goCurrentEnvir.oEntity(mlEntityIndex)

                    Dim oNode As UITreeView.UITreeViewItem = tvwData.oSelectedNode
                    If oNode Is Nothing = False Then
                        Dim lID As Int32 = oNode.lItemData
                        Dim iTypeID As Int16 = CShort(oNode.lItemData2)

                        Dim lCnt As Int32 = 0
                        For Each oTmp As EntityContents In .colHangar
                            If oTmp.lDetailItemID = lID AndAlso oTmp.iDetailItemTypeID = iTypeID Then
                                lCnt += 1
                            End If
                        Next
                        If lCnt < 1 Then Return
                        If lCnt > 255 Then lCnt = 255

                        ReDim yData(15)
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestUndock).CopyTo(yData, 0)
                        System.BitConverter.GetBytes(lID).CopyTo(yData, 2)
                        System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 6)
                        .GetGUIDAsString.CopyTo(yData, 8)
                        yData(14) = 0
                        yData(15) = CByte(lCnt)
                        MyBase.moUILib.SendMsgToRegion(yData)

                    Else
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestUndock).CopyTo(yData, 0)
                        System.BitConverter.GetBytes(CInt(-1)).CopyTo(yData, 2)
                        System.BitConverter.GetBytes(CShort(-1)).CopyTo(yData, 6)
                        .GetGUIDAsString.CopyTo(yData, 8)

                        MyBase.moUILib.SendMsgToRegion(yData)
                    End If

                End With
			End If
		End If
	End Sub

	Public Sub SetEntityRef(ByVal lIndex As Int32)
		If mlEntityIndex = lIndex Then Return
		mlEntityIndex = lIndex

		If goCurrentEnvir Is Nothing Then Return
		If mlEntityIndex = -1 Then Return
		If goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return
		If goCurrentEnvir.oEntity(mlEntityIndex) Is Nothing Then Return

		If goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eSpaceStationSpecial Then
			If glCurrentCycle - mlLastStationRqst > 30 Then
				Dim yMsg(11) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eGetColonyChildList).CopyTo(yMsg, 0)
				goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
				System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 8)
				MyBase.moUILib.SendMsgToPrimary(yMsg)
				mlLastStationRqst = glCurrentCycle
			End If
		End If
        Dim bHasCargo As Boolean = (goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0
        Dim bHasHangar As Boolean = (goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus And elUnitStatus.eHangarOperational) <> 0
        If bHasCargo OrElse bHasHangar Then
            'If goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eFacility = True OrElse goCurrentEnvir.oEntity(mlEntityIndex).oUnitDef.Cargo_Cap = 0 Then
            If goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eFacility = True OrElse bHasCargo = False Then
                btnSwitch.Enabled = False
                mbHasNoCargo = True
                If myCurrView = 1 Then
                    btnSwitch_Click(btnSwitch.ControlName)
                End If
            ElseIf bHasHangar = False Then
                btnSwitch.Enabled = False
                mbHasNoHangar = True
                If myCurrView = 0 Then
                    btnSwitch_Click(btnSwitch.ControlName)
                End If
            End If
        End If
        mbForceRefresh = True

        If lIndex <> -1 Then UpdateEntity()
    End Sub

	Protected Overrides Sub Finalize()
		MyBase.moUILib.RemoveWindow("frmTransfer")
		MyBase.moUILib.RemoveWindow("frmRepair")
		Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
		msw_Delay.Stop()
		msw_Delay = Nothing
		MyBase.Finalize()
	End Sub

	'Private Sub lblShowTransfer_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles lblShowTransfer.OnMouseDown
	'    Dim ofrmTransfer As frmTransfer = CType(MyBase.moUILib.GetWindow("frmTransfer"), frmTransfer)

	'    If ofrmTransfer Is Nothing = False Then
	'        If ofrmTransfer.Visible = True Then
	'            'ok, remove it
	'            MyBase.moUILib.RemoveWindow(ofrmTransfer.ControlName)
	'        Else
	'            'show it
	'            ofrmTransfer.Visible = True
	'            ofrmTransfer.SetEntityIndex(mlEntityIndex)
	'        End If
	'    Else
	'        'create a new one
	'        ofrmTransfer = New frmTransfer(goUILib)
	'        ofrmTransfer.Visible = True
	'        ofrmTransfer.SetEntityIndex(mlEntityIndex)
	'    End If

	'    ofrmTransfer = Nothing

	'    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\MouseOver.wav", False, EpicaSound.SoundUsage.eUserInterface, Nothing, Nothing)
	'End Sub

	Private Sub DoMultiColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point, ByRef oSpr As Sprite)
		Dim rcSrc As Rectangle

		Dim fX As Single
		Dim fY As Single

		If rcDest.Width = 0 OrElse rcDest.Height = 0 Then Exit Sub

		rcSrc.Location = New Point(192, 0)
		rcSrc.Width = 62
		rcSrc.Height = 64

		'Now, draw it...
		With oSpr
			fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
			fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))
			.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFill)
		End With
	End Sub

	Private Sub HandleRepairButtonClick()
		mbRepairDisplayed = Not mbRepairDisplayed
		If mbRepairDisplayed = True Then
			Dim oRepair As frmRepair = CType(MyBase.moUILib.GetWindow("frmRepair"), frmRepair)
			If oRepair Is Nothing Then oRepair = New frmRepair(MyBase.moUILib)
			oRepair.Visible = True

            Try
                If tvwData.oSelectedNode Is Nothing = False Then
                    Dim lID As Int32 = tvwData.oSelectedNode.lItemData
                    Dim lTypeID As Int32 = tvwData.oSelectedNode.lItemData2

                    If lTypeID <> ObjectType.eUnit Then
                        MyBase.moUILib.RemoveWindow("frmRepair")
                        Return
                    End If

                    With goCurrentEnvir.oEntity(mlEntityIndex)
                        For Each oTmp As EntityContents In .colHangar
                            If oTmp.ObjectID = lID Then
                                oRepair.SetFromEntity(.ObjectID, .ObjTypeID, oTmp.ObjectID, oTmp.ObjTypeID)
                                Exit For
                            End If
                        Next
                    End With
                End If

            Catch ex As Exception
                MyBase.moUILib.RemoveWindow("frmRepair")
            End Try
		Else
			MyBase.moUILib.RemoveWindow("frmRepair")
		End If
	End Sub

    'Private Sub HandleAmmoButtonClick()
    '	mbAmmoDisplayed = Not mbAmmoDisplayed

    '	If mbAmmoDisplayed = True Then
    '		Dim oAmmo As frmAmmo = CType(MyBase.moUILib.GetWindow("frmAmmo"), frmAmmo)
    '		If oAmmo Is Nothing Then oAmmo = New frmAmmo(MyBase.moUILib)
    '		oAmmo.Visible = True
    '		Try
    '			If lstData.ListIndex = -1 Then
    '				With goCurrentEnvir.oEntity(mlEntityIndex)
    '					oAmmo.SetFromEntity(.ObjectID, .ObjTypeID, .ObjectID, .ObjTypeID, .EntityName)
    '				End With
    '			Else
    '				Dim lID As Int32 = lstData.ItemData(lstData.ListIndex)
    '				With goCurrentEnvir.oEntity(mlEntityIndex)
    '					For Each oTmp As EntityContents In .colHangar
    '						If oTmp.ObjectID = lID Then
    '							oAmmo.SetFromEntity(.ObjectID, .ObjTypeID, oTmp.ObjectID, oTmp.ObjTypeID, lstData.List(lstData.ListIndex))
    '							Exit For
    '						End If
    '					Next
    '				End With
    '			End If
    '		Catch
    '			MyBase.moUILib.RemoveWindow("frmAmmo")
    '		End Try
    '	Else
    '		MyBase.moUILib.RemoveWindow("frmAmmo")
    '	End If
    'End Sub
    Private lRecycleObjTypeID As Int16
    Private lRecycleObjectID As Int32
    Private Sub HandleRecycleButtonClick()
        lRecycleObjTypeID = -1
        lRecycleObjectID = -1

        If tvwData.oSelectedNode Is Nothing = False Then


            Dim lID As Int32 = tvwData.oSelectedNode.lItemData
            Dim lTypeID As Int32 = tvwData.oSelectedNode.lItemData2

            If lTypeID <> ObjectType.eUnit Then
                Return
            End If

            If HasAliasedRights(AliasingRights.eDismantle) = False Then
                goUILib.AddNotification("You lack rights to dismantle units.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                Return
            End If

            With goCurrentEnvir.oEntity(mlEntityIndex)
                For Each oTmp As EntityContents In .colHangar
                    If oTmp.ObjectID = lID Then
                        lRecycleObjTypeID = oTmp.ObjTypeID
                        lRecycleObjectID = lID
                        Dim oFrm As New frmMsgBox(goUILib, "This will dismantle and recycle the currently selected item. Are you sure you wish to do so?", MsgBoxStyle.YesNo, "Confirm Recycle")
                        oFrm.Visible = True
                        AddHandler oFrm.DialogClosed, AddressOf RecycleButtonResult
                        Exit For
                    End If
                Next
            End With
        End If
    End Sub

    Private Sub RecycleButtonResult(ByVal lResult As MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            If lRecycleObjTypeID = -1 Then Return
            If lRecycleObjectID = -1 Then Return
            If HasAliasedRights(AliasingRights.eDismantle) = False Then
                goUILib.AddNotification("You lack rights to dismantle units.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                Return
            End If

            Dim yMsg(13) As Byte
            Dim lPos As Int32 = 0

            If goCurrentEnvir Is Nothing Then Return
            If mlEntityIndex = -1 Then Return
            If goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return
            If goCurrentEnvir.oEntity(mlEntityIndex) Is Nothing Then Return

            System.BitConverter.GetBytes(GlobalMessageCode.eSetDismantleTarget).CopyTo(yMsg, 0) : lPos += 2
            goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            System.BitConverter.GetBytes(lRecycleObjectID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lRecycleObjTypeID).CopyTo(yMsg, lPos) : lPos += 2

            'Station -> 1204195
            ' Unit -> 1025918
            MyBase.moUILib.SendMsgToPrimary(yMsg)

        End If
    End Sub

    Private Sub btnLaunchToReinforce_Click(ByVal sName As String) Handles btnLaunchToReinforce.Click
        Dim ofrm As frmFleet = CType(goUILib.GetWindow("frmFleet"), frmFleet)
        If ofrm Is Nothing = False Then
            If ofrm.Visible = True Then
                goUILib.RemoveWindow(ofrm.ControlName)
            Else : ofrm.Visible = True
            End If
        Else
            ofrm = New frmFleet(goUILib, goCurrentEnvir.oEntity(mlEntityIndex).ObjectID)
            ofrm.Visible = True
        End If
        ofrm = Nothing
    End Sub

    Private Sub frmContents_WindowMoved() Handles Me.WindowMoved
        muSettings.ContentsLocX = Me.Left
        muSettings.ContentsLocY = Me.Top
    End Sub

    Private Sub HandleRepairAll()
        Dim oNode As UITreeView.UITreeViewItem = tvwData.oSelectedNode
        If oNode Is Nothing = False Then
            Dim lID As Int32 = oNode.lItemData
            Dim iTypeID As Int16 = CShort(oNode.lItemData2)

            CreateAndSubmitRepairMessage(lID, iTypeID, -7)
            CreateAndSubmitRepairMessage(lID, iTypeID, Int32.MaxValue)
        End If
    End Sub

    Private Sub CreateAndSubmitRepairMessage(ByVal lEID As Int32, ByVal iETypeID As Int16, ByVal lRepairItem As Int32)
        Dim yMsg(17) As Byte
        Dim lPos As Int32 = 0

        If goCurrentEnvir Is Nothing Then Return
        If mlEntityIndex = -1 Then Return
        If goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return
        If goCurrentEnvir.oEntity(mlEntityIndex) Is Nothing Then Return

        System.BitConverter.GetBytes(GlobalMessageCode.eRepairOrder).CopyTo(yMsg, lPos) : lPos += 2
        goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(lEID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(iETypeID).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lRepairItem).CopyTo(yMsg, lPos) : lPos += 4

        MyBase.moUILib.SendMsgToPrimary(yMsg)

    End Sub

    Private Sub btnCancelLaunch_Click(ByVal sName As String) Handles btnCancelLaunch.Click
        Dim oNode As UITreeView.UITreeViewItem = tvwData.oSelectedNode
        If oNode Is Nothing = False Then
            Dim lID As Int32 = oNode.lItemData
            Dim iTypeID As Int16 = CShort(oNode.lItemData2)
            Dim yData(14) As Byte

            If oNode.lItemData2 = ObjectType.eUnitDef Then 'Cancel (All' Launch
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestUndock).CopyTo(yData, 0)

                'Send our object's GUID
                System.BitConverter.GetBytes(CInt(-1)).CopyTo(yData, 2)
                System.BitConverter.GetBytes(CShort(-1)).CopyTo(yData, 6)
                'Send our parent object's GUID
                goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yData, 8)

                yData(14) = 1

                MyBase.moUILib.SendMsgToRegion(yData)
            Else

                System.BitConverter.GetBytes(GlobalMessageCode.eRequestUndock).CopyTo(yData, 0)

                'Send our object's GUID
                System.BitConverter.GetBytes(lID).CopyTo(yData, 2)
                System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 6)

                'Send our parent object's GUID
                goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yData, 8)

                yData(14) = 1

                MyBase.moUILib.SendMsgToRegion(yData)
            End If
        End If
    End Sub
End Class
