Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmTransfer
    Inherits UIWindow

    Private mlEntityIndex As Int32 = -1
    Private mbUseChild As Boolean = False
    Private mlChildID As Int32 = -1
    Private miChildTypeID As Int16 = -1

    Private WithEvents lstLeft As UIListBox
    Private WithEvents lstRight As UIListBox
    Private WithEvents btnTransferToColony As UIButton
    Private WithEvents btnTransferFromColony As UIButton
    Private lblTitle As UILabel
    Private WithEvents btnClose As UIButton
    Private WithEvents txtQuantity As UITextBox
    Private lblQty As UILabel

    Private mlPreviousUpdate As Int32 = -1

    Private mbPopulated As Boolean = False

    Public Shared bReRequestResources As Boolean = False
    Private mlDelayedUpdate As Int32 = -1

    Private mlLastTotalRefresh As Int32 = 0

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmNewTransfer initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTransfer
            .ControlName = "frmTransfer"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256
            .Width = 512
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
            .bRoundedBorder = True
        End With

        'lstLeft initial props
        lstLeft = New UIListBox(oUILib)
        With lstLeft
            .ControlName = "lstLeft"
            .Left = 5
            .Top = 30
            .Width = 246
            .Height = 440
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .RenderIcons = True
            .oIconTexture = goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
        End With
        Me.AddChild(CType(lstLeft, UIControl))

        'lstRight initial props
        lstRight = New UIListBox(oUILib)
        With lstRight
            .ControlName = "lstRight"
            .Left = 260
            .Top = 30
            .Width = 246
            .Height = 440
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstRight, UIControl))

        'btnTransferFromColony initial props
        btnTransferFromColony = New UIButton(oUILib)
        With btnTransferFromColony
            .ControlName = "btnTransferFromColony"
            .Left = 5
            .Top = 480
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Xfer From Colony"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to transfer the selected item in the" & vbCrLf & "right list to the selected entity in the left list."
        End With
        Me.AddChild(CType(btnTransferFromColony, UIControl))

        'btnTransferToColony initial props
        btnTransferToColony = New UIButton(oUILib)
        With btnTransferToColony
            .ControlName = "btnTransferToColony"
            .Left = Me.Width - 155
            .Top = 480
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Xfer To Colony"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to transfer the selected item in the" & vbCrLf & "left list to the selected entity in the right list."
        End With
        Me.AddChild(CType(btnTransferToColony, UIControl))

        'txtQuantity initial props
        txtQuantity = New UITextBox(oUILib)
        With txtQuantity
            .ControlName = "txtQuantity"
            .Left = (Me.Width \ 2)
            .Top = 480
            .Width = 96
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
            .MaxLength = 7
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtQuantity, UIControl))

        'lblQty initial props
        lblQty = New UILabel(oUILib)
        With lblQty
            .ControlName = "lblQty"
            .Left = txtQuantity.Left - 65
            .Top = 480
            .Width = 61
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Quantity:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
        End With
        Me.AddChild(CType(lblQty, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 160
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Colony Transfer Cargo"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 487
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

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        'goUILib.AddNotification("frmTransfer newed", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Public Sub SetFromEntity(ByVal lEntityIndex As Int32, ByVal lChildID As Int32, ByVal iChildTypeID As Int16)

        Try
            mlEntityIndex = lEntityIndex
            If goCurrentEnvir Is Nothing Then Return
            If mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return
            Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(mlEntityIndex)
            If oEntity Is Nothing Then Return

            mlChildID = lChildID
            miChildTypeID = iChildTypeID

            mbUseChild = False
            If mlChildID > 0 AndAlso miChildTypeID > 0 Then
                'Ok, find the child object...
                Dim oChild As StationChild = oEntity.GetChild(lChildID, iChildTypeID)

                If oChild Is Nothing = False AndAlso oChild.oChildDef Is Nothing = False Then
                    mbUseChild = True
                End If
            End If

            Dim bRequestResources As Boolean = False
            If goCurrentEnvir Is Nothing = False Then
                If goAvailableResources Is Nothing = False Then
                    If goCurrentEnvir.ObjectID <> goAvailableResources.lColonyParentID OrElse goCurrentEnvir.ObjTypeID = goAvailableResources.iColonyParentTypeID Then
                        Dim lID As Int32 = oEntity.ObjectID
                        Dim iTypeID As Int16 = oEntity.ObjTypeID
                        If goAvailableResources.lColonyParentID <> lID OrElse goAvailableResources.iColonyParentTypeID <> iTypeID Then
                            goAvailableResources = Nothing
                            bRequestResources = True
                        End If

                    End If
                Else : bRequestResources = True
                End If
            End If
            If bRequestResources = True Then
                Dim yMsg(7) As Byte
                If oEntity.ObjTypeID = ObjectType.eFacility Then
                    System.BitConverter.GetBytes(GlobalMessageCode.eGetAvailableResources).CopyTo(yMsg, 0)
                    If mbUseChild = False Then
                        oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
                    Else
                        System.BitConverter.GetBytes(mlChildID).CopyTo(yMsg, 2)
                        System.BitConverter.GetBytes(miChildTypeID).CopyTo(yMsg, 6)
                    End If
                    MyBase.moUILib.SendMsgToPrimary(yMsg)
                End If
            End If

            If oEntity Is Nothing = False Then
                Dim colHangar As Collection = oEntity.colHangar
                If colHangar Is Nothing = False Then
                    For Each oTmp As EntityContents In colHangar
                        'Ok, send request to get contents
                        Dim yData(7) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityContents).CopyTo(yData, 0)
                        System.BitConverter.GetBytes(oTmp.ObjectID).CopyTo(yData, 2)
                        System.BitConverter.GetBytes(oTmp.ObjTypeID).CopyTo(yData, 6)
                        MyBase.moUILib.SendMsgToPrimary(yData)
                    Next
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub FillListFromEntity(ByRef lst As UIListBox, ByRef oEntity As BaseEntity, ByRef lstOther As UIListBox)

        Try
            With oEntity
                'bSpaceTradepost = (oEntity.ObjTypeID = ObjectType.eFacility) AndAlso (goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem) AndAlso (oEntity.yProductionType = ProductionType.eTradePost)
                Dim bResetLastTotalRefresh As Boolean = False
                Dim lEnsureItemVisibleID1 As Int32 = -1
                Dim lEnsureItemVisibleID2 As Int32 = -1
                Dim lEnsureItemVisibleID3 As Int32 = -1
                If mlLastTotalRefresh = Int32.MinValue OrElse glCurrentCycle - mlLastTotalRefresh > 150 Then
                    If lst.ListIndex > -1 Then
                        lEnsureItemVisibleID1 = lst.ItemData(lst.ListIndex)
                        lEnsureItemVisibleID2 = lst.ItemData2(lst.ListIndex)
                        lEnsureItemVisibleID3 = lst.ItemData3(lst.ListIndex)
                    End If
                    lst.Clear()
                    bResetLastTotalRefresh = True
                End If
                'set our update flags to false, in AddList... we set the Update flag to True
                lst.SetAllUpdates(False)

                AddListItemIfNeeded(lst, .EntityName, .ObjectID, .ObjTypeID, -1)
                If (.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 AndAlso .colHangar Is Nothing = False Then
                    For Each oTmp As EntityContents In .colHangar
                        Dim sName As String = " " & GetCacheObjectValue(oTmp.ObjectID, oTmp.ObjTypeID)
                        Dim iDetTypeID As Int16 = 0
                        If oTmp.oDetailItem Is Nothing = False Then iDetTypeID = CType(oTmp.oDetailItem, Base_GUID).ObjTypeID
                        AddListChildItemIfNeeded(lst, sName, oTmp.ObjectID, oTmp.ObjTypeID, iDetTypeID, .ObjectID)
                        If oTmp.colCargo Is Nothing = False Then
                            For Each oCargo As EntityContents In oTmp.colCargo
                                sName = "  "
                                If oCargo.oDetailItem Is Nothing = False Then
                                    If oCargo.ObjTypeID = ObjectType.eMineral OrElse oCargo.ObjTypeID = ObjectType.eMineralCache Then
                                        sName &= CType(oCargo.oDetailItem, Mineral).MineralName
                                    Else
                                        Select Case CType(oCargo.oDetailItem, Base_GUID).ObjTypeID
                                            Case ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
                                                sName &= CType(oCargo.oDetailItem, Base_Tech).GetComponentName
                                            Case Else
                                                Dim lObjID As Int32 = CType(oCargo.oDetailItem, Base_GUID).ObjectID
                                                Dim iObjTypeID As Int16 = CType(oCargo.oDetailItem, Base_GUID).ObjTypeID
                                                sName &= GetCacheObjectValue(lObjID, iObjTypeID)
                                        End Select
                                    End If
                                Else
                                    sName &= GetCacheObjectValue(oCargo.ObjectID, oCargo.ObjTypeID)
                                End If

                                sName &= ": " & oCargo.lQuantity.ToString("#,##0")

                                iDetTypeID = 0
                                If oCargo.oDetailItem Is Nothing = False Then iDetTypeID = CType(oCargo.oDetailItem, Base_GUID).ObjTypeID
                                AddListChildItemIfNeeded(lst, sName, oCargo.ObjectID, oCargo.ObjTypeID, iDetTypeID, oTmp.ObjectID)
                            Next
                        End If
                    Next
                    If (.yProductionType = ProductionType.eTradePost OrElse .yProductionType = ProductionType.eSpaceStationSpecial OrElse (.oUnitDef Is Nothing = False AndAlso (.oUnitDef.ModelID And 255) = 148)) AndAlso (.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 AndAlso .colCargo Is Nothing = False Then

                        Dim bSpaceStationCheck As Boolean = False
                        Try
                            If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                bSpaceStationCheck = .ObjTypeID = ObjectType.eFacility
                            End If
                        Catch
                        End Try

                        For Each oTmp As EntityContents In .colCargo

                            Dim sName As String = " "
                            If oTmp.oDetailItem Is Nothing = False Then
                                If oTmp.ObjTypeID = ObjectType.eMineral OrElse oTmp.ObjTypeID = ObjectType.eMineralCache Then
                                    sName &= CType(oTmp.oDetailItem, Mineral).MineralName
                                Else
                                    Select Case CType(oTmp.oDetailItem, Base_GUID).ObjTypeID
                                        Case ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
                                            sName &= CType(oTmp.oDetailItem, Base_Tech).GetComponentName
                                        Case Else
                                            Dim lObjID As Int32 = CType(oTmp.oDetailItem, Base_GUID).ObjectID
                                            Dim iObjTypeID As Int16 = CType(oTmp.oDetailItem, Base_GUID).ObjTypeID
                                            sName &= GetCacheObjectValue(lObjID, iObjTypeID)
                                    End Select
                                End If
                            Else
                                If oTmp.ObjTypeID = ObjectType.eColonists Then
                                    sName &= "Colonists"
                                ElseIf oTmp.ObjTypeID = ObjectType.eEnlisted Then
                                    sName &= "Enlisted"
                                ElseIf oTmp.ObjTypeID = ObjectType.eOfficers Then
                                    sName &= "Officers"
                                Else : sName &= GetCacheObjectValue(oTmp.ObjectID, oTmp.ObjTypeID)
                                End If
                            End If

                            Dim lTmpQty As Int32 = oTmp.lQuantity
                            If bSpaceStationCheck = True AndAlso lstOther Is Nothing = False Then
                                'ok, gotta check the other list
                                If lstOther Is Nothing = False AndAlso oTmp.oDetailItem Is Nothing = False Then
                                    Dim bFound As Boolean = False
                                    Dim lDetailID As Int32 = CType(oTmp.oDetailItem, Base_GUID).ObjectID
                                    Dim iTypeID As Int16 = CType(oTmp.oDetailItem, Base_GUID).ObjTypeID
                                    For X As Int32 = 0 To lstOther.ListCount - 1
                                        If lDetailID = lstOther.ItemData(X) AndAlso iTypeID = lstOther.ItemData2(X) Then
                                            Dim lIdx As Int32 = lstOther.List(X).LastIndexOf(":"c)
                                            If lIdx > -1 Then
                                                Dim lRtQty As Int32 = CInt(Val(lstOther.List(X).Substring(lIdx + 1).Trim))
                                                If oTmp.lQuantity <= lRtQty Then
                                                    bFound = True
                                                Else
                                                    lTmpQty -= lRtQty
                                                End If
                                            End If
                                            Exit For
                                        End If
                                    Next X
                                    If bFound = True Then Continue For
                                End If
                            End If

                            sName &= ": " & lTmpQty.ToString("#,##0")

                            Dim iDetTypeID As Int16 = 0
                            If oTmp.oDetailItem Is Nothing = False Then iDetTypeID = CType(oTmp.oDetailItem, Base_GUID).ObjTypeID

                            'If bSpaceTradepost = False Then
                            AddListChildItemIfNeeded(lstOther, sName, oTmp.ObjectID, oTmp.ObjTypeID, iDetTypeID, -1)
                            'Else
                            '	AddListChildItemIfNeeded(lstOther, sName, oTmp.ObjectID, oTmp.ObjTypeID, iDetTypeID)
                            'End If
                        Next
                    End If
                End If

                'now, remove any unupdated items
                Dim bDone As Boolean = False
                While bDone = False
                    bDone = True
                    For X As Int32 = 0 To lst.ListCount - 1
                        If lst.ItemUpdated(X) = False Then
                            bDone = False
                            lst.RemoveItem(X)
                            Exit For
                        End If
                    Next X
                End While

                If bResetLastTotalRefresh = True Then
                    lst.RestorePosition()
                    mlLastTotalRefresh = glCurrentCycle
                    'If lEnsureItemVisibleID1 > -1 Then
                    '    For X As Int32 = 0 To lst.ListCount - 1
                    '        If lst.ItemData(X) = lEnsureItemVisibleID1 AndAlso lst.ItemData2(X) = lEnsureItemVisibleID2 AndAlso lst.ItemData3(X) = lEnsureItemVisibleID3 Then
                    '            lst.ListIndex = X
                    '            lst.EnsureItemVisible(X)
                    '            Exit For
                    '        End If
                    '    Next X
                    'End If
                End If

            End With
        Catch
            'do nothing
        End Try
    End Sub

    Private Sub frmTransfer_OnNewFrame() Handles Me.OnNewFrame
        If goCurrentEnvir Is Nothing Then Return
        If mlEntityIndex = -1 Then Return
        If goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return

        If bReRequestResources = True And glCurrentCycle - mlDelayedUpdate > 60 Then
            'bReRequestResources = False
            mlDelayedUpdate = glCurrentCycle

            Dim yMsg(7) As Byte
            If goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eFacility Then
                System.BitConverter.GetBytes(GlobalMessageCode.eGetAvailableResources).CopyTo(yMsg, 0)
                If mbUseChild = False Then
                    goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
                Else
                    System.BitConverter.GetBytes(mlChildID).CopyTo(yMsg, 2)
                    System.BitConverter.GetBytes(miChildTypeID).CopyTo(yMsg, 6)
                End If
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
        ElseIf goAvailableResources Is Nothing = False Then
            'ElseIf glCurrentCycle - mlDelayedUpdate > 10 Then
            bReRequestResources = False
        End If

        'Ok, ensure our lists have everything in them
		Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(mlEntityIndex)
		'Dim bSpaceTradepost As Boolean = False


        FillListFromEntity(lstLeft, oEntity, lstRight)
        Dim bNeedRightListCleared As Boolean = False
        If oEntity Is Nothing = False AndAlso oEntity.oUnitDef Is Nothing = False AndAlso (oEntity.oUnitDef.ModelID And 255) = 148 Then
            'fill right list with cargo contents
            bNeedRightListCleared = False
            'FillListFromEntity(lstRight, oEntity, Nothing)
        ElseIf goAvailableResources Is Nothing = False Then
            With goAvailableResources
                If .LastUpdate <> -1 OrElse 1 = 1 Then

                    If .iColonyParentTypeID = ObjectType.eFacility Then
                        'space station, is the station still selected?
                        If goCurrentEnvir.oEntity(mlEntityIndex).ObjectID <> .lColonyParentID OrElse goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID <> .iColonyParentTypeID Then
                            'apparently not... so clear out the resources
                            .ClearResources()
                            bNeedRightListCleared = True
                        ElseIf mlPreviousUpdate <> .LastUpdate Then
                            lstRight.Clear()
                            'fill our list as needed
                            mlPreviousUpdate = .LastUpdate
                            .FillListBox(lstRight) ', bSpaceTradepost)
                        End If
                    ElseIf goCurrentEnvir.ObjectID <> .lColonyParentID OrElse goCurrentEnvir.ObjTypeID <> .iColonyParentTypeID Then
                        .ClearResources()
                        bNeedRightListCleared = True
                    ElseIf mlPreviousUpdate <> .LastUpdate Then
                        'lstRight.Clear()
                        'fill our list as needed
                        mlPreviousUpdate = .LastUpdate
                        .FillListBox(lstRight) ', bSpaceTradepost)
                        lstRight.RestorePosition()
                    End If
                Else : bNeedRightListCleared = True
                End If
            End With
        End If
        If bNeedRightListCleared = True AndAlso lstRight.ListCount <> 0 Then lstRight.Clear()

        mbPopulated = True
    End Sub

    Private Sub AddListItemIfNeeded(ByRef lstData As UIListBox, ByVal sText As String, ByVal lID1 As Int32, ByVal lID2 As Int32, ByVal lParentID As Int32)
        For X As Int32 = 0 To lstData.ListCount - 1
            If lstData.ItemData(X) = lID1 AndAlso lstData.ItemData2(X) = lID2 AndAlso lstData.ItemData3(X) = lParentID Then
                If lstData.List(X) <> sText Then lstData.List(X) = sText
                lstData.ItemUpdated(X) = True
                Return
            End If
        Next X

        lstData.AddItem(sText, True)
        lstData.ItemData(lstData.NewIndex) = lID1
        lstData.ItemData2(lstData.NewIndex) = lID2
        lstData.ItemData3(lstData.NewIndex) = lParentID
        lstData.ApplyIconOffset(lstData.NewIndex) = False
        lstData.ItemUpdated(lstData.NewIndex) = True

        mlLastTotalRefresh = Int32.MinValue
        'If mbPopulated = True Then
        '    mbPopulated = False
        '    lstData.Clear()
        '    Me.IsDirty = True
        'End If
    End Sub

    Private Sub AddListChildItemIfNeeded(ByRef lstData As UIListBox, ByVal sText As String, ByVal lID1 As Int32, ByVal lID2 As Int32, ByVal iDetailTypeID As Int16, ByVal lParentID As Int32)

        For X As Int32 = 0 To lstData.ListCount - 1
            If lstData.ItemData(X) = lID1 AndAlso lstData.ItemData2(X) = lID2 AndAlso lstData.ItemData3(X) = lParentID Then
                If lstData.List(X) <> sText Then lstData.List(X) = sText
                lstData.ItemUpdated(X) = True
                Return
            End If
        Next X
        Dim lIndex As Int32 = -1
        For x As Int32 = 0 To lstData.ListCount - 1
            If lstData.ItemData(x) = lParentID Then
                lIndex = x
            End If
        Next
        'If lParentID > 0 AndAlso lIndex > -0 Then
        'lstData.AddItemAfterIndex(sText, lIndex, False)
        'Else
        lstData.AddItem(sText, False)
        'End If
        lstData.ItemData(lstData.NewIndex) = lID1
        lstData.ItemData2(lstData.NewIndex) = lID2
        lstData.ItemData3(lstData.NewIndex) = lParentID
        lstData.ApplyIconOffset(lstData.NewIndex) = True
        lstData.ItemUpdated(lstData.NewIndex) = True
        Select Case Math.Abs(iDetailTypeID)
            Case ObjectType.eArmorTech
                lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(0, 0, 16, 16)
                lstData.IconForeColor(lstData.NewIndex) = Color.White
            Case ObjectType.eEngineTech
                lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(0, 16, 16, 16)
                lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
            Case ObjectType.eRadarTech
                lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(48, 0, 16, 16)
                lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Case ObjectType.eShieldTech
                lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 16, 16, 16)
                lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            Case ObjectType.eWeaponTech
                lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(32, 32, 16, 16)
                lstData.IconForeColor(lstData.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            Case ObjectType.eOfficers
                lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 0, 16, 16)
                lstData.IconForeColor(lstData.NewIndex) = Color.FromArgb(255, 192, 192, 0)
            Case ObjectType.eEnlisted
                lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 0, 16, 16)
                lstData.IconForeColor(lstData.NewIndex) = Color.FromArgb(255, 64, 192, 64)
            Case ObjectType.eColonists
                lstData.rcIconRectangle(lstData.NewIndex) = New Rectangle(16, 0, 16, 16)
                lstData.IconForeColor(lstData.NewIndex) = Color.White
        End Select

        'If mbPopulated = True Then
        '    mbPopulated = False
        '    lstData.Clear()
        '    Me.IsDirty = True
        'End If
    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub lstRight_ItemClick(ByVal lIndex As Integer) Handles lstRight.ItemClick
		If lIndex > -1 AndAlso goAvailableResources Is Nothing = False Then
			Dim iTypeID As Int16 = CShort(lstRight.ItemData2(lIndex))
			'MyBase.moUILib.AddNotification("Right: " & lstRight.ItemData(lIndex) & ", " & lstRight.ItemData2(lIndex), Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Select Case iTypeID
				Case ObjectType.eArmorTech, ObjectType.eComponentCache, ObjectType.eEngineTech, ObjectType.eMineral, ObjectType.eMineralCache, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
					Dim lQty As Int32 = goAvailableResources.GetObjectQuantity(lstRight.ItemData(lIndex), iTypeID)
					txtQuantity.Caption = lQty.ToString
				Case ObjectType.eColonists, ObjectType.eEnlisted, ObjectType.eOfficers
					Dim lQty As Int32 = lstRight.ItemData(lIndex)
					txtQuantity.Caption = lQty.ToString
			End Select
		End If
	End Sub

	Private Sub lstLeft_ItemClick(ByVal lIndex As Integer) Handles lstLeft.ItemClick
		If lIndex > -1 Then
			Dim lObjID As Int32 = lstLeft.ItemData(lIndex)
			Dim iTypeID As Int16 = CShort(lstLeft.ItemData2(lIndex))
			'MyBase.moUILib.AddNotification("Left: " & lObjID & ", " & iTypeID, Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Select Case iTypeID
				Case ObjectType.eArmorTech, ObjectType.eComponentCache, ObjectType.eEngineTech, ObjectType.eMineral, ObjectType.eMineralCache, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
					Dim bFound As Boolean = False
					Dim lQty As Int32 = 0
					Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(mlEntityIndex)
					If oEntity Is Nothing Then Return
					If oEntity.colCargo Is Nothing = False Then
						For Each oTmp As EntityContents In oEntity.colCargo
							If oTmp.ObjectID = lObjID AndAlso oTmp.ObjTypeID = iTypeID Then
								lQty = oTmp.lQuantity
								bFound = True
								Exit For
							End If
						Next
					End If
					If bFound = False AndAlso oEntity.colHangar Is Nothing = False Then
						For Each oTmp As EntityContents In oEntity.colHangar
							If oTmp.colCargo Is Nothing = False Then
								For Each oCargo As EntityContents In oTmp.colCargo
									If oCargo.ObjectID = lObjID AndAlso oCargo.ObjTypeID = iTypeID Then
										lQty = oCargo.lQuantity
										bFound = True
										Exit For
									End If
								Next
							End If
							If bFound = True Then Exit For
						Next
					End If

					txtQuantity.Caption = lQty.ToString
				Case ObjectType.eColonists, ObjectType.eEnlisted, ObjectType.eOfficers
					Dim lQty As Int32 = lObjID
					txtQuantity.Caption = (lQty + 1).ToString
			End Select
		End If
	End Sub

	Private Sub txtQuantity_TextChanged() Handles txtQuantity.TextChanged
		If txtQuantity.Caption <> "" Then
			Dim lTemp As Int32 = CInt(Val(txtQuantity.Caption))
			If lTemp.ToString <> txtQuantity.Caption Then
				txtQuantity.Caption = lTemp.ToString
			End If
		End If
	End Sub

	Private Sub btnTransferFromColony_Click(ByVal sName As String) Handles btnTransferFromColony.Click
        btnTransferFromColony.Enabled = False


        Dim lFromID As Int32 = -1
        Dim iFromTypeID As Int16 = -1
        Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(mlEntityIndex)
        If oEntity Is Nothing Then
            MyBase.moUILib.AddNotification("Something bad happened with the selected unit/facility.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            btnTransferFromColony.Enabled = True
            Return
        End If

        If goAvailableResources Is Nothing = False Then
            lFromID = goAvailableResources.lColonyID
            iFromTypeID = ObjectType.eColony
        Else
            'lFromID = oEntity.ObjectID
            'iFromTypeID = oEntity.ObjTypeID
            Return
        End If

        If lstLeft Is Nothing = False AndAlso lstLeft.ListIndex > -1 AndAlso lstRight Is Nothing = False AndAlso lstRight.ListIndex > -1 Then
            'Ok, the selection is what they want to send from the colony
            Dim lObjID As Int32 = lstRight.ItemData(lstRight.ListIndex)
            Dim iObjTypeID As Int16 = CShort(lstRight.ItemData2(lstRight.ListIndex))

            Dim lQuantity As Int32 = CInt(Val(txtQuantity.Caption))
            If lQuantity < 1 Then
                MyBase.moUILib.AddNotification("Please enter a value greater than 0 for quantity to transfer.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                btnTransferFromColony.Enabled = True
                Return
            End If

            'Now, ensure that iobjtypeid is a component cache or mineral cache
            If iObjTypeID <> ObjectType.eUnit AndAlso iObjTypeID <> ObjectType.eFacility AndAlso iObjTypeID <> ObjectType.eColony AndAlso iObjTypeID > 0 Then

                Dim lToID As Int32 = lstLeft.ItemData(lstLeft.ListIndex)
                Dim iToTypeID As Int16 = CShort(lstLeft.ItemData2(lstLeft.ListIndex))
                'Debug.Print("F: " & lFromID & " O:" & lObjID & " T:" & lToID & " FT: " & iFromTypeID & " OT:" & iObjTypeID & " TT:" & iToTypeID)
                If iToTypeID <> ObjectType.eUnit AndAlso iToTypeID <> ObjectType.eFacility Then
                    'Ok, find the item selected's parent and then use that parent
                    Dim bFound As Boolean = False
                    If oEntity.colCargo Is Nothing = False Then
                        For Each oTmp As EntityContents In oEntity.colCargo
                            If oTmp.ObjectID = lToID AndAlso oTmp.ObjTypeID = iToTypeID Then
                                lToID = oEntity.ObjectID
                                iToTypeID = oEntity.ObjTypeID

                                If (oEntity.CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then
                                    MyBase.moUILib.AddNotification("Selection's cargo bay is inoperable.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                    btnTransferFromColony.Enabled = True
                                    Return
                                End If
                                bFound = True
                                Exit For
                            End If
                        Next
                    End If
                    If bFound = False AndAlso oEntity.colHangar Is Nothing = False Then
                        For Each oTmp As EntityContents In oEntity.colHangar
                            If oTmp.colCargo Is Nothing = False Then
                                For Each oCargo As EntityContents In oTmp.colCargo
                                    If oCargo.ObjectID = lToID AndAlso oCargo.ObjTypeID = iToTypeID Then
                                        lToID = oTmp.ObjectID
                                        iToTypeID = oTmp.ObjTypeID
                                        bFound = True
                                        Exit For
                                    End If
                                Next
                            End If
                            If bFound = True Then Exit For
                        Next
                    End If

                    If bFound = False Then
                        MyBase.moUILib.AddNotification("Please select a valid destination for the transfer.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        btnTransferFromColony.Enabled = True
                        Return
                    End If
                End If

                PrepareAndSendTransferMsg(lFromID, iFromTypeID, lToID, iToTypeID, lObjID, iObjTypeID, lQuantity)
                MyBase.moUILib.AddNotification("Transfer request sent. Waiting for shipment to finish...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Else
                MyBase.moUILib.AddNotification("You can only transfer components, minerals and personnel from the colony.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        Else
            MyBase.moUILib.AddNotification("You must select a source item in the right list and a destination in the left list to proceed.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If

		btnTransferFromColony.Enabled = True
	End Sub

	Private Sub btnTransferToColony_Click(ByVal sName As String) Handles btnTransferToColony.Click
		btnTransferToColony.Enabled = False
		'find what item is selected
		If lstLeft Is Nothing = False AndAlso lstLeft.ListIndex > -1 AndAlso goAvailableResources Is Nothing = False Then
			'Ok, the selection is what they want to send to the colony
			Dim lObjID As Int32 = lstLeft.ItemData(lstLeft.ListIndex)
			Dim iObjTypeID As Int16 = CShort(lstLeft.ItemData2(lstLeft.ListIndex))

			Dim lQuantity As Int32 = CInt(Val(txtQuantity.Caption))
			If lQuantity < 1 Then
				MyBase.moUILib.AddNotification("Please enter a value greater than 0 for quantity to transfer.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				btnTransferToColony.Enabled = True
				Return
			End If

			'Now, ensure that iobjtypeid is a component cache or mineral cache
			If iObjTypeID = ObjectType.eComponentCache OrElse iObjTypeID = ObjectType.eMineralCache OrElse iObjTypeID = ObjectType.eColonists OrElse iObjTypeID = ObjectType.eEnlisted OrElse iObjTypeID = ObjectType.eOfficers Then
				Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(mlEntityIndex)
				If oEntity Is Nothing Then
					MyBase.moUILib.AddNotification("Something bad happened with the selected unit/facility.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					btnTransferToColony.Enabled = True
					Return
				End If

				Dim bFound As Boolean = False
				Dim lFromID As Int32 = -1
				Dim iFromTypeID As Int16 = -1
				If oEntity.colCargo Is Nothing = False Then
					For Each oTmp As EntityContents In oEntity.colCargo
						If oTmp.ObjectID = lObjID AndAlso oTmp.ObjTypeID = iObjTypeID Then
							lFromID = oEntity.ObjectID
							iFromTypeID = oEntity.ObjTypeID
							bFound = True
							Exit For
						End If
					Next
				End If
				If bFound = False AndAlso oEntity.colHangar Is Nothing = False Then
					For Each oTmp As EntityContents In oEntity.colHangar
						If oTmp.colCargo Is Nothing = False Then
							For Each oCargo As EntityContents In oTmp.colCargo
								If oCargo.ObjectID = lObjID AndAlso oCargo.ObjTypeID = iObjTypeID Then
									lFromID = oTmp.ObjectID
									iFromTypeID = oTmp.ObjTypeID
									bFound = True
									Exit For
								End If
							Next
						End If
						If bFound = True Then Exit For
					Next
				End If

				If bFound = False Then
					MyBase.moUILib.AddNotification("Unable to determine owner of selection. Please try again or close and reopen the transfer window.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					btnTransferToColony.Enabled = True
					Return
				End If
                'Debug.Print("F: " & lFromID & " O:" & lObjID & " FT: " & iFromTypeID & " OT:" & iObjTypeID)
				PrepareAndSendTransferMsg(lFromID, iFromTypeID, goAvailableResources.lColonyID, ObjectType.eColony, lObjID, iObjTypeID, lQuantity)
				MyBase.moUILib.AddNotification("Transfer request sent. Waiting for shipment to finish...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Else
				MyBase.moUILib.AddNotification("You can only transfer components, minerals and personnel to the colony.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			End If

        Else
            MyBase.moUILib.AddNotification("Select an item to transfer to the colony.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If

		btnTransferToColony.Enabled = True
	End Sub

    Private Sub PrepareAndSendTransferMsg(ByVal lFromID As Int32, ByVal iFromTypeID As Int16, ByVal lToID As Int32, ByVal iToTypeID As Int16, ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal lQuantity As Int32)
        Dim yMsg(23) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eTransferContents).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(lObjID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, 6)
        System.BitConverter.GetBytes(lFromID).CopyTo(yMsg, 8)
        System.BitConverter.GetBytes(iFromTypeID).CopyTo(yMsg, 12)
        System.BitConverter.GetBytes(lToID).CopyTo(yMsg, 14)
        System.BitConverter.GetBytes(iToTypeID).CopyTo(yMsg, 18)
        System.BitConverter.GetBytes(lQuantity).CopyTo(yMsg, 20)

        MyBase.moUILib.SendMsgToPrimary(yMsg)

        'MyBase.moUILib.AddNotification("Transfer (" & lObjID & ", " & iObjTypeID & ") From (" & lFromID & ", " & iFromTypeID & ") To (" & lToID & ", " & iToTypeID & ") x" & lQuantity, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub frmTransfer_WindowClosed() Handles Me.WindowClosed
        'goAvailableResources = Nothing
        bReRequestResources = True
    End Sub
End Class