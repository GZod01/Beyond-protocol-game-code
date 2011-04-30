Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D


'Interface created from Interface Builder
Public Class frmCommand
	Inherits UIWindow

	Private lblTitle As UILabel
	Private WithEvents btnClose As UIButton
	Private lnDiv1 As UILine
	Private WithEvents tvwMain As UITreeView

	Private WithEvents optFacilities As UIOption
    Private WithEvents optUnits As UIOption
    Private WithEvents optUniverse As UIOption

    Private mlLastRefresh As Int32
    Private mbLoading As Boolean = True

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenCommandManagementWindow)

		'frmCommand initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eCommand
            .ControlName = "frmCommand"
            Dim oWin As UIWindow = oUILib.GetWindow("frmQuickBar")
            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If oWin Is Nothing = False Then
                lLeft = oWin.Left + oWin.Width
                If lLeft + 256 > oUILib.oDevice.PresentationParameters.BackBufferWidth Then
                    lLeft = -1
                End If
            End If
            oWin = Nothing

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.CommandX
                lTop = muSettings.CommandY
            End If

            .Left = lLeft
            .Top = lTop

            If lLeft = -1 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            If lTop = -1 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256

            .Left = lLeft
            .Top = lTop
            .Width = 256
            .Height = 512
            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .BorderLineWidth = 2
            .mbAcceptReprocessEvents = True
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 5
			.Width = 170
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Command Management"
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

		'lnDiv1 initial props
		lnDiv1 = New UILine(oUILib)
		With lnDiv1
			.ControlName = "lnDiv1"
            .Left = Me.BorderLineWidth
			.Top = 25
            .Width = Me.Width - Me.BorderLineWidth
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'optFacilities initial props
		optFacilities = New UIOption(oUILib)
		With optFacilities
			.ControlName = "optFacilities"
            .Left = 10
			.Top = 30
            .Width = 66
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "Facilities"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = True
		End With
		Me.AddChild(CType(optFacilities, UIControl))

		'optUnits initial props
		optUnits = New UIOption(oUILib)
		With optUnits
			.ControlName = "optUnits"
            .Left = 100
			.Top = 30
            .Width = 42
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "Units"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
        Me.AddChild(CType(optUnits, UIControl))

        'optUniverse initial props
        optUniverse = New UIOption(oUILib)
        With optUniverse
            .ControlName = "optUniverse"
            .Left = 175
            .Top = 30
            .Width = 70
            .Height = 20
            .Enabled = Not gbAliased
            .Visible = isAdmin()
            .Caption = "Universal"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optUniverse, UIControl))

		'tvwMain initial props
		tvwMain = New UITreeView(oUILib)
		With tvwMain
			.ControlName = "tvwMain"
			.Left = 5
			.Top = 55
			.Width = 246
			.Height = 452
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
		End With
		Me.AddChild(CType(tvwMain, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewUnitsAndFacilities) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack rights to view Units and Facilities.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        mbLoading = False
		FillList()
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub optFacilities_Click() Handles optFacilities.Click
		optFacilities.Value = True
        optUnits.Value = False
        optUniverse.Value = False
        'tvwMain.Clear()
		FillList()
	End Sub

	Private Sub optUnits_Click() Handles optUnits.Click
		optUnits.Value = True
        optFacilities.Value = False
        optUniverse.Value = False
        'tvwMain.Clear()
		FillList()
    End Sub

    Private Sub optUniverse_click() Handles optUniverse.Click
        optUniverse.Value = True
        optUnits.Value = False
        optFacilities.Value = False
        FillList()
    End Sub

	Private Sub frmCommand_OnNewFrame() Handles Me.OnNewFrame
		Dim clrBlack As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 32, 255)
		Dim clrRed As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
		Dim clrOrange As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 128, 64)
		Dim clrYellow As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 0)
        Dim clrGreen As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        Dim clrUnknown As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        Static lCurrentSystemIdx As Int32 = -1
        If goCurrentEnvir Is Nothing = False Then
            If lCurrentSystemIdx <> goCurrentEnvir.ObjectID Then
                tvwMain.Clear()
                tvwMain.IsDirty = True
            End If
        End If

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return
        'If optUniverse.Value = True Then Return

        Try

            'Check the environment for entities that are in the environment but not in our list
            If optUniverse.Value = False AndAlso glCurrentCycle - mlLastRefresh > 150 Then
                If goCurrentEnvir Is Nothing = False Then
                    If lCurrentSystemIdx <> goCurrentEnvir.ObjectID Then
                        FillList()
                    End If
                    lCurrentSystemIdx = goCurrentEnvir.ObjectID
                End If
                mlLastRefresh = glCurrentCycle
                If oEnvir.lEntityIdx Is Nothing = False Then
                    Dim lCurUB As Int32 = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx(0))
                    Dim iReqTypeID As Int16
                    If optFacilities.Value = True Then iReqTypeID = ObjectType.eFacility Else iReqTypeID = ObjectType.eUnit
                    Try
                        For X As Int32 = 0 To lCurUB
                            If oEnvir.lEntityIdx(X) <> -1 AndAlso oEnvir.oEntity(X).ObjTypeID = iReqTypeID AndAlso oEnvir.oEntity(X).OwnerID = glPlayerID Then
                                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                                Dim oChkNode As UITreeView.UITreeViewItem = tvwMain.GetNodeByItemData2(oEntity.ObjectID, oEntity.ObjTypeID)
                                If oChkNode Is Nothing Then
                                    Dim oParentRoot As UITreeView.UITreeViewItem = Nothing
                                    If optFacilities.Value = True Then
                                        Dim yProdType As Byte = oEntity.yProductionType
                                        If yProdType = ProductionType.eEnlisted OrElse yProdType = ProductionType.eOfficers Then yProdType = ProductionType.eEnlistedAndOfficers
                                        oParentRoot = tvwMain.GetNodeByItemData3(-1, yProdType, -1)
                                        If oParentRoot Is Nothing Then
                                            Select Case oEntity.yProductionType
                                                Case ProductionType.eAerialProduction
                                                    oParentRoot = tvwMain.AddNode("Spaceports (1)", -1, ProductionType.eAerialProduction, -1, Nothing, Nothing)
                                                Case ProductionType.eColonists
                                                    oParentRoot = tvwMain.AddNode("Residences (1)", -1, ProductionType.eColonists, -1, Nothing, Nothing)
                                                Case ProductionType.eEnlisted, ProductionType.eOfficers
                                                    oParentRoot = tvwMain.AddNode("Personnel (1)", -1, ProductionType.eEnlistedAndOfficers, -1, Nothing, Nothing)
                                                Case ProductionType.eLandProduction
                                                    oParentRoot = tvwMain.AddNode("Factories (1)", -1, ProductionType.eLandProduction, -1, Nothing, Nothing)
                                                Case ProductionType.eMining
                                                    oParentRoot = tvwMain.AddNode("Mining Facilites (1)", -1, ProductionType.eMining, -1, Nothing, Nothing)
                                                Case ProductionType.eNavalProduction
                                                    oParentRoot = tvwMain.AddNode("Naval Factories (1)", -1, ProductionType.eNavalProduction, -1, Nothing, Nothing)
                                                Case ProductionType.ePowerCenter
                                                    oParentRoot = tvwMain.AddNode("Power Generators (1)", -1, ProductionType.ePowerCenter, -1, Nothing, Nothing)
                                                Case ProductionType.eRefining
                                                    oParentRoot = tvwMain.AddNode("Refineries (1)", -1, ProductionType.eRefining, -1, Nothing, Nothing)
                                                Case ProductionType.eResearch
                                                    oParentRoot = tvwMain.AddNode("Research Labs (1)", -1, ProductionType.eResearch, -1, Nothing, Nothing)
                                                Case ProductionType.eSpaceStationSpecial
                                                    oParentRoot = tvwMain.AddNode("Space Stations (1)", -1, ProductionType.eSpaceStationSpecial, -1, Nothing, Nothing)
                                                Case ProductionType.eWareHouse
                                                    oParentRoot = tvwMain.AddNode("Warehouses (1)", -1, ProductionType.eWareHouse, -1, Nothing, Nothing)
                                            End Select
                                            If (oEntity.oMesh.lModelID And 255) = 11 Then
                                                oParentRoot = tvwMain.AddNode("Warehouses (1)", -1, ProductionType.eWareHouse, -1, Nothing, Nothing)
                                            End If
                                            If oParentRoot Is Nothing = False Then oParentRoot.bItemBold = True
                                        End If
                                    ElseIf optUnits.Value = True Then
                                        'unit
                                        Dim oModelDef As ModelDef = goModelDefs.GetModelDef(oEntity.oUnitDef.ModelID)
                                        If oModelDef Is Nothing Then Continue For
                                        Dim lTestID2 As Int32 = oModelDef.TypeID
                                        Dim lTestID3 As Int32 = oModelDef.SubTypeID
                                        If lTestID2 = 3 OrElse lTestID2 = 4 OrElse lTestID2 = 5 OrElse lTestID2 = 7 Then lTestID3 = -1
                                        oParentRoot = tvwMain.GetNodeByItemData3(-1, lTestID2, lTestID3)
                                        If oParentRoot Is Nothing Then
                                            Select Case lTestID2
                                                Case 0
                                                    If lTestID3 = 0 Then
                                                        oParentRoot = tvwMain.AddNode("Battleships (1)", -1, 0, 0, Nothing, Nothing)
                                                    Else : oParentRoot = tvwMain.AddNode("Battlecruisers (1)", -1, 0, 2, Nothing, Nothing)
                                                    End If
                                                Case 1
                                                    Select Case lTestID3
                                                        Case 0 : oParentRoot = tvwMain.AddNode("Corvettes (1)", -1, 1, 0, Nothing, Nothing)
                                                        Case 1 : oParentRoot = tvwMain.AddNode("Cruisers (1)", -1, 1, 1, Nothing, Nothing)
                                                        Case 2 : oParentRoot = tvwMain.AddNode("Destroyers (1)", -1, 1, 2, Nothing, Nothing)
                                                        Case 3 : oParentRoot = tvwMain.AddNode("Escorts (1)", -1, 1, 4, Nothing, Nothing)
                                                        Case Else
                                                            oParentRoot = tvwMain.AddNode("Frigates (1)", -1, 1, 3, Nothing, Nothing)
                                                    End Select
                                                Case 3 : oParentRoot = tvwMain.AddNode("Fighters (1)", -1, 3, -1, Nothing, Nothing)
                                                Case 4 : oParentRoot = tvwMain.AddNode("Small Vehicles (1)", -1, 4, -1, Nothing, Nothing)
                                                Case 5 : oParentRoot = tvwMain.AddNode("Tanks (1)", -1, 5, -1, Nothing, Nothing)
                                                Case 7 : oParentRoot = tvwMain.AddNode("Utility (1)", -1, 7, -1, Nothing, Nothing)
                                            End Select
                                            If oParentRoot Is Nothing = False Then oParentRoot.bItemBold = True
                                        End If
                                    ElseIf optUniverse.Value = True Then
                                    End If

                                    If oParentRoot Is Nothing = False Then
                                        Dim lParenIdx As Int32 = oParentRoot.sItem.IndexOf("("c)
                                        If lParenIdx <> -1 Then
                                            lParenIdx += 1
                                            Dim lChildCnt As Int32 = CInt(Val(oParentRoot.sItem.Substring(lParenIdx)))
                                            lChildCnt += 1
                                            oParentRoot.sItem = oParentRoot.sItem.Substring(0, lParenIdx) & lChildCnt & ")"
                                        End If
                                    End If

                                    tvwMain.AddNode(" ", oEntity.ObjectID, oEntity.ObjTypeID, X, oParentRoot, Nothing)

                                End If
                            End If
                        Next X
                    Catch
                    End Try
                End If
            End If

            Dim oNode As UITreeView.UITreeViewItem = tvwMain.GetFirstVisibleNode()

            'ok, let's traverse them
            Dim lRenderCnt As Int32 = tvwMain.SingleScreenNodeRenderCnt()
            Dim bDirty As Boolean = False
            While lRenderCnt <> 0
                If oNode Is Nothing Then Exit While
                If oNode.lItemData > 0 AndAlso oNode.lItemData2 > 0 Then
                    If optUniverse.Value = False Then
                        If oNode.lItemData3 > -1 Then
                            If oEnvir.lEntityIdx(oNode.lItemData3) <> -1 Then
                                Dim oEntity As BaseEntity = oEnvir.oEntity(oNode.lItemData3)

                                If oEntity Is Nothing OrElse oEntity.ObjectID <> oNode.lItemData OrElse oEntity.ObjTypeID <> oNode.lItemData2 Then
                                    If oNode.oParentNode Is Nothing = False Then
                                        Dim lParenIdx As Int32 = oNode.oParentNode.sItem.IndexOf("("c)
                                        If lParenIdx <> -1 Then
                                            lParenIdx += 1
                                            Dim lChildCnt As Int32 = CInt(Val(oNode.oParentNode.sItem.Substring(lParenIdx)))
                                            lChildCnt -= 1
                                            oNode.oParentNode.sItem = oNode.oParentNode.sItem.Substring(0, lParenIdx) & lChildCnt & ")"
                                        End If
                                    End If
                                    bDirty = True
                                    tvwMain.RemoveNode(oNode)
                                    oNode = oNode.oNextSibling
                                    Continue While
                                End If
                                Dim clrVal As System.Drawing.Color = clrUnknown
                                If oEntity.bRequestedDetails = False Then
                                    Dim yMsg(13) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yMsg, 0)
                                    oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
                                    goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, 8)
                                    MyBase.moUILib.SendMsgToRegion(yMsg)
                                    oEntity.bRequestedDetails = True
                                End If
                                If oEntity.lLastHPUpdate <= 0 Then
                                    clrVal = clrUnknown
                                    If oEntity.lLastHPUpdate = Int32.MinValue Then
                                        oEntity.lLastHPUpdate = Int32.MinValue + 1
                                        Dim yData(7) As Byte
                                        System.BitConverter.GetBytes(GlobalMessageCode.eUnitHPUpdate).CopyTo(yData, 0)
                                        System.BitConverter.GetBytes(oEntity.ObjectID).CopyTo(yData, 2)
                                        System.BitConverter.GetBytes(oEntity.ObjTypeID).CopyTo(yData, 6)
                                        MyBase.moUILib.SendMsgToRegion(yData)
                                    End If
                                Else
                                    clrVal = clrGreen

                                    If (oEntity.CritList And (elUnitStatus.eEngineOperational Or elUnitStatus.eRadarOperational)) <> 0 Then
                                        '  Black is combat in-capable (if the unit was capable of combat that is)
                                        clrVal = clrBlack
                                    ElseIf oEntity.yStructureHP < 75 Then
                                        '  Red is 75% or less structure remains
                                        clrVal = clrRed
                                    Else
                                        For Z As Int32 = 0 To oEntity.yArmorHP.GetUpperBound(0)
                                            If oEntity.yArmorHP(Z) = 0 Then
                                                '  Orange is at least one side at 0
                                                clrVal = clrOrange
                                                Exit For
                                            ElseIf oEntity.yArmorHP(Z) <> 100 AndAlso oEntity.yArmorHP(Z) <> 255 Then
                                                '  yellow is damaged
                                                clrVal = clrYellow
                                            End If
                                        Next Z
                                    End If
                                    If oNode.clrItemColor <> clrVal Then
                                        oNode.clrItemColor = clrVal
                                        bDirty = True
                                    End If
                                    oNode.bUseItemColor = True
                                End If
                            Else
                                If oNode.oParentNode Is Nothing = False Then
                                    Dim lParenIdx As Int32 = oNode.oParentNode.sItem.IndexOf("("c)
                                    If lParenIdx <> -1 Then
                                        lParenIdx += 1
                                        Dim lChildCnt As Int32 = CInt(Val(oNode.oParentNode.sItem.Substring(lParenIdx)))
                                        lChildCnt -= 1
                                        oNode.oParentNode.sItem = oNode.oParentNode.sItem.Substring(0, lParenIdx) & lChildCnt & ")"
                                    End If
                                End If
                                bDirty = True
                                tvwMain.RemoveNode(oNode)
                                oNode = oNode.oNextSibling
                                Continue While
                            End If
                        End If
                    End If

                    If CShort(oNode.lItemData2) = ObjectType.eWormhole Then
                        If optFacilities.Value = True Then
                            Dim oSystem As SolarSystem = CType(goCurrentEnvir.oGeoObject, SolarSystem)
                            If oSystem.WormholeUB <> -1 Then
                                For X As Int32 = 0 To oSystem.WormholeUB
                                    With oSystem.moWormholes(X)
                                        If oSystem.moWormholes(X).System1.ObjectID = oNode.lItemData Then
                                            Dim sTemp As String = oSystem.moWormholes(X).System1.SystemName
                                            If sTemp Is Nothing OrElse sTemp = "" Then sTemp = "Unknown"
                                            sTemp = "Wormhole To " & sTemp
                                            If oNode.sItem <> sTemp Then
                                                oNode.sItem = sTemp
                                                bDirty = True
                                            End If
                                        ElseIf oSystem.moWormholes(X).System2.ObjectID = oNode.lItemData Then
                                            Dim sTemp As String = oSystem.moWormholes(X).System2.SystemName
                                            If sTemp Is Nothing OrElse sTemp = "" Then sTemp = "Unknown"
                                            sTemp = "Wormhole To " & sTemp
                                            If oNode.sItem <> sTemp Then
                                                oNode.sItem = sTemp
                                                bDirty = True
                                            End If
                                        End If
                                    End With
                                Next X
                            End If
                        End If
                    Else
                        Dim sTemp As String = GetCacheObjectValue(oNode.lItemData, CShort(oNode.lItemData2))
                        If optUniverse.Value = True Then
                            sTemp = sTemp & " (" & oNode.lItemData3.ToString & ")"
                        End If
                        If oNode.sItem <> sTemp Then
                            oNode.sItem = sTemp
                            bDirty = True
                        End If
                    End If
                End If
                lRenderCnt -= 1
                oNode = UITreeView.TraverseNextNode(oNode)
            End While
            If bDirty = True Then Me.IsDirty = True
        Catch
        End Try
        Try
            If glCurrentEnvirView = CurrentView.ePlanetMapView AndAlso goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                Dim oNode As UITreeView.UITreeViewItem = tvwMain.oSelectedNode
                If oNode Is Nothing = False AndAlso oNode.lTop <> Int32.MaxValue Then
                    Dim pt As Point = tvwMain.GetAbsolutePosition()

                    Dim lHalfExtent As Int32 = CType(goCurrentEnvir.oGeoObject, Planet).GetExtent \ 2
                    Dim lFullExtent As Int32 = CType(goCurrentEnvir.oGeoObject, Planet).GetExtent


                    Dim lTempW As Int32 = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth
                    Dim lTempH As Int32 = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight
                    Dim lSmaller As Int32 = CInt(Math.Min(lTempW, lTempH) * 0.8)
                    Dim lLarger As Int32 = Math.Max(lTempW, lTempH)
                    Dim ptPos As Point
                    ptPos.X = CInt((lTempW - lSmaller) / 2)
                    ptPos.Y = CInt((lTempH - lSmaller) / 2)

                    Dim lSrcRectWH As Int32 = muSettings.PlanetModelTextureWH - 1
                    If MyBase.moUILib.oDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = False OrElse muSettings.HiResPlanetTexture = False Then lSrcRectWH = 255


                    For X As Int32 = 0 To oEnvir.lEntityUB
                        If oEnvir.lEntityIdx(X) = oNode.lItemData Then
                            Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                            If oEntity Is Nothing Then Continue For
                            With oEntity
                                If .ObjTypeID = oNode.lItemData2 Then
                                    Dim lLocX As Int32 = CInt(((.LocX + lHalfExtent) / lFullExtent) * lSmaller)
                                    Dim lLocZ As Int32 = CInt(((.LocZ + lHalfExtent) / lFullExtent) * lSmaller)

                                    Dim v2Pts(1) As Vector2
                                    Dim lFinalX As Int32 = (lSmaller - lLocX) + ptPos.X
                                    If lFinalX > pt.X + Me.Width Then pt.X += Me.Width
                                    v2Pts(0) = New Vector2(pt.X, pt.Y + oNode.lTop - tvwMain.Top + 9)
                                    v2Pts(1) = New Vector2(lFinalX, lLocZ + ptPos.Y)

                                    'ValidateBorderLine()
                                    'With moBorderLine
                                    '    .Antialias = True
                                    '    .Width = 2
                                    '    .Begin()
                                    '    .Draw(v2Pts, Color.White)
                                    '    .End()
                                    'End With
                                    BPLine.DrawLine(2, True, v2Pts, Color.White)

                                    Exit For
                                End If
                            End With
                        End If
                    Next X

                End If
            End If
        Catch
        End Try
    End Sub

	Private Sub FillList()
		Dim oEnvir As BaseEnvironment = goCurrentEnvir
		If oEnvir Is Nothing Then Return

		Dim clrBlack As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 32, 255)
		Dim clrRed As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
		Dim clrOrange As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 128, 64)
		Dim clrYellow As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 0)
        Dim clrGreen As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        Dim clrUnknown As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        tvwMain.Clear()
		If tvwMain.oRootNode Is Nothing Then
			'ok, let's set up our base root nodes
			If optFacilities.Value = True Then
				'Ok, facilities... go by productiontype
				Dim lSpacePorts As Int32 = 0
				Dim lResidentials As Int32 = 0
				Dim lBarracksAndOfficers As Int32 = 0
				Dim lFactories As Int32 = 0
				Dim lMines As Int32 = 0
				Dim lNavalFacs As Int32 = 0
				Dim lPowerGens As Int32 = 0
				Dim lRefineries As Int32 = 0
				Dim lResearchLabs As Int32 = 0
				Dim lSpaceStations As Int32 = 0
                Dim lWarehouses As Int32 = 0
                Dim lTurrets As Int32 = 0
                'Dim lWormHoles As Int32 = 0

				Dim oFactoryRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Factories", -1, ProductionType.eLandProduction, -1, Nothing, Nothing)
				Dim oMineRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Mining Facilites", -1, ProductionType.eMining, -1, Nothing, Nothing)
				Dim oNavalFacRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Naval Factories", -1, ProductionType.eNavalProduction, -1, Nothing, Nothing)
				Dim oPersonnelRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Personnel", -1, ProductionType.eEnlistedAndOfficers, -1, Nothing, Nothing)
				Dim oPowerGenRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Power Generators", -1, ProductionType.ePowerCenter, -1, Nothing, Nothing)
				Dim oRefineryRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Refineries", -1, ProductionType.eRefining, -1, Nothing, Nothing)
				Dim oResearchRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Research Labs", -1, ProductionType.eResearch, -1, Nothing, Nothing)
				Dim oResidentialRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Residences", -1, ProductionType.eColonists, -1, Nothing, Nothing)
				Dim oSpacePortRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Spaceports", -1, ProductionType.eAerialProduction, -1, Nothing, Nothing)
				Dim oStationRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Space Stations", -1, ProductionType.eSpaceStationSpecial, -1, Nothing, Nothing)
                Dim oWarehouseRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Warehouses", -1, ProductionType.eWareHouse, -1, Nothing, Nothing)
                Dim oTurretRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Turrets", -1, -1, -1, Nothing, Nothing)
                'Dim oWormholeRoot As UITreeView.UITreeViewItem = tvwMain.AddNode("Wormholes", -1, -1, -1, Nothing, Nothing)

				oFactoryRoot.bItemBold = True
				oMineRoot.bItemBold = True
				oNavalFacRoot.bItemBold = True
				oPersonnelRoot.bItemBold = True
				oPowerGenRoot.bItemBold = True
				oRefineryRoot.bItemBold = True
				oResearchRoot.bItemBold = True
				oResidentialRoot.bItemBold = True
				oSpacePortRoot.bItemBold = True
				oStationRoot.bItemBold = True
                oWarehouseRoot.bItemBold = True
                oTurretRoot.bItemBold = True
                'oWormholeRoot.bItemBold = True

                Dim lCurUB As Int32 = -1
                Dim iLastIndex As Int32 = 0
                If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
				For X As Int32 = 0 To lCurUB
					If oEnvir.lEntityIdx(X) <> -1 Then
						Dim oEntity As BaseEntity = oEnvir.oEntity(X)
						If oEntity Is Nothing = False Then
							If oEntity.OwnerID = glPlayerID AndAlso oEntity.ObjTypeID = ObjectType.eFacility Then

								Dim oParent As UITreeView.UITreeViewItem = Nothing
								Dim oBefore As UITreeView.UITreeViewItem = Nothing

                                Select Case oEntity.yProductionType
                                    Case 0
                                        If oEntity.oMesh Is Nothing = False Then
                                            Select Case oEntity.oMesh.lModelID
                                                Case 138
                                                    lTurrets += 1
                                                    oParent = oTurretRoot : oBefore = Nothing
                                                Case Else
                                                    oParent = Nothing : oBefore = Nothing
                                            End Select
                                        Else
                                            oParent = Nothing : oBefore = Nothing
                                        End If
                                    Case ProductionType.eAerialProduction
                                        lSpacePorts += 1
                                        oParent = oSpacePortRoot : oBefore = Nothing
                                    Case ProductionType.eColonists
                                        lResidentials += 1
                                        oParent = oResidentialRoot : oBefore = Nothing
                                    Case ProductionType.eEnlisted, ProductionType.eOfficers
                                        lBarracksAndOfficers += 1
                                        oParent = oPersonnelRoot : oBefore = Nothing
                                    Case ProductionType.eLandProduction
                                        lFactories += 1
                                        oParent = oFactoryRoot : oBefore = Nothing
                                    Case ProductionType.eMining
                                        lMines += 1
                                        oParent = oMineRoot : oBefore = Nothing
                                    Case ProductionType.eNavalProduction
                                        lNavalFacs += 1
                                        oParent = oNavalFacRoot : oBefore = Nothing
                                    Case ProductionType.ePowerCenter
                                        lPowerGens += 1
                                        oParent = oPowerGenRoot : oBefore = Nothing
                                    Case ProductionType.eRefining
                                        lRefineries += 1
                                        oParent = oRefineryRoot : oBefore = Nothing
                                    Case ProductionType.eResearch
                                        lResearchLabs += 1
                                        oParent = oResearchRoot : oBefore = Nothing
                                    Case ProductionType.eSpaceStationSpecial
                                        lSpaceStations += 1
                                        oParent = oStationRoot : oBefore = Nothing
                                    Case ProductionType.eWareHouse
                                        lWarehouses += 1
                                        oParent = oWarehouseRoot : oBefore = Nothing
                                    Case Else
                                        oParent = Nothing : oBefore = Nothing
                                End Select

                                Dim sTemp As String = oEntity.EntityName
                                If sTemp Is Nothing OrElse sTemp = "" Then sTemp = "Unknown"

                                Dim oNewNode As UITreeView.UITreeViewItem = tvwMain.AddNode(sTemp, oEntity.ObjectID, oEntity.ObjTypeID, X, oParent, oBefore)
                                iLastIndex = X
                                If oNewNode Is Nothing = False Then
                                    Dim clrVal As System.Drawing.Color = clrUnknown
                                    If oEntity.bRequestedDetails = False Then
                                        Dim yMsg(13) As Byte
                                        System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yMsg, 0)
                                        oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
                                        goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, 8)
                                        MyBase.moUILib.SendMsgToRegion(yMsg)
                                        oEntity.bRequestedDetails = True
                                    End If
                                    If oEntity.lLastHPUpdate <= 0 Then
                                        clrVal = clrUnknown
                                        If oEntity.lLastHPUpdate = Int32.MinValue Then
                                            oEntity.lLastHPUpdate = Int32.MinValue + 1
                                            Dim yData(7) As Byte
                                            System.BitConverter.GetBytes(GlobalMessageCode.eUnitHPUpdate).CopyTo(yData, 0)
                                            System.BitConverter.GetBytes(oEntity.ObjectID).CopyTo(yData, 2)
                                            System.BitConverter.GetBytes(oEntity.ObjTypeID).CopyTo(yData, 6)
                                            MyBase.moUILib.SendMsgToRegion(yData)
                                        End If
                                    Else
                                        clrVal = clrGreen

                                        If (oEntity.CritList And (elUnitStatus.eEngineOperational Or elUnitStatus.eRadarOperational)) <> 0 Then
                                            '  Black is combat in-capable (if the unit was capable of combat that is)
                                            clrVal = clrBlack
                                        ElseIf oEntity.yStructureHP < 75 Then
                                            '  Red is 75% or less structure remains
                                            clrVal = clrRed
                                        Else
                                            For Z As Int32 = 0 To oEntity.yArmorHP.GetUpperBound(0)
                                                If oEntity.yArmorHP(Z) = 0 Then
                                                    '  Orange is at least one side at 0
                                                    clrVal = clrOrange
                                                    Exit For
                                                ElseIf oEntity.yArmorHP(Z) <> 100 AndAlso oEntity.yArmorHP(Z) <> 255 Then
                                                    '  yellow is damaged
                                                    clrVal = clrYellow
                                                End If
                                            Next Z
                                        End If
                                        oNewNode.clrItemColor = clrVal
                                        oNewNode.bUseItemColor = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next X

                If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso goCurrentEnvir.oGeoObject Is Nothing = False AndAlso CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                    Dim oSystem As SolarSystem = CType(goCurrentEnvir.oGeoObject, SolarSystem)
                    If oSystem.WormholeUB <> -1 Then
                        For X As Int32 = 0 To oSystem.WormholeUB
                            With oSystem.moWormholes(X)
                                Dim oNewNode As UITreeView.UITreeViewItem = Nothing
                                If .System1 Is Nothing = False AndAlso .System1.ObjectID <> oSystem.ObjectID Then
                                    Dim sTemp As String = oSystem.moWormholes(X).System1.SystemName
                                    If sTemp Is Nothing OrElse sTemp = "" Then sTemp = "Unknown"
                                    sTemp = "Wormhole To " & sTemp
                                    oNewNode = tvwMain.AddNode(sTemp, oSystem.moWormholes(X).ObjectID, oSystem.moWormholes(X).ObjTypeID, -1, Nothing, Nothing)

                                    'lWormHoles += 1
                                Else
                                    Dim sTemp As String = oSystem.moWormholes(X).System2.SystemName
                                    If sTemp Is Nothing OrElse sTemp = "" Then sTemp = "Unknown"
                                    sTemp = "Wormhole To " & sTemp
                                    oNewNode = tvwMain.AddNode(sTemp, oSystem.moWormholes(X).ObjectID, oSystem.moWormholes(X).ObjTypeID, -1, Nothing, Nothing)
                                    'lWormHoles += 1
                                End If
                                If oNewNode Is Nothing = False Then
                                    Dim clrVal As System.Drawing.Color = clrGreen
                                    oNewNode.clrItemColor = clrVal
                                    oNewNode.bUseItemColor = True

                                End If
                            End With
                        Next X

                    End If
                End If

                If lSpacePorts = 0 Then tvwMain.RemoveNode(oSpacePortRoot) Else oSpacePortRoot.sItem = (oSpacePortRoot.sItem & " (" & lSpacePorts & ")") '.ToUpper
                If lResidentials = 0 Then tvwMain.RemoveNode(oResidentialRoot) Else oResidentialRoot.sItem = (oResidentialRoot.sItem & " (" & lResidentials & ")") '.ToUpper
                If lBarracksAndOfficers = 0 Then tvwMain.RemoveNode(oPersonnelRoot) Else oPersonnelRoot.sItem = (oPersonnelRoot.sItem & " (" & lBarracksAndOfficers & ")") '.ToUpper
                If lFactories = 0 Then tvwMain.RemoveNode(oFactoryRoot) Else oFactoryRoot.sItem = (oFactoryRoot.sItem & " (" & lFactories & ")") '.ToUpper
                If lMines = 0 Then tvwMain.RemoveNode(oMineRoot) Else oMineRoot.sItem = (oMineRoot.sItem & " (" & lMines & ")") '.ToUpper
                If lNavalFacs = 0 Then tvwMain.RemoveNode(oNavalFacRoot) Else oNavalFacRoot.sItem = (oNavalFacRoot.sItem & " (" & lNavalFacs & ")") '.ToUpper
                If lPowerGens = 0 Then tvwMain.RemoveNode(oPowerGenRoot) Else oPowerGenRoot.sItem = (oPowerGenRoot.sItem & " (" & lPowerGens & ")") '.ToUpper
                If lRefineries = 0 Then tvwMain.RemoveNode(oRefineryRoot) Else oRefineryRoot.sItem = (oRefineryRoot.sItem & " (" & lRefineries & ")") '.ToUpper
                If lResearchLabs = 0 Then tvwMain.RemoveNode(oResearchRoot) Else oResearchRoot.sItem = (oResearchRoot.sItem & " (" & lResearchLabs & ")") '.ToUpper
                If lSpaceStations = 0 Then tvwMain.RemoveNode(oStationRoot) Else oStationRoot.sItem = (oStationRoot.sItem & " (" & lSpaceStations & ")") '.ToUpper
                If lWarehouses = 0 Then tvwMain.RemoveNode(oWarehouseRoot) Else oWarehouseRoot.sItem = (oWarehouseRoot.sItem & " (" & lWarehouses & ")") '.ToUpper
                If lTurrets = 0 Then tvwMain.RemoveNode(oTurretRoot) Else oTurretRoot.sItem = (oTurretRoot.sItem & " (" & lTurrets & ")") '.ToUpper
                'If lWormHoles = 0 Then tvwMain.RemoveNode(oWormholeRoot) Else oWormholeRoot.sItem = (oWormholeRoot.sItem & " (" & lWormHoles & ")") '.ToUpper
            ElseIf optUnits.Value = True Then
                'I can distinguish between tank, utility, corvette, fighter, etc...

                Dim lCnts(10) As Int32
                Dim oRoots(10) As UITreeView.UITreeViewItem

                For X As Int32 = 0 To lCnts.GetUpperBound(0)
                    lCnts(X) = 0
                Next X
                oRoots(0) = tvwMain.AddNode("Battlecruisers", -1, 0, 2, Nothing, Nothing)
                oRoots(1) = tvwMain.AddNode("Battleships", -1, 0, 0, Nothing, Nothing)
                oRoots(2) = tvwMain.AddNode("Corvettes", -1, 1, 0, Nothing, Nothing)
                oRoots(3) = tvwMain.AddNode("Cruisers", -1, 1, 1, Nothing, Nothing)
                oRoots(4) = tvwMain.AddNode("Destroyers", -1, 1, 2, Nothing, Nothing)
                oRoots(5) = tvwMain.AddNode("Escorts", -1, 1, 4, Nothing, Nothing)
                oRoots(6) = tvwMain.AddNode("Fighters", -1, 3, -1, Nothing, Nothing)
                oRoots(7) = tvwMain.AddNode("Frigates", -1, 1, 3, Nothing, Nothing)
                oRoots(8) = tvwMain.AddNode("Small Vehicles", -1, 4, -1, Nothing, Nothing)
                oRoots(9) = tvwMain.AddNode("Tanks", -1, 5, -1, Nothing, Nothing)
                oRoots(10) = tvwMain.AddNode("Utility", -1, 7, -1, Nothing, Nothing)

                Dim lCurUB As Int32 = -1
                If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
                For X As Int32 = 0 To lCurUB
                    If oEnvir.lEntityIdx(X) <> -1 Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False Then
                            If oEntity.OwnerID = glPlayerID AndAlso oEntity.ObjTypeID = ObjectType.eUnit Then
                                If oEntity.oUnitDef Is Nothing = False Then
                                    Dim oModelDef As ModelDef = goModelDefs.GetModelDef(oEntity.oUnitDef.ModelID)
                                    If oModelDef Is Nothing = False Then

                                        Dim oParentNode As UITreeView.UITreeViewItem = Nothing
                                        For Y As Int32 = 0 To lCnts.GetUpperBound(0)
                                            If oRoots(Y).lItemData2 = oModelDef.TypeID Then
                                                If oRoots(Y).lItemData3 = oModelDef.SubTypeID OrElse oRoots(Y).lItemData3 = -1 Then
                                                    oParentNode = oRoots(Y)
                                                    lCnts(Y) += 1
                                                    Exit For
                                                End If
                                            End If
                                        Next Y

                                        Dim sTemp As String = oEntity.EntityName
                                        If sTemp Is Nothing OrElse sTemp = "" Then sTemp = "Unknown"

                                        Dim oNewNode As UITreeView.UITreeViewItem = tvwMain.AddNode(sTemp, oEntity.ObjectID, oEntity.ObjTypeID, X, oParentNode, Nothing)
                                        If oNewNode Is Nothing = False Then
                                            Dim clrVal As System.Drawing.Color = clrUnknown
                                            If oEntity.bRequestedDetails = False Then
                                                Dim yMsg(13) As Byte
                                                System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yMsg, 0)
                                                oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
                                                goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, 8)
                                                MyBase.moUILib.SendMsgToRegion(yMsg)
                                                oEntity.bRequestedDetails = True
                                            End If
                                            If oEntity.lLastHPUpdate <= 0 Then
                                                clrVal = clrUnknown
                                                If oEntity.lLastHPUpdate = Int32.MinValue Then
                                                    oEntity.lLastHPUpdate = Int32.MinValue + 1
                                                    Dim yData(7) As Byte
                                                    System.BitConverter.GetBytes(GlobalMessageCode.eUnitHPUpdate).CopyTo(yData, 0)
                                                    System.BitConverter.GetBytes(oEntity.ObjectID).CopyTo(yData, 2)
                                                    System.BitConverter.GetBytes(oEntity.ObjTypeID).CopyTo(yData, 6)
                                                    MyBase.moUILib.SendMsgToRegion(yData)
                                                End If
                                            Else

                                                clrVal = clrGreen
                                                If (oEntity.CritList And (elUnitStatus.eEngineOperational Or elUnitStatus.eRadarOperational)) <> 0 Then
                                                    '  Black is combat in-capable (if the unit was capable of combat that is)
                                                    clrVal = clrBlack
                                                ElseIf oEntity.yStructureHP < 75 Then
                                                    '  Red is 75% or less structure remains
                                                    clrVal = clrRed
                                                Else
                                                    For Z As Int32 = 0 To oEntity.yArmorHP.GetUpperBound(0)
                                                        If oEntity.yArmorHP(Z) = 0 Then
                                                            '  Orange is at least one side at 0
                                                            clrVal = clrOrange
                                                            Exit For
                                                        ElseIf oEntity.yArmorHP(Z) <> 100 AndAlso oEntity.yArmorHP(Z) <> 255 Then
                                                            '  yellow is damaged
                                                            clrVal = clrYellow
                                                        End If
                                                    Next Z
                                                End If
                                                oNewNode.clrItemColor = clrVal
                                                oNewNode.bUseItemColor = True
                                            End If
                                        End If
                                    End If
                                End If

                            End If
                        End If
                    End If
                Next X

                For X As Int32 = 0 To lCnts.GetUpperBound(0)
                    If lCnts(X) = 0 Then
                        tvwMain.RemoveNode(oRoots(X))
                    Else
                        oRoots(X).sItem = (oRoots(X).sItem & " (" & lCnts(X) & ")") '.ToUpper
                        oRoots(X).bItemBold = True
                    End If
                Next X
            ElseIf optUniverse.Value = True Then
                If mbUniverseLoaded = True Then 'Load it up
                    FillListForUniverse()
                Else
                    ReadUniverseInventory()
                End If
            End If
        End If
	End Sub

    'Universal Inventory
    Private Shared mbUniverseRequested As Boolean = False
    Private Shared mbUniverseLoaded As Boolean = False
    Private Structure T_Inventory
        Dim iEnvirTypeID As Int16
        Dim lEnvirID As Int32
        Dim lParentStarID As Int32
        Dim iParentTypeID As Int16
        Dim iModelDefTypeID As Int32
        Dim lModelSubTypeID As Int32
        Dim oEntityDef As EntityDef

        Dim sModelName As String
        Dim lUnitCount As Int32
    End Structure
    Private Shared moUniverse() As T_Inventory

    Public Sub HandleUniverseInventory(ByRef yData() As Byte)
        'Read from packet
        Dim lPos As Int32 = 2   '2 for msgcode

        Dim EntityUB As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ReDim moUniverse(EntityUB)
        Dim x As Int32 = 0
        Try
            For x = 0 To EntityUB
                With moUniverse(x)
                    .iEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lParentStarID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .iParentTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .iModelDefTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .lModelSubTypeID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lUnitCount = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .sModelName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                    Debug.Print(.lParentStarID & "," & .iEnvirTypeID.ToString & "," & .lEnvirID.ToString & "," & .iModelDefTypeID.ToString & "," & .lModelSubTypeID.ToString & "," & .lUnitCount & "," & .sModelName.ToString)
                End With
            Next x
        Catch
        End Try

        'Sort


        'Write to cache file
        SaveUniverseInventory()

        mbUniverseLoaded = True

        'Just in case I closed, open, reset, etc.. Pop my window back into shape
        optFacilities.Value = False
        optUnits.Value = False
        optUniverse.Value = True
        tvwMain.Clear()
        FillListForUniverse()

    End Sub

    Private Sub ReadUniverseInventory()
        If goCurrentPlayer Is Nothing Then Return

        'Read from file
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= goCurrentPlayer.PlayerName & ".inv"

        If My.Computer.FileSystem.FileExists(sFile) = True Then
            Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Open)
            Dim oReader As IO.BinaryReader = New IO.BinaryReader(fsFile)
            Dim lPos As Int32 = 0
            Dim lCnt As Int32 = 0
            Try
                While fsFile.Position < fsFile.Length
                    Dim yHdr() As Byte = oReader.ReadBytes(42)
                    If yHdr Is Nothing = False Then
                        lPos = 0
                        ReDim Preserve moUniverse(lCnt)
                        With moUniverse(lCnt)
                            .iEnvirTypeID = System.BitConverter.ToInt16(yHdr, lPos) : lPos += 2
                            .lEnvirID = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .lParentStarID = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .iParentTypeID = System.BitConverter.ToInt16(yHdr, lPos) : lPos += 2
                            .iModelDefTypeID = System.BitConverter.ToInt16(yHdr, lPos) : lPos += 2
                            .lModelSubTypeID = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .lUnitCount = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .sModelName = GetStringFromBytes(yHdr, lPos, 20) : lPos += 20
                            lCnt += 1
                        End With
                    Else
                        Exit While
                    End If
                End While
            Catch
            End Try
            oReader.Close()
            oReader = Nothing
            fsFile.Close()
            fsFile.Dispose()
            fsFile = Nothing

            mbUniverseLoaded = True
            FillListForUniverse()
        End If

        If mbUniverseRequested = False Then 'Request a fresh copy from the server
            Dim yMsg(1) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestUniversalAssets).CopyTo(yMsg, 0)
            MyBase.moUILib.GetMsgSys.SendToPrimary(yMsg)
            mbUniverseRequested = True
        End If
    End Sub

    Private Sub SaveUniverseInventory()
        If goCurrentPlayer Is Nothing Then Return

        Dim UniverseUB As Int32 = moUniverse.GetUpperBound(0)

        'Write to file
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= goCurrentPlayer.PlayerName & ".inv"
        Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Create)

        Dim lPos As Int32 = 0
        Dim x As Int32 = 0
        Try
            For X = 0 To UniverseUB
                Dim yShortcutSave(41) As Byte
                lPos = 0
                With moUniverse(X)
                    System.BitConverter.GetBytes(.iEnvirTypeID).CopyTo(yShortcutSave, lPos) : lPos += 2
                    System.BitConverter.GetBytes(.lEnvirID).CopyTo(yShortcutSave, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lParentStarID).CopyTo(yShortcutSave, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.iParentTypeID).CopyTo(yShortcutSave, lPos) : lPos += 2
                    System.BitConverter.GetBytes(.iModelDefTypeID).CopyTo(yShortcutSave, lPos) : lPos += 2
                    System.BitConverter.GetBytes(.lModelSubTypeID).CopyTo(yShortcutSave, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lUnitCount).CopyTo(yShortcutSave, lPos) : lPos += 4
                    System.Text.ASCIIEncoding.ASCII.GetBytes(.sModelName).CopyTo(yShortcutSave, lPos) : lPos += 20
                    fsFile.Write(yShortcutSave, 0, yShortcutSave.Length)
                End With
            Next X
        Catch
        End Try
        fsFile.Close()
        fsFile.Dispose()
        fsFile = Nothing
    End Sub

    Private mlUniverseCurrentEnvirTypeID As Int16 = -1
    Private mlUniverseCurrentEnvirID As Int32 = -1
    Private mlUniverseIDX As Int32 = 0 'Used to save time in the second environment loop
    Private Sub FillListForUniverse()
        If moUniverse Is Nothing = True OrElse mbUniverseLoaded = False Then Exit Sub
        mlUniverseIDX = 0

        For X As Int32 = 0 To moUniverse.GetUpperBound(0)
            With moUniverse(X)
                If mlUniverseCurrentEnvirTypeID <> .iEnvirTypeID OrElse mlUniverseCurrentEnvirID <> .lEnvirID Then
                    Debug.Print("Fill for " & X.ToString & " (" & .iEnvirTypeID.ToString & "," & .lEnvirID.ToString & ") " & mlUniverseIDX.ToString)
                    Dim oRootNode As UITreeView.UITreeViewItem = tvwMain.GetNodeByItemData2(.lEnvirID, .iEnvirTypeID)
                    If oRootNode Is Nothing = True Then
                        Dim oParentNode As UITreeView.UITreeViewItem = Nothing
                        Dim oBeforeNode As UITreeView.UITreeViewItem = Nothing
                        If .iEnvirTypeID = ObjectType.ePlanet Then
                            oParentNode = tvwMain.GetNodeByItemData2(.lParentStarID, ObjectType.eSolarSystem)
                            If oParentNode Is Nothing = True Then
                                Dim sParentName As String = GetCacheObjectValue(.lParentStarID, ObjectType.eSolarSystem)
                                oBeforeNode = tvwMain.oRootNode
                                Dim bDone As Boolean = False
                                While bDone = False
                                    If oBeforeNode Is Nothing = False AndAlso sParentName > oBeforeNode.sItem Then
                                        oBeforeNode = oBeforeNode.oNextSibling
                                    Else : bDone = True
                                    End If
                                End While
                                oParentNode = tvwMain.AddNode(sParentName, .lParentStarID, ObjectType.eSolarSystem, 0, Nothing, oBeforeNode)
                            End If
                            oBeforeNode = oParentNode.oFirstChild
                        Else : oBeforeNode = tvwMain.oRootNode
                        End If
                        Dim sEnvirName As String = GetCacheObjectValue(.lEnvirID, .iEnvirTypeID)
                        If oBeforeNode Is Nothing = False Then
                            Dim bDone As Boolean = False
                            While bDone = False
                                If oBeforeNode Is Nothing = False Then
                                    If .iEnvirTypeID = ObjectType.ePlanet Then
                                        If .lEnvirID > oBeforeNode.lItemData Then
                                            oBeforeNode = oBeforeNode.oNextSibling
                                        Else : bDone = True
                                        End If
                                    Else
                                        If sEnvirName > oBeforeNode.sItem Then
                                            oBeforeNode = oBeforeNode.oNextSibling
                                        Else : bDone = True
                                        End If
                                    End If
                                Else : bDone = True
                                End If
                            End While
                        End If
                        oRootNode = tvwMain.AddNode(sEnvirName, .lEnvirID, .iEnvirTypeID, 0, oParentNode, oBeforeNode)
                    End If

                    mlUniverseCurrentEnvirTypeID = .iEnvirTypeID
                    mlUniverseCurrentEnvirID = .lEnvirID

                    FillListForUniversePerEnvironment(oRootNode)
                    oRootNode.sItem = oRootNode.sItem & " (" & oRootNode.lItemData3.ToString & ")"
                End If
            End With
        Next
    End Sub

    Private Sub FillListForUniversePerEnvironment(ByRef oRootNode As UITreeView.UITreeViewItem)
        Dim lCnts(10) As Int32
        Dim oRoots(10) As UITreeView.UITreeViewItem
        Dim clrGreen As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        Dim clrYellow As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 0)
        Dim oBeforeNode As UITreeView.UITreeViewItem = oRootNode.oFirstChild
        For X As Int32 = 0 To lCnts.GetUpperBound(0)
            lCnts(X) = 0
        Next X
        If oBeforeNode Is Nothing = False AndAlso oBeforeNode.lItemData = -1 Then
            'Stop
        End If
        oRoots(0) = tvwMain.AddNode("Battlecruisers", -1, 0, 2, oRootNode, oBeforeNode)
        oRoots(1) = tvwMain.AddNode("Battleships", -1, 0, 0, oRootNode, oBeforeNode)
        oRoots(2) = tvwMain.AddNode("Corvettes", -1, 1, 0, oRootNode, oBeforeNode)
        oRoots(3) = tvwMain.AddNode("Cruisers", -1, 1, 1, oRootNode, oBeforeNode)
        oRoots(4) = tvwMain.AddNode("Destroyers", -1, 1, 2, oRootNode, oBeforeNode)
        oRoots(5) = tvwMain.AddNode("Escorts", -1, 1, 4, oRootNode, oBeforeNode)
        oRoots(6) = tvwMain.AddNode("Fighters", -1, 3, -1, oRootNode, oBeforeNode)
        oRoots(7) = tvwMain.AddNode("Frigates", -1, 1, 3, oRootNode, oBeforeNode)
        oRoots(8) = tvwMain.AddNode("Small Vehicles", -1, 4, -1, oRootNode, oBeforeNode)
        oRoots(9) = tvwMain.AddNode("Tanks", -1, 5, -1, oRootNode, oBeforeNode)
        oRoots(10) = tvwMain.AddNode("Utility", -1, 7, -1, oRootNode, oBeforeNode)

        For X As Int32 = mlUniverseIDX To moUniverse.GetUpperBound(0)
            With moUniverse(X)
                If .iEnvirTypeID = mlUniverseCurrentEnvirTypeID AndAlso .lEnvirID = mlUniverseCurrentEnvirID Then
                    Dim oParentNode As UITreeView.UITreeViewItem = Nothing
                    For Y As Int32 = 0 To lCnts.GetUpperBound(0)
                        If oRoots(Y).lItemData2 = .iModelDefTypeID Then
                            'If oRoots(Y).lItemData3 = .lModelSubTypeID OrElse oRoots(Y).lItemData3 = -1 Then
                            oParentNode = oRoots(Y)
                            lCnts(Y) += .lUnitCount
                            oRootNode.lItemData3 += .lUnitCount
                            If oRootNode.oParentNode Is Nothing = False Then
                                oRootNode.oParentNode.lItemData3 += .lUnitCount
                            End If
                            Exit For
                            'End If
                        End If
                    Next Y

                    Dim oNewNode As UITreeView.UITreeViewItem = tvwMain.AddNode(.sModelName & " (" & .lUnitCount.ToString & ")", -1, -1, -1, oParentNode, Nothing)
                    If oNewNode Is Nothing = False Then
                        If .iEnvirTypeID <> .iParentTypeID Then
                            oNewNode.sItem &= "(Docked)"
                            oNewNode.clrItemColor = clrYellow
                        Else : oNewNode.clrItemColor = clrGreen
                        End If
                        oNewNode.bUseItemColor = True
                    End If
                    mlUniverseIDX += 1
                Else 'We are out of units in this environment
                    Exit For
                End If
            End With

        Next X

        For X As Int32 = 0 To lCnts.GetUpperBound(0)
            If lCnts(X) = 0 Then
                tvwMain.RemoveNode(oRoots(X))
            Else
                oRoots(X).sItem = (oRoots(X).sItem & " (" & lCnts(X) & ")")
                oRoots(X).bItemBold = True
            End If
        Next X

    End Sub

    Private Sub tvwMain_NodeSelected(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwMain.NodeSelected
        If oNode Is Nothing = False Then
            Dim lID As Int32 = oNode.lItemData
            Dim iTypeID As Int16 = CShort(oNode.lItemData2)

            If goCamera Is Nothing = False Then goCamera.TrackingIndex = -1
            'MyBase.moUILib.ClearSelection()

            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                Dim lCurUB As Int32 = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))

                oEnvir.DeselectAll()

                If lID < 1 Then Return
                If iTypeID < 1 Then Return

                If iTypeID = ObjectType.eWormhole Then
                    If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso goCurrentEnvir.oGeoObject Is Nothing = False AndAlso CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                        Dim oSystem As SolarSystem = CType(goCurrentEnvir.oGeoObject, SolarSystem)
                        If oSystem.WormholeUB <> -1 Then
                            For X As Int32 = 0 To oSystem.WormholeUB
                                With oSystem.moWormholes(X)
                                    If glCurrentEnvirView <> CurrentView.eSystemView Then glCurrentEnvirView = CurrentView.eSystemView
                                    If .ObjectID = lID Then
                                        If .System1 Is Nothing = False AndAlso .System1.ObjectID = oSystem.ObjectID Then
                                            goCamera.SimplyPlaceCamera(CInt(.LocX1), 0, CInt(.LocY1)) 'goCamera.TrackingIndex = X
                                            Exit For
                                        ElseIf .System2 Is Nothing = False AndAlso .System2.ObjectID = oSystem.ObjectID Then
                                            goCamera.SimplyPlaceCamera(CInt(.LocX2), 0, CInt(.LocY2)) 'goCamera.TrackingIndex = X
                                            Exit For
                                        End If
                                    End If
                                End With
                            Next X
                        End If
                    End If
                Else

                    For X As Int32 = 0 To lCurUB
                        If oEnvir.lEntityIdx(X) = lID Then
                            Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                            If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iTypeID Then
                                If goCurrentEnvir Is Nothing = False Then
                                    If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                        If glCurrentEnvirView <> CurrentView.eSystemView Then glCurrentEnvirView = CurrentView.eSystemView
                                    End If
                                End If

                                oEntity.bSelected = True
                                goCamera.SimplyPlaceCamera(CInt(oEntity.LocX), CInt(oEntity.LocY), CInt(oEntity.LocZ)) 'goCamera.TrackingIndex = X

                                Exit For
                            End If
                        End If
                    Next X
                End If
            End If
        End If
        If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
        tvwMain.HasFocus = False
        Me.HasFocus = False
        goUILib.FocusedControl = Nothing
    End Sub

    Private Sub frmCommand_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.CommandX = Me.Left
            muSettings.CommandY = Me.Top
        End If
    End Sub

    Private Sub tvwMain_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles tvwMain.OnMouseMove
        Dim oNode As UITreeView.UITreeViewItem = tvwMain.GetNodeUnderMouse(lMouseX, lMouseY)
        If oNode Is Nothing = False Then
            If oNode.lItemData2 = ObjectType.eUnit Then
                Dim sTemp As String = ""

                Dim Exp_Level As Byte
                For x As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(x) <> -1 Then
                        If goCurrentEnvir.lEntityIdx(x) = oNode.lItemData Then
                            Exp_Level = goCurrentEnvir.oEntity(x).Exp_Level
                            Exit For
                        End If
                    End If
                Next
                Dim lXPRank As Int32 = Math.Abs((CInt(Exp_Level) - 1) \ 25)
                sTemp = BaseEntity.GetExperienceLevelNameAndBenefits(Exp_Level)
                sTemp &= vbCrLf & "CP Usage: " & Math.Max((10 - lXPRank), 1)
                'Dim lWPUpkeep As Int32 = GetUnitWarpointUpkeep(oNode.lItemData)
                'If lWPUpkeep > -1 Then
                '    sTemp &= vbCrLf & "WP Upkeep Cost: " & lWPUpkeep.ToString("#,##0")
                'End If
                If sTemp <> "" Then goUILib.SetToolTip(sTemp, lMouseX, lMouseY)
            End If
        End If
    End Sub
End Class