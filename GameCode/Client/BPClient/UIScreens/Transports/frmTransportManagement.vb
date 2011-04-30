Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmTransportManagement
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private WithEvents btnClose As UIButton
    Private lblTransList As UILabel
    Private mfraUnitDetails As fraUnitDetails
    Private WithEvents btnAdd As UIButton
    Private WithEvents btnDecommission As UIButton
    Private WithEvents vscTransList As UIScrollBar

    Private moCurrentTrans As Transport = Nothing

    Private ctlTransItem() As ctlTransport

#Region "  Transport Management  "
    Private Shared moTransports() As Transport
    Private Shared mlTransportUB As Int32 = -1
    Public Shared Sub HandleRequestTransports(ByVal yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lUB As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lUB -= 1
        Dim oTemp(lUB) As Transport
        For X As Int32 = 0 To lUB
            oTemp(X) = New Transport()
            With oTemp(X)
                .TransportID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lSecs As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .TransFlags = yData(lPos) : lPos += 1

                .LocationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .LocationTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                If (.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 AndAlso lSecs > 0 Then
                    .ETA = Now.AddSeconds(lSecs)
                End If
                .ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            End With
        Next X

        mlTransportUB = -1
        moTransports = oTemp
        mlTransportUB = lUB
    End Sub

    Public Shared Function GetTransport(ByVal lID As Int32) As Transport
        If lID = -1 Then Return Nothing
        For X As Int32 = 0 To mlTransportUB
            If moTransports(X) Is Nothing = False AndAlso moTransports(X).TransportID = lID Then Return moTransports(X)
        Next X
        Return Nothing
    End Function

    Public Shared Sub DecommissionTransport(ByVal lID As Int32)
        'upon receipt of the decommission msg - remove the transport from the moTransports list
        For X As Int32 = 0 To mlTransNameUB
            If moTransports(X) Is Nothing = False Then
                If moTransports(X).TransportID = lID Then
                    For Y As Int32 = X To mlTransportUB - 1
                        moTransports(Y) = moTransports(Y + 1)
                    Next Y
                    mlTransportUB -= 1
                    Exit For
                End If
            End If
        Next X
    End Sub

    Public Shared Sub HandleSetTransportStatus(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lTransID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yStatus As Byte = yData(lPos) : lPos += 1

        Dim oTrans As Transport = GetTransport(lTransID)
        If oTrans Is Nothing Then Return

        Select Case yStatus
            Case frmTransportOrders.eyStatusCode.eBeginRoute
                Dim yResult As Byte = yData(lPos) : lPos += 1
                If yResult = 255 Then
                    oTrans.TransFlags = oTrans.TransFlags Or Transport.elTransportFlags.eEnRoute
                    If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) <> 0 Then oTrans.TransFlags = oTrans.TransFlags Xor Transport.elTransportFlags.eIdle

                    oTrans.LocationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    oTrans.LocationTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    oTrans.TransFlags = yData(lPos) : lPos += 1
                    Dim lSecs As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    oTrans.ETA = Now.AddSeconds(lSecs)

                    goUILib.AddNotification("Route started", Color.Teal, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Else
                    goUILib.AddNotification("Unable to start route.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            Case frmTransportOrders.eyStatusCode.eClearRoute
                Dim yResult As Byte = yData(lPos) : lPos += 1
                If yResult = 255 Then
                    oTrans.lRouteUB = -1
                Else
                    goUILib.AddNotification("Unable to clear route.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            Case frmTransportOrders.eyStatusCode.eDeleteItem
                Dim yRoute As Byte = yData(lPos) : lPos += 1
                Dim yAction As Byte = yData(lPos) : lPos += 1

                'remove the route or action
                If yAction = 255 Then
                    'remove action item of route item
                    For X As Int32 = yRoute To oTrans.lRouteUB - 1
                        oTrans.oRoute(X) = oTrans.oRoute(X + 1)
                        oTrans.oRoute(X).OrderNum -= CByte(1)
                    Next X
                    oTrans.lRouteUB -= 1
                    ReDim Preserve oTrans.oRoute(oTrans.lRouteUB)
                Else
                    'remove route item
                    Dim oRoute As TransportRoute = oTrans.oRoute(yRoute)
                    For X As Int32 = yAction To oRoute.lActionUB - 1
                        oRoute.oActions(X) = oRoute.oActions(X + 1)
                        oRoute.oActions(X).ActionOrderNum -= CByte(1)
                    Next X
                    oRoute.lActionUB -= 1
                    ReDim Preserve oRoute.oActions(oRoute.lActionUB)
                    If oTrans.CurrentWaypoint > yRoute Then
                        Dim lNewVal As Int32 = oTrans.CurrentWaypoint
                        lNewVal -= 1
                        If lNewVal < 0 Then lNewVal = 0
                        oTrans.CurrentWaypoint = CByte(lNewVal)
                    End If
                End If
            Case frmTransportOrders.eyStatusCode.eMoveItemDown
                Dim yRoute As Byte = yData(lPos) : lPos += 1
                Dim yAction As Byte = yData(lPos) : lPos += 1

                'move the item down
                If oTrans.lRouteUB < yRoute Then Return
                Dim oRoute As TransportRoute = oTrans.oRoute(yRoute)
                If oRoute Is Nothing = False Then
                    'moving down is increasing this index by 1
                    If yAction = 255 Then
                        'movign the route
                        If yRoute = oTrans.lRouteUB Then Return

                        Dim lToIdx As Int32 = CInt(yRoute) + 1
                        Dim lFromIdx As Int32 = CInt(yRoute)

                        Dim oTo As TransportRoute = oTrans.oRoute(lToIdx)
                        oTrans.oRoute(lToIdx) = oTrans.oRoute(lFromIdx)
                        oTrans.oRoute(lFromIdx) = oTo

                        oTrans.oRoute(lToIdx).OrderNum = CByte(lToIdx)
                        oTrans.oRoute(lFromIdx).OrderNum = CByte(lFromIdx)
                    Else
                        'moving the action
                        If yAction = oRoute.lActionUB Then Return

                        Dim lToIdx As Int32 = CInt(yAction) + 1
                        Dim lFromIdx As Int32 = CInt(yAction)
                        Dim oTo As TransportRouteAction = oRoute.oActions(lToIdx)
                        oRoute.oActions(lToIdx) = oRoute.oActions(lFromIdx)
                        oRoute.oActions(lFromIdx) = oTo

                        oRoute.oActions(lToIdx).ActionOrderNum = CByte(lToIdx)
                        oRoute.oActions(lFromIdx).ActionOrderNum = CByte(lFromIdx)
                    End If
                End If
            Case frmTransportOrders.eyStatusCode.eMoveItemUp
                Dim yRoute As Byte = yData(lPos) : lPos += 1
                Dim yAction As Byte = yData(lPos) : lPos += 1

                'move the item up
                If oTrans.lRouteUB < yRoute Then Return
                Dim oRoute As TransportRoute = oTrans.oRoute(yRoute)
                If oRoute Is Nothing = False Then
                    'moving up is reducing this index by 1
                    If yAction = 255 Then
                        'movign the route
                        If yRoute = 0 Then Return

                        Dim lToIdx As Int32 = CInt(yRoute) - 1
                        Dim lFromIdx As Int32 = CInt(yRoute)

                        Dim oTo As TransportRoute = oTrans.oRoute(lToIdx)
                        oTrans.oRoute(lToIdx) = oTrans.oRoute(lFromIdx)
                        oTrans.oRoute(lFromIdx) = oTo

                        oTrans.oRoute(lToIdx).OrderNum = CByte(lToIdx)
                        oTrans.oRoute(lFromIdx).OrderNum = CByte(lFromIdx)
                    Else
                        'moving the action
                        If yAction = 0 Then Return

                        Dim lToIdx As Int32 = CInt(yAction) - 1
                        Dim lFromIdx As Int32 = CInt(yAction)
                        Dim oTo As TransportRouteAction = oRoute.oActions(lToIdx)
                        oRoute.oActions(lToIdx) = oRoute.oActions(lFromIdx)
                        oRoute.oActions(lFromIdx) = oTo

                        oRoute.oActions(lToIdx).ActionOrderNum = CByte(lToIdx)
                        oRoute.oActions(lFromIdx).ActionOrderNum = CByte(lFromIdx)
                    End If
                End If
            Case frmTransportOrders.eyStatusCode.ePauseRoute
                Dim yIsPaused As Byte = yData(lPos) : lPos += 1
                If yIsPaused = 0 Then
                    If (oTrans.TransFlags And Transport.elTransportFlags.ePaused) <> 0 Then oTrans.TransFlags = oTrans.TransFlags Xor Transport.elTransportFlags.ePaused
                    oTrans.ETA = Date.MinValue
                    oTrans.TransFlags = oTrans.TransFlags Or Transport.elTransportFlags.eEnRoute
                Else
                    oTrans.TransFlags = oTrans.TransFlags Or Transport.elTransportFlags.ePaused
                End If
            Case frmTransportOrders.eyStatusCode.eRecall
                oTrans.TransFlags = yData(lPos) : lPos += 1
                oTrans.LocationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                oTrans.LocationTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                oTrans.TransFlags = oTrans.TransFlags Or Transport.elTransportFlags.eEnRoute
                oTrans.ETA = DateTime.MinValue
            Case frmTransportOrders.eyStatusCode.eLoopOrders
                Dim yLoop As Byte = yData(lPos) : lPos += 1
                If yLoop = 0 Then
                    If (oTrans.TransFlags And Transport.elTransportFlags.eLoop) <> 0 Then oTrans.TransFlags = oTrans.TransFlags Xor Transport.elTransportFlags.eLoop
                Else
                    oTrans.TransFlags = oTrans.TransFlags Or Transport.elTransportFlags.eLoop
                End If
            Case frmTransportOrders.eyStatusCode.eAddDest
                Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim oRoute As New TransportRoute()
                With oRoute
                    .lActionUB = -1
                    ReDim .oActions(-1)
                    .OrderNum = CByte(oTrans.lRouteUB + 1)
                    .oTransport = oTrans
                    .WaypointFlags = 0
                    .DestinationID = lDestID
                    .DestinationTypeID = iDestTypeID
                End With
                ReDim Preserve oTrans.oRoute(oTrans.lRouteUB + 1)
                oTrans.oRoute(oTrans.lRouteUB + 1) = oRoute
                oTrans.lRouteUB += 1
            Case frmTransportOrders.eyStatusCode.eDiscardCargo
                Dim lCargoID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iCargoTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                oTrans.RemoveCargo(lCargoID, iCargoTypeID, lOwnerID)
            Case frmTransportOrders.eyStatusCode.eAddAction
                Dim yRoute As Byte = yData(lPos) : lPos += 1
                Dim yActionTypeID As Byte = yData(lPos) : lPos += 1

                Dim oRoute As TransportRoute = oTrans.oRoute(yRoute)
                Dim oAction As New TransportRouteAction()

                With oAction
                    .ActionOrderNum = CByte(oRoute.lActionUB + 1)
                    .ActionTypeID = CType(yActionTypeID, TransportRouteAction.TransportRouteActionType)
                    .Extended1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Extended2 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .Extended3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .oParentRoute = oRoute
                End With
                ReDim Preserve oRoute.oActions(oRoute.lActionUB + 1)
                oRoute.oActions(oRoute.lActionUB + 1) = oAction
                oRoute.lActionUB += 1
            Case frmTransportOrders.eyStatusCode.eRenameTransport
                'Handled by eRequestTransportName
        End Select
        Dim oTmpWin As frmTransportOrders = CType(goUILib.GetWindow("frmTransportOrders"), frmTransportOrders)
        If oTmpWin Is Nothing = False Then oTmpWin.FillOrders()

        oTrans.lLastSetStatusMsg = glCurrentCycle + 1
    End Sub

    Public Shared Sub HandleAddTransport(ByVal yData() As Byte)

        Dim lPos As Int32 = 2

        Dim yResult As Byte = yData(lPos) : lPos += 1
        Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If yResult = 0 Then
            Dim oTrans As New Transport()
            With oTrans
                .TransportID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .TransFlags = yData(lPos) : lPos += 1
                .LocationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .LocationTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .lCargoUB = -1
                .lRouteUB = -1
            End With

            ReDim Preserve moTransports(mlTransportUB + 1)
            moTransports(mlTransportUB + 1) = oTrans
            mlTransportUB += 1

            frmAddTransport.HandleAddTransport(lUnitID)
        ElseIf yResult = 1 Then
            'too many transports
            goUILib.AddNotification("Unable to create transport: Too many transports (max 100)", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            'could not create transport
            goUILib.AddNotification("Unable to create transport.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Public Shared Sub HandleDecommissionTransport(ByVal yData() As Byte)
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)

        For X As Int32 = 0 To mlTransportUB
            If moTransports(X) Is Nothing = False AndAlso moTransports(X).TransportID = lID Then
                For Y As Int32 = X To mlTransportUB - 1
                    moTransports(Y) = moTransports(Y + 1)
                Next Y
                mlTransportUB -= 1
                Exit For
            End If
        Next X

    End Sub
#End Region

#Region "  Ryan's Idea - store the names on the form  "
    Private Shared mlTransNameID() As Int32
    Private Shared msTransName() As String
    Private Shared mlTransNameUB As Int32 = -1

    Public Shared Function GetTransportName(ByVal lID As Int32) As String
        Try
            For X As Int32 = 0 To mlTransNameUB
                If mlTransNameID(X) = lID Then
                    Return msTransName(X)
                End If
            Next X
            'we're here so add it and request it
            Dim lIdx As Int32 = mlTransNameUB + 1
            ReDim Preserve mlTransNameID(lIdx)
            ReDim Preserve msTransName(lIdx)
            mlTransNameID(lIdx) = lID
            msTransName(lIdx) = "Unknown"
            mlTransNameUB += 1

            'Request the name from the server
            Dim yMsg(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestTransportName).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(lID).CopyTo(yMsg, 2)
            goUILib.SendMsgToPrimary(yMsg)
        Catch
            'just fail out
        End Try
        Return "Unknown"
    End Function
    Public Shared Sub SetTransportName(ByVal yData() As Byte)

        Dim lPos As Int32 = 2
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

        Dim lUB As Int32 = -1
        If mlTransNameID Is Nothing = False Then lUB = Math.Min(mlTransNameUB, mlTransNameID.GetUpperBound(0))
        For X As Int32 = 0 To lUB
            If mlTransNameID(X) = lID Then
                msTransName(X) = sName
                Return
            End If
        Next X
    End Sub
#End Region
    Dim mbLoading As Boolean = True
    Private Shared mlLastPlayerID As Int32 = -1
    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmTransportManagement initial props
        With Me
            .ControlName = "frmTransportManagement"
            .DoWindowInitialPosition(259, 55, 512, 512, muSettings.TransportManagementX, muSettings.TransportManagementY, -1, -1, False)
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 167
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Transport Management"
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

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = Me.BorderLineWidth \ 2
            .Top = 27
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblTransList initial props
        lblTransList = New UILabel(oUILib)
        With lblTransList
            .ControlName = "lblTransList"
            .Left = 5
            .Top = 30
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Transport List"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTransList, UIControl))

        'fraUnitDetails initial props
        mfraUnitDetails = New fraUnitDetails(oUILib)
        With mfraUnitDetails
            .ControlName = "mfraUnitDetails"
            .Left = 255
            .Top = 35
            .Width = 250
            .Height = 470
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
        End With
        Me.AddChild(CType(mfraUnitDetails, UIControl))

        'btnAdd initial props
        btnAdd = New UIButton(oUILib)
        With btnAdd
            .ControlName = "btnAdd"
            .Left = 5
            .Top = 483
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Add Transport"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAdd, UIControl))

        'btnDecommission initial props
        btnDecommission = New UIButton(oUILib)
        With btnDecommission
            .ControlName = "btnDecommission"
            .Left = 130
            .Top = 483
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Decommission"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDecommission, UIControl))

        'vscTransList initial props
        vscTransList = New UIScrollBar(oUILib, True)
        With vscTransList
            .ControlName = "vscTransList"
            .Left = 228
            .Top = 50
            .Width = 24
            .Height = 428
            .Enabled = False
            .Visible = True
            .Value = 0
            .MaxValue = 0
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 1
            .ReverseDirection = True
        End With
        Me.AddChild(CType(vscTransList, UIControl))

        Dim lCnt As Int32 = vscTransList.Height \ 45
        ReDim ctlTransItem(lCnt - 1)
        For X As Int32 = 0 To lCnt - 1
            ctlTransItem(X) = New ctlTransport(goUILib)
            With ctlTransItem(X)
                .ControlName = "ctlTransItem(" & X & ")"
                .Left = 5
                .Top = 50 + (X * 45)
                .Width = 220
                .Height = 45
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
            End With
            Me.AddChild(CType(ctlTransItem(X), UIControl))

            AddHandler ctlTransItem(X).TransportSelected, AddressOf ctlTransItem_TransportSelected
        Next X

        If gbAliased = True Then
            If (glAliasRights And (AliasingRights.eViewBattleGroups Or AliasingRights.eViewColonyStats)) <> (AliasingRights.eViewBattleGroups Or AliasingRights.eViewColonyStats) Then
                goUILib.AddNotification("You lack alias rights to view the Transport Management window.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                goUILib.AddNotification("Request to have View Battlegroups and View Colony Stats rights.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
        End If

        'if we dont have a transport array - request it
        If mlLastPlayerID <> glPlayerID Then mlTransportUB = -1
        mlLastPlayerID = glPlayerID
        If mlTransportUB = -1 Then
            Dim yMsg(1) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestTransports).CopyTo(yMsg, 0)
            goUILib.SendMsgToPrimary(yMsg)
        End If

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
    End Sub

    Private Sub btnAdd_Click(ByVal sName As String) Handles btnAdd.Click

        If gbAliased = True Then
            goUILib.AddNotification("Aliased accounts cannot create transports.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To mlTransportUB
            If moTransports(X) Is Nothing = False Then lCnt += 1
        Next
        If lCnt >= 100 Then
            goUILib.AddNotification("You can only have 100 transports.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim ofrm As New frmAddTransport(goUILib)
        ofrm.Visible = True
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        goUILib.RemoveWindow("frmTransportOrders")
        goUILib.RemoveWindow("frmAddTransport")
    End Sub

    Private Sub btnDecommission_Click(ByVal sName As String) Handles btnDecommission.Click

        If gbAliased = True Then
            goUILib.AddNotification("Aliased accounts cannot decommission transports.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If moCurrentTrans Is Nothing Then
            goUILib.AddNotification("You must select a transport to decommission!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'Check our status
        If (moCurrentTrans.TransFlags And Transport.elTransportFlags.eIdle) = 0 Then
            goUILib.AddNotification("The transport must be idle before it can be decommissioned!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'prompt with msgbox
        Dim oFrm As New frmMsgBox(goUILib, "Decommission of a transport removes the transport permanently. This process cannot be undone. The transport is lost forever." & vbCrLf & vbCrLf & "Are you sure?", MsgBoxStyle.YesNo, "Confirm Decommission")
        AddHandler oFrm.DialogClosed, AddressOf DecommisionMsgBox
    End Sub

    Dim miLastTransportSelected As Int32 = -1
    Private Sub ctlTransItem_TransportSelected(ByVal lID As Int32)
        Dim oTrans As Transport = GetTransport(lID)

        If moCurrentTrans Is Nothing = False AndAlso oTrans Is Nothing = False AndAlso moCurrentTrans.TransportID = oTrans.TransportID AndAlso glCurrentCycle - miLastTransportSelected < 15 Then

            Dim ofrm As New frmTransportOrders(goUILib)
            ofrm.Visible = True
            ofrm.SetTransport(oTrans.TransportID)
        End If
        miLastTransportSelected = glCurrentCycle
        moCurrentTrans = oTrans
        If mfraUnitDetails Is Nothing = False Then
            If oTrans Is Nothing = False Then mfraUnitDetails.SetCurrentTransport(oTrans.TransportID) Else mfraUnitDetails.SetCurrentTransport(-1)
        End If
        If ctlTransItem Is Nothing = False Then
            For X As Int32 = 0 To ctlTransItem.GetUpperBound(0)
                ctlTransItem(X).SetSelectedState(ctlTransItem(X).GetID = lID AndAlso lID <> -1)
            Next X
        End If
    End Sub

    Private Sub vscTransList_ValueChange() Handles vscTransList.ValueChange
        If ctlTransItem Is Nothing = False Then
            Dim lIdx As Int32 = vscTransList.Value
            For X As Int32 = 0 To ctlTransItem.GetUpperBound(0)
                If ctlTransItem(X) Is Nothing = False Then
                    If mlTransportUB < lIdx Then
                        ctlTransItem(X).SetData("", DateTime.MinValue, -1)
                    Else
                        Dim oTrans As Transport = moTransports(lIdx)
                        If oTrans Is Nothing = False Then
                            ctlTransItem(X).SetData(oTrans.GetStatusText(), oTrans.ETA, oTrans.TransportID)
                        End If
                    End If
                End If
            Next X
        End If
    End Sub

    Private Sub frmTransportManagement_OnNewFrame() Handles Me.OnNewFrame
        Dim lScrollCnt As Int32 = 0
        If ctlTransItem Is Nothing = False Then
            lScrollCnt = Math.Max(0, (mlTransportUB + 1) - (ctlTransItem.GetUpperBound(0) + 1))

            Dim lCurrID As Int32 = -1
            If moCurrentTrans Is Nothing = False Then lCurrID = moCurrentTrans.TransportID
            Dim lStartIdx As Int32 = 0
            If lScrollCnt <> 0 Then lStartIdx = vscTransList.Value
            Dim sName As String = "Transport List (" & (mlTransportUB + 1).ToString & "/100)"
            If lblTransList.Caption <> sName Then lblTransList.Caption = sName
            For X As Int32 = 0 To ctlTransItem.GetUpperBound(0)
                Dim lTransID As Int32 = -1
                Dim lIdx As Int32 = (X + lStartIdx)
                If moTransports Is Nothing = False AndAlso lIdx <= mlTransportUB AndAlso lIdx <= moTransports.GetUpperBound(0) Then
                    lTransID = moTransports(lIdx).TransportID
                End If
                If ctlTransItem(X) Is Nothing = False Then
                    If lTransID < 1 Then
                        If ctlTransItem(X).Visible = True Then
                            ctlTransItem(X).SetData("", DateTime.MinValue, -1)
                            ctlTransItem(X).Visible = False
                            Me.IsDirty = True
                        End If
                    ElseIf ctlTransItem(X).GetID <> lTransID Then
                        If lTransID = -1 Then
                            If ctlTransItem(X).Visible = True Then
                                ctlTransItem(X).SetData("", DateTime.MinValue, -1)
                                ctlTransItem(X).Visible = False
                                Me.IsDirty = True
                            End If
                        Else
                            ctlTransItem(X).SetData(moTransports(lIdx).GetStatusText, moTransports(lIdx).ETA, moTransports(lIdx).TransportID)
                            If ctlTransItem(X).Visible = False Then
                                Me.IsDirty = True
                                ctlTransItem(X).Visible = True
                            End If
                        End If
                    End If
                    If lCurrID > 0 Then
                        ctlTransItem(X).SetSelectedState(ctlTransItem(X).GetID = lCurrID)
                    End If
                    ctlTransItem(X).ctlTransport_OnNewFrame()
                End If
            Next X

        End If
        If vscTransList Is Nothing = False AndAlso vscTransList.MaxValue <> lScrollCnt Then
            vscTransList.MaxValue = lScrollCnt
            If vscTransList.Value > vscTransList.MaxValue Then vscTransList.Value = vscTransList.MaxValue
            If lScrollCnt = 0 Then
                If vscTransList.Enabled = True Then vscTransList.Enabled = False
            ElseIf vscTransList.Enabled = False Then
                vscTransList.Enabled = True
            End If
        End If
        If mfraUnitDetails Is Nothing = False Then mfraUnitDetails.fraUnitDetails_OnNewFrame()
    End Sub

    Private Sub DecommisionMsgBox(ByVal lRes As MsgBoxResult)
        'if yes - send msg
        If lRes = MsgBoxResult.Yes AndAlso moCurrentTrans Is Nothing = False Then
            Dim yMsg(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eDecommissionTransport).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(moCurrentTrans.TransportID).CopyTo(yMsg, 2)
            goUILib.SendMsgToPrimary(yMsg)

            goUILib.AddNotification("Decommission Requested...", Color.Teal, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub frmTransportManagement_OnRenderEnd() Handles Me.OnRenderEnd
        If mfraUnitDetails Is Nothing = False Then mfraUnitDetails.fraUnitDetails_OnRenderEnd()
    End Sub

    Private Sub frmTransportManagement_WindowClosed() Handles Me.WindowClosed
        goUILib.RemoveWindow("frmTransportOrders")
        goUILib.RemoveWindow("frmAddTransport")
    End Sub

    Private Sub frmTransportManagement_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.TransportManagementX = Me.Left
            muSettings.TransportManagementY = Me.Top
        End If
    End Sub
End Class