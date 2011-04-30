Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmTransportOrders
    Inherits UIWindow

    Public Enum eyStatusCode As Byte
        eBeginRoute = 1
        eClearRoute = 2
        eDeleteItem = 3
        eMoveItemDown = 4
        eMoveItemUp = 5
        ePauseRoute = 6
        eRecall = 7
        eLoopOrders = 8
        eAddDest = 9
        eDiscardCargo = 10
        eAddAction = 11
        eRenameTransport = 12
    End Enum

    Private mlTransportID As Int32 = -1
    Private mlLastDestRcvd As Int32 = 0
    Private mlLastRouteListCheck As Int32 = -1
    Private mlTransLastSetStatus As Int32 = -1

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblCurrentOrders As UILabel
    Private lblDests As UILabel
    Private lblWaypointAction As UILabel
    Private lblTripDuration As UILabel
    Private tvwDests As UITreeView
    Private mfraWaypointAction As fraWaypointAction
    Private tvwOrders As UITreeView

    Private WithEvents btnClose As UIButton
    Private WithEvents btnAddDest As UIButton
    Private WithEvents btnMoveUp As UIButton
    Private WithEvents btnMoveDown As UIButton
    Private WithEvents chkLoopOrders As UICheckBox
    Private WithEvents btnClear As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnPause As UIButton
    Private WithEvents btnBegin As UIButton
    Private WithEvents btnRecall As UIButton

#Region "  Route Dest Management  "

    Private Structure RouteDest
        Public lColonyID As Int32
        Public lSystemID As Int32
    End Structure
    Private Shared mlLastRouteDestListRequest As Int32 = -1
    Private Shared muRouteDest() As RouteDest
    Private Shared mlRouteDestUB As Int32 = -1
    Private Shared mlLastRouteDestListReceived As Int32 = -1
    Public Shared Sub HandleTransportRouteDestList(ByVal yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lUB As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lUB -= 1             'to make it a UB

        Dim uList(lUB) As RouteDest
        For X As Int32 = 0 To lUB
            With uList(X)
                .lColonyID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSystemID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End With
        Next X

        mlRouteDestUB = -1
        muRouteDest = uList
        mlRouteDestUB = lUB

        mlLastRouteDestListReceived = glCurrentCycle + 1
    End Sub
    Private Shared Sub RequestTransportRouteDestList()
        If glCurrentCycle - mlLastRouteDestListRequest > 90 Then        '3 seconds
            Dim yMsg(1) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestTransportRouteDestList).copyto(yMsg, 0)
            goUILib.SendMsgToPrimary(yMsg)

            mlLastRouteDestListRequest = glCurrentCycle
        End If
    End Sub
#End Region

    Private Sub FillDests()
        tvwDests.Clear()

        Dim lUB As Int32 = -1
        If muRouteDest Is Nothing = False Then lUB = Math.Min(muRouteDest.GetUpperBound(0), mlRouteDestUB)

        Dim lSystemID(-1) As Int32
        Dim sSystemName(-1) As String
        Dim lSysUB As Int32 = -1

        'Get our list of systems...
        For X As Int32 = 0 To lUB
            Dim bFound As Boolean = False
            For Y As Int32 = 0 To lSysUB
                If lSystemID(Y) = muRouteDest(X).lSystemID Then
                    bFound = True
                    Exit For
                End If
            Next Y
            If bFound = False Then
                lSysUB += 1
                ReDim Preserve lSystemID(lSysUB)
                ReDim Preserve sSystemName(lSysUB)
                lSystemID(lSysUB) = muRouteDest(X).lSystemID

                For Y As Int32 = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(Y).ObjectID = muRouteDest(X).lSystemID Then
                        sSystemName(lSysUB) = goGalaxy.moSystems(Y).SystemName
                        Exit For
                    End If
                Next Y
            End If
        Next X

        'Now, sort the list of systems
        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        For X As Int32 = 0 To lSysUB
            Dim lIdx As Int32 = -1
            Dim sName As String = sSystemName(X).ToUpper
            For Y As Int32 = 0 To lSortedUB
                If sSystemName(lSorted(Y)).ToUpper > sName Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedUB += 1
            ReDim Preserve lSorted(lSortedUB)
            If lIdx = -1 Then
                lSorted(lSortedUB) = X
            Else
                For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                    lSorted(Y) = lSorted(Y - 1)
                Next Y
                lSorted(lIdx) = X
            End If
        Next X

        'Now, add our systme nodes
        For X As Int32 = 0 To lSysUB
            tvwDests.AddNode(sSystemName(lSorted(X)), lSystemID(lSorted(X)), ObjectType.eSolarSystem, -1, Nothing, Nothing)
        Next X

        'Now, go through our dests again
        For X As Int32 = 0 To lUB
            With muRouteDest(X)
                Dim oParent As UITreeView.UITreeViewItem = tvwDests.GetNodeByItemData2(.lSystemID, ObjectType.eSolarSystem)
                Dim sText As String = GetCacheObjectValue(.lColonyID, ObjectType.eColony)
                tvwDests.AddNode(sText, .lColonyID, ObjectType.eColony, -1, oParent, Nothing)
            End With
        Next X
    End Sub

    Private mbLoading As Boolean = True
    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmTransportOrders initial props
        With Me
            .ControlName = "frmTransportOrders"
            .DoWindowInitialPosition(280, 85, 511, 511, muSettings.TransportOrdersX, muSettings.TransportOrdersY, -1, -1, False)
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
            .Width = 120
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Transport Orders"
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

        'lblCurrentOrders initial props
        lblCurrentOrders = New UILabel(oUILib)
        With lblCurrentOrders
            .ControlName = "lblCurrentOrders"
            .Left = 5
            .Top = 30
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Current Orders   (max 30)"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCurrentOrders, UIControl))

        'tvwOrders initial props
        tvwOrders = New UITreeView(oUILib)
        With tvwOrders
            .ControlName = "tvwOrders"
            .Left = 5
            .Top = 50
            .Width = 260
            .Height = 266
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(tvwOrders, UIControl))

        'lblDests initial props
        lblDests = New UILabel(oUILib)
        With lblDests
            .ControlName = "lblDests"
            .Left = 280
            .Top = 30
            .Width = 80
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Destinations"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDests, UIControl))

        'tvwDests initial props
        tvwDests = New UITreeView(oUILib)
        With tvwDests
            .ControlName = "tvwDests"
            .Left = 280
            .Top = 50
            .Width = 225
            .Height = 177
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(tvwDests, UIControl))

        'btnAddDest initial props
        btnAddDest = New UIButton(oUILib)
        With btnAddDest
            .ControlName = "btnAddDest"
            .Left = 343
            .Top = 233
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Add"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAddDest, UIControl))

        'lblWaypointAction initial props
        lblWaypointAction = New UILabel(oUILib)
        With lblWaypointAction
            .ControlName = "lblWaypointAction"
            .Left = 280
            .Top = 264
            .Width = 180
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Waypoint Action   (max 10)"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblWaypointAction, UIControl))

        'fraWaypointAction initial props
        mfraWaypointAction = New fraWaypointAction(oUILib)
        With mfraWaypointAction
            .ControlName = "mfraWaypointAction"
            .Left = 280
            .Top = 285
            .Width = 225
            .Height = 220
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
        End With
        Me.AddChild(CType(mfraWaypointAction, UIControl))

        'btnMoveUp initial props
        btnMoveUp = New UIButton(oUILib)
        With btnMoveUp
            .ControlName = "btnMoveUp"
            .Left = 30
            .Top = 325
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Move Up"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnMoveUp, UIControl))

        'btnMoveDown initial props
        btnMoveDown = New UIButton(oUILib)
        With btnMoveDown
            .ControlName = "btnMoveDown"
            .Left = 140
            .Top = 325
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Move Down"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnMoveDown, UIControl))

        'chkLoopOrders initial props
        chkLoopOrders = New UICheckBox(oUILib)
        With chkLoopOrders
            .ControlName = "chkLoopOrders"
            .Left = 90
            .Top = 460
            .Width = 95
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Loop Orders"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkLoopOrders, UIControl))

        'lblTripDuration initial props
        lblTripDuration = New UILabel(oUILib)
        With lblTripDuration
            .ControlName = "lblTripDuration"
            .Left = 5
            .Top = 487
            .Width = 260
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "One-Way Trip Duration: DD:HH:MM:SS"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTripDuration, UIControl))

        'btnClear initial props
        btnClear = New UIButton(oUILib)
        With btnClear
            .ControlName = "btnClear"
            .Left = 30
            .Top = 360
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Clear"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClear, UIControl))

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Left = 140
            .Top = 360
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Delete Item"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDelete, UIControl))

        'btnPause initial props
        btnPause = New UIButton(oUILib)
        With btnPause
            .ControlName = "btnPause"
            .Left = 30
            .Top = 395
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Pause"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnPause, UIControl))

        'btnBegin initial props
        btnBegin = New UIButton(oUILib)
        With btnBegin
            .ControlName = "btnBegin"
            .Left = 140
            .Top = 395
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Begin"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnBegin, UIControl))

        'btnRecall initial props
        btnRecall = New UIButton(oUILib)
        With btnRecall
            .ControlName = "btnRecall"
            .Left = 86
            .Top = 430
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Recall"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRecall, UIControl))

        RequestTransportRouteDestList()

        AddHandler mfraWaypointAction.AddAction, AddressOf WPOrder_AddAction

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Public Sub SetTransport(ByVal lID As Int32)
        mlTransportID = lID
    End Sub

    Private Sub btnAddDest_Click(ByVal sName As String) Handles btnAddDest.Click
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then
            Dim oDest As UITreeView.UITreeViewItem = tvwDests.oSelectedNode
            If oDest Is Nothing Then
                goUILib.AddNotification("Select a destination to add.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            If oTrans.lRouteUB > 28 Then
                goUILib.AddNotification("Unable to assign more than 30 destinations to a transport.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            Dim yMsg(12) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.eAddDest : lPos += 1
            System.BitConverter.GetBytes(oDest.lItemData).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CShort(oDest.lItemData2)).CopyTo(yMsg, lPos) : lPos += 2

            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Private Sub btnBegin_Click(ByVal sName As String) Handles btnBegin.Click
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then
            If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) = 0 Then
                goUILib.AddNotification("The transport must be Idle in order to Begin a route.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            Dim yMsg(6) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).copyto(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.eBeginRoute : lPos += 1

            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sName As String) Handles btnClear.Click
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then
            If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) = 0 Then
                goUILib.AddNotification("The transport must be Idle in order to clear the route list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            Dim yMsg(6) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.eClearRoute : lPos += 1

            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then

            Dim oNode As UITreeView.UITreeViewItem = tvwOrders.oSelectedNode
            If oNode Is Nothing Then Return

            Dim oRoute As TransportRoute = Nothing
            Dim oAction As TransportRouteAction = Nothing

            If oNode.oParentNode Is Nothing Then
                'Deleting a Route Item
                oRoute = CType(oNode.oRelatedObject, TransportRoute)
                oAction = Nothing
            Else
                'Deleting a Route Action Item
                oRoute = CType(oNode.oParentNode.oRelatedObject, TransportRoute)
                oAction = CType(oNode.oRelatedObject, TransportRouteAction)
            End If
            If oRoute Is Nothing Then Return

            If (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                'Ok, check if the transport is currently on the way to that item
                If oRoute.OrderNum = oTrans.CurrentWaypoint Then
                    goUILib.AddNotification("Cannot alter orders of the current waypoint.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
            End If

            Dim yMsg(8) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.eDeleteItem : lPos += 1

            yMsg(lPos) = oRoute.OrderNum : lPos += 1
            If oAction Is Nothing = False Then yMsg(lPos) = oAction.ActionOrderNum Else yMsg(lPos) = 255
            lPos += 1

            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Private Sub btnMoveDown_Click(ByVal sName As String) Handles btnMoveDown.Click
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then

            Dim oNode As UITreeView.UITreeViewItem = tvwOrders.oSelectedNode
            If oNode Is Nothing Then Return

            Dim oRoute As TransportRoute = Nothing
            Dim oAction As TransportRouteAction = Nothing

            If oNode.oParentNode Is Nothing Then
                'Deleting a Route Item
                oRoute = CType(oNode.oRelatedObject, TransportRoute)
                oAction = Nothing
            Else
                'Deleting a Route Action Item
                oRoute = CType(oNode.oParentNode.oRelatedObject, TransportRoute)
                oAction = CType(oNode.oRelatedObject, TransportRouteAction)
            End If
            If oRoute Is Nothing Then Return

            If (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                'Ok, check if the transport is currently on the way to that item
                If oRoute.OrderNum = oTrans.CurrentWaypoint Then
                    goUILib.AddNotification("Cannot alter orders of the current waypoint.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If

                If oAction Is Nothing Then
                    'move down so check if order is next
                    Dim lNext As Int32 = oTrans.CurrentWaypoint
                    lNext += 1
                    If oRoute.OrderNum = lNext Then
                        goUILib.AddNotification("Cannot alter orders of the current waypoint.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If

                    If oRoute.OrderNum = oTrans.lRouteUB Then Return
                End If
            End If

            Dim yMsg(8) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.eMoveItemDown : lPos += 1

            yMsg(lPos) = oRoute.OrderNum : lPos += 1
            If oAction Is Nothing = False Then yMsg(lPos) = oAction.ActionOrderNum Else yMsg(lPos) = 255
            lPos += 1

            goUILib.SendMsgToPrimary(yMsg)

        End If
    End Sub

    Private Sub btnMoveUp_Click(ByVal sName As String) Handles btnMoveUp.Click
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then

            Dim oNode As UITreeView.UITreeViewItem = tvwOrders.oSelectedNode
            If oNode Is Nothing Then Return

            Dim oRoute As TransportRoute = Nothing
            Dim oAction As TransportRouteAction = Nothing

            If oNode.oParentNode Is Nothing Then
                'Deleting a Route Item
                oRoute = CType(oNode.oRelatedObject, TransportRoute)
                oAction = Nothing
            Else
                'Deleting a Route Action Item
                oRoute = CType(oNode.oParentNode.oRelatedObject, TransportRoute)
                oAction = CType(oNode.oRelatedObject, TransportRouteAction)
            End If
            If oRoute Is Nothing Then Return

            If (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                'Ok, check if the transport is currently on the way to that item
                If oRoute.OrderNum = oTrans.CurrentWaypoint Then
                    goUILib.AddNotification("Cannot alter orders of the current waypoint.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If

                If oAction Is Nothing Then
                    'move down so check if order is next
                    Dim lNext As Int32 = oTrans.CurrentWaypoint
                    lNext -= 1
                    If oRoute.OrderNum = lNext Then
                        goUILib.AddNotification("Cannot alter orders of the current waypoint.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If

                    If oRoute.OrderNum = oTrans.lRouteUB Then Return
                End If
            End If

            Dim yMsg(8) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.eMoveItemUp : lPos += 1

            yMsg(lPos) = oRoute.OrderNum : lPos += 1
            If oAction Is Nothing = False Then yMsg(lPos) = oAction.ActionOrderNum Else yMsg(lPos) = 255
            lPos += 1

            goUILib.SendMsgToPrimary(yMsg)

        End If
    End Sub

    Private Sub btnPause_Click(ByVal sName As String) Handles btnPause.Click
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then
            Dim yMsg(6) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.ePauseRoute : lPos += 1

            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Private Sub btnRecall_Click(ByVal sName As String) Handles btnRecall.Click
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then
            Dim yMsg(6) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.eRecall : lPos += 1

            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Private Sub chkLoopOrders_Click() Handles chkLoopOrders.Click
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then
            If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) = 0 Then
                goUILib.AddNotification("The transport must be Idle in order to Begin a route.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            Dim yMsg(6) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.eLoopOrders : lPos += 1

            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Private Sub frmTransportOrders_OnNewFrame() Handles Me.OnNewFrame
        If mlLastDestRcvd <> mlLastRouteDestListReceived Then
            mlLastDestRcvd = mlLastRouteDestListReceived
            FillDests()
        Else
            Dim lUB As Int32 = -1
            If muRouteDest Is Nothing = False Then lUB = Math.Min(mlRouteDestUB, muRouteDest.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                With muRouteDest(X)
                    Dim oNode As UITreeView.UITreeViewItem = tvwDests.GetNodeByItemData2(.lColonyID, ObjectType.eColony)
                    If oNode Is Nothing = False Then
                        Dim sText As String = GetCacheObjectValue(.lColonyID, ObjectType.eColony)
                        If oNode.sItem <> sText Then
                            oNode.sItem = sText
                            tvwDests.IsDirty = True
                        End If
                    End If
                End With
            Next X
        End If

        'Now double-check our orders...
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then

            If (oTrans.TransFlags And Transport.elTransportFlags.ePaused) <> 0 Then
                If btnPause.Caption <> "Unpause" Then btnPause.Caption = "Unpause"
            ElseIf btnPause.Caption <> "Pause" Then
                btnPause.Caption = "Pause"
            End If

            If glCurrentCycle - mlLastRouteListCheck > 90 OrElse oTrans.lLastSetStatusMsg <> mlTransLastSetStatus Then
                mlTransLastSetStatus = oTrans.lLastSetStatusMsg
                mlLastRouteListCheck = glCurrentCycle

                If chkLoopOrders.Value <> ((oTrans.TransFlags And Transport.elTransportFlags.eLoop) <> 0) Then
                    chkLoopOrders.Value = ((oTrans.TransFlags And Transport.elTransportFlags.eLoop) <> 0)
                End If
                'Ok, first check route content vs. list
                For X As Int32 = 0 To oTrans.lRouteUB
                    Dim oItem As TransportRoute = oTrans.oRoute(X)
                    If oItem Is Nothing = False Then
                        Dim oDestNode As UITreeView.UITreeViewItem = tvwOrders.GetNodeByItemData3(oItem.DestinationID, oItem.DestinationTypeID, oItem.OrderNum)
                        If oDestNode Is Nothing Then
                            FillOrders()
                            Return
                        Else
                            'Ok, the dest node exists...
                            Dim sText As String = GetCacheObjectValue(oItem.DestinationID, oItem.DestinationTypeID)
                            If oDestNode.sItem <> sText Then
                                oDestNode.sItem = sText
                                tvwOrders.IsDirty = True
                            End If
                            'Now, go through the actions
                            Dim oActNode As UITreeView.UITreeViewItem = oDestNode.oFirstChild
                            Dim lIdx As Int32 = 0
                            While oActNode Is Nothing = False
                                If oItem.lActionUB < lIdx Then
                                    FillOrders()
                                    Return
                                End If
                                'is this the same item?
                                If oItem.oActions(lIdx).ActionOrderNum <> oActNode.lItemData OrElse oItem.oActions(lIdx).Extended1 <> oActNode.lItemData2 OrElse oItem.oActions(lIdx).Extended2 <> oActNode.lItemData3 Then
                                    FillOrders()
                                    Return
                                End If
                                'Ok, check the text
                                sText = oItem.oActions(lIdx).GetDisplay()
                                If oActNode.sItem <> sText Then
                                    oActNode.sItem = sText
                                    tvwOrders.IsDirty = True
                                End If

                                lIdx += 1
                                oActNode = oActNode.oNextSibling
                            End While
                            If lIdx = oItem.lActionUB Then
                                FillOrders()
                                Return
                            End If
                        End If
                    End If
                Next X

                If oTrans.lRouteUB = -1 AndAlso tvwOrders.oRootNode Is Nothing = False Then
                    tvwOrders.Clear()
                    Me.IsDirty = True
                End If

                ''Now, check list content vs. Route
                'Dim oCurNode As UITreeView.UITreeViewItem = tvwOrders.oRootNode
                'While oCurNode Is Nothing = False
                '    If oCurNode.oParentNode Is Nothing = False Then
                '        'ordernum, ext1, ext2 for Action - if parentnode = nothing = false

                '    Else
                '        'destid, desttypeid, ordernum for route - if parentnode = nothing
                '    End If

                '    oCurNode = UITreeView.FindNextParentSibling(oCurNode)
                'End While

            End If
        ElseIf tvwOrders.TotalNodeCount <> 0 Then
            tvwOrders.Clear()
        End If
       
    End Sub

    Public Sub FillOrders()
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)

        tvwOrders.Clear()

        If oTrans Is Nothing = False Then
            For X As Int32 = 0 To oTrans.lRouteUB
                Dim oItem As TransportRoute = oTrans.oRoute(X)
                If oItem Is Nothing = False Then
                    With oItem
                        Dim sDest As String = GetCacheObjectValue(.DestinationID, .DestinationTypeID)
                        Dim oDestNode As UITreeView.UITreeViewItem = tvwOrders.AddNode(sDest, .DestinationID, .DestinationTypeID, X, Nothing, Nothing)
                        oDestNode.oRelatedObject = oItem

                        Dim bHasChild As Boolean = False

                        For Y As Int32 = 0 To .lActionUB
                            If .oActions(Y) Is Nothing = False Then
                                Dim oTemp As UITreeView.UITreeViewItem = tvwOrders.AddNode(.oActions(Y).GetDisplay, .oActions(Y).ActionOrderNum, .oActions(Y).Extended1, .oActions(Y).Extended2, oDestNode, Nothing)
                                oTemp.oRelatedObject = .oActions(Y)
                                bHasChild = True
                            End If
                        Next Y
                        If bHasChild = True Then oDestNode.bExpanded = True
                    End With
                End If
            Next X
        End If
    End Sub

    Private Sub WPOrder_AddAction(ByVal yActionTypeID As Byte, ByVal lExt1 As Int32, ByVal iExt2 As Int16, ByVal lExt3 As Int32)
        Dim oTrans As Transport = frmTransportManagement.GetTransport(mlTransportID)
        If oTrans Is Nothing = False Then

            Dim oNode As UITreeView.UITreeViewItem = tvwOrders.oSelectedNode
            If oNode Is Nothing Then
                MyBase.moUILib.AddNotification("Select a waypoint to add the action to.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            Dim oRoute As TransportRoute = Nothing

            If oNode.oParentNode Is Nothing Then
                'Deleting a Route Item
                oRoute = CType(oNode.oRelatedObject, TransportRoute)
            Else
                'Deleting a Route Action Item
                oRoute = CType(oNode.oParentNode.oRelatedObject, TransportRoute)
            End If
            If oRoute Is Nothing Then Return
            If oRoute.OrderNum = oTrans.CurrentWaypoint AndAlso (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                MyBase.moUILib.AddNotification("Unable to alter current actions for the current waypoint.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            If oRoute.lActionUB > 8 Then
                MyBase.moUILib.AddNotification("Unable to add more than 10 actions to a waypoint.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            Dim yMsg(18) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(mlTransportID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyStatusCode.eAddAction : lPos += 1
            yMsg(lPos) = oRoute.OrderNum : lPos += 1
            yMsg(lPos) = yActionTypeID : lPos += 1
            System.BitConverter.GetBytes(lExt1).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iExt2).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lExt3).CopyTo(yMsg, lPos) : lPos += 4

            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Private Sub frmTransportOrders_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.TransportOrdersX = Me.Left
            muSettings.TransportOrdersY = Me.Top
        End If
    End Sub
End Class