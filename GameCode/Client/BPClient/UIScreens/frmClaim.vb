Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmClaim
    Inherits UIWindow

    Private lstItems As UIListBox
    Private WithEvents btnViewDetails As UIButton
    Private WithEvents btnClaim As UIButton
    Private WithEvents btnClose As UIButton
    Private lnDiv1 As UILine
    Private lblTitle As UILabel

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmClaim initial props
        With Me
            .ControlName = "frmClaim"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            .Width = 511
            .Height = 255
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
        End With

        'lstItems initial props
        lstItems = New UIListBox(oUILib)
        With lstItems
            .ControlName = "lstItems"
            .Left = 5
            .Top = 30
            .Width = 500
            .Height = 180
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstItems, UIControl))

        'btnViewDetails initial props
        btnViewDetails = New UIButton(oUILib)
        With btnViewDetails
            .ControlName = "btnViewDetails"
            .Left = 10
            .Top = 223
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "View Details"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnViewDetails, UIControl))

        'btnClaim initial props
        btnClaim = New UIButton(oUILib)
        With btnClaim
            .ControlName = "btnClaim"
            .Left = Me.Width - 110
            .Top = 223
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Claim"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClaim, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 21 - Me.BorderLineWidth
            .Top = 3
            .Width = 22
            .Height = 22
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
            .Left = 2
            .Top = 25
            .Width = Me.Width - 4
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 0
            .Width = 105
            .Height = 25
            .Enabled = True
            .Visible = True
            .Caption = "Claim Rewards"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        'ReDim guClaimables(2)
        'With guClaimables(0)
        '    .lID = 87
        '    .iTypeID = ObjectType.eHullTech
        '    .sName = "Reward Hull 1"
        '    .yClaimFlag = 0
        '    .lOfferCode = 1
        'End With
        'With guClaimables(1)
        '    .lID = 89
        '    .iTypeID = ObjectType.eHullTech
        '    .sName = "Reward Hull 2"
        '    .yClaimFlag = 0
        '    .lOfferCode = 1
        'End With
        'With guClaimables(2)
        '    .lID = 82
        '    .iTypeID = ObjectType.eHullTech
        '    .sName = "Reward Hull 3"
        '    .yClaimFlag = 0
        '    .lOfferCode = 1
        'End With
        'RefreshList()

        Dim yMsg(11) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eClaimItem).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(-1S).CopyTo(yMsg, 6)
        System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 8)
        MyBase.moUILib.SendMsgToPrimary(yMsg)


    End Sub

    Private Sub RefreshList()
        lstItems.Clear()

        If guClaimables Is Nothing = False Then
            For X As Int32 = 0 To guClaimables.GetUpperBound(0)
                With guClaimables(X)
                    Dim sName As String = .sName
                    
                    If (.yClaimFlag And eyClaimFlags.eClaimed) <> 0 Then
                        sName = sName.PadRight(32, " "c) & "CLAIMED"
                    End If
                    lstItems.AddItem(sName, True)
                    lstItems.ItemData(lstItems.NewIndex) = .lID
                    lstItems.ItemData2(lstItems.NewIndex) = .iTypeID
                    lstItems.ItemData3(lstItems.NewIndex) = .lOfferCode
                End With
            Next X
        End If
    End Sub

    Private Sub btnClaim_Click(ByVal sName As String) Handles btnClaim.Click
        If lstItems.ListIndex > -1 Then
            Dim lID As Int32 = lstItems.ItemData(lstItems.ListIndex)
            Dim iTypeID As Int16 = CShort(lstItems.ItemData2(lstItems.ListIndex))
            Dim lOfferCode As Int32 = lstItems.ItemData3(lstItems.ListIndex)

            Dim sGiveup As String = ""

            If guClaimables Is Nothing = False Then
                Dim sHeader As String = ""
                For X As Int32 = 0 To guClaimables.GetUpperBound(0)
                    If guClaimables(X).lOfferCode = lOfferCode AndAlso guClaimables(X).lID = lID AndAlso guClaimables(X).iTypeID = iTypeID Then
                        sHeader = "Accepting " & guClaimables(X).sName & " as your reward will forfeit access to these rewards:"
                        Exit For
                    End If
                Next X
                For X As Int32 = 0 To guClaimables.GetUpperBound(0)
                    If guClaimables(X).lOfferCode = lOfferCode Then
                        If guClaimables(X).lID <> lID OrElse guClaimables(X).iTypeID <> iTypeID Then
                            If sGiveup = "" Then
                                sGiveup = sHeader
                            End If
                            sGiveup &= vbCrLf & "  " & guClaimables(X).sName
                        End If
                    End If
                Next X
            End If

            If sGiveup <> "" Then
                sGiveup &= vbCrLf & vbCrLf & "Are you sure you want this reward?"
                Dim oMsgBox As New frmMsgBox(MyBase.moUILib, sGiveup, MsgBoxStyle.YesNo, "Confirm Reward")
                AddHandler oMsgBox.DialogClosed, AddressOf MsgBoxRes
            Else
                SubmitClaim()
            End If
        Else
            MyBase.moUILib.AddNotification("You must select an item in the list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub SubmitClaim()
        If lstItems.ListIndex > -1 Then
            Dim lID As Int32 = lstItems.ItemData(lstItems.ListIndex)
            Dim iTypeID As Int16 = CShort(lstItems.ItemData2(lstItems.ListIndex))
            Dim lOfferCode As Int32 = lstItems.ItemData3(lstItems.ListIndex)

            Dim lExtID As Int32 = -1

            If iTypeID = ObjectType.eSpecialTech Then
                Try
                    If goCurrentEnvir Is Nothing = False Then
                        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                            If goCurrentEnvir.lEntityIdx(X) > 0 Then
                                Dim oEnt As BaseEntity = goCurrentEnvir.oEntity(X)
                                If oEnt Is Nothing = False Then
                                    If oEnt.bSelected = True AndAlso oEnt.OwnerID = glPlayerID Then
                                        If oEnt.ObjTypeID = ObjectType.eFacility Then
                                            If oEnt.yProductionType = ProductionType.eResearch Then
                                                lExtID = oEnt.ObjectID
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next X
                    End If
                Catch
                    Return
                End Try

                If lExtID = -1 Then
                    Dim oMsgBox As New frmMsgBox(goUILib, "You must select a research facility that is currently researching a project to apply this claimable item to.", MsgBoxStyle.OkOnly, "Select a Research Lab")
                    Return
                End If
            End If

            Dim yMsg(15) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eClaimItem).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(lID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, 6)
            System.BitConverter.GetBytes(lOfferCode).CopyTo(yMsg, 8)
            System.BitConverter.GetBytes(lExtID).CopyTo(yMsg, 12)
            MyBase.moUILib.SendMsgToPrimary(yMsg)

        End If
    End Sub

    Private Sub MsgBoxRes(ByVal lResult As MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            SubmitClaim()
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnViewDetails_Click(ByVal sName As String) Handles btnViewDetails.Click
        If lstItems.ListIndex > -1 Then
            Dim lID As Int32 = lstItems.ItemData(lstItems.ListIndex)
            Dim iTypeID As Int16 = CShort(lstItems.ItemData2(lstItems.ListIndex))

            Select Case iTypeID
                Case ObjectType.eHullTech
                    'lID is the model ID - now, we show the hull builder with the model
                    Dim ofrmHull As frmHullBuilder = CType(MyBase.moUILib.GetWindow("frmHullBuilder"), frmHullBuilder)
                    If ofrmHull Is Nothing Then ofrmHull = New frmHullBuilder(MyBase.moUILib)
                    ofrmHull.SetFromClaimableModel(lID)
                    ofrmHull = Nothing
                Case ObjectType.eSpecialTech
                    Dim oMsgBox As New frmMsgBox(goUILib, "This applies the amount of research production, to the selected research facility's current research project.  You must select a facility in order to be able to claim this item.", MsgBoxStyle.OkOnly, "Research Production")
            End Select
        Else
            MyBase.moUILib.AddNotification("You must select an item in the list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub frmClaim_OnNewFrame() Handles Me.OnNewFrame
        Try
            If guClaimables Is Nothing = False Then
                If guClaimables.Length <> lstItems.ListCount Then RefreshList()
            Else
                If lstItems.ListCount <> 0 Then lstItems.Clear()
            End If
        Catch
            lstItems.Clear()
        End Try
    End Sub
End Class