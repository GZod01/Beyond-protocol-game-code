Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class fraDeliverBuyOrder
    Inherits UIWindow

    Private lblSelectCargo As UILabel
    Private lstCargo As UIListBox
    Private WithEvents btnDeliver As UIButton
    Private WithEvents btnCancel As UIButton

    Private mlTradepostID As Int32 = -1
    Private moOrder As BuyOrder = Nothing

    Public Event CancelClicked()
    Public Event DeliverClicked(ByVal lCacheID As Int32, ByVal iCacheTypeID As Int16)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'fraDeliverBuyOrder initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eDeliverBuyOrder
            .ControlName = "fraDeliverBuyOrder"
            .Left = 129
            .Top = 116
            .Width = 255
            .Height = 128
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .BorderLineWidth = 2
        End With

        'lblSelectCargo initial props
        lblSelectCargo = New UILabel(oUILib)
        With lblSelectCargo
            .ControlName = "lblSelectCargo"
            .Left = 5
            .Top = 3
            .Width = 185
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select Cargo To Deliver"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSelectCargo, UIControl))

        'lstCargo initial props
        lstCargo = New UIListBox(oUILib)
        With lstCargo
            .ControlName = "lstCargo"
            .Left = 5
            .Top = 23
            .Width = 245
            .Height = 75
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstCargo, UIControl))

        'btnDeliver initial props
        btnDeliver = New UIButton(oUILib)
        With btnDeliver
            .ControlName = "btnDeliver"
            .Left = 25
            .Top = 102
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Deliver"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDeliver, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = 130
            .Top = 102
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Cancel"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCancel, UIControl))
    End Sub

    Public Sub SetDetails(ByVal lTradepostID As Int32, ByRef oOrder As BuyOrder)
        mlTradepostID = lTradepostID
        moOrder = oOrder

        Dim yMsg(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetTradePostTradeables).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(lTradepostID).CopyTo(yMsg, 2)
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		RaiseEvent CancelClicked()
	End Sub

	Private Sub btnDeliver_Click(ByVal sName As String) Handles btnDeliver.Click
		If HasAliasedRights(AliasingRights.eAlterTrades) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter/place trades.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		If lstCargo.ListIndex <> -1 Then
			RaiseEvent DeliverClicked(lstCargo.ItemData(lstCargo.ListIndex), CShort(lstCargo.ItemData2(lstCargo.ListIndex)))
		End If
	End Sub

    Private Sub fraDeliverBuyOrder_OnNewFrame() Handles Me.OnNewFrame
        If mlTradepostID = -1 OrElse moOrder Is Nothing Then Return

        Dim oTPC As TradePostContents = Nothing
        For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
            If TradePostContents.lTradePostContentsIdx(X) = mlTradepostID Then
                oTPC = TradePostContents.oTradePostContents(X)
                Exit For
            End If
        Next X
        If oTPC Is Nothing = False Then
            Dim lID As Int32 = -1
            Dim iTypeID As Int16 = -1
            If lstCargo.ListIndex <> -1 Then
                lID = lstCargo.ItemData(lstCargo.ListIndex) : iTypeID = CShort(lstCargo.ItemData2(lstCargo.ListIndex))
            End If
            oTPC.FillDeliverList(moOrder.iTradeTypeID, moOrder.blQty, lstCargo)
            For X As Int32 = 0 To lstCargo.ListCount - 1
                If lstCargo.ItemData(X) = lID AndAlso lstCargo.ItemData2(X) = iTypeID Then
                    lstCargo.ListIndex = X
                    Exit For
                End If
            Next X
        End If
    End Sub
End Class