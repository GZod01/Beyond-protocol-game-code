'Option Strict On

'Imports Microsoft.DirectX
'Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
'Public Class frmTradeMain
'    Inherits UIWindow

'    Private lstPlayers As UIListBox
'    Private lblPlayers As UILabel
'    Private lblTitle As UILabel
'    Private lnDiv1 As UILine
'    Private btnClose As UIButton
'    Private lblListHeader As UILabel
'    Private btnClose2 As UIButton
'    Private btnNewTrade As UIButton
'    Private lstTrades As UIListBox
'    Private btnView As UIButton

'    Private mlLastUpdate As Int32

'    Public Sub New(ByRef oUILib As UILib)
'        MyBase.New(oUILib)

'        'frmTradeMain initial props
'        With Me
'            .ControlName = "frmTradeMain"
'            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 350
'            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 175
'            .Width = 700
'            .Height = 350
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = False
'            .Moveable = True
'            .BorderLineWidth = 1
'            .mbAcceptReprocessEvents = True
'        End With

'        'lstPlayers initial props
'        lstPlayers = New UIListBox(oUILib)
'        With lstPlayers
'            .ControlName = "lstPlayers"
'            .Left = 5
'            .Top = 50
'            .Width = 166
'            .Height = 265
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(lstPlayers, UIControl))

'        'lblPlayers initial props
'        lblPlayers = New UILabel(oUILib)
'        With lblPlayers
'            .ControlName = "lblPlayers"
'            .Left = 5
'            .Top = 30
'            .Width = 100
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Known Players:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblPlayers, UIControl))

'        'btnNewTrade initial props
'        btnNewTrade = New UIButton(oUILib)
'        With btnNewTrade
'            .ControlName = "btnNewTrade"
'            .Left = 15
'            .Top = 320
'            .Width = 150
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "Propose New Trade"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnNewTrade, UIControl))

'        'lblTitle initial props
'        lblTitle = New UILabel(oUILib)
'        With lblTitle
'            .ControlName = "lblTitle"
'            .Left = 5
'            .Top = 5
'            .Width = 155
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Trade Management"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblTitle, UIControl))

'        'lnDiv1 initial props
'        lnDiv1 = New UILine(oUILib)
'        With lnDiv1
'            .ControlName = "lnDiv1"
'            .Left = 0
'            .Top = 25
'            .Width = 700
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv1, UIControl))

'        'btnClose initial props
'        btnClose = New UIButton(oUILib)
'        With btnClose
'            .ControlName = "btnClose"
'            .Left = 676
'            .Top = 1
'            .Width = 24
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "X"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnClose, UIControl))

'        'lstTrades initial props
'        lstTrades = New UIListBox(oUILib)
'        With lstTrades
'            .ControlName = "lstTrades"
'            .Left = 185
'            .Top = 50
'            .Width = 510
'            .Height = 265
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(lstTrades, UIControl))

'        'btnView initial props
'        btnView = New UIButton(oUILib)
'        With btnView
'            .ControlName = "btnView"
'            .Left = 475
'            .Top = 320
'            .Width = 100
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "View Trade"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnView, UIControl))

'        'lblListHeader initial props
'        lblListHeader = New UILabel(oUILib)
'        With lblListHeader
'            .ControlName = "lblListHeader"
'            .Left = 187
'            .Top = 31
'            .Width = 500
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Player                Status"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblListHeader, UIControl))

'        'btnClose2 initial props
'        btnClose2 = New UIButton(oUILib)
'        With btnClose2
'            .ControlName = "btnClose2"
'            .Left = 595
'            .Top = 320
'            .Width = 100
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "Close"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnClose2, UIControl))

'        FillPlayerList()

'        AddHandler btnClose.Click, AddressOf CloseButtonClick
'        AddHandler btnClose2.Click, AddressOf CloseButtonClick
'        AddHandler btnNewTrade.Click, AddressOf btnNewTrade_Click
'        AddHandler btnView.Click, AddressOf btnView_Click
'        AddHandler lstTrades.ItemDblClick, AddressOf lstTrades_ItemDblClick

'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'        MyBase.moUILib.AddWindow(Me)
'    End Sub

'    Private Sub FillPlayerList()
'        lstPlayers.Clear()

'        'Let's load the player's rels now...
'        Dim oTmpRel As PlayerRel
'        Dim sName As String

'        If goCurrentPlayer Is Nothing = False Then

'            For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
'                oTmpRel = goCurrentPlayer.GetPlayerRelByIndex(X)
'                If oTmpRel Is Nothing = False Then
'                    sName = GetCacheObjectValue(oTmpRel.lThisPlayer, ObjectType.ePlayer)
'                    lstPlayers.AddItem(sName)
'                    lstPlayers.ItemData(lstPlayers.NewIndex) = oTmpRel.lThisPlayer
'                End If
'            Next X

'        End If
'    End Sub

'    Private Sub CloseButtonClick()
'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'    End Sub

'    Private Sub btnNewTrade_Click()
'        If lstPlayers.ListIndex > -1 Then
'            Dim ofrmTrade As New frmTrade(goUILib)
'            ofrmTrade.Visible = True
'            ofrmTrade.CreateNewTrade(lstPlayers.ItemData(lstPlayers.ListIndex))
'            ofrmTrade = Nothing
'        Else
'            goUILib.AddNotification("Please select a player in the list to begin a trade with.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'        End If
'    End Sub

'    Private Sub btnView_Click()
'        If lstTrades.ListIndex > -1 Then
'            Dim lID As Int32 = lstTrades.ItemData(lstTrades.ListIndex)
'            For X As Int32 = 0 To goCurrentPlayer.mlTradeUB
'                If goCurrentPlayer.mlTradeIdx(X) = lID Then
'                    Dim ofrmTrade As New frmTrade(goUILib)
'                    ofrmTrade.Visible = True
'                    ofrmTrade.SetFromTradeObject(goCurrentPlayer.moTrades(X))
'                    ofrmTrade = Nothing

'                    Exit For
'                End If
'            Next X
'        Else
'            goUILib.AddNotification("Please select a trade in the list to view and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'        End If
'    End Sub

'    Private Sub lstTrades_ItemDblClick(ByVal lIndex As Integer)
'        btnView_Click()
'    End Sub

'    Private Sub frmTradeMain_OnNewFrame() Handles Me.OnNewFrame
'        If glCurrentCycle - mlLastUpdate > 15 Then
'            mlLastUpdate = glCurrentCycle

'            For X As Int32 = 0 To lstPlayers.ListCount - 1
'                Dim sValue As String = GetCacheObjectValue(lstPlayers.ItemData(X), ObjectType.ePlayer)
'                If lstPlayers.List(X) <> sValue Then lstPlayers.List(X) = sValue
'            Next X

'            'Now, update what we got in the trade list
'            For X As Int32 = 0 To goCurrentPlayer.mlTradeUB
'                If goCurrentPlayer.mlTradeIdx(X) <> -1 Then
'                    Dim lID As Int32 = goCurrentPlayer.moTrades(X).ObjectID
'                    Dim sTemp As String = goCurrentPlayer.moTrades(X).GetTradeMainListText()
'                    Dim bFound As Boolean = False

'                    For Y As Int32 = 0 To lstTrades.ListCount - 1
'                        If lstTrades.ItemData(Y) = lID Then
'                            If lstTrades.List(Y) <> sTemp Then lstTrades.List(Y) = sTemp
'                            bFound = True
'                            Exit For
'                        End If
'                    Next Y

'                    If bFound = False Then
'                        FillTradeList()
'                        Exit For
'                    End If
'                End If
'            Next X
'        End If
'    End Sub

'    Private Sub FillTradeList()
'        lstTrades.Clear()

'        For X As Int32 = 0 To goCurrentPlayer.mlTradeUB
'            If goCurrentPlayer.mlTradeIdx(X) <> -1 Then
'                lstTrades.AddItem(goCurrentPlayer.moTrades(X).GetTradeMainListText())
'                lstTrades.ItemData(lstTrades.NewIndex) = goCurrentPlayer.moTrades(X).ObjectID
'            End If
'        Next X
'    End Sub
'End Class
