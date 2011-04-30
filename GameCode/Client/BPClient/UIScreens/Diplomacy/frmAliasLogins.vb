Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAliasLogins
    Inherits UIWindow

    Private WithEvents lstList As UIListBox
    Private WithEvents btnView As UIButton
    Private WithEvents btnLogin As UIButton
    Private WithEvents btnClose As UIButton
    Private lblTitle As UILabel

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmAliasLogins initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eMyAliasesWindow
            .ControlName = "frmAliasLogins"
            .Left = 267
            .Top = 152
            .Width = 255
            .Height = 512 '165
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
        End With

        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmAliasing")
        If ofrm Is Nothing = False Then
            Me.Height = ofrm.Height
            Me.Left = ofrm.Left - Me.Width
            Me.Top = ofrm.Top
        Else
            Dim oofrm As UIWindow = MyBase.moUILib.GetWindow("frmOptions")
            If oofrm Is Nothing = False Then
                Me.Height = oofrm.Height
                Me.Left = oofrm.Left - Me.Width
                Me.Top = oofrm.Top
            End If
            oofrm = Nothing
        End If


        'lstList initial props
        lstList = New UIListBox(oUILib)
        With lstList
            .ControlName = "lstList"
            .Left = 5
            .Top = 25
            .Width = 245
            .Height = Me.Height - .Top - 30 '100
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstList, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 26
            .Top = 2
            .Width = 24
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

        'btnView initial props
        btnView = New UIButton(oUILib)
        With btnView
            .ControlName = "btnView"
            .Width = 100
            .Height = 22
            .Left = 5
            .Top = Me.Height - .Height - 5 ' 130
            .Enabled = gbAliased = False
            .Visible = True
            '.Visible = gbAliased = False
            .Caption = "View Rights"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnView, UIControl))

        'btnLogin initial props
        btnLogin = New UIButton(oUILib)
        With btnLogin
            .ControlName = "btnLogin"
            .Width = 100
            .Height = 22
            .Left = 150
            .Top = Me.Height - .Height - 5 '130
            .Enabled = True
            .Visible = True
            .Caption = "Login"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnLogin, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 94
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Your Aliases"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

    End Sub

    Private Sub btnLogin_Click(ByVal sName As String) Handles btnLogin.Click
        If lstList.ListIndex > -1 Then
            Dim lIdx As Int32 = lstList.ItemData(lstList.ListIndex)
            If lIdx = Int32.MinValue Then
                'log in as myself
                glCurrentEnvirView = CurrentView.eStartupLogin
                MyBase.moUILib.RemoveAllWindows()
                MyBase.moUILib.GetMsgSys.DisconnectAll()
                glPlayerIntelUB = -1
                glPlayerTechKnowledgeUB = -1
                glItemIntelUB = -1
                glEntityDefUB = -1
                glMineralPropertyUB = -1
                glMineralUB = -1


                goCurrentPlayer = Nothing
                goCurrentEnvir = Nothing
                goControlGroups = Nothing
                goMinerals = Nothing
                goMineralProperty = Nothing
                goAvailableResources = Nothing

                gyTriggerFired = Nothing

                Application.DoEvents()
                frmMain.bForceRequestDetails = True
                frmMain.mlLastMineralPropRequest = glCurrentCycle + 300
                frmMain.mbDoMineralPropRequests = True

                gfrmMain.Login_StartLogin(gsUserName, gsPassword, False, "", "", Nothing)
            Else
                If lIdx > goCurrentPlayer.mlAllowanceUB OrElse lIdx < 0 Then Return
                Dim uItem As PlayerAlias = goCurrentPlayer.muAllowances(lIdx)

                glCurrentEnvirView = CurrentView.eStartupLogin
                MyBase.moUILib.RemoveAllWindows()
                MyBase.moUILib.GetMsgSys.DisconnectAll()
                glPlayerIntelUB = -1
                glPlayerTechKnowledgeUB = -1
                glItemIntelUB = -1
                glEntityDefUB = -1
                glMineralPropertyUB = -1
                glMineralUB = -1

                goCurrentPlayer = Nothing
                goCurrentEnvir = Nothing
                goControlGroups = Nothing
                goMinerals = Nothing
                goMineralProperty = Nothing
                goAvailableResources = Nothing

                gyTriggerFired = Nothing

                Application.DoEvents()
                frmMain.bForceRequestDetails = True
                frmMain.mlLastMineralPropRequest = glCurrentCycle + 300
                frmMain.mbDoMineralPropRequests = True
                frmMain.Login_StartLogin(gsUserName, gsPassword, True, uItem.sUserName, uItem.sPassword, Nothing)
            End If
            HACK_SetCacheEntries()


        Else
            MyBase.moUILib.AddNotification("Select an item in the list first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub btnView_Click(ByVal sName As String) Handles btnView.Click
        If lstList.ListIndex > -1 Then
            Dim lIdx As Int32 = lstList.ItemData(lstList.ListIndex)
            If lIdx > goCurrentPlayer.mlAllowanceUB OrElse lIdx < 0 Then Return
            Dim uItem As PlayerAlias = goCurrentPlayer.muAllowances(lIdx)
            Dim oFrm As frmAliasing = CType(goUILib.GetWindow("frmAliasing"), frmAliasing)
            If oFrm Is Nothing = False Then
                oFrm.SetFromItem(uItem, True)
            End If
            oFrm = Nothing
        Else
            MyBase.moUILib.AddNotification("Select an item in the list first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub FillList()
        lstList.Clear()
        If goCurrentPlayer Is Nothing = False Then
            If gbAliased = True Then
                lstList.AddItem("<Your Account>", False)
                lstList.ItemData(lstList.NewIndex) = Int32.MinValue
            End If
            For X As Int32 = 0 To goCurrentPlayer.mlAllowanceUB
                With goCurrentPlayer.muAllowances(X)
                    lstList.AddItem(.sPlayerName, False)
                    lstList.ItemData(lstList.NewIndex) = X
                End With
            Next X
        End If
        lstList.SortList(False, False)
    End Sub

    Private Sub frmAliasLogins_OnNewFrame() Handles Me.OnNewFrame
        If goCurrentPlayer Is Nothing = False Then
            Dim lCnt As Int32 = 1
            If gbAliased = True Then lCnt = 2
            If goCurrentPlayer.mlAllowanceUB + lCnt <> lstList.ListCount Then
                FillList()
            End If
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub frmAliasLogins_WindowMoved() Handles Me.WindowMoved
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmAliasing")
        If ofrm Is Nothing = False Then
            ofrm.Left = Me.Left + Me.Width
            ofrm.Top = Me.Top
        Else
            Dim oofrm As UIWindow = MyBase.moUILib.GetWindow("frmOptions")
            If oofrm Is Nothing = False Then
                oofrm.Left = Me.Left + Me.Width
                oofrm.Top = Me.Top
            End If
            oofrm = Nothing
        End If
    End Sub

    Private Sub lstList_ItemDblClick(ByVal lIndex As Integer) Handles lstList.ItemDblClick
        btnLogin_Click(Nothing)
    End Sub
End Class