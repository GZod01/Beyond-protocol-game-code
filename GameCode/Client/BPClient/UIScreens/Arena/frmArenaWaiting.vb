Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmArenaWaiting
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblPlayers As UILabel
    Private lblAvail As UILabel
    Private lblUnitsAssigned As UILabel
    Private lblChangeSide As UILabel
    Private lnDiv2 As UILine
    Private lnDiv3 As UILine
    Private txtSummary As UITextBox
    Private tvwUnitsAvail As UITreeView
    Private tvwUnitsAssign As UITreeView
    Private lstPlayers As UIListBox
    
    Private WithEvents cboCurrentSide As UIComboBox
    Private WithEvents btnViewRules As UIButton
    Private WithEvents btnAssign As UIButton
    Private WithEvents btnRemoveFromArena As UIButton
    Private WithEvents chkReady As UICheckBox
    Private WithEvents btnLeaveArena As UIButton

    Private mbLoading As Boolean = True
    Private moArena As Arena = Nothing

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmArenaWaiting initial props
        With Me
            .ControlName = "frmArenaWaiting"
            
            .Width = 512
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 2

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1
            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.ArenaWaitX
                lTop = muSettings.ArenaWaitY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256
            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Left = lLeft
            .Top = lTop
        End With

        'lstPlayers initial props
        lstPlayers = New UIListBox(oUILib)
        With lstPlayers
            .ControlName = "lstPlayers"
            .Left = 5
            .Top = 50
            .Width = 245
            .Height = 139
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstPlayers, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 145
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Arena Waiting Room"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 25
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblPlayers initial props
        lblPlayers = New UILabel(oUILib)
        With lblPlayers
            .ControlName = "lblPlayers"
            .Left = 5
            .Top = 30
            .Width = 49
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Players"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPlayers, UIControl))

        'lblAvail initial props
        lblAvail = New UILabel(oUILib)
        With lblAvail
            .ControlName = "lblAvail"
            .Left = 5
            .Top = 205
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Units Available"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAvail, UIControl))

        'tvwUnitsAvail initial props
        tvwUnitsAvail = New UITreeView(oUILib)
        With tvwUnitsAvail
            .ControlName = "tvwUnitsAvail"
            .Left = 5
            .Top = 225
            .Width = 245
            .Height = 215
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(tvwUnitsAvail, UIControl))

        'tvwUnitsAssign initial props
        tvwUnitsAssign = New UITreeView(oUILib)
        With tvwUnitsAssign
            .ControlName = "tvwUnitsAssign"
            .Left = 260
            .Top = 225
            .Width = 245
            .Height = 215
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(tvwUnitsAssign, UIControl))

        'lblUnitsAssigned initial props
        lblUnitsAssigned = New UILabel(oUILib)
        With lblUnitsAssigned
            .ControlName = "lblUnitsAssigned"
            .Left = 260
            .Top = 205
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Units Assigned"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblUnitsAssigned, UIControl))

        'lblChangeSide initial props
        lblChangeSide = New UILabel(oUILib)
        With lblChangeSide
            .ControlName = "lblChangeSide"
            .Left = 260
            .Top = 50
            .Width = 83
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Current Side:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblChangeSide, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 1
            .Top = 200
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'btnViewRules initial props
        btnViewRules = New UIButton(oUILib)
        With btnViewRules
            .ControlName = "btnViewRules"
            .Left = 305
            .Top = 170
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "View Arena Config"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnViewRules, UIControl))

        'txtSummary initial props
        txtSummary = New UITextBox(oUILib)
        With txtSummary
            .ControlName = "txtSummary"
            .Left = 260
            .Top = 75
            .Width = 245
            .Height = 85
            .Enabled = True
            .Visible = True
            .Caption = "Arena Summary goes here"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
            .Locked = True
        End With
        Me.AddChild(CType(txtSummary, UIControl))

        'btnAssign initial props
        btnAssign = New UIButton(oUILib)
        With btnAssign
            .ControlName = "btnAssign"
            .Left = 45
            .Top = 450
            .Width = 160
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Assign to Arena"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAssign, UIControl))

        'btnRemoveFromArena initial props
        btnRemoveFromArena = New UIButton(oUILib)
        With btnRemoveFromArena
            .ControlName = "btnRemoveFromArena"
            .Left = 305
            .Top = 450
            .Width = 160
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Remove From Arena"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRemoveFromArena, UIControl))

        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = 1
            .Top = 480
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'chkReady initial props
        chkReady = New UICheckBox(oUILib)
        With chkReady
            .ControlName = "chkReady"
            .Left = 140
            .Top = 485
            .Width = 219
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Check This Box When Ready"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkReady, UIControl))

        'btnLeaveArena initial props
        btnLeaveArena = New UIButton(oUILib)
        With btnLeaveArena
            .ControlName = "btnLeaveArena"
            .Left = 387
            .Top = 1
            .Width = 125
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Leave Arena"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnLeaveArena, UIControl))

        'cboCurrentSide initial props
        cboCurrentSide = New UIComboBox(oUILib)
        With cboCurrentSide
            .ControlName = "cboCurrentSide"
            .Left = 355
            .Top = 50
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboCurrentSide, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        mbLoading = False
    End Sub

    Private Sub FillPlayerList()
        lstPlayers.Clear()

        For X As Int32 = 0 To moArena.lSideCnt - 1
            Dim oSide As ArenaSide = moArena.oSides(X)
            If oSide Is Nothing = False Then
                lstPlayers.AddItem("Side: " & oSide.lSideID, True)
                lstPlayers.ItemData(lstPlayers.NewIndex) = -1
                lstPlayers.ItemData2(lstPlayers.NewIndex) = oSide.lSideID
                lstPlayers.ItemLocked(lstPlayers.NewIndex) = True

                If oSide.oSidePlayers Is Nothing = False Then
                    For Y As Int32 = 0 To oSide.oSidePlayers.GetUpperBound(0)
                        Dim oASP As ArenaSidePlayer = oSide.oSidePlayers(Y)
                        If oASP Is Nothing = False Then
                            Dim bReady As Boolean = (oASP.yFlags And eyArenaSidePlayerFlag.PlayerReady) <> 0

                            lstPlayers.AddItem(GetCacheObjectValue(oASP.lPlayerID, ObjectType.ePlayer), bReady)
                            If bReady = True Then
                                lstPlayers.ItemCustomColor(lstPlayers.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                            Else
                                lstPlayers.ItemCustomColor(lstPlayers.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                            End If
                            lstPlayers.ItemData(lstPlayers.NewIndex) = oASP.lPlayerID
                            lstPlayers.ItemData2(lstPlayers.NewIndex) = oASP.yFlags
                        End If
                    Next Y
                End If
            End If
        Next X
    End Sub

    Private Sub frmArenaWaiting_OnNewFrame() Handles Me.OnNewFrame
        If moArena Is Nothing Then Return

        Try
            'Ok, find the player in the sides...
            Dim oASP As ArenaSidePlayer = moArena.GetPlayerSidePlayer(glPlayerID)
            If oASP Is Nothing = False Then
                Dim lCurSide As Int32 = -1
                If cboCurrentSide.ListIndex > -1 Then lCurSide = cboCurrentSide.ItemData(cboCurrentSide.ListIndex)
                If oASP.oSide.lSideID <> lCurSide Then
                    mbLoading = True
                    cboCurrentSide.FindComboItemData(lCurSide)
                    mbLoading = False
                End If
            End If

            Dim sText As String = moArena.GetArenaSummaryText()
            If sText <> txtSummary.Caption Then txtSummary.Caption = sText

            If moArena.lPlayerCnt <> lstPlayers.ListCount Then
                FillPlayerList()
            Else
                Dim lSideID As Int32 = 1
                Dim bGood As Boolean = True

                For X As Int32 = 0 To lstPlayers.ListCount - 1
                    Dim lID As Int32 = lstPlayers.ItemData(X)
                    If lID = -1 Then
                        lSideID = lstPlayers.ItemData2(X)
                    Else
                        oASP = moArena.GetPlayerSidePlayer(lID)
                        If oASP Is Nothing OrElse oASP.oSide.lSideID <> lSideID Then
                            bGood = False
                            Exit For
                        Else
                            If oASP.yFlags <> lstPlayers.ItemData2(X) Then
                                bGood = False
                                Exit For
                            Else
                                sText = GetCacheObjectValue(lID, ObjectType.ePlayer)
                                If lstPlayers.List(X) <> sText Then
                                    lstPlayers.List(X) = sText
                                    lstPlayers.IsDirty = True
                                End If
                                If oASP.lPlayerID = glPlayerID Then
                                    Dim bValue As Boolean = (oASP.yFlags And eyArenaSidePlayerFlag.PlayerReady) <> 0
                                    If chkReady.Value <> bValue Then
                                        mbLoading = True
                                        chkReady.Value = bValue
                                        mbLoading = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next X
            End If

        Catch
        End Try
        


    End Sub

    Private Sub frmArenaWaiting_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.ArenaWaitX = Me.Left
            muSettings.ArenaWaitY = Me.Top
        End If
    End Sub

    Private Sub btnAssign_Click(ByVal sName As String) Handles btnAssign.Click
        If tvwUnitsAvail.oSelectedNode Is Nothing = False Then
            Dim lID As Int32 = tvwUnitsAvail.oSelectedNode.lItemData
            Dim iTypeID As Int16 = CShort(tvwUnitsAvail.oSelectedNode.lItemData2)
            If iTypeID > -1 Then
                Dim yMsg(12) As Byte
                Dim lPos As Int32 = 0

                System.BitConverter.GetBytes(GlobalMessageCode.eArenaUnitAssignment).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(moArena.lArenaID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = 1 : lPos += 1      '1 to indicate add the unit to arena
                System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
        End If
    End Sub

    Private Sub btnLeaveArena_Click(ByVal sName As String) Handles btnLeaveArena.Click
        If btnLeaveArena.Caption.ToUpper = "CONFIRM" Then
            Dim yMsg(6) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eArenaPlayerStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(moArena.lArenaID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyArenaPlayerActionCode.PlayerLeaveArena : lPos += 1
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            MyBase.moUILib.RemoveWindow(Me.ControlName)
        Else
            btnLeaveArena.Caption = "Confirm"
        End If
    End Sub
 
    Private Sub btnRemoveFromArena_Click(ByVal sName As String) Handles btnRemoveFromArena.Click

        If tvwUnitsAssign.oSelectedNode Is Nothing = False Then
            Dim lID As Int32 = tvwUnitsAssign.oSelectedNode.lItemData
            Dim iTypeID As Int16 = CShort(tvwUnitsAssign.oSelectedNode.lItemData2)
            If iTypeID > -1 Then
                Dim yMsg(12) As Byte
                Dim lPos As Int32 = 0

                System.BitConverter.GetBytes(GlobalMessageCode.eArenaUnitAssignment).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(moArena.lArenaID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = 0 : lPos += 1      '0 to indicate remove the unit from the arena
                System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
        End If
    End Sub

    Private Sub btnViewRules_Click(ByVal sName As String) Handles btnViewRules.Click
        Dim oFrm As frmArenaConfig = CType(goUILib.GetWindow("frmArenaConfig"), frmArenaConfig)
        If oFrm Is Nothing = False Then goUILib.RemoveWindow(oFrm.ControlName)
        oFrm = New frmArenaConfig(goUILib, CType(moArena.yGameMode, eyGameMode))
        oFrm.SetFromArena(moArena)
    End Sub

    Private Sub cboCurrentSide_ItemSelected(ByVal lItemIndex As Integer) Handles cboCurrentSide.ItemSelected
        If cboCurrentSide.ListIndex > -1 AndAlso mbLoading = False Then
            Dim lSideID As Int32 = cboCurrentSide.ItemData(cboCurrentSide.ListIndex)
            Dim yMsg(10) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eArenaPlayerStatus).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(moArena.lArenaID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = eyArenaPlayerActionCode.PlayerChangesSides : lPos += 1
            System.BitConverter.GetBytes(lSideID).CopyTo(yMsg, lPos) : lPos += 4
            MyBase.moUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Private Sub chkReady_Click() Handles chkReady.Click
        Dim yMsg(6) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eArenaPlayerStatus).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(moArena.lArenaID).CopyTo(yMsg, lPos) : lPos += 4
        If chkReady.Value = True Then
            yMsg(lPos) = eyArenaPlayerActionCode.SetPlayerReady
        Else
            yMsg(lPos) = eyArenaPlayerActionCode.SetPlayerNotReady
        End If
        lPos += 1

        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Public Sub SetArena(ByRef oArena As Arena)
        moArena = oArena
        'Request the Arena's Details, etc...
        moArena.RequestDetails()
    End Sub
End Class