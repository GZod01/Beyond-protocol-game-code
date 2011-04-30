Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmArenaMain
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblStats As UILabel
    Private txtStatistics As UITextBox
    Private lblCurrent As UILabel
    Private lstArenas As UIListBox
    Private WithEvents btnClose As UIButton
    Private WithEvents btnEnterArena As UIButton 
    Private WithEvents btnCreateNew As UIButton

    Private moArenaItems(-1) As Arena
    Private mlArenaItemUB As Int32 = -1

    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmArenaMain initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAgentView

            .ControlName = "frmArenaMain"
            .Left = 166
            .Top = 50
            .Width = 512
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = Debugger.IsAttached
            .Moveable = True
            .BorderLineWidth = 2

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1
            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.ArenaMainX
                lTop = muSettings.ArenaMainY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 300
            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Left = lLeft
            .Top = lTop
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 137
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Arena Management"
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

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 487
            .Top = 1
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

        'lblStats initial props
        lblStats = New UILabel(oUILib)
        With lblStats
            .ControlName = "lblStats"
            .Left = 5
            .Top = 30
            .Width = 138
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Your Statistics in Arena"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblStats, UIControl))

        'txtStatistics initial props
        txtStatistics = New UITextBox(oUILib)
        With txtStatistics
            .ControlName = "txtStatistics"
            .Left = 5
            .Top = 50
            .Width = 500
            .Height = 240
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .Locked = True
            .MultiLine = True
        End With
        Me.AddChild(CType(txtStatistics, UIControl))

        'lblCurrent initial props
        lblCurrent = New UILabel(oUILib)
        With lblCurrent
            .ControlName = "lblCurrent"
            .Left = 5
            .Top = 295
            .Width = 138
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Active Arenas"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCurrent, UIControl))

        'lstArenas initial props
        lstArenas = New UIListBox(oUILib)
        With lstArenas
            .ControlName = "lstArenas"
            .Left = 5
            .Top = 315
            .Width = 500
            .Height = 155
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .sHeaderRow = Arena.GetArenaMainListBoxHeader()
        End With
        Me.AddChild(CType(lstArenas, UIControl))

        'btnEnterArena initial props
        btnEnterArena = New UIButton(oUILib)
        With btnEnterArena
            .ControlName = "btnEnterArena"
            .Left = 5
            .Top = 480
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Enter Arena"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnEnterArena, UIControl))

        'btnCreateNew initial props
        btnCreateNew = New UIButton(oUILib)
        With btnCreateNew
            .ControlName = "btnCreateNew"
            .Left = 406
            .Top = 480
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Create New"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCreateNew, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        DoTestData()
        RequestArenaList()

        mbLoading = False
    End Sub

    Private Sub frmArenaMain_OnNewFrame() Handles Me.OnNewFrame
        If lstArenas.ListCount <> mlArenaItemUB + 1 Then
            lstArenas.Clear()
            Try
                For X As Int32 = 0 To mlArenaItemUB
                    If moArenaItems(X) Is Nothing = False Then
                        lstArenas.AddItem(moArenaItems(X).GetArenaMainListText, False)
                        lstArenas.ItemData(lstArenas.NewIndex) = moArenaItems(X).lArenaID
                        lstArenas.ItemData2(lstArenas.NewIndex) = moArenaItems(X).lCreatorID
                    End If
                Next X
            Catch
                lstArenas.Clear()
            End Try
        Else
            Try
                For X As Int32 = 0 To mlArenaItemUB
                    If moArenaItems(X).lArenaID <> lstArenas.ItemData(X) Then
                        lstArenas.Clear()
                        Exit For
                    Else
                        Dim sValue As String = moArenaItems(X).GetArenaMainListText
                        If lstArenas.List(X) <> sValue Then lstArenas.List(X) = sValue
                    End If
                Next X
            Catch
                lstArenas.Clear()
            End Try
        End If
    End Sub

    Private Sub frmArenaMain_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.ArenaMainX = Me.Left
            muSettings.ArenaMainY = Me.Top
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnCreateNew_Click(ByVal sName As String) Handles btnCreateNew.Click
        'TODO: Since DM is the only mode right now, we'll do a straight to arena config. Eventually, we will want an interim form to select game mode
        Dim oFrm As New frmArenaConfig(goUILib, eyGameMode.eDeathMatch)
        oFrm.Visible = True
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnEnterArena_Click(ByVal sName As String) Handles btnEnterArena.Click
        If lstArenas.ListIndex > -1 Then
            Dim lArenaID As Int32 = lstArenas.ItemData(lstArenas.ListIndex)
            Dim oArena As Arena = Nothing
            For X As Int32 = 0 To mlArenaItemUB
                If moArenaItems(X) Is Nothing = False Then
                    If moArenaItems(X).lArenaID = lArenaID Then
                        oArena = moArenaItems(X)
                        Exit For
                    End If
                End If
            Next X
            If oArena Is Nothing Then Return
            'Ok, what state is the arena in?

            Select Case oArena.yArenaState
                Case eyArenaState.eClosed
                    MyBase.moUILib.AddNotification("That arena is closed and cannot be entered at this time.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case eyArenaState.eForming, eyArenaState.eInitial, eyArenaState.eStarting, eyArenaState.eValidation
                    Dim oFrm As New frmArenaWaiting(goUILib)
                    oFrm.SetArena(oArena)
                    btnClose_Click(btnClose.ControlName)
                Case Else       'in progress, process results, round finished
                    'in all these cases, try to ENTER the arena by changing to Arena mode
            End Select
        Else
            MyBase.moUILib.AddNotification("You must select an arena to enter.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If

    End Sub

    Private Sub DoTestData()
        Dim lUB As Int32 = 2
        Dim oItems(lUB) As DM_Arena
        For X As Int32 = 0 To lUB
            oItems(X) = New DM_Arena
            With oItems(X)
                .lArenaID = X + 1
                Select Case X
                    Case 0
                        .lCreatorID = 1
                    Case 1
                        .lCreatorID = 2
                    Case 2
                        .lCreatorID = 6
                End Select

                .lDuration = 480
                .lMapID = 1
                .lSideCnt = CInt(Rnd() * 2) + 2

                ReDim .oSides(.lSideCnt - 1)
                For Y As Int32 = 0 To .lSideCnt - 1
                    .oSides(Y) = New ArenaSide()
                    .oSides(Y).lSideID = Y + 1
                Next Y

                .lPlayerCnt = CInt(Rnd() * 4) + 3

                For Y As Int32 = 0 To .lPlayerCnt - 1
                    .AddPlayer(CInt(Rnd() * 555), CInt(Rnd() * .lSideCnt) + 1)
                Next

                If CInt(Rnd() * 100) < 50 Then
                    .yBaseArenaFlags = eyGeneralArenaFlags.eResultsTracked
                Else : .yBaseArenaFlags = 0
                End If
                .yGameMode = eyGameMode.eDeathMatch
                .yPublicity = eyPublicity.ePublicToAll
                .yArenaState = eyArenaState.eForming
            End With
        Next X

        moArenaItems = oItems
        mlArenaItemUB = lUB

    End Sub
    Private Sub RequestArenaList()
        Dim yMsg(7) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(ObjectType.eArena).CopyTo(yMsg, lPos) : lPos += 2
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Public Sub HandleArenasListItem(ByVal yData() As Byte)
        Dim lPos As Int32 = 2   'for msg code

        Dim yGameMode As Byte = yData(lPos) : lPos += 1
        Dim lArenaID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oArena As Arena = Nothing

        SyncLock moArenaItems
            Dim bFound As Boolean = False
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To mlArenaItemUB
                If moArenaItems(X) Is Nothing = False Then
                    If moArenaItems(X).lArenaID = lArenaID Then
                        oArena = moArenaItems(X)
                        bFound = True
                        Exit For
                    End If
                ElseIf lIdx = -1 Then
                    lIdx = X
                End If
            Next X
            If bFound = False Then
                If lIdx = -1 Then
                    mlArenaItemUB += 1
                    lIdx = mlArenaItemUB
                    ReDim Preserve moArenaItems(mlArenaItemUB)
                End If

                Select Case CType(yGameMode, eyGameMode)
                    Case eyGameMode.eDeathMatch
                        moArenaItems(lIdx) = New DM_Arena()
                    Case Else
                        Return
                End Select

                oArena = moArenaItems(lIdx)
            End If
        End SyncLock

        If oArena Is Nothing Then Return

        With oArena
            .lArenaID = lArenaID
            .lCreatorID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lDuration = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lMapID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSideCnt = yData(lPos) : lPos += 1
            .lPlayerCnt = yData(lPos) : lPos += 1
            .yBaseArenaFlags = yData(lPos) : lPos += 1
            .yGameMode = yGameMode
            .yPublicity = yData(lPos) : lPos += 1
            .yArenaState = CType(yData(lPos), eyArenaState) : lPos += 1

            'NOT INCLUDED
            '.lInstanceCreateTime()
            '.lInstanceStartTime()
            '.lMaxFlyUnits()
            '.lMaxGroundUnits()
            '.lMaxUnitCount()
            '.lRespawnDelay()
            '.lRespawnLimit()
            '.lSideCnt()
            '.lUnitLimitUB()
            '.Map
            '.oSides()
            '.oUnitLimits()            
        End With
    End Sub
End Class