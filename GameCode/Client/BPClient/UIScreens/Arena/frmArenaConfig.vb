Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmArenaConfig
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private fraMap As UIWindow
    Private lblDuration As UILabel
    Private lblHrs As UILabel
    Private lblLimits As UILabel
    Private lblHullType As UILabel
    Private lblMaxHull As UILabel
    Private lblMaxCnt As UILabel
    Private txtMaxCnt As UITextBox
    Private lnDiv2 As UILine
    Private lnDiv3 As UILine
    Private txtHullSize As UITextBox
    Private txtDuration As UITextBox
    Private cboHullType As UIComboBox
    Private chkTrackResults As UICheckBox

    Private WithEvents btnClose As UIButton
    Private WithEvents lstMaps As UIListBox
    Private WithEvents optPublic As UIOption
    Private WithEvents optGuildOnly As UIOption
    Private WithEvents optInviteOnly As UIOption
    Private WithEvents lstLimits As UIListBox
    Private WithEvents btnRemoveLimit As UIButton
    Private WithEvents btnAddLimit As UIButton
    Private WithEvents btnSubmit As UIButton
    Private WithEvents btnCancel As UIButton

    Private fraAdvanced As fraArenaSubFrame
    Private mbLoading As Boolean = True
    Private myGameMode As Byte
    Private mlArenaID As Int32 = -1

    Public Sub New(ByRef oUILib As UILib, ByVal yType As eyGameMode)
        MyBase.New(oUILib)

        myGameMode = yType
        'frmArenaConfig initial props
        With Me
            .ControlName = "frmArenaConfig"
            .Left = 165
            .Top = 55
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
                lLeft = muSettings.ArenaConfigX
                lTop = muSettings.ArenaConfigY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256
            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Left = lLeft
            .Top = lTop
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 8
            .Top = 5
            .Width = 139
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Arena Configuration"
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

        'fraMap initial props
        fraMap = New UIWindow(oUILib)
        With fraMap
            .ControlName = "fraMap"
            .Left = 5
            .Top = 30
            .Width = 120
            .Height = 120
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .BorderLineWidth = 2
            .FillWindow = False
            .bRoundedBorder = False
        End With
        Me.AddChild(CType(fraMap, UIControl))

        'lstMaps initial props
        lstMaps = New UIListBox(oUILib)
        With lstMaps
            .ControlName = "lstMaps"
            .Left = 130
            .Top = 30
            .Width = 220
            .Height = 120
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstMaps, UIControl))

        'optPublic initial props
        optPublic = New UIOption(oUILib)
        With optPublic
            .ControlName = "optPublic"
            .Left = 360
            .Top = 30
            .Width = 54
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Public"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
        End With
        Me.AddChild(CType(optPublic, UIControl))

        'optGuildOnly initial props
        optGuildOnly = New UIOption(oUILib)
        With optGuildOnly
            .ControlName = "optGuildOnly"
            .Left = 360
            .Top = 50
            .Width = 78
            .Height = 18
            .Enabled = goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False
            .Visible = True
            .Caption = "Guild Only"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optGuildOnly, UIControl))

        'optInviteOnly initial props
        optInviteOnly = New UIOption(oUILib)
        With optInviteOnly
            .ControlName = "optInviteOnly"
            .Left = 360
            .Top = 70
            .Width = 78
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Invite Only"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optInviteOnly, UIControl))

        'chkTrackResults initial props
        chkTrackResults = New UICheckBox(oUILib)
        With chkTrackResults
            .ControlName = "chkTrackResults"
            .Left = 360
            .Top = 90
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Track Results"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkTrackResults, UIControl))

        'lblDuration initial props
        lblDuration = New UILabel(oUILib)
        With lblDuration
            .ControlName = "lblDuration"
            .Left = 358
            .Top = 115
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Duration:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDuration, UIControl))

        'txtDuration initial props
        txtDuration = New UITextBox(oUILib)
        With txtDuration
            .ControlName = "txtDuration"
            .Left = 416
            .Top = 115
            .Width = 30
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
            .MaxLength = 2
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtDuration, UIControl))

        'lblHrs initial props
        lblHrs = New UILabel(oUILib)
        With lblHrs
            .ControlName = "lblHrs"
            .Left = 455
            .Top = 115
            .Width = 37
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "hours"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHrs, UIControl))

        'lblLimits initial props
        lblLimits = New UILabel(oUILib)
        With lblLimits
            .ControlName = "lblLimits"
            .Left = 5
            .Top = 165
            .Width = 109
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Unit Limitations"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblLimits, UIControl))

        'lstLimits initial props
        lstLimits = New UIListBox(oUILib)
        With lstLimits
            .ControlName = "lstLimits"
            .Left = 5
            .Top = 185
            .Width = 345
            .Height = 130
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstLimits, UIControl))

        'lblHullType initial props
        lblHullType = New UILabel(oUILib)
        With lblHullType
            .ControlName = "lblHullType"
            .Left = 360
            .Top = 185
            .Width = 88
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hull Type:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHullType, UIControl))

        'txtHullSize initial props
        txtHullSize = New UITextBox(oUILib)
        With txtHullSize
            .ControlName = "txtHullSize"
            .Left = 370
            .Top = 250
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "0"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 7
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtHullSize, UIControl))

        'lblMaxHull initial props
        lblMaxHull = New UILabel(oUILib)
        With lblMaxHull
            .ControlName = "lblMaxHull"
            .Left = 360
            .Top = 230
            .Width = 88
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Max Hull Size:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMaxHull, UIControl))

        'lblMaxCnt initial props
        lblMaxCnt = New UILabel(oUILib)
        With lblMaxCnt
            .ControlName = "lblMaxCnt"
            .Left = 360
            .Top = 275
            .Width = 88
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Max Count:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMaxCnt, UIControl))

        'txtMaxCnt initial props
        txtMaxCnt = New UITextBox(oUILib)
        With txtMaxCnt
            .ControlName = "txtMaxCnt"
            .Left = 370
            .Top = 295
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "0"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 7
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtMaxCnt, UIControl))

        'btnRemoveLimit initial props
        btnRemoveLimit = New UIButton(oUILib)
        With btnRemoveLimit
            .ControlName = "btnRemoveLimit"
            .Left = 250
            .Top = 320
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Remove"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRemoveLimit, UIControl))

        'btnAddLimit initial props
        btnAddLimit = New UIButton(oUILib)
        With btnAddLimit
            .ControlName = "btnAddLimit"
            .Left = 380
            .Top = 320
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
        Me.AddChild(CType(btnAddLimit, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 1
            .Top = 160
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = 1
            .Top = 350
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'fraAdvanced initial props
        'fraAdvanced = New UIWindow(oUILib)
        mlArenaID = -1
        Select Case yType
            Case eyGameMode.eDeathMatch
                fraAdvanced = New frmArenaConfig.fraDMArena(oUILib)
        End Select

        If fraAdvanced Is Nothing = False Then
            With fraAdvanced
                .ControlName = "fraAdvanced"
                .Left = 5
                .Top = 360
                .Width = 500
                .Height = 115
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .Caption = "Advanced Configuration"
            End With
            Me.AddChild(CType(fraAdvanced, UIControl))
        End If

        'btnSubmit initial props
        btnSubmit = New UIButton(oUILib)
        With btnSubmit
            .ControlName = "btnSubmit"
            .Left = 405
            .Top = 482
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Submit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSubmit, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = 300
            .Top = 482
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

        'cboHullType initial props
        cboHullType = New UIComboBox(oUILib)
        With cboHullType
            .ControlName = "cboHullType"
            .Left = 370
            .Top = 205
            .Width = 135
            .Height = 20
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboHullType, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        FillHullTypesList()
        FillMapList()

        mbLoading = False
    End Sub

    Public Sub SetFromArena(ByRef oArena As Arena)
        'Fill values from the arena object here
        With oArena
            mlArenaID = .lArenaID
            txtDuration.Caption = .lDuration.ToString
            For X As Int32 = 0 To lstMaps.ListCount - 1
                If lstMaps.ItemData(X) = .lMapID Then
                    lstMaps.ListIndex = X
                End If
            Next X

            chkTrackResults.Value = (.yBaseArenaFlags And eyGeneralArenaFlags.eResultsTracked) <> 0

            mbLoading = True
            optPublic.Value = False
            optGuildOnly.Value = False
            optInviteOnly.Value = False
            Select Case .yPublicity
                Case eyPublicity.ePublicToAll
                    optPublic.Value = True
                Case eyPublicity.eInviteOnly
                    optInviteOnly.Value = True
                Case eyPublicity.eGuildOnly
                    optGuildOnly.Value = True
            End Select
            mbLoading = False

            lstLimits.Clear()
            For X As Int32 = 0 To .lUnitLimitUB
                With .oUnitLimits(X)
                    Dim sValue As String = GetLimitListText(.yHullType, .lHullSize, .lMaxCnt)
                    lstLimits.AddItem(sValue, False)
                    lstLimits.ItemData(lstLimits.NewIndex) = .yHullType
                    lstLimits.ItemData2(lstLimits.NewIndex) = .lHullSize
                    lstLimits.ItemData3(lstLimits.NewIndex) = .lMaxCnt
                End With
            Next X 
        End With

        'use creator and arenastate to determine whether to diable the controls
        If oArena.lCreatorID <> glPlayerID Then
            DisableControls()
        Else
            If oArena.yArenaState <> eyArenaState.eForming AndAlso oArena.yArenaState <> eyArenaState.eInitial Then
                DisableControls()
            End If
        End If

        If fraAdvanced Is Nothing = False Then
            fraAdvanced.SetFromArena(oArena)
        End If
    End Sub

    Private Sub DisableControls()
        txtMaxCnt.Enabled = False
        txtHullSize.Enabled = False
        txtDuration.Enabled = False
        cboHullType.Enabled = False
        chkTrackResults.Locked = True
        lstMaps.Enabled = False
        optPublic.Locked = True
        optGuildOnly.Locked = True
        optInviteOnly.Locked = True
        btnRemoveLimit.Enabled = False
        btnAddLimit.Enabled = False
        btnSubmit.Enabled = False
    End Sub

    Private Sub FillHullTypesList()
        cboHullType.Clear()

        '254 indicates All Unit limit, 253 indicates Ground unit, 252 indicates Flying Unit
        cboHullType.AddItem("All Units") : cboHullType.ItemData(cboHullType.NewIndex) = 254
        cboHullType.AddItem("Ground Units") : cboHullType.ItemData(cboHullType.NewIndex) = 253
        cboHullType.AddItem("Flying Units") : cboHullType.ItemData(cboHullType.NewIndex) = 252

        cboHullType.AddItem("Battlecruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.BattleCruiser
        cboHullType.AddItem("Battleship") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Battleship
        cboHullType.AddItem("Corvette") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Corvette
        cboHullType.AddItem("Cruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Cruiser
        cboHullType.AddItem("Destroyer") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Destroyer
        cboHullType.AddItem("Escort") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Escort
        'cboHullType.AddItem("Facility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Facility
        cboHullType.AddItem("Fighter (Light)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.LightFighter
        cboHullType.AddItem("Fighter (Medium)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.MediumFighter
        cboHullType.AddItem("Fighter (Heavy)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.HeavyFighter
        cboHullType.AddItem("Frigate") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Frigate
        cboHullType.AddItem("Naval Battleship") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalBattleship
        cboHullType.AddItem("Naval Carrier") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalCarrier
        cboHullType.AddItem("Naval Cruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalCruiser
        cboHullType.AddItem("Naval Destroyer") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalDestroyer
        cboHullType.AddItem("Naval Frigate") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalFrigate
        cboHullType.AddItem("Naval Submarine") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalSub
        cboHullType.AddItem("Naval Utility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Utility
        cboHullType.AddItem("Small Vehicle") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.SmallVehicle
        'cboHullType.AddItem("Space Station") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.SpaceStation
        cboHullType.AddItem("Tank") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Tank
        cboHullType.AddItem("Utility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Utility
    End Sub

    Private Sub FillMapList()
        'Add each map for this game type
        lstMaps.Clear()
        For X As Int32 = 0 To ArenaMap.AllMapUB
            Dim oMap As ArenaMap = ArenaMap.AllMaps(X)
            If oMap Is Nothing = False Then
                If oMap.yGameMode = myGameMode Then
                    lstMaps.AddItem(("(" & oMap.lMinPlayerCnt & "-" & oMap.lMaxPlayerCnt & ")").PadRight(7, " "c) & oMap.sMapName, False)
                    lstMaps.ItemData(lstMaps.NewIndex) = oMap.lMapID
                End If
            End If
        Next X
    End Sub

    Private Function GetLimitListText(ByVal lHullType As Int32, ByVal lHullSize As Int32, ByVal lMaxCnt As Int32) As String
        Dim sValue As String = ""
        Select Case lHullType
            Case 254
                sValue = "All Units: "
            Case 253
                sValue = "Ground Units: "
            Case 252
                sValue = "Flying Units: "
            Case Else
                sValue = Base_Tech.GetHullTypeName(CByte(lHullType)) & ": "
        End Select

        If lMaxCnt > 0 Then
            sValue &= " Max " & lMaxCnt.ToString
        End If
        If lHullSize > 0 Then
            sValue &= " Max Hull " & lHullSize.ToString
        End If

        Return sValue
    End Function

    Private Sub btnAddLimit_Click(ByVal sName As String) Handles btnAddLimit.Click
        If cboHullType.ListIndex > -1 Then
            Dim lHullType As Int32 = cboHullType.ItemData(cboHullType.ListIndex)
            Dim lHullSize As Int32 = CInt(Val(txtHullSize.Caption))
            Dim lMaxCnt As Int32 = CInt(Val(txtMaxCnt.Caption))

            If lHullType < 0 Then lHullType = 0
            If lHullType > 255 Then lHullType = 255
            If lHullSize < 1 Then lHullSize = 0
            If lMaxCnt < 0 Then lMaxCnt = 0

            Dim sValue As String = GetLimitListText(lHullType, lHullSize, lMaxCnt)
           
            lstLimits.AddItem(sValue)
            lstLimits.ItemData(lstLimits.NewIndex) = lHullType
            lstLimits.ItemData2(lstLimits.NewIndex) = lHullSize
            lstLimits.ItemData3(lstLimits.NewIndex) = lMaxCnt
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click, btnCancel.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnRemoveLimit_Click(ByVal sName As String) Handles btnRemoveLimit.Click
        If lstLimits.ListIndex > -1 Then
            lstLimits.RemoveItem(lstLimits.ListIndex)
        Else
            MyBase.moUILib.AddNotification("Select a limit in the list to remove.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click
        'Verify the configuration
        Dim lDuration As Int32 = CInt(Val(txtDuration.Caption))
        Dim lMapID As Int32 = -1
        If lstMaps.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select a map for this Arena match.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        lMapID = lstMaps.ItemData(lstMaps.ListIndex)

        Dim yBaseFlags As Byte = 0
        If chkTrackResults.Value = True Then
            yBaseFlags = yBaseFlags Or eyGeneralArenaFlags.eResultsTracked
        End If
        Dim yPublicity As Byte = 0
        If optGuildOnly.Value = True Then
            yPublicity = eyPublicity.eGuildOnly
        ElseIf optInviteOnly.Value = True Then
            yPublicity = eyPublicity.eInviteOnly
        End If
        If fraAdvanced Is Nothing = False Then
            If fraAdvanced.AdvancedConfigValid = False Then Return
        Else
            MyBase.moUILib.AddNotification("An unexpected error occurred, please close and reopen this window.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'Create the configuration message
        Dim lLen As Int32 = 20 + (lstLimits.ListCount * 12) + fraAdvanced.AdvancedConfigMsgLen
        Dim yMsg(lLen) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eSubmitArenaConfig).CopyTo(yMsg, lPos) : lPos += 2
        yMsg(lPos) = myGameMode : lPos += 1
        System.BitConverter.GetBytes(mlArenaID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDuration).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMapID).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = yBaseFlags : lPos += 1
        yMsg(lPos) = yPublicity : lPos += 1

        System.BitConverter.GetBytes(lstLimits.ListCount).CopyTo(yMsg, lPos) : lPos += 4
        For X As Int32 = 0 To lstLimits.ListCount - 1
            System.BitConverter.GetBytes(lstLimits.ItemData(X)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lstLimits.ItemData2(X)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lstLimits.ItemData3(X)).CopyTo(yMsg, lPos) : lPos += 4
        Next X

        fraAdvanced.AdvancedConfigMsgAppend.CopyTo(yMsg, lPos) : lPos += fraAdvanced.AdvancedConfigMsgLen

        'Send it to the Server
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Private Sub lstLimits_ItemClick(ByVal lIndex As Integer) Handles lstLimits.ItemClick
        If lstLimits.ListIndex > -1 Then
            cboHullType.FindComboItemData(lstLimits.ItemData(lstLimits.ListIndex))
            txtHullSize.Caption = lstLimits.ItemData2(lstLimits.ListIndex).ToString
            txtMaxCnt.Caption = lstLimits.ItemData3(lstLimits.ListIndex).ToString
        End If
    End Sub

    Private Sub lstMaps_ItemClick(ByVal lIndex As Integer) Handles lstMaps.ItemClick
        If lstMaps.ListIndex > -1 Then
            Dim lID As Int32 = lstMaps.ItemData(lstMaps.ListIndex)
            Dim oMap As ArenaMap = ArenaMap.AllMapByID(lID)
            If oMap Is Nothing = False Then

                'do any Unit Limit config here...
                Dim bAll As Boolean = False
                'Dim bGrnd As Boolean = False
                'Dim bAir As Boolean = False
                '254 indicates All Unit limit, 253 indicates Ground unit, 252 indicates Flying Unit
                For X As Int32 = 0 To lstLimits.ListCount - 1
                    Select Case lstLimits.ItemData(X)
                        Case 254
                            bAll = True
                            If lstLimits.ItemData3(X) > oMap.lMaxUnitCntPerSide Then
                                lstLimits.ItemData3(X) = oMap.lMaxUnitCntPerSide
                                Dim sValue As String = GetLimitListText(254, lstLimits.ItemData2(X), oMap.lMaxUnitCntPerSide)
                                lstLimits.List(X) = sValue
                            End If
                            Exit For
                            'Case 253
                            'Case 252
                    End Select
                Next X

                If bAll = False AndAlso oMap.lMaxUnitCntPerSide > 0 Then
                    Dim sValue As String = GetLimitListText(254, 0, oMap.lMaxUnitCntPerSide)
                    lstLimits.AddItem(sValue, False)
                    lstLimits.ItemData(lstLimits.NewIndex) = 254
                    lstLimits.ItemData2(lstLimits.NewIndex) = 0
                    lstLimits.ItemData3(lstLimits.NewIndex) = oMap.lMaxUnitCntPerSide
                End If
            End If
        End If
    End Sub

    Private Sub optGuildOnly_Click() Handles optGuildOnly.Click
        If mbLoading = True Then Return
        mbLoading = True
        optInviteOnly.Value = False
        optPublic.Value = False
        optGuildOnly.Value = True
        mbLoading = False
    End Sub

    Private Sub optInviteOnly_Click() Handles optInviteOnly.Click
        If mbLoading = True Then Return
        mbLoading = True
        optInviteOnly.Value = True
        optPublic.Value = False
        optGuildOnly.Value = False
        mbLoading = False
    End Sub

    Private Sub optPublic_Click() Handles optPublic.Click
        If mbLoading = True Then Return
        mbLoading = True
        optInviteOnly.Value = False
        optPublic.Value = True
        optGuildOnly.Value = False
        mbLoading = False
    End Sub

    Private Sub frmArenaConfig_OnRender() Handles Me.OnRender

        If lstMaps.ListIndex > -1 Then
            Dim lID As Int32 = lstMaps.ItemData(lstMaps.ListIndex)
            Dim oMap As ArenaMap = ArenaMap.AllMapByID(lID)
            If oMap Is Nothing = False Then
                Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                If sPath.EndsWith("\") = False Then sPath &= "\"
                sPath &= "Arenas\"
                sPath &= oMap.sThumbnail
                If Exists(sPath) = True Then
                    Dim oTex As Texture = Nothing

                    Try
                        oTex = TextureLoader.FromFile(GFXEngine.moDevice, sPath)

                        BPSprite.Draw2DOnce(GFXEngine.moDevice, oTex, New Rectangle(0, 0, 128, 128), New Rectangle(fraMap.Left + 1, fraMap.Top + 1, fraMap.Width - 2, fraMap.Height - 2), Color.White, 128, 128)
                    Catch
                    End Try

                    If oTex Is Nothing = False Then oTex.Dispose()
                    oTex = Nothing

                End If

            End If
        End If
    End Sub
 

    Private Sub frmArenaConfig_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.ArenaConfigX = Me.Left
            muSettings.ArenaConfigY = Me.Top
        End If
    End Sub



    Private MustInherit Class fraArenaSubFrame
        Inherits UIWindow

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)
        End Sub

        Public MustOverride Sub SetFromArena(ByRef oArena As Arena)
        Public MustOverride Function AdvancedConfigValid() As Boolean
        Public MustOverride Function AdvancedConfigMsgLen() As Int32
        Public MustOverride Function AdvancedConfigMsgAppend() As Byte()

    End Class
End Class