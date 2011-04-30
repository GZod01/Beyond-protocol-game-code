Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'eEnvirRelations
' Int32 - Get or Set            (-1 to Get, 1 to Set)
' Int32 - Player ID             (-1 for all, or ObjectID of Player)
' Int32 - Enviromnet ID         (-1 for all, or ObjectID of environment)
' Int16 - Relation Value        (0 - Enemy, 1 - Friendly)
Public Class frmEnvirRelations
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private WithEvents btnClose As UIButton
    Private WithEvents tvwRelations As UITreeView
    Private Shared mbExp() As String

    Private mySortBy As Int16
    Private WithEvents optSortByPlayer As UIOption
    Private WithEvents optSortByRelations As UIOption
    Private lnDiv2 As UILine

    Private lblNewRelation As UILabel
    Private cboNewEnvironments As UIComboBox
    Private cboNewPlayers As UIComboBox
    Private WithEvents btnNewRelation As UIButton
    Private lnDiv3 As UILine

    Private lblExistingPlayers As UILabel
    Private cboExistingPlayers As UIComboBox
    Private cboExistingEnvironments As UIComboBox
    Private WithEvents btnSetFriendly As UIButton
    Private WithEvents btnSetEnemy As UIButton
    Private lnDiv4 As UILine

    Private lblGuilds As UILabel
    Private cboGuilds As UIComboBox
    Private cboGuildsEnvironments As UIComboBox
    Private WithEvents btnSetFriendlyGuild As UIButton
    Private WithEvents btnSetEnemyGuild As UIButton
    Private lnDiv5 As UILine


    Private myViewState As Byte = 0     'Environment = 0, Player = 1
    Private rcByEnvironment As Rectangle
    Private rcByPlayer As Rectangle

    Private mbLoading As Boolean = True
    Private mbHasUnknowns As Boolean = False
    Private mlLastUpdate As Int32 = -1

    Private Structure TEnvirRelations
        Dim lEnvirID As Int32
        Dim lPlayerID As Int32
        Dim iRelation As Int16 '0 Foe - 1 Friend 
    End Structure
    Private Shared moEnvirRelations() As TEnvirRelations
    Private Shared mlEnvirRelationsUB As Int32 = -1
    Dim mbRelationsRequested As Boolean = False

    Private Structure TGuild
        Dim ObjectID As Int32
        Dim sName As String
    End Structure
    Private moGuilds() As TGuild
    Private moGuildsUB As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmEnvirRelations initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAvailableResources
            .ControlName = "frmEnvirRelations"
            Dim lLeft As Int32 = muSettings.EnvirRelationsX
            Dim lTop As Int32 = muSettings.EnvirRelationsY

            .Width = 512
            .Height = 512

            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - (.Width \ 2)
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - (.Height \ 2)
            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height

            .Left = lLeft
            .Top = lTop

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .bRoundedBorder = True
            .BorderLineWidth = 2
            .mbAcceptReprocessEvents = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 300
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Environment Relations Manager"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
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
            .Top = btnClose.Top + btnClose.Height + 2
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'tvwRelations initial props
        tvwRelations = New UITreeView(oUILib)
        With tvwRelations
            .ControlName = "tvwRelations"
            .Left = 5
            .Top = lnDiv1.Top + 5
            .Width = 300
            .Height = Me.Height - lnDiv1.Top - 5 - (Me.BorderLineWidth * 2)
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(tvwRelations, UIControl))

        'optSortByPlayer initial props
        optSortByPlayer = New UIOption(oUILib)
        With optSortByPlayer
            .ControlName = "optSortByPlayer"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = 30
            .Width = 110
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Sort By Player"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
        End With
        Me.AddChild(CType(optSortByPlayer, UIControl))

        'optSortByRelations initial props
        optSortByRelations = New UIOption(oUILib)
        With optSortByRelations
            .ControlName = "optSortByRelations"
            .Width = 132
            .Height = 18
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = optSortByPlayer.Top + .Height + 5
            .Enabled = True
            .Visible = True
            .Caption = "Sort By Relations"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optSortByRelations, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = tvwRelations.Left + tvwRelations.Width
            .Top = optSortByRelations.Top + optSortByRelations.Height + 1
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        ''''' New Relations '''''
        'lblNewRelation initial props
        lblNewRelation = New UILabel(oUILib)
        With lblNewRelation
            .ControlName = "lblNewRelation"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = lnDiv2.Top + 10
            .Width = 300
            .Height = 44
            .Enabled = True
            .Visible = True
            .Caption = "New Relations for " & vbCrLf & "Environments you have no units."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblNewRelation, UIControl))

        'cboNewEnvironments initial props
        cboNewEnvironments = New UIComboBox(oUILib)
        With cboNewEnvironments
            .ControlName = "cboNewEnvironments"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = lblNewRelation.Top + lblNewRelation.Height + 5
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 200
        End With
        Me.AddChild(CType(cboNewEnvironments, UIControl))

        'cboNewPlayers initial props
        cboNewPlayers = New UIComboBox(oUILib)
        With cboNewPlayers
            .ControlName = "cboNewPlayers"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = cboNewEnvironments.Top + cboNewEnvironments.Height + 5
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 200
        End With
        Me.AddChild(CType(cboNewPlayers, UIControl))

        'btnNewRelation initial props
        btnNewRelation = New UIButton(oUILib)
        With btnNewRelation
            .ControlName = "btnNewRelation"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = cboNewPlayers.Top + cboNewPlayers.Height + 5
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Add A Relation"
            .ToolTipText = "Click here to add a relation for an environment you do not currenty have units in."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnNewRelation, UIControl))

        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = tvwRelations.Left + tvwRelations.Width
            .Top = btnNewRelation.Top + btnNewRelation.Height + 1
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))
        ''''' End New Relations '''''

        ''''' Existing Players '''''
        'lblExistingPlayers initial props
        lblExistingPlayers = New UILabel(oUILib)
        With lblExistingPlayers
            .ControlName = "lblExistingPlayers"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = lnDiv3.Top + 10
            .Width = 300
            .Height = 44
            .Enabled = True
            .Visible = True
            .Caption = "Alter Current Relations" & vbCrLf & " Or entire players" & vbCrLf & " Or entire environments"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblExistingPlayers, UIControl))

        'cboExistingPlayers initial props
        cboExistingPlayers = New UIComboBox(oUILib)
        With cboExistingPlayers
            .ControlName = "cboExistingPlayers"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = lblExistingPlayers.Top + lblExistingPlayers.Height + 5
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 200
        End With
        Me.AddChild(CType(cboExistingPlayers, UIControl))

        'cboExistingEnviromnets initial props
        cboExistingEnvironments = New UIComboBox(oUILib)
        With cboExistingEnvironments
            .ControlName = "cboExistingEnvironments"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = cboExistingPlayers.Top + cboExistingPlayers.Height + 5
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 200
        End With
        Me.AddChild(CType(cboExistingEnvironments, UIControl))

        'btnSetFriendly initial props
        btnSetFriendly = New UIButton(oUILib)
        With btnSetFriendly
            .ControlName = "btnSetFriendly"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = cboExistingEnvironments.Top + cboExistingEnvironments.Height + 5
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Set Friendly"
            '.ToolTipText = "Click here to add a relation for an environment you do not currenty have units in."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSetFriendly, UIControl))

        'btnSetEnemy initial props
        btnSetEnemy = New UIButton(oUILib)
        With btnSetEnemy
            .ControlName = "btnSetEnemy"
            .Left = btnSetFriendly.Left + btnSetFriendly.Width + 1
            .Top = btnSetFriendly.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Set Enemy"
            '.ToolTipText = "Click here to add a relation for an environment you do not currenty have units in."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSetEnemy, UIControl))

        'lnDiv4 initial props
        lnDiv4 = New UILine(oUILib)
        With lnDiv4
            .ControlName = "lnDiv4"
            .Left = tvwRelations.Left + tvwRelations.Width
            .Top = btnSetEnemy.Top + btnSetEnemy.Height + 1
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv4, UIControl))
        ''''' End Existing Players '''''

        ''''' Guild-Wide '''''
        'lblGuilds initial props
        lblGuilds = New UILabel(oUILib)
        With lblGuilds
            .ControlName = "lblGuilds"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = lnDiv4.Top + 10
            .Width = 300
            .Height = 44
            .Enabled = True
            .Visible = True
            .Caption = "Alter current relations toward entire" & vbCrLf & " guilds."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblGuilds, UIControl))

        'cboGuilds initial props
        cboGuilds = New UIComboBox(oUILib)
        With cboGuilds
            .ControlName = "cboExistingPlayers"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = lblGuilds.Top + lblGuilds.Height + 5
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 200
        End With
        Me.AddChild(CType(cboGuilds, UIControl))

        'cboGuildsEnvironments initial props
        cboGuildsEnvironments = New UIComboBox(oUILib)
        With cboGuildsEnvironments
            .ControlName = "cboGuildsEnvironments"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = cboGuilds.Top + cboGuilds.Height + 5
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 200
        End With
        Me.AddChild(CType(cboGuildsEnvironments, UIControl))

        'btnSetFriendlyGuild initial props
        btnSetFriendlyGuild = New UIButton(oUILib)
        With btnSetFriendlyGuild
            .ControlName = "btnSetFriendlyGuild"
            .Left = tvwRelations.Left + tvwRelations.Width + 5
            .Top = cboGuildsEnvironments.Top + cboGuildsEnvironments.Height + 5
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Set Friendly"
            '.ToolTipText = "Click here to add a relation for an environment you do not currenty have units in."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSetFriendlyGuild, UIControl))

        'btnSetEnemyGuild initial props
        btnSetEnemyGuild = New UIButton(oUILib)
        With btnSetEnemyGuild
            .ControlName = "btnSetEnemyGuild"
            .Left = btnSetFriendlyGuild.Left + btnSetFriendlyGuild.Width + 1
            .Top = btnSetFriendlyGuild.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Set Enemy"
            '.ToolTipText = "Click here to add a relation for an environment you do not currenty have units in."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSetEnemyGuild, UIControl))

        'lnDiv5 initial props
        lnDiv5 = New UILine(oUILib)
        With lnDiv5
            .ControlName = "lnDiv5"
            .Left = tvwRelations.Left + tvwRelations.Width
            .Top = btnSetEnemyGuild.Top + btnSetEnemyGuild.Height + 1
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv5, UIControl))
        ''''' End Existing Guild-Wide '''''


        rcByEnvironment = New Rectangle(btnClose.Left - 222, btnClose.Top, 120, btnClose.Height)
        rcByPlayer = New Rectangle(btnClose.Left - 102, btnClose.Top, 100, btnClose.Height)


        myViewState = 0

        'Now, add me...
        oUILib.RemoveWindow(Me.ControlName)
        oUILib.AddWindow(Me)
        mbLoading = False

        FillPlayerList()
        FillEnvironmentList()

        'if we dont have a transport array - request it
        If mlEnvirRelationsUB = -1 Then
            Dim yMsg(10) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eEnvirRelations).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(-1).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(-1).CopyTo(yMsg, 4)
            System.BitConverter.GetBytes(-1).CopyTo(yMsg, 6)
            System.BitConverter.GetBytes(-1).CopyTo(yMsg, 7)
            'goUILib.SendMsgToPrimary(yMsg)

            'Temp: fake load
            LoadRelations()
        Else
            LoadRelations()
        End If

    End Sub

    Private Sub frmEnvirRelations_OnNewFrame() Handles Me.OnNewFrame
        If mlLastUpdate = -1 Then mlLastUpdate = glCurrentCycle
        If mbHasUnknowns = True AndAlso glCurrentCycle > mlLastUpdate + 30 Then
            mbHasUnknowns = False
            mlLastUpdate = glCurrentCycle
            FillPlayerList()
            FillEnvironmentList()
            LoadRelations()
        End If
    End Sub

    Private Sub frmEnvirRelations_OnRenderEnd() Handles Me.OnRenderEnd
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

        Using oSprite As New Sprite(MyBase.moUILib.oDevice)
            Dim oSelColor As System.Drawing.Color
            oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)
            oSprite.Begin(SpriteFlags.AlphaBlend)
            If myViewState = 0 Then
                DoMultiColorFill(rcByEnvironment, oSelColor, rcByEnvironment.Location, oSprite)
            Else
                DoMultiColorFill(rcByPlayer, oSelColor, rcByPlayer.Location, oSprite)
            End If
            oSprite.End()
            oSprite.Dispose()
        End Using

        Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                oTextSpr.Begin(SpriteFlags.AlphaBlend)
                MyBase.RenderRoundedBorder(rcByEnvironment, 1, muSettings.InterfaceBorderColor)
                MyBase.RenderRoundedBorder(rcByPlayer, 1, muSettings.InterfaceBorderColor)

                Try
                    oFont.DrawText(oTextSpr, "By Environment", rcByEnvironment, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                    oFont.DrawText(oTextSpr, "By Player", rcByPlayer, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                Catch
                End Try
                oTextSpr.End()
                oTextSpr.Dispose()
            End Using
            oFont.Dispose()
        End Using

    End Sub

    Private Sub frmEnvirRelations_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        If mbLoading = True Then Exit Sub
        Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
        Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y

        If rcByEnvironment.Contains(lX, lY) Then
            myViewState = 0
        ElseIf rcByPlayer.Contains(lX, lY) Then
            myViewState = 1
        Else
            Exit Sub
        End If
        goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUnitSpeech, Nothing, Nothing)
        Me.IsDirty = True
        LoadRelations()
    End Sub

    Private Sub frmEnvirRelations_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.EnvirRelationsX = Me.Left
            muSettings.EnvirRelationsY = Me.Top
        End If
    End Sub

    Private Sub optSortByPlayer_Click() Handles optSortByPlayer.Click
        If mbLoading = True Then Exit Sub
        optSortByRelations.Value = False
        optSortByPlayer.Value = True
        mySortBy = 0
        LoadRelations()
    End Sub

    Private Sub optSortByRelations_Click() Handles optSortByRelations.Click
        If mbLoading = True Then Exit Sub
        optSortByPlayer.Value = False
        optSortByRelations.Value = True
        mySortBy = 1
        LoadRelations()
    End Sub

    Private Sub DoMultiColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point, ByRef oSpr As Sprite)
        Dim rcSrc As Rectangle

        Dim fX As Single
        Dim fY As Single

        If rcDest.Width = 0 OrElse rcDest.Height = 0 Then Exit Sub

        rcSrc.Location = New Point(192, 0)
        rcSrc.Width = 62
        rcSrc.Height = 64

        'Now, draw it...
        With oSpr
            fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
            fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))
            .Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFill)
        End With
    End Sub

    Public Sub HandleAddEnvirRelations(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim oRelation As New TEnvirRelations()
        With oRelation
            .lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .iRelation = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        End With

        ReDim Preserve moEnvirRelations(mlEnvirRelationsUB + 1)
        moEnvirRelations(mlEnvirRelationsUB + 1) = oRelation
        mlEnvirRelationsUB += 1
        LoadRelations()
    End Sub

    Private Sub LoadRelations()
        If goGalaxy Is Nothing Then Return
        If tvwRelations.TotalNodeCount > 0 Then SaveListExpansion()
        tvwRelations.Clear()
        If myViewState = 0 Then
            LoadRelationsByEnvironment()
        Else
            LoadRelationsByPlayer()
        End If
    End Sub

    Private Sub LoadRelationsByEnvironment()
        Dim lSortedEnv() As Int32 = Nothing
        Dim lSortedEnvUB As Int32 = -1
        Dim bTutorial As Boolean = NewTutorialManager.TutorialOn
        For X As Int32 = 0 To goGalaxy.mlSystemUB

            If goGalaxy.moSystems(X).ObjTypeID = ObjectType.eSolarSystem AndAlso ((bTutorial = False AndAlso goGalaxy.moSystems(X).ObjectID <> 36 AndAlso goGalaxy.moSystems(X).ObjectID <> 88) OrElse (bTutorial = True AndAlso goGalaxy.moSystems(X).ObjectID = 36)) Then
                Dim lIdx As Int32 = -1
                If EnvironmentHasARelation(goGalaxy.moSystems(X).ObjectID) = False Then Continue For
                Dim sName As String = goGalaxy.moSystems(X).SystemName

                For Y As Int32 = 0 To lSortedEnvUB
                    Dim sOtherName As String = goGalaxy.moSystems(lSortedEnv(Y)).SystemName
                    If sOtherName > sName Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedEnvUB += 1
                ReDim Preserve lSortedEnv(lSortedEnvUB)
                If lIdx = -1 Then
                    lSortedEnv(lSortedEnvUB) = X
                Else
                    For Y As Int32 = lSortedEnvUB To lIdx + 1 Step -1
                        lSortedEnv(Y) = lSortedEnv(Y - 1)
                    Next Y
                    lSortedEnv(lIdx) = X
                End If
            End If
        Next X

        If goCurrentPlayer Is Nothing Then Exit Sub

        Dim lSortedPlayer() As Int32 = Nothing
        Dim sSortedPlayerVal() As String = Nothing
        Dim lSortedPlayerUB As Int32 = -1

        For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
            Dim lIdx As Int32 = -1
            Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(X)
            If oRel Is Nothing = False Then
                Dim lPID As Int32 = oRel.lThisPlayer
                If lPID = goCurrentPlayer.ObjectID Then lPID = oRel.lPlayerRegards
                Dim sName As String = GetCacheObjectValue(lPID, ObjectType.ePlayer)
                If sName = "Unknown" Then mbHasUnknowns = True
                For Y As Int32 = 0 To lSortedPlayerUB
                    If sSortedPlayerVal(Y).ToUpper > sName.ToUpper Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedPlayerUB += 1
                ReDim Preserve lSortedPlayer(lSortedPlayerUB)
                ReDim Preserve sSortedPlayerVal(lSortedPlayerUB)
                If lIdx = -1 Then
                    sSortedPlayerVal(lSortedPlayerUB) = sName
                    lSortedPlayer(lSortedPlayerUB) = oRel.lThisPlayer
                Else
                    For Y As Int32 = lSortedPlayerUB To lIdx + 1 Step -1
                        lSortedPlayer(Y) = lSortedPlayer(Y - 1)
                        sSortedPlayerVal(Y) = sSortedPlayerVal(Y - 1)
                    Next Y
                    sSortedPlayerVal(lIdx) = sName
                    lSortedPlayer(lIdx) = oRel.lThisPlayer
                End If
            End If
        Next X

        For X As Int32 = 0 To lSortedEnvUB

            With goGalaxy.moSystems(lSortedEnv(X))
                Dim sText As String = .SystemName
                Dim oSystemNode As UITreeView.UITreeViewItem = tvwRelations.AddNode(.SystemName, .ObjectID, -1, -1, Nothing, Nothing)
                If bTutorial = True OrElse WasExpanded(oSystemNode) = True Then oSystemNode.bExpanded = True
                oSystemNode.bItemBold = True
                If mySortBy = 1 Then 'By Player name
                    For y As Int32 = 0 To lSortedPlayerUB

                        If GetRelation(lSortedPlayer(y), .ObjectID) = 1 Then
                            Dim oPlayerNode As UITreeView.UITreeViewItem = tvwRelations.AddNode(sSortedPlayerVal(y), lSortedPlayer(y), -1, -1, oSystemNode, Nothing)
                            oPlayerNode.clrItemColor = Color.Green
                            oPlayerNode.bUseItemColor = True
                        End If
                    Next y
                End If
                For y As Int32 = 0 To lSortedPlayerUB
                    Dim PlayerRel As Int16 = GetRelation(lSortedPlayer(y), .ObjectID)
                    If mySortBy = 0 OrElse PlayerRel = 0 Then
                        Dim oPlayerNode As UITreeView.UITreeViewItem = tvwRelations.AddNode(sSortedPlayerVal(y), lSortedPlayer(y), -1, -1, oSystemNode, Nothing)
                        If PlayerRel = 0 Then
                            oPlayerNode.clrItemColor = Color.Red
                        Else
                            oPlayerNode.clrItemColor = Color.Green
                        End If
                        oPlayerNode.bUseItemColor = True
                    End If
                Next y
            End With
        Next X
    End Sub

    Private Function GetRelation(ByVal lPlayerID As Int32, ByVal lEnvirID As Int32) As Int16
        For x As Int32 = 0 To mlEnvirRelationsUB
            If moEnvirRelations(x).lPlayerID = lPlayerID AndAlso moEnvirRelations(x).lEnvirID = lEnvirID Then
                Return moEnvirRelations(x).iRelation
            End If
        Next
        Return 0
    End Function

    Private Function EnvironmentHasARelation(ByVal lEnvirID As Int32) As Boolean
        Select Case lEnvirID
            Case 36, 88, 79, 81, 100
                Return True
        End Select

        For x As Int32 = 0 To mlEnvirRelationsUB
            If moEnvirRelations(x).lEnvirID = lEnvirID AndAlso moEnvirRelations(x).iRelation = 1 Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function PlayerHasARelation(ByVal lPlayerID As Int32) As Boolean
        Select Case lPlayerID
            Case 3510, 20611, 19466
                Return True
        End Select

        For x As Int32 = 0 To mlEnvirRelationsUB
            If moEnvirRelations(x).lPlayerID = lPlayerID AndAlso moEnvirRelations(x).iRelation = 1 Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Sub LoadRelationsForEnvironment()

    End Sub

    Private Sub LoadRelationsByPlayer()
        Dim lSortedEnv() As Int32 = Nothing
        Dim lSortedEnvUB As Int32 = -1
        Dim bTutorial As Boolean = NewTutorialManager.TutorialOn
        For X As Int32 = 0 To goGalaxy.mlSystemUB

            If goGalaxy.moSystems(X).ObjTypeID = ObjectType.eSolarSystem AndAlso ((bTutorial = False AndAlso goGalaxy.moSystems(X).ObjectID <> 36 AndAlso goGalaxy.moSystems(X).ObjectID <> 88) OrElse (bTutorial = True AndAlso goGalaxy.moSystems(X).ObjectID = 36)) Then
                Dim lIdx As Int32 = -1
                Dim sName As String = goGalaxy.moSystems(X).SystemName

                For Y As Int32 = 0 To lSortedEnvUB
                    Dim sOtherName As String = goGalaxy.moSystems(lSortedEnv(Y)).SystemName
                    If sOtherName > sName Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedEnvUB += 1
                ReDim Preserve lSortedEnv(lSortedEnvUB)
                If lIdx = -1 Then
                    lSortedEnv(lSortedEnvUB) = X
                Else
                    For Y As Int32 = lSortedEnvUB To lIdx + 1 Step -1
                        lSortedEnv(Y) = lSortedEnv(Y - 1)
                    Next Y
                    lSortedEnv(lIdx) = X
                End If
            End If
        Next X

        If goCurrentPlayer Is Nothing Then Exit Sub

        Dim lSortedPlayer() As Int32 = Nothing
        Dim sSortedPlayerVal() As String = Nothing
        Dim lSortedPlayerUB As Int32 = -1

        For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
            Dim lIdx As Int32 = -1
            Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(X)
            If oRel Is Nothing = False Then
                Dim lPID As Int32 = oRel.lThisPlayer
                If lPID = goCurrentPlayer.ObjectID Then lPID = oRel.lPlayerRegards
                If PlayerHasARelation(lPID) = False Then Continue For

                Dim sName As String = GetCacheObjectValue(lPID, ObjectType.ePlayer)
                If sName = "Unknown" Then mbHasUnknowns = True
                For Y As Int32 = 0 To lSortedPlayerUB
                    If sSortedPlayerVal(Y).ToUpper > sName.ToUpper Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedPlayerUB += 1
                ReDim Preserve lSortedPlayer(lSortedPlayerUB)
                ReDim Preserve sSortedPlayerVal(lSortedPlayerUB)
                If lIdx = -1 Then
                    sSortedPlayerVal(lSortedPlayerUB) = sName
                    lSortedPlayer(lSortedPlayerUB) = oRel.lThisPlayer
                Else
                    For Y As Int32 = lSortedPlayerUB To lIdx + 1 Step -1
                        lSortedPlayer(Y) = lSortedPlayer(Y - 1)
                        sSortedPlayerVal(Y) = sSortedPlayerVal(Y - 1)
                    Next Y
                    sSortedPlayerVal(lIdx) = sName
                    lSortedPlayer(lIdx) = oRel.lThisPlayer
                End If
            End If
        Next X

        For y As Int32 = 0 To lSortedPlayerUB
            Dim oPlayerNode As UITreeView.UITreeViewItem = tvwRelations.AddNode(sSortedPlayerVal(y), lSortedPlayer(y), -1, -1, Nothing, Nothing)
            For X As Int32 = 0 To lSortedEnvUB
                With goGalaxy.moSystems(lSortedEnv(X))

                    Dim PlayerRel As Int16 = GetRelation(lSortedPlayer(y), .ObjectID)
                    Dim sText As String = .SystemName
                    Dim oSystemNode As UITreeView.UITreeViewItem = tvwRelations.AddNode(.SystemName, .ObjectID, -1, -1, oPlayerNode, Nothing)
                    If PlayerRel = 0 Then
                        oPlayerNode.clrItemColor = Color.Red
                    Else
                        oPlayerNode.clrItemColor = Color.Green
                    End If
                    oPlayerNode.bUseItemColor = True

                End With
            Next X
        Next y
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        If tvwRelations.TotalNodeCount > 0 Then SaveListExpansion()
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub FillPlayerList()
        cboNewPlayers.Clear()
        cboExistingPlayers.Clear()
        cboExistingPlayers.AddItem("*** All ***")
        cboExistingPlayers.ItemData(cboExistingPlayers.NewIndex) = -1

        Dim lSorted() As Int32 = GetSortedPlayerRelIdxArray()
        If lSorted Is Nothing = False Then
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                Dim oTmpRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lSorted(X))
                If oTmpRel Is Nothing = False Then
                    Dim lPlayerID As Int32 = -1
                    If oTmpRel.lPlayerRegards = glPlayerID Then
                        lPlayerID = oTmpRel.lThisPlayer
                    Else : lPlayerID = oTmpRel.lPlayerRegards
                    End If

                    If lPlayerID <> -1 Then
                        Dim sName As String = GetCacheObjectValue(lPlayerID, ObjectType.ePlayer)
                        If sName = "Unknown" Then mbHasUnknowns = True
                        cboNewPlayers.AddItem(sName)
                        cboNewPlayers.ItemData(cboNewPlayers.NewIndex) = lPlayerID
                        cboExistingPlayers.AddItem(sName)
                        cboExistingPlayers.ItemData(cboExistingPlayers.NewIndex) = lPlayerID
                        CheckAddMoGuild(oTmpRel.oPlayerIntel.lGuildID)
                    End If
                End If
            Next X
        End If
        cboGuilds.AddItem("*** None ***")
        cboGuilds.ItemData(cboExistingPlayers.NewIndex) = -1
        For x As Int32 = 0 To moGuildsUB
            cboGuilds.AddItem(moGuilds(x).sName)
            cboGuilds.ItemData(cboGuilds.NewIndex) = moGuilds(x).ObjectID
        Next
        cboExistingPlayers.ListIndex = 0
        cboGuilds.ListIndex = 0
    End Sub

    Private Sub CheckAddMoGuild(ByVal lGuildID As Int32)
        For x As Int32 = 0 To moGuildsUB
            If moGuilds(x).ObjectID = lGuildID Then
                Exit Sub
            End If
        Next
        moGuildsUB += 1

        ReDim Preserve moGuilds(moGuildsUB)
        With moGuilds(moGuildsUB)
            .ObjectID = lGuildID
            .sName = GetCacheObjectValue(lGuildID, ObjectType.eGuild)
            If .sName = "Unknown" Then mbHasUnknowns = True
        End With
    End Sub

    Private Sub FillEnvironmentList()
        If goGalaxy Is Nothing Then Return
        cboNewEnvironments.Clear()
        cboExistingEnvironments.Clear()
        cboGuildsEnvironments.Clear()
        Dim lSortedEnv() As Int32 = Nothing
        Dim lSortedEnvUB As Int32 = -1
        Dim bTutorial As Boolean = NewTutorialManager.TutorialOn

        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X).ObjTypeID = ObjectType.eSolarSystem AndAlso ((bTutorial = False AndAlso goGalaxy.moSystems(X).ObjectID <> 36 AndAlso goGalaxy.moSystems(X).ObjectID <> 88) OrElse (bTutorial = True AndAlso goGalaxy.moSystems(X).ObjectID = 36)) Then
                Dim lIdx As Int32 = -1
                If EnvironmentHasARelation(goGalaxy.moSystems(X).ObjectID) = False Then Continue For
                Dim sName As String = goGalaxy.moSystems(X).SystemName

                For Y As Int32 = 0 To lSortedEnvUB
                    Dim sOtherName As String = goGalaxy.moSystems(lSortedEnv(Y)).SystemName
                    If sOtherName > sName Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedEnvUB += 1
                ReDim Preserve lSortedEnv(lSortedEnvUB)
                If lIdx = -1 Then
                    lSortedEnv(lSortedEnvUB) = X
                Else
                    For Y As Int32 = lSortedEnvUB To lIdx + 1 Step -1
                        lSortedEnv(Y) = lSortedEnv(Y - 1)
                    Next Y
                    lSortedEnv(lIdx) = X
                End If
            End If
        Next X

        For X As Int32 = 0 To lSortedEnvUB
            With goGalaxy.moSystems(lSortedEnv(X))
                'Force cache to load
                'SetCacheObjectValue(.ObjectID, ObjectType.eSolarSystem, .SystemName)
                Dim sSystemName As String = GetCacheObjectValue(.ObjectID, ObjectType.eSolarSystem)
                If sSystemName = "Unknown" Then mbHasUnknowns = True
                cboNewEnvironments.AddItem(.SystemName)
                cboNewEnvironments.ItemData(cboNewEnvironments.NewIndex) = .ObjectID
            End With
        Next X

        cboExistingEnvironments.AddItem("*** All ***")
        cboExistingEnvironments.ItemData(cboExistingEnvironments.NewIndex) = -1
        For X As Int32 = 0 To lSortedEnvUB
            With goGalaxy.moSystems(lSortedEnv(X))
                'Force cache to load
                'SetCacheObjectValue(.ObjectID, ObjectType.eSolarSystem, .SystemName)
                Dim sSystemName As String = GetCacheObjectValue(.ObjectID, ObjectType.eSolarSystem)
                If sSystemName = "Unknown" Then mbHasUnknowns = True
                cboExistingEnvironments.AddItem(.SystemName)
                cboExistingEnvironments.ItemData(cboExistingEnvironments.NewIndex) = .ObjectID
            End With
        Next X

        cboGuildsEnvironments.AddItem("*** All ***")
        cboGuildsEnvironments.ItemData(cboGuildsEnvironments.NewIndex) = -1
        For X As Int32 = 0 To lSortedEnvUB
            With goGalaxy.moSystems(lSortedEnv(X))
                'Force cache to load
                'SetCacheObjectValue(.ObjectID, ObjectType.eSolarSystem, .SystemName)
                Dim sSystemName As String = GetCacheObjectValue(.ObjectID, ObjectType.eSolarSystem)
                If sSystemName = "Unknown" Then mbHasUnknowns = True
                cboGuildsEnvironments.AddItem(.SystemName)
                cboGuildsEnvironments.ItemData(cboGuildsEnvironments.NewIndex) = .ObjectID
            End With
        Next X
        cboExistingEnvironments.ListIndex = 0
        cboGuildsEnvironments.ListIndex = 0
    End Sub

    Private Sub SaveListExpansion()
        ReDim mbExp(-1)
        Dim lIdx As Int32 = 0
        Dim lListCount As Int32 = tvwRelations.TotalNodeCount
        Dim oActNode As UITreeView.UITreeViewItem = tvwRelations.GetNodeByItemText("", False, False, lIdx - 1)

        While oActNode Is Nothing = False AndAlso lIdx < lListCount
            If oActNode.bExpanded = True Then
                ReDim Preserve mbExp(mbExp.GetUpperBound(0) + 1)
                mbExp(mbExp.GetUpperBound(0)) = oActNode.lItemData & "." & oActNode.lItemData2 & "." & oActNode.lItemData3
            End If
            lIdx += 1
            oActNode = tvwRelations.GetNodeByItemText("", False, False, lIdx - 1)
        End While
    End Sub

    Private Function WasExpanded(ByVal oNode As UITreeView.UITreeViewItem) As Boolean
        If mbExp Is Nothing = True OrElse mbExp.GetUpperBound(0) = -1 Then Exit Function
        For x As Int32 = 0 To mbExp.GetUpperBound(0)
            If mbExp(x) = oNode.lItemData & "." & oNode.lItemData2 & "." & oNode.lItemData3 Then
                Return True
            End If
        Next
    End Function

    Private Sub btnNewRelation_Click(ByVal sName As String) Handles btnNewRelation.Click
        If cboNewEnvironments.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select an environment!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        If cboNewPlayers.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select a player!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        Dim lOtherEnviromnetID As Int32 = cboNewEnvironments.ItemData(cboNewEnvironments.ListIndex)
        Dim lOtherPlayerID As Int32 = cboNewPlayers.ItemData(cboNewPlayers.ListIndex)
        Dim sSystemName As String = GetCacheObjectValue(lOtherEnviromnetID, ObjectType.eSolarSystem)
        Dim sPlayerName As String = GetCacheObjectValue(lOtherPlayerID, ObjectType.ePlayer)

        Dim oFrm As New frmMsgBox(goUILib, "Are you sure you wish to add friendly relations with player [" & sPlayerName & "] in the environment [" & sSystemName & "]?", MsgBoxStyle.YesNo, "Confirm Relation")
        oFrm.Visible = True
        AddHandler oFrm.DialogClosed, AddressOf ConfirmAddRelations
    End Sub

    Private Sub btnSetFriendly_Click(ByVal sName As String) Handles btnSetFriendly.Click
        If cboExistingEnvironments.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select an environment!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        If cboExistingPlayers.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select a player!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        If cboExistingEnvironments.ListIndex = 0 AndAlso cboExistingPlayers.ListIndex = 0 Then
            MyBase.moUILib.AddNotification("You cannot set relations for all players in all places at this time.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        Dim lOtherEnviromnetID As Int32 = cboExistingEnvironments.ItemData(cboExistingEnvironments.ListIndex)
        Dim lOtherPlayerID As Int32 = cboExistingPlayers.ItemData(cboExistingPlayers.ListIndex)
        Dim sSystemName As String = "*** All ***"
        Dim sPlayerName As String = "*** All ***"
        If lOtherEnviromnetID <> -1 Then GetCacheObjectValue(lOtherEnviromnetID, ObjectType.eSolarSystem)
        If lOtherPlayerID <> -1 Then sPlayerName = GetCacheObjectValue(lOtherPlayerID, ObjectType.ePlayer)


        Dim oFrm As New frmMsgBox(goUILib, "Are you sure you wish to set friendly relations with player [" & sPlayerName & "] in the environment [" & sSystemName & "]?", MsgBoxStyle.YesNo, "Confirm Relation")
        oFrm.Visible = True
        AddHandler oFrm.DialogClosed, AddressOf ConfirmSetFriendly
    End Sub

    Private Sub btnSetEnemy_Click(ByVal sName As String) Handles btnSetEnemy.Click
        If cboExistingEnvironments.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select an environment!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        If cboExistingPlayers.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select a player!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        If cboExistingEnvironments.ListIndex = 0 AndAlso cboExistingPlayers.ListIndex = 0 Then
            MyBase.moUILib.AddNotification("You cannot set relations for all players in all places at this time.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        Dim lOtherEnviromnetID As Int32 = cboExistingEnvironments.ItemData(cboExistingEnvironments.ListIndex)
        Dim lOtherPlayerID As Int32 = cboExistingPlayers.ItemData(cboExistingPlayers.ListIndex)
        Dim sSystemName As String = "*** All ***"
        Dim sPlayerName As String = "*** All ***"
        If lOtherEnviromnetID <> -1 Then GetCacheObjectValue(lOtherEnviromnetID, ObjectType.eSolarSystem)
        If lOtherPlayerID <> -1 Then sPlayerName = GetCacheObjectValue(lOtherPlayerID, ObjectType.ePlayer)

        Dim oFrm As New frmMsgBox(goUILib, "Are you sure you wish to set enemy relations with player [" & sPlayerName & "] in the environment [" & sSystemName & "]?", MsgBoxStyle.YesNo, "Confirm Relation")
        oFrm.Visible = True
        AddHandler oFrm.DialogClosed, AddressOf ConfirmSetEnemy
    End Sub

    Private Sub btnSetFriendlyGuild_Click(ByVal sName As String) Handles btnSetFriendlyGuild.Click
        If cboGuilds.ListIndex <= 0 Then
            MyBase.moUILib.AddNotification("You must select a guild!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        If cboGuildsEnvironments.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select an environment!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        Dim lOtherEnviromnetID As Int32 = cboGuildsEnvironments.ItemData(cboGuildsEnvironments.ListIndex)
        Dim lOtherGuildID As Int32 = cboGuilds.ItemData(cboGuilds.ListIndex)
        Dim sSystemName As String = "*** All ***"
        If lOtherEnviromnetID <> -1 Then sSystemName = GetCacheObjectValue(lOtherEnviromnetID, ObjectType.eSolarSystem)
        Dim sGuildName As String = GetCacheObjectValue(lOtherGuildID, ObjectType.eGuild)

        Dim oFrm As New frmMsgBox(goUILib, "Are you sure you wish to set friendly relations with the entire guild [" & sGuildName & "] in the environment [" & sSystemName & "]?", MsgBoxStyle.YesNo, "Confirm Relation")
        oFrm.Visible = True
        AddHandler oFrm.DialogClosed, AddressOf ConfirmRelationsSetFriendlyGuild
    End Sub

    Private Sub btnSetEnemyGuild_Click(ByVal sName As String) Handles btnSetEnemyGuild.Click
        If cboGuilds.ListIndex <= 0 Then
            MyBase.moUILib.AddNotification("You must select an environment!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        If cboGuildsEnvironments.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select a player!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        Dim lOtherEnviromnetID As Int32 = cboGuildsEnvironments.ItemData(cboGuildsEnvironments.ListIndex)
        Dim lOtherGuildID As Int32 = cboGuilds.ItemData(cboGuilds.ListIndex)
        Dim sSystemName As String = "*** All ***"
        If lOtherEnviromnetID <> -1 Then sSystemName = GetCacheObjectValue(lOtherEnviromnetID, ObjectType.eSolarSystem)
        Dim sGuildName As String = GetCacheObjectValue(lOtherGuildID, ObjectType.eGuild)

        Dim oFrm As New frmMsgBox(goUILib, "Are you sure you wish to set enemy relations with the entire guild [" & sGuildName & "] in the environment [" & sSystemName & "]?", MsgBoxStyle.YesNo, "Confirm Relation")
        oFrm.Visible = True
        AddHandler oFrm.DialogClosed, AddressOf ConfirmRelationsSetEnemyGuild
    End Sub

    Private Sub ConfirmAddRelations(ByVal lResult As MsgBoxResult)
        If lResult <> MsgBoxResult.Yes Then Exit Sub
        If cboNewEnvironments.ListIndex = -1 Then Return
        If cboNewPlayers.ListIndex = -1 Then Return
        Dim lOtherEnviromnetID As Int32 = cboNewEnvironments.ItemData(cboNewEnvironments.ListIndex)
        Dim lOtherPlayerID As Int32 = cboNewPlayers.ItemData(cboNewPlayers.ListIndex)

        UpdateRelations(lOtherEnviromnetID, lOtherPlayerID, 1)
    End Sub

    Private Sub ConfirmSetFriendly(ByVal lResult As MsgBoxResult)
        If lResult <> MsgBoxResult.Yes Then Exit Sub
        If cboExistingEnvironments.ListIndex = -1 Then Return
        If cboExistingPlayers.ListIndex = -1 Then Return
        Dim lOtherEnviromnetID As Int32 = cboExistingEnvironments.ItemData(cboExistingEnvironments.ListIndex)
        Dim lOtherPlayerID As Int32 = cboExistingPlayers.ItemData(cboExistingPlayers.ListIndex)

        If lOtherEnviromnetID > 0 Then
            UpdateRelations(lOtherEnviromnetID, lOtherPlayerID, 1)
        Else
            For y As Int32 = 0 To goGalaxy.mlSystemUB
                If goGalaxy.moSystems(y).ObjTypeID = ObjectType.eSolarSystem Then
                    UpdateRelations(goGalaxy.moSystems(y).ObjectID, lOtherPlayerID, 1)
                End If
            Next
        End If
    End Sub
    Private Sub ConfirmSetEnemy(ByVal lResult As MsgBoxResult)
        If lResult <> MsgBoxResult.Yes Then Exit Sub
        If cboExistingEnvironments.ListIndex = -1 Then Return
        If cboExistingPlayers.ListIndex = -1 Then Return
        Dim lOtherEnviromnetID As Int32 = cboExistingEnvironments.ItemData(cboExistingEnvironments.ListIndex)
        Dim lOtherPlayerID As Int32 = cboExistingPlayers.ItemData(cboExistingPlayers.ListIndex)

        If lOtherEnviromnetID > 0 Then
            UpdateRelations(lOtherEnviromnetID, lOtherPlayerID, 1)
        Else
            For y As Int32 = 0 To goGalaxy.mlSystemUB
                If goGalaxy.moSystems(y).ObjTypeID = ObjectType.eSolarSystem Then
                    UpdateRelations(goGalaxy.moSystems(y).ObjectID, lOtherPlayerID, 0)
                End If
            Next
        End If
    End Sub

    Private Sub ConfirmRelationsSetFriendlyGuild(ByVal lResult As MsgBoxResult)
        If lResult <> MsgBoxResult.Yes Then Exit Sub
        If cboGuildsEnvironments.ListIndex <= 0 Then Return
        If cboGuilds.ListIndex = -1 Then Return
        Dim lOtherEnviromnetID As Int32 = cboGuildsEnvironments.ItemData(cboGuildsEnvironments.ListIndex)
        Dim lOtherGuildID As Int32 = cboGuilds.ItemData(cboGuilds.ListIndex)

        Dim lSorted() As Int32 = GetSortedPlayerRelIdxArray()
        If lSorted Is Nothing = False Then
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                Dim oTmpRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lSorted(X))
                If oTmpRel Is Nothing = False Then
                    Dim lPlayerID As Int32 = -1
                    If oTmpRel.lPlayerRegards = glPlayerID Then
                        lPlayerID = oTmpRel.lThisPlayer
                    Else : lPlayerID = oTmpRel.lPlayerRegards
                    End If

                    If lPlayerID <> -1 Then
                        If lOtherEnviromnetID > 0 Then
                            UpdateRelations(lOtherEnviromnetID, lPlayerID, 1)
                        Else
                            For y As Int32 = 0 To goGalaxy.mlSystemUB
                                If goGalaxy.moSystems(y).ObjTypeID = ObjectType.eSolarSystem Then
                                    UpdateRelations(goGalaxy.moSystems(y).ObjectID, lPlayerID, 1)
                                End If
                            Next
                        End If
                    End If
                End If
            Next X
        End If
    End Sub

    Private Sub ConfirmRelationsSetEnemyGuild(ByVal lResult As MsgBoxResult)
        If lResult <> MsgBoxResult.Yes Then Exit Sub
        If cboGuildsEnvironments.ListIndex <= 0 Then Return
        If cboGuilds.ListIndex = -1 Then Return
        Dim lOtherEnviromnetID As Int32 = cboGuildsEnvironments.ItemData(cboGuildsEnvironments.ListIndex)
        Dim lOtherGuildID As Int32 = cboGuilds.ItemData(cboGuilds.ListIndex)

        Dim lSorted() As Int32 = GetSortedPlayerRelIdxArray()
        If lSorted Is Nothing = False Then
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                Dim oTmpRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lSorted(X))
                If oTmpRel Is Nothing = False Then
                    Dim lPlayerID As Int32 = -1
                    If oTmpRel.lPlayerRegards = glPlayerID Then
                        lPlayerID = oTmpRel.lThisPlayer
                    Else : lPlayerID = oTmpRel.lPlayerRegards
                    End If

                    If lPlayerID <> -1 Then
                        If lOtherEnviromnetID > 0 Then
                            UpdateRelations(lOtherEnviromnetID, lPlayerID, 1)
                        Else
                            For y As Int32 = 0 To goGalaxy.mlSystemUB
                                If goGalaxy.moSystems(y).ObjTypeID = ObjectType.eSolarSystem Then
                                    UpdateRelations(goGalaxy.moSystems(y).ObjectID, lPlayerID, 0)
                                End If
                            Next
                        End If
                    End If
                End If
            Next X
        End If
    End Sub

    Private Sub UpdateRelations(ByVal EnvirID As Int32, ByVal PlayerID As Int32, ByVal Relation As Int16)
        Dim bFound As Boolean = False
        '(check array)
        For x As Int32 = 0 To mlEnvirRelationsUB
            If moEnvirRelations(x).lPlayerID = PlayerID AndAlso moEnvirRelations(x).lEnvirID = EnvirID Then
                moEnvirRelations(x).iRelation = Relation
                bFound = True
                Exit For
            End If
        Next

        If bFound = False Then
            mlEnvirRelationsUB += 1
            ReDim Preserve moEnvirRelations(mlEnvirRelationsUB)
            With moEnvirRelations(mlEnvirRelationsUB)
                .lEnvirID = EnvirID
                .lPlayerID = PlayerID
                .iRelation = Relation
            End With
        End If

        Dim yMsg(10) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eEnvirRelations).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(1).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(EnvirID).CopyTo(yMsg, 4)
        System.BitConverter.GetBytes(PlayerID).CopyTo(yMsg, 6)
        System.BitConverter.GetBytes(Relation).CopyTo(yMsg, 7)
        'goUILib.SendMsgToPrimary(yMsg)

        LoadRelations()
    End Sub
End Class
