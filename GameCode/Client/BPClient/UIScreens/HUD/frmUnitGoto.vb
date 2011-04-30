Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D


Public Class frmUnitGoto
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents btnClose As UIButton
    Private lnDiv1 As UILine
    Private WithEvents tvwMain As UITreeView
    Private Shared miObjectID As Int32 = -1
    Private Shared miObjTypeID As Int16 = -1
    Private Shared miIndex As Int32 = -1
    Private mbExp() As String
    Private mbIgnoreSelection As Boolean = False

    Private Shared myAlertValue As Byte = 0       '0 for none, 1 for first, 2 for all
    Private WithEvents optAlertNone As UIOption
    Private WithEvents optAlertFirst As UIOption
    Private WithEvents optAlertAll As UIOption
    Private lnDiv2 As UILine

    Private lblNearNum As UILabel
    Private lblDist As UILabel
    Private txtDistance As UITextBox
    Private Shared miDist As Int32 = 100

    Private Shared myNearDirection As Byte = 5
    Private rcOptNear1 As Rectangle
    Private rcOptNear2 As Rectangle
    Private rcOptNear3 As Rectangle
    Private rcOptNear4 As Rectangle
    Private rcOptNear5 As Rectangle
    Private rcOptNear6 As Rectangle
    Private rcOptNear7 As Rectangle
    Private rcOptNear8 As Rectangle
    Private rcOptNear9 As Rectangle
    Private WithEvents btnGoNear As UIButton

    Private lnDiv3 As UILine
    Private lblGridNum As UILabel
    Private Shared myGridLocation As Byte = 5
    Private rcOptGrid1 As Rectangle
    Private rcOptGrid2 As Rectangle
    Private rcOptGrid3 As Rectangle
    Private rcOptGrid4 As Rectangle
    Private rcOptGrid5 As Rectangle
    Private rcOptGrid6 As Rectangle
    Private rcOptGrid7 As Rectangle
    Private rcOptGrid8 As Rectangle
    Private rcOptGrid9 As Rectangle
    Private WithEvents btnGoEnter As UIButton

    Private lnDiv4 As UILine
    Private WithEvents btnGoDirect As UIButton
    Private WithEvents btnAttackTargets As UIButton

    Private mbLoading As Boolean = True
    Private mbRequestingEnvironment As Boolean = False
    Private mlRequestingEnvironmentID As Int32 = -1
    Private mlLastCheck As Int32 = Int32.MinValue

    Private mbMultiDepth As Boolean = False 'Enable this to allow automatic route creation of remote locations :)

    Private Structure TMoveQueue
        Dim MoveType As Integer
        Dim locX As Integer
        Dim locY As Integer
        Dim locZ As Integer
        Dim lDestID As Integer
        Dim iDestTypeID As Integer
        Dim WormholeID As Integer
    End Structure

    Private MoveQueue() As TMoveQueue

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmUnitGoto initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTransfer
            .ControlName = "frmUnitGoto"
            If muSettings.UnitGotoLocX <> -1 Then .Left = muSettings.UnitGotoLocX
            If muSettings.UnitGotoLocY <> -1 Then .Top = muSettings.UnitGotoLocY
            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Width = 512
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
            .bRoundedBorder = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 170
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Unit Goto Manager"
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
            .Left = Me.BorderLineWidth
            .Top = 25
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'tvwMain initial props
        tvwMain = New UITreeView(oUILib)
        With tvwMain
            .ControlName = "tvwMain"
            .Left = 5
            .Top = lnDiv1.Top + 5
            .Width = 246
            .Height = Me.Height - lnDiv1.Top - 8
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(tvwMain, UIControl))

        'optAlertNone initial props
        optAlertNone = New UIOption(oUILib)
        With optAlertNone
            .ControlName = "optAlertNone"
            .Left = tvwMain.Left + tvwMain.Width + 5
            .Top = lnDiv1.Top + 5
            .Width = 78
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Alert None"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = (myAlertValue = 0)
        End With
        Me.AddChild(CType(optAlertNone, UIControl))

        'optAlertFirst initial props
        optAlertFirst = New UIOption(oUILib)
        With optAlertFirst
            .ControlName = "optAlertFirst"
            .Left = tvwMain.Left + tvwMain.Width + 5
            .Top = optAlertNone.Top + optAlertNone.Height + 1
            .Width = 95
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Alert First Unit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = (myAlertValue = 1)
            .ToolTipText = "If set, first selected unit will have their Alert On Arrival toggle set."
        End With
        Me.AddChild(CType(optAlertFirst, UIControl))

        'optAlertAll initial props
        optAlertAll = New UIOption(oUILib)
        With optAlertAll
            .ControlName = "optAlertAll"
            .Left = tvwMain.Left + tvwMain.Width + 5
            .Top = optAlertFirst.Top + optAlertFirst.Height + 1
            .Width = 92
            .Height = 18
            .Enabled = False
            .Visible = True
            .Caption = "Alert All Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = (myAlertValue = 2)
            .ToolTipText = "If set, all selected unit(s) will have their Alert On Arrival toggle set."
        End With
        Me.AddChild(CType(optAlertAll, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = tvwMain.Left + tvwMain.Width
            .Top = optAlertAll.Top + optAlertAll.Height + 2
            .Width = Me.Width - tvwMain.Left - tvwMain.Width
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'lblNum  initial props
        lblNearNum = New UILabel(oUILib)
        With lblNearNum
            .ControlName = "lblNearNum"
            .Left = tvwMain.Left + tvwMain.Width + 2
            .Top = lnDiv2.Top + 5
            .Width = 170
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Near Direction"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblNearNum, UIControl))

        rcOptNear7 = New Rectangle(tvwMain.Left + tvwMain.Width + 5, lblNearNum.Top + lblNearNum.Height + 2, 16, 16)
        rcOptNear8 = New Rectangle(rcOptNear7.X + 16 + 1, rcOptNear7.Y, 16, 16)
        rcOptNear9 = New Rectangle(rcOptNear8.X + 16 + 1, rcOptNear7.Y, 16, 16)

        rcOptNear4 = New Rectangle(rcOptNear7.X, rcOptNear7.Y + 16 + 1, 16, 16)
        rcOptNear5 = New Rectangle(rcOptNear8.X, rcOptNear4.Y, 16, 16)
        rcOptNear6 = New Rectangle(rcOptNear9.X, rcOptNear4.Y, 16, 16)

        rcOptNear1 = New Rectangle(rcOptNear7.X, rcOptNear4.Y + 16 + 1, 16, 16)
        rcOptNear2 = New Rectangle(rcOptNear8.X, rcOptNear1.Y, 16, 16)
        rcOptNear3 = New Rectangle(rcOptNear9.X, rcOptNear1.Y, 16, 16)
        rcOptNear5.X += 1

        'btnGoNear initial props
        btnGoNear = New UIButton(oUILib)
        With btnGoNear
            .ControlName = "btnGoNear"
            .Width = 125
            .Height = 24
            .Left = Me.Width - .Width - Me.BorderLineWidth
            .Top = lnDiv2.Top + 5
            .Enabled = False
            .Visible = True
            .Caption = "Go Near Location"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Go Near the desired Planet or Wormhole." & vbCrLf & "You can either choose the center, or a direction from the object to go nearby." & vbCrLf & "You can also adjust the distance (100 - 15000)"
        End With
        Me.AddChild(CType(btnGoNear, UIControl))

        'lblDist initial props
        lblDist = New UILabel(oUILib)
        With lblDist
            .ControlName = "lblDist"
            .Left = tvwMain.Left + tvwMain.Width + 5
            .Top = 163
            .Width = 170
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Distance From Location"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDist, UIControl))

        'txtDistance initial props
        txtDistance = New UITextBox(oUILib)
        With txtDistance
            .ControlName = "txtDistance"
            .Left = tvwMain.Left + tvwMain.Width + 5
            .Top = lblDist.Top + lblDist.Height + 2
            .Width = 65
            .Height = 18
            .Enabled = Not (myNearDirection = 5)
            .Visible = True
            .Caption = miDist.ToString
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 5
            .BorderColor = muSettings.InterfaceBorderColor
            .Locked = False
            .ToolTipText = "Sepcifies how far away from the location to move to."
        End With
        Me.AddChild(CType(txtDistance, UIControl))


        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = tvwMain.Left + tvwMain.Width
            .Top = 205
            .Width = Me.Width - tvwMain.Left - tvwMain.Width
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'lblGridNum initial props
        lblGridNum = New UILabel(oUILib)
        With lblGridNum
            .ControlName = "lblGridNum"
            .Left = tvwMain.Left + tvwMain.Width + 2
            .Top = lnDiv3.Top + 5
            .Width = 170
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Grid Location"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblGridNum, UIControl))

        'btnGoEnter initial props
        btnGoEnter = New UIButton(oUILib)
        With btnGoEnter
            .ControlName = "btnGoEnter"
            .Width = 125
            .Height = 24
            .Left = Me.Width - .Width - Me.BorderLineWidth
            '.Top = btnGoNear.Top + btnGoNear.Height + 5
            .Top = lnDiv3.Top + 5
            .Enabled = False
            .Visible = True
            .Caption = "Enter Atmosphere"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Go to the desired Planet." & vbCrLf & "You are able to choose a numpad like " & vbCrLf & " grid location to assist placement."
        End With
        Me.AddChild(CType(btnGoEnter, UIControl))

        rcOptGrid7 = New Rectangle(tvwMain.Left + tvwMain.Width + 5, lblGridNum.Top + lblGridNum.Height + 2, 16, 16)
        rcOptGrid8 = New Rectangle(rcOptGrid7.X + 16 + 1, rcOptGrid7.Y, 16, 16)
        rcOptGrid9 = New Rectangle(rcOptGrid8.X + 16 + 1, rcOptGrid7.Y, 16, 16)

        rcOptGrid4 = New Rectangle(rcOptGrid7.X, rcOptGrid7.Y + 16 + 1, 16, 16)
        rcOptGrid5 = New Rectangle(rcOptGrid8.X, rcOptGrid4.Y, 16, 16)
        rcOptGrid6 = New Rectangle(rcOptGrid9.X, rcOptGrid4.Y, 16, 16)

        rcOptGrid1 = New Rectangle(rcOptGrid7.X, rcOptGrid4.Y + 16 + 1, 16, 16)
        rcOptGrid2 = New Rectangle(rcOptGrid8.X, rcOptGrid1.Y, 16, 16)
        rcOptGrid3 = New Rectangle(rcOptGrid9.X, rcOptGrid1.Y, 16, 16)


        'lnDiv4 initial props
        lnDiv4 = New UILine(oUILib)
        With lnDiv4
            .ControlName = "lnDiv4"
            .Left = tvwMain.Left + tvwMain.Width
            .Top = rcOptGrid2.Top + rcOptGrid2.Height + 20 'lnDiv2.Top + 75
            .Width = Me.Width - tvwMain.Left - tvwMain.Width
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv4, UIControl))

        'btnGoNear initial props
        btnGoDirect = New UIButton(oUILib)
        With btnGoDirect
            .ControlName = "btnGoDirect"
            .Width = 125
            .Height = 24
            .Left = Me.Width - .Width - Me.BorderLineWidth
            .Top = lnDiv4.Top + 5 'btnGoNear.Top + btnGoNear.Height + 5
            .Enabled = False
            .Visible = True
            .Caption = "Goto Location"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Go Directly to the shortcut, or jump the selected wormhole."
        End With
        Me.AddChild(CType(btnGoDirect, UIControl))

        'btnAttackTargets initial props
        btnAttackTargets = New UIButton(oUILib)
        With btnAttackTargets
            .ControlName = "btnAttackTargets"
            .Width = 155
            .Height = 24
            .Left = Me.Width - .Width - Me.BorderLineWidth
            .Top = btnGoDirect.Top + btnGoDirect.Height + 5
            .Enabled = True
            .Visible = False
            .Caption = "Attack Known Targets"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Split forces amongst known enemy contacts."
        End With
        Me.AddChild(CType(btnAttackTargets, UIControl))

        If isAdmin() Then mbMultiDepth = True
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eMoveUnits) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            'goUILib.AddNotification("You lack rights to change unit environments.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            'If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        mbLoading = False
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub frmUnitGoto_OnRenderEnd() Handles Me.OnRenderEnd
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        Dim clrVal As System.Drawing.Color
        Dim moDevice As Device = GFXEngine.moDevice
        Using oSprite As New Sprite(MyBase.moUILib.oDevice)
            oSprite.Begin(SpriteFlags.AlphaBlend)

            If myNearDirection = 1 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            BPSprite.Draw2DOnceRotation(moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(rcOptNear1.X, rcOptNear1.Y, 16, 16), clrVal, 256, 256, 225)

            If myNearDirection = 2 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            BPSprite.Draw2DOnceRotation(moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(rcOptNear2.X, rcOptNear2.Y, 16, 16), clrVal, 256, 256, 270)

            If myNearDirection = 3 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            BPSprite.Draw2DOnceRotation(moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(rcOptNear3.X, rcOptNear3.Y, 16, 16), clrVal, 256, 256, 315)

            If myNearDirection = 4 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            BPSprite.Draw2DOnceRotation(moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(rcOptNear4.X, rcOptNear4.Y, 16, 16), clrVal, 256, 256, 180)

            If myNearDirection = 5 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eOption_Marked), New Rectangle(0, 0, 16, 16), New Point(103, 47), 0, rcOptNear5.Location, clrVal)
            Else
                clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eOption_Normal), New Rectangle(0, 0, 16, 16), New Point(103, 47), 0, rcOptNear5.Location, clrVal)
            End If

            If myNearDirection = 6 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            BPSprite.Draw2DOnceRotation(moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(rcOptNear6.X, rcOptNear6.Y, 16, 16), clrVal, 256, 256, 0)

            If myNearDirection = 7 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            BPSprite.Draw2DOnceRotation(moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(rcOptNear7.X, rcOptNear7.Y, 16, 16), clrVal, 256, 256, 135)

            If myNearDirection = 8 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            BPSprite.Draw2DOnceRotation(moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(rcOptNear8.X, rcOptNear8.Y, 16, 16), clrVal, 256, 256, 90)

            If myNearDirection = 9 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            BPSprite.Draw2DOnceRotation(moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(rcOptNear9.X, rcOptNear9.Y, 16, 16), clrVal, 256, 256, 45)

            Dim iOffsetX As Int32 = 97
            Dim iOffsetY As Int32 = 85
            If myGridLocation = 1 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eCheck_Unchecked), rcOptGrid1, Point.Empty, 0, New Point(rcOptGrid1.X - iOffsetX - 0, rcOptGrid1.Y - iOffsetY - 6 - 6), clrVal)

            If myGridLocation = 2 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eCheck_Unchecked), rcOptGrid2, Point.Empty, 0, New Point(rcOptGrid2.X - iOffsetX - 6, rcOptGrid2.Y - iOffsetY - 6 - 6), clrVal)

            If myGridLocation = 3 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eCheck_Unchecked), rcOptGrid3, Point.Empty, 0, New Point(rcOptGrid3.X - iOffsetX - 6 - 6, rcOptGrid3.Y - iOffsetY - 6 - 6), clrVal)

            If myGridLocation = 4 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eCheck_Unchecked), rcOptGrid4, Point.Empty, 0, New Point(rcOptGrid4.X - iOffsetX - 0, rcOptGrid4.Y - iOffsetY - 6), clrVal)

            If myGridLocation = 5 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eCheck_Unchecked), rcOptGrid5, Point.Empty, 0, New Point(rcOptGrid5.X - iOffsetX - 6, rcOptGrid5.Y - iOffsetY - 6), clrVal)

            If myGridLocation = 6 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eCheck_Unchecked), rcOptGrid6, Point.Empty, 0, New Point(rcOptGrid6.X - iOffsetX - 6 - 6, rcOptGrid6.Y - iOffsetY - 6), clrVal)

            If myGridLocation = 7 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eCheck_Unchecked), rcOptGrid7, Point.Empty, 0, New Point(rcOptGrid7.X - iOffsetX - 0, rcOptGrid7.Y - iOffsetY - 0), clrVal)

            If myGridLocation = 8 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eCheck_Unchecked), rcOptGrid8, Point.Empty, 0, New Point(rcOptGrid8.X - iOffsetX - 6, rcOptGrid8.Y - iOffsetY - 0), clrVal)

            If myGridLocation = 9 Then
                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else : clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
            End If
            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eCheck_Unchecked), rcOptGrid9, Point.Empty, 0, New Point(rcOptGrid9.X - iOffsetX - 6 - 6, rcOptGrid9.Y - iOffsetY - 0), clrVal)

            oSprite.End()
        End Using
    End Sub

    Private Sub frmUnitGoto_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Try
            Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
            Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y
            Dim sNearMsg As String = "Sets which direction from the object to move toward. Numpad Loc "
            Dim sGridMsg As String = "Sets which grid location in the environment to move to. Numpad Loc "
            Select Case True
                Case rcOptNear1.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sNearMsg & "1", lMouseX, lMouseY)
                Case rcOptNear2.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sNearMsg & "2", lMouseX, lMouseY)
                Case rcOptNear3.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sNearMsg & "3", lMouseX, lMouseY)
                Case rcOptNear4.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sNearMsg & "4", lMouseX, lMouseY)
                Case rcOptNear5.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sNearMsg & "5", lMouseX, lMouseY)
                Case rcOptNear6.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sNearMsg & "6", lMouseX, lMouseY)
                Case rcOptNear7.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sNearMsg & "7", lMouseX, lMouseY)
                Case rcOptNear8.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sNearMsg & "8", lMouseX, lMouseY)
                Case rcOptNear9.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sNearMsg & "9", lMouseX, lMouseY)
                Case rcOptGrid1.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sGridMsg & "1", lMouseX, lMouseY)
                Case rcOptGrid2.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sGridMsg & "2", lMouseX, lMouseY)
                Case rcOptGrid3.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sGridMsg & "3", lMouseX, lMouseY)
                Case rcOptGrid4.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sGridMsg & "4", lMouseX, lMouseY)
                Case rcOptGrid5.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sGridMsg & "5", lMouseX, lMouseY)
                Case rcOptGrid6.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sGridMsg & "6", lMouseX, lMouseY)
                Case rcOptGrid7.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sGridMsg & "7", lMouseX, lMouseY)
                Case rcOptGrid8.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sGridMsg & "8", lMouseX, lMouseY)
                Case rcOptGrid9.Contains(lX, lY)
                    MyBase.moUILib.SetToolTip(sGridMsg & "9", lMouseX, lMouseY)
            End Select
        Catch
        End Try
    End Sub

    Private Sub frmUnitGoto_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        Try
            Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
            Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y
            Select Case True
                Case rcOptNear1.Contains(lX, lY)
                    myNearDirection = 1
                Case rcOptNear2.Contains(lX, lY)
                    myNearDirection = 2
                Case rcOptNear3.Contains(lX, lY)
                    myNearDirection = 3
                Case rcOptNear4.Contains(lX, lY)
                    myNearDirection = 4
                Case rcOptNear5.Contains(lX, lY)
                    myNearDirection = 5
                Case rcOptNear6.Contains(lX, lY)
                    myNearDirection = 6
                Case rcOptNear7.Contains(lX, lY)
                    myNearDirection = 7
                Case rcOptNear8.Contains(lX, lY)
                    myNearDirection = 8
                Case rcOptNear9.Contains(lX, lY)
                    myNearDirection = 9
                Case rcOptGrid1.Contains(lX, lY)
                    myGridLocation = 1
                Case rcOptGrid2.Contains(lX, lY)
                    myGridLocation = 2
                Case rcOptGrid3.Contains(lX, lY)
                    myGridLocation = 3
                Case rcOptGrid4.Contains(lX, lY)
                    myGridLocation = 4
                Case rcOptGrid5.Contains(lX, lY)
                    myGridLocation = 5
                Case rcOptGrid6.Contains(lX, lY)
                    myGridLocation = 6
                Case rcOptGrid7.Contains(lX, lY)
                    myGridLocation = 7
                Case rcOptGrid8.Contains(lX, lY)
                    myGridLocation = 8
                Case rcOptGrid9.Contains(lX, lY)
                    myGridLocation = 9
                Case Else
                    Exit Sub
            End Select
            Me.IsDirty = True
            If myNearDirection = 5 Then txtDistance.Enabled = False Else txtDistance.Enabled = True
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
        Catch
        End Try
    End Sub

    Private Sub frmUnitGoto_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.UnitGotoLocX = Me.Left
            muSettings.UnitGotoLocY = Me.Top
        End If
    End Sub

    'Solar system and planets
    Private Structure Environments
        Dim iIndex As Int32
        Dim iEnvirTypeID As Int16
        Dim lEnvirID As Int32
        Dim lParentEnvirID As Int32
        Dim sName As String
        Dim iLocX As Int32
        Dim iLocZ As Int32
        Dim yPlanetSizeID As Byte
        Dim oNode As UITreeView.UITreeViewItem
        Dim fDistance As Single
    End Structure
    Private moEnvironments() As Environments

    '' Shortcuts
    Private Structure EnvirShortcut
        Dim iIndex As Int32
        Dim iEnvirTypeID As Int16
        Dim lEnvirID As Int32
        Dim iLocX As Int32
        'Dim iLocY As Int32
        Dim iLocZ As Int32
        Dim sName As String
        Dim bNoSave As Boolean
        Dim bInList As Boolean
    End Structure
    Private moShortCuts() As EnvirShortcut
    Private Sub LoadShortcutFile()
        If goCurrentPlayer Is Nothing Then Return

        'Read from file
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= goCurrentPlayer.PlayerName & ".sht"

        If My.Computer.FileSystem.FileExists(sFile) = True Then
            Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Open)
            Dim oReader As IO.BinaryReader = New IO.BinaryReader(fsFile)
            Dim lPos As Int32 = 0
            Dim lCnt As Int32 = 0
            Try
                While fsFile.Position < fsFile.Length
                    Dim yHdr() As Byte = oReader.ReadBytes(34)
                    If yHdr Is Nothing = False Then
                        lPos = 0
                        ReDim Preserve moShortCuts(lCnt)
                        With moShortCuts(lCnt)
                            .iIndex = lCnt
                            .iEnvirTypeID = System.BitConverter.ToInt16(yHdr, lPos) : lPos += 2
                            .lEnvirID = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .iLocX = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            '.iLocY = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .iLocZ = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .sName = GetStringFromBytes(yHdr, lPos, 20) : lPos += 20
                            .bInList = False
                            lCnt += 1
                        End With
                    Else
                        Exit While
                    End If
                End While
            Catch
            End Try
            oReader.Close()
            oReader = Nothing
            fsFile.Close()
            fsFile.Dispose()
            fsFile = Nothing
        End If
    End Sub

    Public Sub FillList()
        If mbRequestingEnvironment = True Then
            mbRequestingEnvironment = False
            tvwMain.Enabled = True
        End If
        If tvwMain.TotalNodeCount > 0 Then SaveListExpansion()
        tvwMain.Clear()
        btnGoNear.Enabled = False
        btnGoEnter.Enabled = False
        btnGoDirect.Enabled = False

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return

        Dim oSystem As SolarSystem
        If oEnvir.ObjTypeID = ObjectType.eSolarSystem Then
            oSystem = CType(goCurrentEnvir.oGeoObject, SolarSystem)
        ElseIf oEnvir.ObjTypeID = ObjectType.ePlanet Then
            oSystem = CType(CType(goCurrentEnvir.oGeoObject, Planet).ParentSystem, SolarSystem)
        Else : Return
        End If
        If oSystem Is Nothing Then Return


        ReDim moEnvironments(0)
        With moEnvironments(0)
            .iIndex = 0
            .iEnvirTypeID = oSystem.ObjTypeID
            .lEnvirID = oSystem.ObjectID
            .sName = oSystem.SystemName
            .iLocX = -1
            .iLocZ = -1
            .fDistance = Int32.MaxValue
            .yPlanetSizeID = 0
        End With
        Dim oSystemNode As UITreeView.UITreeViewItem = tvwMain.AddNode(oSystem.SystemName, oSystem.ObjectID, oSystem.ObjTypeID, -1, Nothing, Nothing)
        oSystemNode.bExpanded = True
        oSystemNode.bItemBold = True
        moEnvironments(0).oNode = oSystemNode


        LoadShortcutFile()
        Dim iSystemPath(-1) As Int32
        FillForEnvironment(oSystem, oSystemNode, 0, iSystemPath)
        Dim oNode As UITreeView.UITreeViewItem = Nothing
        If miObjTypeID > 0 AndAlso miObjectID > 0 Then
            oNode = tvwMain.GetNodeByItemData3(miObjectID, miObjTypeID, miIndex)
            If oNode Is Nothing = False Then
                tvwMain.oSelectedNode = oNode
            Else 'User must have changed environments and this item is nolonger relevent
                miObjTypeID = -1
                miObjectID = -1
                miIndex = -1
            End If
        End If
        If oNode Is Nothing = True Then ' User has yet to select something, so lets select closest (but not store that as selection)
            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To moEnvironments.GetUpperBound(0)
                Dim lIdx As Int32 = -1
                For Y As Int32 = 0 To lSortedUB
                    If moEnvironments(X).fDistance < moEnvironments(lSorted(Y)).fDistance Then
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
            If moEnvironments(lSorted(0)).oNode Is Nothing = False Then
                Debug.Print("1st Lowest is " & moEnvironments(lSorted(0)).oNode.sItem & " - " & moEnvironments(lSorted(0)).fDistance.ToString("#,###.##"))
                Debug.Print("2nd Lowest is " & moEnvironments(lSorted(1)).oNode.sItem & " - " & moEnvironments(lSorted(1)).fDistance.ToString("#,###.##"))
                Debug.Print("3rd Lowest is " & moEnvironments(lSorted(2)).oNode.sItem & " - " & moEnvironments(lSorted(2)).fDistance.ToString("#,###.##"))
                mbIgnoreSelection = True
                tvwMain.oSelectedNodeNoExpand = moEnvironments(lSorted(0)).oNode
                mbIgnoreSelection = False
            ElseIf miObjTypeID > 0 AndAlso miObjectID > 0 Then
                tvwMain.oSelectedNode = tvwMain.GetNodeByItemData3(miObjectID, miObjTypeID, miIndex)
            End If
        End If
    End Sub

    Private Function FillForEnvironment(ByVal oSystem As SolarSystem, ByVal oSystemNode As UITreeView.UITreeViewItem, ByVal iDepth As Int32, ByVal iSystemPath() As Int32) As Boolean
        If mbMultiDepth = False AndAlso iDepth > 0 Then Return False
        If iDepth > 25 Then Return False
        If moEnvironments.GetUpperBound(0) > 1000 Then Return False
        With oSystem
            ReDim Preserve iSystemPath(iSystemPath.GetUpperBound(0) + 1)
            iSystemPath(iSystemPath.GetUpperBound(0)) = oSystem.ObjectID
            If oSystem.PlanetUB = -1 Then Return False

            'Load Planets
            Dim lIdx As Int32 = moEnvironments.GetUpperBound(0) + 1
            ReDim Preserve moEnvironments(moEnvironments.GetUpperBound(0) + .PlanetUB + 1)
            For X As Int32 = 0 To .PlanetUB
                With .moPlanets(X)
                    Dim sPlanetName As String = oSystem.moPlanets(X).PlanetName
                    Dim oPlanetNode As UITreeView.UITreeViewItem = tvwMain.AddNode(.PlanetName, .ObjectID, .ObjTypeID, -1, oSystemNode, Nothing)
                    With moEnvironments(lIdx)
                        .iIndex = lIdx
                        .iEnvirTypeID = oSystem.moPlanets(X).ObjTypeID
                        .lEnvirID = oSystem.moPlanets(X).ObjectID
                        .sName = oSystem.moPlanets(X).PlanetName
                        .iLocX = oSystem.moPlanets(X).LocX
                        .iLocZ = oSystem.moPlanets(X).LocZ
                        .yPlanetSizeID = oSystem.moPlanets(X).PlanetSizeID
                        .oNode = oPlanetNode
                        .lParentEnvirID = oSystem.ObjectID
                        If .lEnvirID = goCurrentEnvir.ObjectID AndAlso .iEnvirTypeID = ObjectType.ePlanet AndAlso miObjTypeID = ObjectType.ePlanet Then
                            miObjectID = .lEnvirID
                        End If
                        If iDepth = 0 Then
                            If oSystem.moPlanets(X).ObjectID = goCurrentEnvir.ObjectID Then
                                .fDistance = -1
                            Else
                                Dim fDist As Single
                                Dim fTX As Single = oSystem.moPlanets(X).LocX - goCamera.mlCameraX
                                Dim fTZ As Single = oSystem.moPlanets(X).LocZ - goCamera.mlCameraZ
                                fTX *= fTX
                                fTZ *= fTZ
                                fDist = CSng(Math.Sqrt(fTX + fTZ))
                                .fDistance = fDist
                                'Debug.Print(.sName & " - " & fDist.ToString("#,###.##"))
                            End If
                        Else
                            .fDistance = Int32.MaxValue
                        End If
                    End With
                End With
                lIdx += 1
            Next

            'Load Wormholes
            For X As Int32 = 0 To .WormholeUB
                With .moWormholes(X)
                    'Traverse Wormholes...
                    'If system is T2, do not traverse back to T3 if depth > 0..1..2
                    'If system is T2, do not traverse other T2 if depth > 0..1..
                    'Z -> E -> C -> P
                    'Z -> E -> D -> C -> P
                    'NOT:
                    'Z -> E -> D -> E
                    'Z -> E -> D -> C -> E...
                    'Z -> E -> D -> C -> D...

                    Dim oWormholeNode As UITreeView.UITreeViewItem = Nothing
                    Dim oRemoteSystem As SolarSystem = Nothing
                    If .System1 Is Nothing = False AndAlso .System1.ObjectID = oSystem.ObjectID Then
                        For Y As Int32 = 0 To goGalaxy.mlSystemUB
                            If goGalaxy.moSystems(Y).ObjectID = .System2.ObjectID Then
                                oRemoteSystem = goGalaxy.moSystems(Y)
                                Exit For
                            End If
                        Next
                        Dim sWormholeName As String = .System2.SystemName
                        If sWormholeName Is Nothing OrElse sWormholeName = "" Then sWormholeName = "Unknown"
                        sWormholeName = "Wormhole To " & sWormholeName
                        oWormholeNode = tvwMain.AddNode(sWormholeName, .ObjectID, .ObjTypeID, .System1.ObjectID, oSystemNode, Nothing)
                        If WasExpanded(oWormholeNode) = True Then oWormholeNode.bExpanded = True

                        If oRemoteSystem Is Nothing = False AndAlso NotInPath(iSystemPath, oRemoteSystem) = True Then
                            If iDepth = iDepth AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then

                                ReDim Preserve moEnvironments(moEnvironments.GetUpperBound(0) + 1)
                                With moEnvironments(moEnvironments.GetUpperBound(0))
                                    .iIndex = moEnvironments.GetUpperBound(0)
                                    .iEnvirTypeID = oSystem.moWormholes(X).ObjTypeID
                                    .lEnvirID = oSystem.moWormholes(X).ObjectID
                                    .sName = oSystem.moWormholes(X).System2.SystemName
                                    .iLocX = oSystem.moWormholes(X).LocX1
                                    .iLocZ = oSystem.moWormholes(X).LocY1
                                    .yPlanetSizeID = Byte.MaxValue - 1
                                    .oNode = oWormholeNode
                                    .lParentEnvirID = oRemoteSystem.ObjectID
                                    Dim fDist As Single
                                    Dim fTX As Single = .iLocX - goCamera.mlCameraX
                                    Dim fTZ As Single = .iLocZ - goCamera.mlCameraZ
                                    fTX *= fTX
                                    fTZ *= fTZ
                                    fDist = CSng(Math.Sqrt(fTX + fTZ))
                                    .fDistance = fDist
                                    'Debug.Print("Wh To " & .sName & " - " & fDist.ToString("#,###.##"))
                                End With
                            End If
                            If mbMultiDepth = True Then
                                If FillForEnvironment(oRemoteSystem, oWormholeNode, iDepth + 1, iSystemPath) = False Then
                                    tvwMain.AddNode("Loading...", oWormholeNode.lItemData, ObjectType.eSolarSystem, .System2.ObjectID, oWormholeNode, Nothing)
                                End If
                            End If
                        End If
                    ElseIf .System2 Is Nothing = False AndAlso .System2.ObjectID = oSystem.ObjectID Then
                        For Y As Int32 = 0 To goGalaxy.mlSystemUB
                            If goGalaxy.moSystems(Y).ObjectID = .System1.ObjectID Then
                                oRemoteSystem = goGalaxy.moSystems(Y)
                                Exit For
                            End If
                        Next
                        Dim sWormholeName As String = .System1.SystemName
                        If sWormholeName Is Nothing OrElse sWormholeName = "" Then sWormholeName = "Unknown"
                        sWormholeName = "Wormhole To " & sWormholeName
                        oWormholeNode = tvwMain.AddNode(sWormholeName, .ObjectID, .ObjTypeID, .System2.ObjectID, oSystemNode, Nothing)
                        If WasExpanded(oWormholeNode) = True Then oWormholeNode.bExpanded = True
                        If oRemoteSystem Is Nothing = False AndAlso NotInPath(iSystemPath, oRemoteSystem) = True Then
                            If iDepth = iDepth AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                ReDim Preserve moEnvironments(moEnvironments.GetUpperBound(0) + 1)
                                With moEnvironments(moEnvironments.GetUpperBound(0))
                                    .iIndex = moEnvironments.GetUpperBound(0)
                                    .iEnvirTypeID = oSystem.moWormholes(X).ObjTypeID
                                    .lEnvirID = oSystem.moWormholes(X).ObjectID
                                    .sName = oSystem.moWormholes(X).System1.SystemName
                                    .iLocX = oSystem.moWormholes(X).LocX2
                                    .iLocZ = oSystem.moWormholes(X).LocY2
                                    .yPlanetSizeID = Byte.MaxValue - 1
                                    .oNode = oWormholeNode
                                    .lParentEnvirID = oRemoteSystem.ObjectID
                                    Dim fDist As Single
                                    Dim fTX As Single = .iLocX - goCamera.mlCameraX
                                    Dim fTZ As Single = .iLocZ - goCamera.mlCameraZ
                                    fTX *= fTX
                                    fTZ *= fTZ
                                    fDist = CSng(Math.Sqrt(fTX + fTZ))
                                    .fDistance = fDist
                                    'Debug.Print("Wh To " & .sName & " - " & fDist.ToString("#,###.##"))
                                End With
                            End If
                            If mbMultiDepth = True Then
                                If FillForEnvironment(oRemoteSystem, oWormholeNode, iDepth + 1, iSystemPath) = False Then
                                    tvwMain.AddNode("Loading...", oWormholeNode.lItemData, ObjectType.eSolarSystem, .System1.ObjectID, oWormholeNode, Nothing)
                                End If
                            End If
                        End If
                    End If
                End With
            Next

            'Load Shortcuts
            If iDepth = 0 AndAlso moShortCuts Is Nothing = False Then
                For X As Int32 = 0 To moShortCuts.GetUpperBound(0)
                    'Debug.Print("Shortcut: " & X.ToString & " " & moShortCuts(X).iEnvirTypeID & " " & moShortCuts(X).lEnvirID & " " & moShortCuts(X).sName)
                    For y As Int32 = 0 To moEnvironments.GetUpperBound(0)
                        'Debug.Print("ENV: " & moEnvironments(y).iEnvirTypeID & " " & moEnvironments(y).lEnvirID & " " & moEnvironments(y).lParentEnvirID & " " & moEnvironments(y).sName)
                        If moShortCuts(X).bInList = False Then
                            If (moShortCuts(X).iEnvirTypeID = moEnvironments(y).iEnvirTypeID AndAlso moShortCuts(X).lEnvirID = moEnvironments(y).lEnvirID) OrElse (moEnvironments(y).iEnvirTypeID = ObjectType.eWormhole AndAlso moShortCuts(X).iEnvirTypeID = ObjectType.eSolarSystem AndAlso moEnvironments(y).lParentEnvirID = moShortCuts(X).lEnvirID) Then
                                moShortCuts(X).bInList = True
                                If mbMultiDepth = True OrElse (mbMultiDepth = False AndAlso moEnvironments(y).oNode.lItemData2 <> ObjectType.eWormhole) Then
                                    Dim oShortcutNode As UITreeView.UITreeViewItem = tvwMain.AddNode("Shortcut: " & moShortCuts(X).sName, .ObjectID, .ObjTypeID, moShortCuts(X).iIndex, moEnvironments(y).oNode, Nothing)
                                    If (moEnvironments(y).iEnvirTypeID = ObjectType.eSolarSystem AndAlso moEnvironments(y).lEnvirID = goCurrentEnvir.ObjectID) OrElse (moEnvironments(y).iEnvirTypeID = ObjectType.ePlanet AndAlso moEnvironments(y).lParentEnvirID = goCurrentEnvir.ObjectID) Then moEnvironments(y).oNode.bExpanded = True
                                    If iDepth = 0 AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                        ReDim Preserve moEnvironments(moEnvironments.GetUpperBound(0) + 1)
                                        With moEnvironments(moEnvironments.GetUpperBound(0))
                                            .iIndex = moEnvironments.GetUpperBound(0)
                                            .iEnvirTypeID = oSystem.ObjTypeID
                                            .lEnvirID = oSystem.ObjectID
                                            .sName = moShortCuts(X).sName
                                            .iLocX = moShortCuts(X).iLocX
                                            .iLocZ = moShortCuts(X).iLocZ
                                            .yPlanetSizeID = Byte.MaxValue - 2
                                            .oNode = oShortcutNode

                                            Dim fDist As Single
                                            Dim fTX As Single = .iLocX - goCamera.mlCameraX
                                            Dim fTZ As Single = .iLocZ - goCamera.mlCameraZ
                                            fTX *= fTX
                                            fTZ *= fTZ
                                            fDist = CSng(Math.Sqrt(fTX + fTZ))
                                            .fDistance = fDist
                                            'Debug.Print("Sh To " & .sName & " - " & fDist.ToString("#,###.##"))
                                        End With
                                    End If
                                End If
                            End If
                        End If
                    Next
                Next
            End If
        End With
        Return True
    End Function

    Private Function NotInPath(ByVal iSystemsPath() As Int32, ByVal oSystem As SolarSystem) As Boolean
        For x As Int32 = 0 To iSystemsPath.GetUpperBound(0)
            If iSystemsPath(x) = oSystem.ObjectID Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Sub SaveListExpansion()
        ReDim mbExp(-1)
        Dim lIdx As Int32 = 0
        Dim lListCount As Int32 = tvwMain.TotalNodeCount
        Dim oActNode As UITreeView.UITreeViewItem = tvwMain.GetNodeByItemText("", False, False, lIdx - 1)

        While oActNode Is Nothing = False AndAlso lIdx < lListCount
            If oActNode.bExpanded = True Then
                ReDim Preserve mbExp(mbExp.GetUpperBound(0) + 1)
                mbExp(mbExp.GetUpperBound(0)) = oActNode.lItemData & "." & oActNode.lItemData2 & "." & oActNode.lItemData3
            End If
            lIdx += 1
            oActNode = tvwMain.GetNodeByItemText("", False, False, lIdx - 1)
        End While
    End Sub

    Private Function WasExpanded(ByVal oNode As UITreeView.UITreeViewItem) As Boolean
        If mbExp Is Nothing = True OrElse mbExp.GetUpperBound(0) = 0 Then Exit Function
        For x As Int32 = 0 To mbExp.GetUpperBound(0)
            If mbExp(x) = oNode.lItemData & "." & oNode.lItemData2 & "." & oNode.lItemData3 Then
                Return True
            End If
        Next
    End Function

    Private Sub tvwMain_NodeExpanded(ByVal oNode As UITreeView.UITreeViewItem) Handles tvwMain.NodeExpanded
        If mbRequestingEnvironment = True Then Exit Sub
        If oNode Is Nothing = False Then
            Dim oChildNode As UITreeView.UITreeViewItem = oNode.oFirstChild
            If oChildNode Is Nothing = False Then
                Dim lID As Int32 = oChildNode.lItemData
                Dim iTypeID As Int16 = CShort(oChildNode.lItemData2)
                Dim lObjectID As Int32 = oChildNode.lItemData3

                If iTypeID <> ObjectType.eSolarSystem OrElse oChildNode.sItem <> "Loading..." Then Exit Sub
                For Y As Int32 = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(Y).ObjectID = lObjectID Then
                        If goGalaxy.moSystems(Y).PlanetUB <> -1 Then Exit Sub
                        Exit For
                    End If
                Next
                mbRequestingEnvironment = True
                mlRequestingEnvironmentID = lObjectID
                Dim yData(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestSystemDetails).CopyTo(yData, 0)
                System.BitConverter.GetBytes(lObjectID).CopyTo(yData, 2)
                goUILib.SendMsgToPrimary(yData)
                tvwMain.Enabled = False
                Me.IsDirty = True
            End If
        End If
    End Sub

    Private Sub tvwMain_NodeSelected(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwMain.NodeSelected
        btnGoNear.Enabled = False
        btnGoEnter.Enabled = False
        btnGoDirect.Enabled = False
        Dim lNewObjectID As Int32
        Dim lNewObjTypeID As Int16
        Dim iNewIndex As Int32

        If oNode Is Nothing = False Then
            Dim lID As Int32 = oNode.lItemData
            Dim iTypeID As Int16 = CShort(oNode.lItemData2)
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                Dim lCurUB As Int32 = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))

                If lID < 1 Then Return
                If iTypeID < 1 Then Return
                lNewObjectID = lID
                lNewObjTypeID = iTypeID
                iNewIndex = oNode.lItemData3

                If oNode.lTreeViewIndex = 0 Then

                ElseIf iTypeID = ObjectType.eWormhole Then
                    btnGoNear.Enabled = True
                    btnGoDirect.Enabled = True
                    'End If
                ElseIf iTypeID = ObjectType.eSolarSystem Then
                    btnGoDirect.Enabled = True
                ElseIf iTypeID = ObjectType.ePlanet Then
                    btnGoNear.Enabled = True
                    btnGoEnter.Enabled = True
                Else
                    btnGoDirect.Enabled = True
                End If
            End If
            If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
            tvwMain.HasFocus = False
            Me.HasFocus = False
            goUILib.FocusedControl = Nothing
            lNewObjectID = lID
            lNewObjTypeID = iTypeID
        End If
        If mbIgnoreSelection = False Then
            miObjectID = lNewObjectID
            miObjTypeID = lNewObjTypeID
            miIndex = iNewIndex
        End If
    End Sub

    Private Sub ExecuteAlertSettings(ByRef yMultiMsg() As Byte)
        Dim iLen As Int16 = 9
        Dim lPos As Int32 = 0

        Select Case myAlertValue
            Case 1 'Set first
                Dim oFirstUnit As BaseEntity
                If goCurrentEnvir Is Nothing = False Then
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                            oFirstUnit = goCurrentEnvir.oEntity(X)
                            frmSingleSelect.AddUnitAlert(goCurrentEnvir.oEntity(X).ObjectID, goCurrentEnvir.oEntity(X).ObjTypeID)

                            ReDim yMultiMsg(((iLen + 2) * 1) - 1)
                            System.BitConverter.GetBytes(iLen).CopyTo(yMultiMsg, lPos) : lPos += 2
                            System.BitConverter.GetBytes(GlobalMessageCode.eAlertDestinationReached).CopyTo(yMultiMsg, lPos) : lPos += 2
                            oFirstUnit.GetGUIDAsString.CopyTo(yMultiMsg, lPos) : lPos += 6
                            yMultiMsg(lPos) = 1 : lPos += 1
                            Exit For
                        End If
                    Next X
                End If
            Case 0, 2 'Set all
                Dim lCnt As Int32 = 0
                If goCurrentEnvir Is Nothing = False Then
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                            If myAlertValue = 0 Then
                                frmSingleSelect.UnitAlertReceived(goCurrentEnvir.oEntity(X).ObjectID, goCurrentEnvir.oEntity(X).ObjTypeID)
                            Else
                                frmSingleSelect.AddUnitAlert(goCurrentEnvir.oEntity(X).ObjectID, goCurrentEnvir.oEntity(X).ObjTypeID)
                            End If
                            lCnt += 1
                        End If
                    Next X

                    ReDim yMultiMsg(((iLen + 2) * lCnt) - 1)
                    lPos = 0
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                            System.BitConverter.GetBytes(iLen).CopyTo(yMultiMsg, lPos) : lPos += 2
                            System.BitConverter.GetBytes(GlobalMessageCode.eAlertDestinationReached).CopyTo(yMultiMsg, lPos) : lPos += 2
                            goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yMultiMsg, lPos) : lPos += 6
                            If myAlertValue = 0 Then
                                yMultiMsg(lPos) = 0 : lPos += 1
                            Else
                                yMultiMsg(lPos) = 1 : lPos += 1
                            End If

                        End If
                    Next X
                    'goUILib.SendMsgToRegion(yMultiMsg)
                End If
        End Select
    End Sub

    Private Sub optAlertNone_Click() Handles optAlertNone.Click
        optAlertNone.Value = True
        If myAlertValue = 0 Then Return
        optAlertFirst.Value = False
        optAlertAll.Value = False

        myAlertValue = 0
    End Sub

    Private Sub optAlertFirst_Click() Handles optAlertFirst.Click
        optAlertFirst.Value = True
        If myAlertValue = 1 Then Return
        optAlertNone.Value = False
        optAlertAll.Value = False

        myAlertValue = 1
    End Sub

    Private Sub optAlertAll_Click() Handles optAlertAll.Click
        optAlertAll.Value = True
        If myAlertValue = 2 Then Return
        optAlertNone.Value = False
        optAlertFirst.Value = False

        myAlertValue = 2
    End Sub

    Private Sub btnGoNear_Click(ByVal sName As String) Handles btnGoNear.Click

        Dim oNode As UITreeView.UITreeViewItem = tvwMain.oSelectedNode
        If oNode Is Nothing = True Then Return

        If Val(txtDistance.Caption) < 1 Then
            MyBase.moUILib.AddNotification("Distance cannot be less than 1.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf Val(txtDistance.Caption) > 15000 Then
            MyBase.moUILib.AddNotification("Distance cannot be greater than 15,000.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        miDist = CInt(Val(txtDistance.Caption))
        Dim iDist As Int32 = miDist * gl_FINAL_GRID_SQUARE_SIZE

        'Dim oEnvir As BaseEnvironment = goCurrentEnvir
        'If oEnvir Is Nothing Then Return

        'Dim oSystem As SolarSystem
        'If oEnvir.ObjTypeID = ObjectType.eSolarSystem Then
        '    oSystem = CType(goCurrentEnvir.oGeoObject, SolarSystem)
        'ElseIf oEnvir.ObjTypeID = ObjectType.ePlanet Then
        '    oSystem = CType(CType(goCurrentEnvir.oGeoObject, Planet).ParentSystem, SolarSystem)
        'Else : Return
        'End If
        ''If oSystem Is Nothing Then Return

        'now, generate our movement requests... determine if we have a multiple selection (of say > 3)
        Dim lOrderCnt As Int32 = 0
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    lOrderCnt += 1
                End If
            End If
        Next X
        If lOrderCnt = 0 Then Return

        Dim lDestID As Int32 = oNode.lItemData
        Dim iDestTypeID As Int16 = CShort(oNode.lItemData2)

        Dim bFound As Boolean = False
        Dim lDestX As Int32 = 0
        Dim lDestZ As Int32 = 0
        If oNode.sItem.StartsWith("Shortcut: ") Then
        ElseIf oNode.sItem.StartsWith("Wormhole To ") Then
            For x As Int32 = 0 To moEnvironments.GetUpperBound(0)
                If moEnvironments(x).iEnvirTypeID = iDestTypeID Then
                    If moEnvironments(x).lEnvirID = lDestID Then
                        'If oSystem.moWormholes(x).ObjectID = lDestID AndAlso oSystem.moWormholes(x).ObjTypeID = iDestTypeID Then
                        bFound = True
                        iDestTypeID = ObjectType.eSolarSystem
                        lDestID = oNode.lItemData3
                        'If oSystem.moWormholes(x).System1 Is Nothing = False AndAlso oSystem.moWormholes(x).System1.ObjectID = oSystem.ObjectID Then
                        'lDestX = oSystem.moWormholes(x).LocX1
                        'lDestZ = oSystem.moWormholes(x).LocY1
                        'ElseIf oSystem.moWormholes(x).System2 Is Nothing = False AndAlso oSystem.moWormholes(x).System2.ObjectID = oSystem.ObjectID Then
                        'lDestX = oSystem.moWormholes(x).LocX2
                        'lDestZ = oSystem.moWormholes(x).LocY2
                        'End If
                        lDestX = moEnvironments(x).iLocX
                        lDestZ = moEnvironments(x).iLocZ
                        Exit For
                    End If
                End If
            Next x
        Else 'must be planet
            If moEnvironments Is Nothing Then Return
            For x As Int32 = 0 To moEnvironments.GetUpperBound(0)
                If moEnvironments(x).lEnvirID = lDestID AndAlso moEnvironments(x).iEnvirTypeID = iDestTypeID Then
                    lDestX = moEnvironments(x).iLocX
                    lDestZ = moEnvironments(x).iLocZ
                    bFound = True
                    iDestTypeID = ObjectType.eSolarSystem
                    lDestID = moEnvironments(x).lParentEnvirID
                    Exit For
                End If
            Next x
        End If
        If bFound = False Then Return
        Dim iAngle As Int32 = 0
        Dim iX As Int32 = 0
        Dim iY As Int32 = 0
        Dim iA As Int32 = 0
        Dim iB As Int32 = 0
        Select Case myNearDirection
            Case 1 ' Down and Left
                iAngle = 225
            Case 2 ' Down
                iAngle = 180
            Case 3 ' Down and Right
                iAngle = 135
            Case 4 ' Left
                iAngle = 270
            Case 5 ' Center
                iAngle = 0
                iDist = 0
            Case 6 ' Right
                iAngle = 90
            Case 7 ' Up and Left
                iAngle = 315
            Case 8 ' Up
                iAngle = 0
            Case 9 ' Up and Right
                iAngle = 45
        End Select

        If iAngle < 0 Or iAngle > 359 Then iAngle = 0
        Dim R As Double = iAngle * (Math.PI / 180)
        lDestX += iDist * CInt(Math.Sin(R))
        lDestZ += iDist * CInt(Math.Cos(R))

        ReDim MoveQueue(0)
        With MoveQueue(0)
            .MoveType = GlobalMessageCode.eMoveObjectRequest
            .locX = lDestX
            .locZ = lDestZ
            .lDestID = lDestID 'oSystem.ObjectID
            .iDestTypeID = iDestTypeID 'oSystem.ObjTypeID
        End With

        ProcessMoveQueue()
    End Sub

    Private Sub btnGoEnter_Click(ByVal sName As String) Handles btnGoEnter.Click
        Dim oNode As UITreeView.UITreeViewItem = tvwMain.oSelectedNode
        If oNode Is Nothing = True Then Return

        'now, generate our movement requests... determine if we have a multiple selection (of say > 3)
        Dim lOrderCnt As Int32 = 0
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    lOrderCnt += 1
                End If
            End If
        Next X
        If lOrderCnt = 0 Then Return

        Dim lDestID As Int32 = oNode.lItemData
        Dim iDestTypeID As Int16 = CShort(oNode.lItemData2)

        Dim bFound As Boolean = False
        Dim lDestX As Int32 = 0
        Dim lDestZ As Int32 = 0
        Dim lSize As Int32 = 0
        If oNode.sItem.StartsWith("Shortcut: ") Then
            If moShortCuts Is Nothing = False Then
                Dim iIndex As Int32 = oNode.lItemData3
                Try
                    With moShortCuts(iIndex)
                        lDestX = .iLocX
                        lDestZ = .iLocZ
                        lDestID = .lEnvirID
                        iDestTypeID = .iEnvirTypeID
                    End With
                Catch
                    Return
                End Try
            End If
        ElseIf oNode.sItem.StartsWith("Wormhole To ") Then
        Else 'must be planet
            If moEnvironments Is Nothing Then Return
            For x As Int32 = 0 To moEnvironments.GetUpperBound(0)
                If moEnvironments(x).lEnvirID = lDestID AndAlso moEnvironments(x).iEnvirTypeID = iDestTypeID Then
                    'Calculate dest based upon Grid
                    lDestX = 0
                    lDestZ = 0
                    Select Case moEnvironments(x).yPlanetSizeID
                        Case 0 'Tiny (24000)
                            lSize = gl_TINY_PLANET_CELL_SPACING * TerrainClass.Width \ 2
                        Case 1 'Normal
                            lSize = gl_SMALL_PLANET_CELL_SPACING * TerrainClass.Width \ 2
                        Case 2 'medium
                            lSize = gl_MEDIUM_PLANET_CELL_SPACING * TerrainClass.Width \ 2
                        Case 3 'Large 
                            lSize = gl_LARGE_PLANET_CELL_SPACING * TerrainClass.Width \ 2
                        Case 4 'Hige
                            lSize = gl_HUGE_PLANET_CELL_SPACING * TerrainClass.Width \ 2
                    End Select
                    '+x-Z         -x-z
                    '+x+z         -x+z
                    Select Case myGridLocation
                        Case 1 ' Down and Left
                            lDestX = CInt(lSize * (2 / 3))
                            lDestZ = CInt(lSize * (2 / 3))
                        Case 2 ' Down
                            lDestX = 0
                            lDestZ = CInt(lSize * (2 / 3))
                        Case 3 ' Down and Right
                            lDestX = -CInt(lSize * (2 / 3))
                            lDestZ = CInt(lSize * (2 / 3))
                        Case 4 ' Left
                            lDestX = CInt(lSize * (2 / 3))
                            lDestZ = 0
                        Case 5 ' Center
                            lDestX = 0
                            lDestZ = 0
                        Case 6 ' Right
                            lDestX = -CInt(lSize * (2 / 3))
                            lDestZ = 0
                        Case 7 ' Up and Left
                            lDestX = CInt(lSize * (2 / 3))
                            lDestZ = -CInt(lSize * (2 / 3))
                        Case 8 ' Up
                            lDestX = 0
                            lDestZ = -CInt(lSize * (2 / 3))
                        Case 9 ' Up and Right
                            lDestX = -CInt(lSize * (2 / 3))
                            lDestZ = -CInt(lSize * (2 / 3))
                    End Select
                    bFound = True
                    Exit For
                End If
            Next x
        End If
        If bFound = False Then Return

        ReDim MoveQueue(0)
        With MoveQueue(0)
            .MoveType = GlobalMessageCode.eMoveObjectRequest
            .locX = lDestX
            .locZ = lDestZ
            .lDestID = lDestID
            .iDestTypeID = iDestTypeID
        End With

        ProcessMoveQueue()
    End Sub

    Private Sub btnGoDirect_Click(ByVal sName As String) Handles btnGoDirect.Click
        Dim oNode As UITreeView.UITreeViewItem = tvwMain.oSelectedNode
        If oNode Is Nothing = True Then Return
        Dim X As Int32

        'now, generate our movement requests... determine if we have a multiple selection (of say > 3)
        Dim lOrderCnt As Int32 = 0
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    lOrderCnt += 1
                End If
            End If
        Next X

        If lOrderCnt = 0 Then Return
        Dim lDestID As Int32 = oNode.lItemData
        Dim iDestTypeID As Int16 = CShort(oNode.lItemData2)

        If CShort(oNode.lItemData2) = ObjectType.eWormhole Then
            ReDim MoveQueue(0)
            With MoveQueue(0)
                .MoveType = GlobalMessageCode.eJumpTarget
                .WormholeID = oNode.lItemData3
                .lDestID = lDestID
                .iDestTypeID = iDestTypeID
                For z As Int32 = 0 To moEnvironments.GetUpperBound(0)
                    If moEnvironments(z).iEnvirTypeID = ObjectType.eWormhole AndAlso moEnvironments(z).lEnvirID = .lDestID Then
                        .locX = moEnvironments(z).iLocX
                        .locY = moEnvironments(z).iLocZ
                        Exit For
                    End If
                Next
            End With
        Else
            Dim lDestX As Int32 = 0
            Dim lDestZ As Int32 = 0
            If oNode.sItem.StartsWith("Shortcut: ") Then
                If moShortCuts Is Nothing = False Then
                    Dim iIndex As Int32 = oNode.lItemData3
                    Try
                        With moShortCuts(iIndex)
                            lDestX = .iLocX
                            lDestZ = .iLocZ
                            lDestID = .lEnvirID
                            iDestTypeID = .iEnvirTypeID
                        End With
                    Catch
                        Return
                    End Try
                End If
            ElseIf oNode.sItem.StartsWith("Wormhole To ") Then
                Dim sss As Int32 = -1
            Else
                Return
            End If

            ReDim MoveQueue(0)
            With MoveQueue(0)
                .MoveType = GlobalMessageCode.eMoveObjectRequest
                .locX = lDestX
                .locZ = lDestZ
                .lDestID = lDestID
                .iDestTypeID = iDestTypeID
            End With
        End If

        ProcessMoveQueue()
    End Sub

    Private Sub DoTheMove(ByVal yMultiMove() As Byte)
        Dim yAlertMsg() As Byte
        ReDim yAlertMsg(0)
        ExecuteAlertSettings(yAlertMsg)
        MyBase.moUILib.SendLenAppendedMsgToRegion(yAlertMsg)
        Dim oThread As New Threading.Thread(AddressOf SleepAndMove)
        Dim oParms As FuncCallParms = New FuncCallParms(yMultiMove)
        oThread.Start(oParms)
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub SleepAndMove(ByVal oArgs As Object)
        Dim oParms As FuncCallParms = CType(oArgs, FuncCallParms)
        Dim yMsg() As Byte = oParms.yMsg
        System.Threading.Thread.Sleep(1000 * oParms.iSleepDelay)
        If yMsg Is Nothing = False Then MyBase.moUILib.SendMsgToRegion(yMsg)
    End Sub

    Private Sub SendClearEntityProdQueue(ByRef oEntity As BaseEntity)
        Dim yMsg(7) As Byte
        Dim lPos As Int32 = 0

        If goCurrentEnvir Is Nothing = False Then
            If goCurrentEnvir.ClearUnitsProdQueue(oEntity.ObjectID, oEntity.ObjTypeID) = False Then Return
        End If

        System.BitConverter.GetBytes(GlobalMessageCode.eClearEntityProdQueue).CopyTo(yMsg, lPos) : lPos += 2
        oEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Private Sub frmUnitGoto_WindowClosed() Handles Me.WindowClosed
        Try
            miDist = Math.Max(Math.Min(CInt(Val(txtDistance.Caption)), 15000), 100)
        Catch
        End Try
    End Sub

    Private Sub ProcessMoveQueue()
        Dim lOrderCnt As Int32 = 0
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    lOrderCnt += 1
                End If
            End If
        Next X
        Dim oNode As UITreeView.UITreeViewItem = tvwMain.oSelectedNode
        If oNode Is Nothing = True Then Return
        If oNode.oParentNode.lItemData3 = -1 then 
            Dim yMultiMove() As Byte        'used for group move, attack, mining, orders etc...
            Dim lPos As Int32 = 0
            With MoveQueue(0)
                If .MoveType = GlobalMessageCode.eJumpTarget Then
                    ReDim yMultiMove(13 + (lOrderCnt * 6))
                    System.BitConverter.GetBytes(GlobalMessageCode.eJumpTarget).CopyTo(yMultiMove, 0) : lPos += 2
                    System.BitConverter.GetBytes(.WormholeID).CopyTo(yMultiMove, lPos) : lPos += 4
                    System.BitConverter.GetBytes(ObjectType.eSolarSystem).CopyTo(yMultiMove, lPos) : lPos += 2
                    System.BitConverter.GetBytes(.lDestID).CopyTo(yMultiMove, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.iDestTypeID).CopyTo(yMultiMove, lPos) : lPos += 2
                Else

                    ReDim yMultiMove(17 + (lOrderCnt * 6))
                    System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMultiMove, 0) : lPos += 2
                    System.BitConverter.GetBytes(.locX).CopyTo(yMultiMove, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.locZ).CopyTo(yMultiMove, lPos) : lPos += 4
                    ''TODO: eventually, we will want to determine the final facing if the player wants to specify one
                    System.BitConverter.GetBytes(CShort(-1)).CopyTo(yMultiMove, lPos) : lPos += 2
                    System.BitConverter.GetBytes(.lDestID).CopyTo(yMultiMove, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.iDestTypeID).CopyTo(yMultiMove, lPos) : lPos += 2
                End If
            End With
            'Now, go thru each entity and add them to the list
            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                    If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                        goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yMultiMove, lPos) : lPos += 6
                        SendClearEntityProdQueue(goCurrentEnvir.oEntity(X))
                    End If
                End If
            Next X
            If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eMoveRequest, lOrderCnt))
            DoTheMove(yMultiMove)
        Else
            Dim yAlertMsg() As Byte
            ReDim yAlertMsg(0)
            ExecuteAlertSettings(yAlertMsg)
            MyBase.moUILib.SendLenAppendedMsgToRegion(yAlertMsg)
            BuildMoveQueue()
        End If
    End Sub

    Private Sub BuildMoveQueue()
        Dim oNode As UITreeView.UITreeViewItem = tvwMain.oSelectedNode
        If oNode Is Nothing = True Then Return
        Dim yMultiMove() As Byte        'used for group move, attack, mining, orders etc...
        Dim iLen As Int32 = 0
        Dim lPos As Int32 = 0

        Dim lOrderCnt As Int32 = 0
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    lOrderCnt += 1
                End If
            End If
        Next X
        If lOrderCnt = 0 Then Return

        'Build Route
        Dim oParentNode As UITreeView.UITreeViewItem = oNode.oParentNode
        Dim idx As Int32 = 1
        While oParentNode Is Nothing = False
            Dim bGood As Boolean = False
            If oParentNode.lItemData3 = -1 Then Exit While
            ReDim Preserve MoveQueue(MoveQueue.GetUpperBound(0) + 1)
            With MoveQueue(idx)
                Dim lDestID As Int32 = oParentNode.lItemData
                Dim iDestTypeID As Int16 = CShort(oParentNode.lItemData2)

                If oParentNode.sItem.StartsWith("Wormhole To ") Then
                    .MoveType = GlobalMessageCode.eJumpTarget
                    .WormholeID = oParentNode.lItemData3
                    .lDestID = lDestID
                    .iDestTypeID = iDestTypeID
                    For z As Int32 = 0 To moEnvironments.GetUpperBound(0)
                        If moEnvironments(z).iEnvirTypeID = ObjectType.eWormhole AndAlso moEnvironments(z).lEnvirID = .lDestID Then
                            .locX = moEnvironments(z).iLocX
                            .locY = moEnvironments(z).iLocZ
                            bGood = True
                            Exit For
                        End If
                    Next
                    If bGood = False Then
                        Exit Sub
                    End If
                Else
                    'Error!
                End If
            End With
            idx += 1
            oParentNode = oParentNode.oParentNode
        End While

        'Clear Route
        iLen = 10
        ReDim yMultiMove(((iLen + 2) * lOrderCnt) - 1)
        lPos = 0
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    System.BitConverter.GetBytes(iLen).CopyTo(yMultiMove, lPos) : lPos += 2
                    System.BitConverter.GetBytes(GlobalMessageCode.eRemoveRouteItem).CopyTo(yMultiMove, lPos) : lPos += 2
                    System.BitConverter.GetBytes(goCurrentEnvir.oEntity(X).ObjectID).CopyTo(yMultiMove, lPos) : lPos += 4
                    System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yMultiMove, lPos) : lPos += 4
                End If
            End If
        Next
        MyBase.moUILib.SendLenAppendedMsgToPrimary(yMultiMove)

        'Build Route
        idx = 0
        For y As Int32 = MoveQueue.GetUpperBound(0) To 1 Step -1
            iLen = 24
            ReDim yMultiMove(((iLen + 2) * lOrderCnt) - 1)
            lPos = 0
            With MoveQueue(y)
                Dim bGood As Boolean = False
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                        If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                            System.BitConverter.GetBytes(iLen).CopyTo(yMultiMove, lPos) : lPos += 2
                            System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMultiMove, lPos) : lPos += 2
                            System.BitConverter.GetBytes(goCurrentEnvir.oEntity(X).ObjectID).CopyTo(yMultiMove, lPos) : lPos += 4
                            System.BitConverter.GetBytes(idx).CopyTo(yMultiMove, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.lDestID).CopyTo(yMultiMove, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.iDestTypeID).CopyTo(yMultiMove, lPos) : lPos += 2
                            System.BitConverter.GetBytes(.locX).CopyTo(yMultiMove, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.locY).CopyTo(yMultiMove, lPos) : lPos += 4
                            'For z As Int32 = 0 To moEnvironments.GetUpperBound(0)
                            '    If moEnvironments(z).iEnvirTypeID = ObjectType.eWormhole AndAlso moEnvironments(z).lEnvirID = .lDestID Then
                            '        System.BitConverter.GetBytes(moEnvironments(z).iLocX).CopyTo(yMultiMove, lPos) : lPos += 4
                            '        System.BitConverter.GetBytes(moEnvironments(z).iLocZ).CopyTo(yMultiMove, lPos) : lPos += 4
                            '        bGood = True
                            '        Exit For
                            '    End If
                            'Next
                            'If bGood = False Then
                            '    Exit Sub
                            'End If
                        End If
                    End If
                Next X
                MyBase.moUILib.SendLenAppendedMsgToPrimary(yMultiMove)
            End With
            idx += 1
        Next y

        iLen = 24
        ReDim yMultiMove(((iLen + 2) * lOrderCnt) - 1)
        lPos = 0
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    With MoveQueue(0)
                        System.BitConverter.GetBytes(iLen).CopyTo(yMultiMove, lPos) : lPos += 2
                        System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMultiMove, lPos) : lPos += 2
                        System.BitConverter.GetBytes(goCurrentEnvir.oEntity(X).ObjectID).CopyTo(yMultiMove, lPos) : lPos += 4
                        System.BitConverter.GetBytes(idx).CopyTo(yMultiMove, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.lDestID).CopyTo(yMultiMove, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.iDestTypeID).CopyTo(yMultiMove, lPos) : lPos += 2
                        System.BitConverter.GetBytes(.locX).CopyTo(yMultiMove, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.locY).CopyTo(yMultiMove, lPos) : lPos += 4
                        'For z As Int32 = 0 To moEnvironments.GetUpperBound(0)
                        '    If moEnvironments(z).iEnvirTypeID = ObjectType.eWormhole AndAlso moEnvironments(z).lEnvirID = .lDestID Then
                        '        System.BitConverter.GetBytes(moEnvironments(z).iLocX).CopyTo(yMultiMove, lPos) : lPos += 4
                        '        System.BitConverter.GetBytes(moEnvironments(z).iLocZ).CopyTo(yMultiMove, lPos) : lPos += 4
                        '        Exit For
                        '    End If
                        'Next
                    End With
                End If
            End If
        Next X
        MyBase.moUILib.SendLenAppendedMsgToPrimary(yMultiMove)

        'Begin Routes
        iLen = 10
        ReDim yMultiMove(((iLen + 2) * lOrderCnt) - 1)
        lPos = 0
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    System.BitConverter.GetBytes(iLen).CopyTo(yMultiMove, lPos) : lPos += 2
                    System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRouteStatus).CopyTo(yMultiMove, lPos) : lPos += 2
                    System.BitConverter.GetBytes(goCurrentEnvir.oEntity(X).ObjectID).CopyTo(yMultiMove, lPos) : lPos += 4
                    System.BitConverter.GetBytes(-4I).CopyTo(yMultiMove, lPos) : lPos += 4
                End If
            End If
        Next
        MyBase.moUILib.SendLenAppendedMsgToPrimary(yMultiMove)

    End Sub

    Private Structure TAttackers
        Dim oEntity As BaseEntity
    End Structure

    Private Structure TTargets
        Dim oEntity As BaseEntity
        Dim oAttackers() As TAttackers
    End Structure

    Private Sub btnAttackTargets_Click(ByVal sName As String) Handles btnAttackTargets.Click
        Dim iUnitCount As Int32 = 0
        Dim iTargetCount As Int32 = 0
        Dim yRel As Byte
        Dim oAttackers() As TAttackers
        Dim oTargets() As TTargets
        ReDim oAttackers(0)
        ReDim oTargets(0)

        'Now, go thru each entity and add them to the list
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    ReDim Preserve oAttackers(iUnitCount)
                    oAttackers(iUnitCount).oEntity = goCurrentEnvir.oEntity(X)
                    iUnitCount += 1
                End If
            End If
        Next X
        If iUnitCount = 0 Then Exit Sub
        ReDim oTargets(0)
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).OwnerID <> glPlayerID Then
                    yRel = goCurrentPlayer.GetPlayerRelScore(goCurrentEnvir.oEntity(X).OwnerID)
                    If yRel <= elRelTypes.eWar Then
                        If goCurrentEnvir.oEntity(X).yVisibility > eVisibilityType.NoVisibility Then
                            iTargetCount += 1
                            ReDim Preserve oTargets(iTargetCount)
                            oTargets(iTargetCount).oEntity = goCurrentEnvir.oEntity(X)
                            ReDim oTargets(iTargetCount).oAttackers(0)
                        End If
                    End If
                End If
            End If
        Next X
        If iTargetCount = 0 Then Exit Sub

        ' 100 attackers, 1 target, 1 bucket of oAttackers w/ them all
        ' 1 attacker, 100 targets, 1 bucket of oAttackers w/ 1
        ' 100 attackers, 50 targets, 50 buckets of 2
        ' 100 attackers, 49 targets, 49 buckets of 2, except first bucket has 3
        ' 10 attackers, 3 targets, 3 buckets of 2, except first bucket has 3

        'The smaller of the two numbers is the max bucket.  100 units 1 target 1 bucket, 1 unit 100 targets 1 bucket.
        Dim iBuckets As Int32 = Math.Min(iTargetCount, iUnitCount)
        Dim iBucket As Int32 = 1
        For X As Int32 = 0 To iUnitCount - 1
            ReDim Preserve oTargets(iBucket).oAttackers(oTargets(iBucket).oAttackers.GetUpperBound(0) + 1)
            oTargets(iBucket).oAttackers(oTargets(iBucket).oAttackers.GetUpperBound(0)) = oAttackers(X)
            iBucket += 1
            If iBucket > iBuckets Then iBucket = 1
        Next
        Dim yMsg(23) As Byte
        Dim lPos As Int32 = 0

        For X As Int32 = 1 To iBuckets
            lPos = 0
            ReDim yMsg(17 + (oTargets(iBucket).oAttackers.GetUpperBound(0) * 6))
            ReDim yMsg(17 + (oTargets(X).oAttackers.GetUpperBound(0) * 6))
            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMsg, 0) : lPos += 2
            System.BitConverter.GetBytes(CInt(oTargets(X).oEntity.LocX)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CInt(oTargets(X).oEntity.LocZ)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CShort(-1)).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(goCurrentEnvir.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(goCurrentEnvir.ObjTypeID).CopyTo(yMsg, lPos) : lPos += 2
            For Y As Int32 = 1 To oTargets(X).oAttackers.GetUpperBound(0)
                goUILib.AddNotification("Bucket: " & X.ToString & " UnitID: " & oTargets(X).oAttackers(Y).oEntity.ObjectID & " To (" & CInt(oTargets(X).oEntity.LocX).ToString & "," & CInt(oTargets(X).oEntity.LocZ).ToString & ")", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                oTargets(X).oAttackers(Y).oEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            Next Y
            Dim oThread As New Threading.Thread(AddressOf SleepAndMove)
            Dim oParms As FuncCallParms = New FuncCallParms(yMsg)
            oParms.iSleepDelay = X - 1
            oThread.Start(oParms)
        Next
    End Sub

    Private Class FuncCallParms
        Public yMsg() As Byte
        Public iSleepDelay As Int32 = 1
        Public Sub New(ByVal yMultiMove() As Byte)
            yMsg = yMultiMove
        End Sub
    End Class

End Class
