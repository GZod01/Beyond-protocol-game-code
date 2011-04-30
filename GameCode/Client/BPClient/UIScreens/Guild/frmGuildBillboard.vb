Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmGuildBillboard
    Inherits UIWindow

    Private fraIcon As UIWindow
    Private lblName As UILabel
    Private txtBillboard As UITextBox
    Private btnApply As UIButton

    'Private fraRecruitIcon As UIWindow

    Private rcBack As Rectangle
    Private rcFore1 As Rectangle
    Private rcFore2 As Rectangle
    Private clrBack As System.Drawing.Color
    Private clrFore1 As System.Drawing.Color
    Private clrFore2 As System.Drawing.Color
    Private mlCurrIcon As Int32 = -1

    Private mlIcon As Int32 = -1
    Private mlGuildID As Int32 = -1
    Private miRecruitment As Int16 = 0

    Private mbApplyClicked As Boolean = False
    Private mlSlotID As Int32 = -1

    Public Sub New(ByRef oUILib As UILib, ByVal lSlotNum As Int32)
        MyBase.New(oUILib)

        mlSlotID = lSlotNum

        Dim lTop As Int32
        Dim lLeft As Int32
        Select Case lSlotNum
            Case 1  'up left
                lTop = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 260
                lLeft = 5
            Case 2  'bottom left
                lLeft = 5
                lTop = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) + 130
            Case 3  'top right
                'lTop = 78
                lTop = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 260
                lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 210
            Case 4  'midup right
                'lTop = 211
                lTop = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 130
                lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 210
            Case 5  'middn right
                'lTop = 344
                lTop = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) + 1
                lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 210
            Case Else  'bottom right
                lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 210
                'lTop = 477
                lTop = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) + 130
        End Select

        'frmGuildBillboard initial props
        With Me
            .ControlName = "frmGuildBillboard" & lSlotNum
            .Left = lLeft '130
            .Top = lTop '128
            .Width = 205
            .Height = 127
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
        End With

        'fraIcon initial props
        fraIcon = New UIWindow(oUILib)
        With fraIcon
            .ControlName = "fraIcon"
            .Left = 5
            .Top = 3
            .Width = 67
            .Height = 67
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .BorderLineWidth = 1
        End With
        Me.AddChild(CType(fraIcon, UIControl))

        'lblName initial props
        lblName = New UILabel(oUILib)
        With lblName
            .ControlName = "lblName"
            .Left = 75
            .Top = 2
            .Width = 126
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Space Available"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(1, DrawTextFormat) Or DrawTextFormat.WordBreak
        End With
        Me.AddChild(CType(lblName, UIControl))

        'txtBillboard initial props
        txtBillboard = New UITextBox(oUILib)
        With txtBillboard
            .ControlName = "txtBillboard"
            .Left = 3
            .Top = 75
            .Width = 199
            .Height = 49
            .Enabled = True
            .Visible = True
            .Caption = "Rent this Space to advertise your guild!"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .Locked = True
            .MultiLine = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtBillboard, UIControl))

        btnApply = New UIButton(oUILib)
        With btnApply
            .ControlName = "btnApply"
            .Left = lblName.Left
            .Top = lblName.Top + lblName.Height + 18
            .Width = lblName.Width
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Apply"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnApply, UIControl))

        ''fraRecruitIcon initial props
        'fraRecruitIcon = New UIWindow(oUILib)
        'With fraRecruitIcon
        '    .ControlName = "fraRecruitIcon"
        '    .Left = 76
        '    .Top = 43
        '    .Width = 124
        '    .Height = 27
        '    .Enabled = True
        '    .Visible = True
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceFillColor
        '    .FullScreen = False
        'End With
        'Me.AddChild(CType(fraRecruitIcon, UIControl))

        AddHandler btnApply.Click, AddressOf btnApply_Click

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    'Public Sub SetValues(ByVal lGuildID As Int32, ByVal iRecruit As eiRecruitmentFlags, ByVal sBillboard As String, ByVal lIcon As Int32)
    '    mlGuildID = lGuildID
    '    lblName.Caption = GetCacheObjectValue(lGuildID, ObjectType.eGuild)
    '    txtBillboard.Caption = sBillboard
    '    mlIcon = lIcon
    '    miRecruitment = iRecruit
    'End Sub

    Private Sub frmGuildBillboard_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim pt As Point = Me.GetAbsolutePosition()
        Dim lX As Int32 = lMouseX - pt.X
        Dim lY As Int32 = lMouseY - pt.Y

        If lblName Is Nothing Then Return
        Dim lTop As Int32 = lblName.Top + lblName.Height + 2

        If lY >= lTop AndAlso lY < lTop + 16 Then
            If lX > 75 Then
                lY = (lX - 75) \ 17
                If (miRecruitment And eiRecruitmentFlags.RecruitDiplomacy) <> 0 Then
                    If lY = 0 Then
                        MyBase.moUILib.SetToolTip("Recruiting Diplomatic Specialists", lMouseX, lMouseY)
                        Return
                    End If
                    lY -= 1
                End If
                If (miRecruitment And eiRecruitmentFlags.RecruitEspionage) <> 0 Then
                    If lY = 0 Then
                        MyBase.moUILib.SetToolTip("Recruiting Espionage Specialists", lMouseX, lMouseY)
                        Return
                    End If
                    lY -= 1
                End If
                If (miRecruitment And eiRecruitmentFlags.RecruitMilitary) <> 0 Then
                    If lY = 0 Then
                        MyBase.moUILib.SetToolTip("Recruiting Military Specialists", lMouseX, lMouseY)
                        Return
                    End If
                    lY -= 1
                End If

                If (miRecruitment And eiRecruitmentFlags.RecruitProduction) <> 0 Then
                    If lY = 0 Then
                        MyBase.moUILib.SetToolTip("Recruiting Production Specialists", lMouseX, lMouseY)
                        Return
                    End If
                    lY -= 1
                End If
                If (miRecruitment And eiRecruitmentFlags.RecruitResearch) <> 0 Then
                    If lY = 0 Then
                        MyBase.moUILib.SetToolTip("Recruiting Research Specialists", lMouseX, lMouseY)
                        Return
                    End If
                    lY -= 1
                End If
                If (miRecruitment And eiRecruitmentFlags.RecruitTrade) <> 0 Then
                    If lY = 0 Then
                        MyBase.moUILib.SetToolTip("Recruiting Trade Specialists", lMouseX, lMouseY)
                        Return
                    End If
                    lY -= 1
                End If
            End If
        End If
    End Sub

    Private Sub frmGuildBillboard_OnNewFrame() Handles Me.OnNewFrame
        Try
            If glCurrentEnvirView <> CurrentView.ePlanetMapView Then
                MyBase.moUILib.RemoveWindow(Me.ControlName)
                Return
            End If

            Dim oPlanet As Planet = Nothing
            If goCurrentEnvir Is Nothing Then Return
            If goCurrentEnvir.oGeoObject Is Nothing Then Return
            If goCurrentEnvir.ObjTypeID <> ObjectType.ePlanet Then
                If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx > -1 Then
                    oPlanet = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx)
                End If
            Else
                If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID <> ObjectType.ePlanet Then Return
                oPlanet = CType(goCurrentEnvir.oGeoObject, Planet)
            End If

            If oPlanet Is Nothing Then Return
            If mlSlotID = -1 Then Return

            With oPlanet.uBillboards(mlSlotID - 1)
                miRecruitment = .iRecruitFlags
                mlGuildID = .lGuildID
                mlIcon = .lIcon
                If mlIcon <> mlCurrIcon Then Me.IsDirty = True
                If txtBillboard.Caption <> .BillboardText Then txtBillboard.Caption = .BillboardText
            End With

            If (miRecruitment And eiRecruitmentFlags.GuildRecruiting) <> 0 Then
                If btnApply.Enabled = False AndAlso mbApplyClicked = False Then
                    btnApply.Enabled = True
                    btnApply.Caption = "Apply"
                    btnApply.ToolTipText = "This guild is recruiting, click to apply."
                End If
            ElseIf btnApply.Enabled = True Then
                btnApply.Enabled = False
                btnApply.Caption = "Not Recruiting" 'ToolTipText = "This guild is not presently recruiting"
            End If

            Dim sName As String = GetCacheObjectValue(mlGuildID, ObjectType.eGuild)
            If sName <> lblName.Caption Then lblName.Caption = sName
        Catch
        End Try
    End Sub

    Private Sub frmGuildBillboard_OnRenderEnd() Handles Me.OnRenderEnd

        If mlCurrIcon <> mlIcon Then
            SetIconDetails()
        End If
        Device.IsUsingEventHandlers = False
        If ctlDiplomacy.moIconBack Is Nothing OrElse ctlDiplomacy.moIconBack.Disposed = True Then ctlDiplomacy.moIconBack = goResMgr.GetTexture("DipPlayerBack.bmp", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
        If ctlDiplomacy.moIconFore Is Nothing OrElse ctlDiplomacy.moIconFore.Disposed = True Then ctlDiplomacy.moIconFore = goResMgr.GetTexture("DipPlayerFore.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
        Device.IsUsingEventHandlers = True
        If ctlDiplomacy.moSprite Is Nothing OrElse ctlDiplomacy.moSprite.Disposed = True Then
            Device.IsUsingEventHandlers = False
            ctlDiplomacy.moSprite = New Sprite(MyBase.moUILib.oDevice)
            Device.IsUsingEventHandlers = True
        End If
        If ctlDiplomacy.moIconBack Is Nothing OrElse ctlDiplomacy.moIconFore Is Nothing OrElse ctlDiplomacy.moIconBack.Disposed = True OrElse ctlDiplomacy.moIconFore.Disposed = True Then Return
        ctlDiplomacy.moSprite.Begin(SpriteFlags.AlphaBlend)
        Dim rcDest As Rectangle = New Rectangle(0, 0, 64, 64)
        Dim ptDest As Point
        ptDest.X = Me.Left + fraIcon.Left + 3
        ptDest.Y = Me.Top + fraIcon.Top + 3

        ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moIconBack, rcBack, rcDest, ptDest, clrBack)
        ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moIconFore, rcFore1, rcDest, ptDest, clrFore1)
        ctlDiplomacy.moSprite.Draw2D(ctlDiplomacy.moIconFore, rcFore2, rcDest, ptDest, clrFore2)


        Dim lCurrLeft As Int32 = 75
        Dim lTop As Int32 = lblName.Top + lblName.Height + 2
        'rcSrc(4) = New Rectangle(32, 128, 32, 32)      - diplomatic
        If (miRecruitment And eiRecruitmentFlags.RecruitDiplomacy) <> 0 Then
            ctlDiplomacy.moSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eQuickbar_Diplomacy), New Rectangle(lCurrLeft, lTop, 16, 16), New Point(lCurrLeft * 2, lTop * 2), System.Drawing.Color.FromArgb(255, 255, 255, 255))
            lCurrLeft += 17
        End If
        'rcSrc(8) = New Rectangle(64, 160, 32, 32)      - espionage
        If (miRecruitment And eiRecruitmentFlags.RecruitEspionage) <> 0 Then
            ctlDiplomacy.moSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eQuickbar_Agent), New Rectangle(lCurrLeft, lTop, 16, 16), New Point(lCurrLeft * 2, lTop * 2), System.Drawing.Color.FromArgb(255, 255, 255, 255))
            lCurrLeft += 17
        End If
        'rcSrc(2) = New Rectangle(0, 160, 32, 32)       - military
        If (miRecruitment And eiRecruitmentFlags.RecruitMilitary) <> 0 Then
            ctlDiplomacy.moSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eQuickbar_Battlegroup), New Rectangle(lCurrLeft, lTop, 16, 16), New Point(lCurrLeft * 2, lTop * 2), System.Drawing.Color.FromArgb(255, 255, 255, 255))
            lCurrLeft += 17
        End If
        'rcSrc(7) = New Rectangle(64, 128, 32, 32)      - production
        If (miRecruitment And eiRecruitmentFlags.RecruitProduction) <> 0 Then
            ctlDiplomacy.moSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eQuickbar_Mining), New Rectangle(lCurrLeft, lTop, 16, 16), New Point(lCurrLeft * 2, lTop * 2), System.Drawing.Color.FromArgb(255, 255, 255, 255))
            lCurrLeft += 17
        End If
        'rcSrc(10) = New Rectangle(96, 192, 32, 32)     - research
        If (miRecruitment And eiRecruitmentFlags.RecruitResearch) <> 0 Then
            ctlDiplomacy.moSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eQuickbar_ColonyResearch), New Rectangle(lCurrLeft, lTop, 16, 16), New Point(lCurrLeft * 2, lTop * 2), System.Drawing.Color.FromArgb(255, 255, 255, 255))
            lCurrLeft += 17
        End If
        'rcSrc(3) = New Rectangle(32, 96, 32, 32)       - trade
        If (miRecruitment And eiRecruitmentFlags.RecruitTrade) <> 0 Then
            ctlDiplomacy.moSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eQuickbar_Trade), New Rectangle(lCurrLeft, lTop, 16, 16), New Point(lCurrLeft * 2, lTop * 2), System.Drawing.Color.FromArgb(255, 255, 255, 255))
            lCurrLeft += 17
        End If

        ctlDiplomacy.moSprite.End()
    End Sub
    Private Sub SetIconDetails()
        Dim yBackImg As Byte
        Dim yBackClr As Byte
        Dim yFore1Img As Byte
        Dim yFore1Clr As Byte
        Dim yFore2Img As Byte
        Dim yFore2Clr As Byte

        mlCurrIcon = mlIcon

        PlayerIconManager.FillIconValues(mlIcon, yBackImg, yBackClr, yFore1Img, yFore1Clr, yFore2Img, yFore2Clr)

        rcBack = PlayerIconManager.ReturnImageRectangle(yBackImg, True)
        rcFore1 = PlayerIconManager.ReturnImageRectangle(yFore1Img, False)
        rcFore2 = PlayerIconManager.ReturnImageRectangle(yFore2Img, False)

        clrBack = PlayerIconManager.GetColorValue(yBackClr)
        clrFore1 = PlayerIconManager.GetColorValue(yFore1Clr)
        clrFore2 = PlayerIconManager.GetColorValue(yFore2Clr)
    End Sub

    Private Sub btnApply_Click(ByVal sName As String)
        If (miRecruitment And eiRecruitmentFlags.GuildRecruiting) = 0 Then
            MyBase.moUILib.AddNotification("That guild is not currently recruiting.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'confirming... send our acceptance
        mbApplyClicked = True
        SendGuildMemberStatusMsg(glPlayerID, GuildMemberState.Applied, mlGuildID)
        btnApply.Enabled = False

        MyBase.moUILib.AddNotification("Application sent.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub SendGuildMemberStatusMsg(ByVal lMemberID As Int32, ByVal lStatusUpdate As Int32, ByVal lGuildID As Int32)
        Dim yMsg(13) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eGuildMemberStatus).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lGuildID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMemberID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lStatusUpdate).CopyTo(yMsg, lPos) : lPos += 4
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub
End Class