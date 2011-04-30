Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmCustomTitle
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents btnClose As UIButton
    Private lnDiv1 As UILine
    Private lblConquest As UILabel
    Private lblTrade As UILabel
    Private lblScience As UILabel
    Private lblDiplomacy As UILabel

    Private mrcItems() As Rectangle
    Private myItemRank() As Byte
    Private mlItemPerm() As Int32
    Private mlItemUB As Int32 = 27

    Private mclrForeColor As System.Drawing.Color
    Private mclrDisabledFill As System.Drawing.Color
    Private mclrDisabledBorder As System.Drawing.Color

    Private mclrTypeFill() As System.Drawing.Color
    Private mclrTypeBorder() As System.Drawing.Color

    Private mlLastForcedRefresh As Int32 = 0

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmCustomTitle initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eCustomTitle
            .ControlName = "frmCustomTitle"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256
            .Width = 512
            .Height = 390 '512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 2
            .bRoundedBorder = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 2
            .Width = 186
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Custom Title Management"
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
            .Left = 487
            .Top = 2
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
            .Left = 1
            .Top = 25
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblConquest initial props
        lblConquest = New UILabel(oUILib)
        With lblConquest
            .ControlName = "lblConquest"
            .Left = 10
            .Top = 45
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "CONQUEST"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblConquest, UIControl))

        'lblTrade initial props
        lblTrade = New UILabel(oUILib)
        With lblTrade
            .ControlName = "lblTrade"
            .Left = 140
            .Top = 45
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "TRADE"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTrade, UIControl))

        'lblScience initial props
        lblScience = New UILabel(oUILib)
        With lblScience
            .ControlName = "lblScience"
            .Left = 270
            .Top = 45
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "SCIENCE"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblScience, UIControl))

        'lblDiplomacy initial props
        lblDiplomacy = New UILabel(oUILib)
        With lblDiplomacy
            .ControlName = "lblDiplomacy"
            .Left = 400
            .Top = 45
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "DIPLOMACY"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDiplomacy, UIControl))

        FillBoxRects()

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub FillBoxRects()
        Dim lBoxWidth As Int32 = 100
        Dim lBoxHeight As Int32 = 32

        mclrDisabledFill = System.Drawing.Color.FromArgb(255, 102, 102, 102)
        mclrDisabledBorder = System.Drawing.Color.FromArgb(255, 192, 192, 192)
        mclrForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)

        ReDim mclrTypeBorder(4)
        ReDim mclrTypeFill(4)
        mclrTypeBorder(0) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        mclrTypeFill(0) = System.Drawing.Color.FromArgb(255, 128, 0, 0)
        mclrTypeBorder(1) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
        mclrTypeFill(1) = System.Drawing.Color.FromArgb(255, 92, 92, 0)
        mclrTypeBorder(2) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        mclrTypeFill(2) = System.Drawing.Color.FromArgb(255, 0, 92, 0)
        mclrTypeBorder(3) = System.Drawing.Color.FromArgb(255, 0, 0, 255)
        mclrTypeFill(3) = System.Drawing.Color.FromArgb(255, 0, 0, 128)
        mclrTypeBorder(4) = System.Drawing.Color.FromArgb(255, 192, 0, 0)
        mclrTypeFill(4) = System.Drawing.Color.FromArgb(255, 128, 0, 0)

        'Now, our other items.... 0-6
        ReDim mrcItems(mlItemUB)
        ReDim myItemRank(mlItemUB)
        ReDim mlItemPerm(mlItemUB)
        Dim lIdx As Int32 = -1
        Dim lBaseTop As Int32 = lblConquest.Top + lblConquest.Height + 2

        For lType As Int32 = 0 To 3
            Dim yBaseVal As Byte = 0
            Dim lBasePower As Int32 = 21

            Dim lLeft As Int32 = 10 + (lType * 130)

            Select Case lType
                Case 0 'Rank
                Case 1 'Trade rank
                    yBaseVal = Player.eyCustomRank.TraderShift
                    lBasePower = 14
                Case 2 'Research rank
                    yBaseVal = Player.eyCustomRank.ResearcherShift
                    lBasePower = 0
                Case 3 'Diplomacy Rank
                    yBaseVal = Player.eyCustomRank.DiplomacyShift
                    lBasePower = 7
            End Select

            'Now, loop through 0 to 6 for the different levels
            For lLevel As Int32 = 0 To 6
                lIdx += 1

                Dim yRank As Byte = CByte(yBaseVal Or lLevel)
                Dim lFinalPower As Int32 = lBasePower + lLevel
                Dim lResultingPermission As Int32 = CInt(Math.Pow(2, lFinalPower))

                mrcItems(lIdx) = New Rectangle(lLeft, lBaseTop + ((6 - lLevel) * 45), lBoxWidth, lBoxHeight)
                myItemRank(lIdx) = yRank
                mlItemPerm(lIdx) = lResultingPermission
            Next lLevel
        Next lType
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Function GetClickedOnRank(ByVal lMouseX As Int32, ByVal lMouseY As Int32) As Int32

        lMouseX -= Me.GetAbsolutePosition.X
        lMouseY -= Me.GetAbsolutePosition.Y

        For X As Int32 = 0 To mlItemUB
            If mrcItems(X).Contains(lMouseX, lMouseY) = True Then
                Return X
            End If
        Next X
        Return -1
    End Function

    Private Sub frmCustomTitle_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        Dim lClickedItem As Int32 = GetClickedOnRank(lMouseX, lMouseY)
        If lClickedItem = -1 Then Return

        'Ok, clicked an item, check if it is enabled


    End Sub

    Private Sub frmCustomTitle_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim lClickedItem As Int32 = GetClickedOnRank(lMouseX, lMouseY)
        If lClickedItem = -1 Then
            MyBase.moUILib.SetToolTip(False)
        Else
            MyBase.moUILib.SetToolTip(GetNextTitleRequirements(myItemRank(lClickedItem)), lMouseX, lMouseY)
        End If
    End Sub

    Private Function GetNextTitleRequirements(ByVal yNextTitle As Byte) As String

        Dim bDiplomacy As Boolean = (yNextTitle And Player.eyCustomRank.DiplomacyShift) <> 0
        Dim bResearch As Boolean = (yNextTitle And Player.eyCustomRank.ResearcherShift) <> 0
        Dim bTrade As Boolean = (yNextTitle And Player.eyCustomRank.TraderShift) <> 0

        If bDiplomacy = True Then yNextTitle = yNextTitle Xor Player.eyCustomRank.DiplomacyShift
        If bResearch = True Then yNextTitle = yNextTitle Xor Player.eyCustomRank.ResearcherShift
        If bTrade = True Then yNextTitle = yNextTitle Xor Player.eyCustomRank.TraderShift

        Select Case yNextTitle
            Case Player.eyCustomRank.Magistrate
                If bDiplomacy = True Then
                    Return "First Ally Relationship or Better" & vbCrLf & "with a non guildmate"
                ElseIf bResearch = True Then
                    Return "1 Research Lab outside the tutorial"
                ElseIf bTrade = True Then
                    Return "A Tradepost outside the tutorial"
                Else : Return "Default: players start with this"
                End If
            Case Player.eyCustomRank.Governor
                If bDiplomacy = True Then
                    Return "5 Allies and 1 Blood Ally" & vbCrLf & "All of which are not guildmates"
                ElseIf bResearch = True Then
                    Return "2 Research Labs and 10 Minerals Discovered"
                ElseIf bTrade = True Then
                    Return "5 Tradeposts in same system"
                Else : Return "At least 200,000 Total Population"
                End If
            Case Player.PlayerRank.Overseer     '2
                If bDiplomacy = True Then
                    Return "A full faction member list"
                ElseIf bResearch = True Then
                    Return "2 High Production Research Labs" & vbCrLf & "At least 30 Discovered Minerals" & vbCrLf & "1 Alloy Designed and Researched"
                ElseIf bTrade = True Then
                    Return "10 concurrent sell orders of 5 market types"
                Else : Return "At least 500,000 Total Population"
                End If
            Case Player.PlayerRank.Duke         '3
                If bDiplomacy = True Then
                    Return "Endorse 20 legislations in the Emperor's Chamber"
                ElseIf bResearch = True Then
                    Return "1 of each Component Type that exceeds starting techs"
                ElseIf bTrade = True Then
                    Return "10 Sell Orders of 10 Market Types in single tradepost"
                Else : Return "3/4 Population on a single planet (majority owner)" & vbCrLf & "At least 2 planet-based colonies"
                End If
            Case Player.PlayerRank.Baron        '4
                If bDiplomacy = True Then
                    Return "Your Proposal passes on the Senate Floor"
                ElseIf bResearch = True Then
                    'Return "1 of each Component Type that exceeds starting techs"
                    Return "Researching a Single Technology on a massive scale"
                ElseIf bTrade = True Then
                    Return "20 Sell Orders of 10 Market Types in single tradepost"
                Else : Return "3/4 Population on two planets (majority owner)" & vbCrLf & "At least 3 planet-based colonies"
                End If
                'Return "3/4 Population on at least 2 planets (majority owner)" & vbCrLf & "At least 3 planet-based colonies"
            Case Player.PlayerRank.King
                If bDiplomacy = True Then
                    Return "Non Guilded Blood Allies combined control 1 system"
                ElseIf bResearch = True Then
                    Return "2 of each Component Type that exceeds starting techs" & vbCrLf & "20 Special Projects Researched"
                ElseIf bTrade = True Then
                    Return "Tradepost on all planets in 1 system with 5 sell orders each"
                Else : Return "3/4 Population on more than half" & vbCrLf & "  of the planets in the same system"
                End If
            Case Player.PlayerRank.Emperor      '6
                If bDiplomacy = True Then
                    Return "Non Guilded Blood Allies combined control 5 systems"
                ElseIf bResearch = True Then
                    Return "5 of each Component Type that exceeds starting techs" & vbCrLf & "50 Special Projects Researched"
                ElseIf bTrade = True Then
                    Return "Tradepost on every planet in 3 systems with 5 sell orders each"
                Else : Return "3/4 Population on all planets in the same system" & vbCrLf & "3/4 Population on a planet outside of the owned system"
                End If
            Case Else
                Return ""
        End Select
    End Function

    Private Sub frmCustomTitle_OnNewFrame() Handles Me.OnNewFrame
        If mlLastForcedRefresh = 0 OrElse glCurrentCycle - mlLastForcedRefresh > 150 Then
            Me.IsDirty = True
            mlLastForcedRefresh = glCurrentCycle
        End If
    End Sub

    Private Sub frmCustomTitle_OnRender() Handles Me.OnRender
        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return

        Using oSprite As New Sprite(MyBase.moUILib.oDevice)
            Dim bBegun As Boolean = False
            For X As Int32 = 0 To mlItemUB
                'ok, let's do it
                If bBegun = False Then
                    bBegun = True
                    oSprite.Begin(SpriteFlags.AlphaBlend)
                End If

                Dim clrFill As System.Drawing.Color = mclrDisabledFill
                If (goCurrentPlayer.lCustomTitlePermission And mlItemPerm(X)) <> 0 Then
                    'ok, go permission
                    If (myItemRank(X) And Player.eyCustomRank.DiplomacyShift) <> 0 Then
                        clrFill = mclrTypeFill(3)
                    ElseIf (myItemRank(X) And Player.eyCustomRank.ResearcherShift) <> 0 Then
                        clrFill = mclrTypeFill(2)
                    ElseIf (myItemRank(X) And Player.eyCustomRank.TraderShift) <> 0 Then
                        clrFill = mclrTypeFill(1)
                    Else
                        clrFill = mclrTypeFill(0)
                    End If
                ElseIf goCurrentPlayer.yPlayerTitle = X Then
                    clrFill = mclrTypeFill(4)
                End If
                DoMultiColorFill(mrcItems(X), clrFill, mrcItems(X).Location, oSprite)
            Next X
            If bBegun = True Then oSprite.End()
        End Using

        Using oFont As New Font(MyBase.moUILib.oDevice, lblConquest.GetFont)
            'ok, now, we need to render our text...
            For X As Int32 = 0 To mlItemUB
                Dim clrBorder As System.Drawing.Color = mclrDisabledBorder
                If (goCurrentPlayer.lCustomTitlePermission And mlItemPerm(X)) <> 0 Then 'OrElse True = True Then
                    'ok, go permission
                    If (myItemRank(X) And Player.eyCustomRank.DiplomacyShift) <> 0 Then
                        clrBorder = mclrTypeBorder(3)
                    ElseIf (myItemRank(X) And Player.eyCustomRank.ResearcherShift) <> 0 Then
                        clrBorder = mclrTypeBorder(2)
                    ElseIf (myItemRank(X) And Player.eyCustomRank.TraderShift) <> 0 Then
                        clrBorder = mclrTypeBorder(1)
                    Else
                        clrBorder = mclrTypeBorder(0)
                    End If
                ElseIf goCurrentPlayer.yPlayerTitle = X Then
                    clrBorder = mclrTypeBorder(4)
                End If


                Dim sItem As String = Player.GetPlayerTitle(myItemRank(X), goCurrentPlayer.bIsMale)
                RenderBox(mrcItems(X), 2, clrBorder)
                oFont.DrawText(Nothing, sItem, mrcItems(X), DrawTextFormat.Center Or DrawTextFormat.VerticalCenter Or DrawTextFormat.WordBreak, mclrForeColor)
            Next X
        End Using
    End Sub

    Private Sub DoMultiColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point, ByRef oSpr As Sprite)
        Dim rcSrc As Rectangle

        Dim fX As Single
        Dim fY As Single

        If rcDest.Width = 0 OrElse rcDest.Height = 0 OrElse MyBase.moUILib.oInterfaceTexture Is Nothing Then Exit Sub

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
End Class
