Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmSelectRespawn
    Inherits UIWindow

    Private Class SystemItemData
        Public lID As Int32
        Public sName As String
        Public lPlanetCount As Int32
        Public lActiveWars As Int32

        Public lMinCacheCount As Int32
        Public lAvgMinCacheConc As Int32
        Public blTotalMineralQty As Int64

        Public lMagistrateCnt As Int32
        Public lGovernorCnt As Int32
        Public lOverseerCnt As Int32
        Public lDukeCnt As Int32
        Public lBaronCnt As Int32
        Public lKingCnt As Int32
        Public lEmperorCnt As Int32

        Public lWHLinks As Int32

        Public lPlanetOwnerUB As Int32 = -1
        Public lPlanetOwnerID(-1) As Int32
        Public yPlanetOwnerCnt(-1) As Byte

    End Class
    Private mcolData As New Collection

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblDesc As UILabel
    Private lnDiv2 As UILine
    'Private lblAvail As UILabel
    Private moSystemItem() As ctlSystemItem
    Private WithEvents vscrScroll As UIScrollBar
    Private WithEvents btnSubmit As UIButton
    Private WithEvents btnCancel As UIButton

    Private Shared mrcWH As Rectangle = New Rectangle(5, 30, 18, 18)            'guild's O icon
    Private Shared mrcWars As Rectangle = New Rectangle(70, 30, 18, 18)         'Agent sword icon
    Private Shared mrcMinConc As Rectangle = New Rectangle(5, 50, 18, 18)       'avail resources quickbar
    Private Shared mrcMinCount As Rectangle = New Rectangle(70, 50, 18, 18)     'hull builder cargo icon (num of min caches)
    Private Shared mrcTotalMins As Rectangle = New Rectangle(5, 70, 18, 18)     'mining window quickbar (sum of mineral cache qty)

    Private mlSelectedID As Int32 = Int32.MinValue

    Private moAgentIcons As Texture = Nothing

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmSelectRespawn initial props
        With Me
            .ControlName = "frmSelectRespawn"
            .Left = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 256
            .Top = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 256
            .Width = 511
            .Height = 511
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .Resizable = False
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 171
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select Respawn System"
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
            .Left = Me.BorderLineWidth \ 2
            .Top = 25
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblDesc initial props
        lblDesc = New UILabel(oUILib)
        With lblDesc
            .ControlName = "lblDesc"
            .Left = 1
            .Top = 30
            .Width = 511
            .Height = 54
            .Enabled = True
            .Visible = True
            .Caption = "You select what system you wish to respawn in. Systems listed are typically not empty. You will be given a temporary space engineer that can build your initial command center. Once built, your death fund will begin and the space engineer will be lost."
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.WordBreak Or DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblDesc, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = Me.BorderLineWidth \ 2
            .Top = 90
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        ''lblAvail initial props
        'lblAvail = New UILabel(oUILib)
        'With lblAvail
        '    .ControlName = "lblAvail"
        '    .Left = 5
        '    .Top = 95
        '    .Width = 130
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Available Systems"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(4, DrawTextFormat)
        'End With
        'Me.AddChild(CType(lblAvail, UIControl))

        'vscrScroll initial props
        vscrScroll = New UIScrollBar(oUILib, True)
        With vscrScroll
            .ControlName = "vscrScroll"
            .Left = 485
            .Top = 95
            .Width = 22
            .Height = 384
            .Enabled = True
            .Visible = True
            .Value = 0
            .MaxValue = 0
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 1
            .ReverseDirection = True
        End With
        Me.AddChild(CType(vscrScroll, UIControl))

        'btnSubmit initial props
        btnSubmit = New UIButton(oUILib)
        With btnSubmit
            .ControlName = "btnSubmit"
            .Left = 146
            .Top = 483
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
            .Left = 266
            .Top = btnSubmit.Top
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

        'Ok, get our count
        Dim lCnt As Int32 = vscrScroll.Height \ 96
        ReDim moSystemItem(lCnt - 1)
        For X As Int32 = 0 To moSystemItem.GetUpperBound(0)
            moSystemItem(X) = New ctlSystemItem(oUILib)
            With moSystemItem(X)
                .ControlName = "moSystemItem(" & X & ")"
                .Left = 5
                .Top = vscrScroll.Top + (X * 96)
                .Width = 475
                .Height = 96
                .Enabled = True
                .Visible = False
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .Resizable = False
            End With
            AddHandler moSystemItem(X).OnSelect, AddressOf moSystemItem_OnSelect
            Me.AddChild(CType(moSystemItem(X), UIControl))
        Next X

        oUILib.RemoveWindow(Me.ControlName)
        oUILib.AddWindow(Me)

        'Now, send off a msg to request the respawn list
        Dim yMsg(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRespawnSelection).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 2)
        goUILib.SendMsgToPrimary(yMsg)

        goUILib.RetainTooltip = False
    End Sub

    Private Sub frmSelectRespawn_OnNewFrame() Handles Me.OnNewFrame
        GFXEngine.bRenderInProgress = False
        If vscrScroll Is Nothing = False AndAlso moSystemItem Is Nothing = False Then
            If mcolData Is Nothing = False Then
                Dim lCurCnt As Int32 = vscrScroll.MaxValue + moSystemItem.Length
                If lCurCnt <> mcolData.Count Then
                    vscrScroll.MaxValue = Math.Max(0, mcolData.Count - moSystemItem.Length)
                End If
            End If

            For X As Int32 = 0 To moSystemItem.GetUpperBound(0)
                If moSystemItem(X) Is Nothing = False AndAlso moSystemItem(X).Visible = True Then
                    moSystemItem(X).Process_OnNewFrame()
                End If
            Next X
        End If

        Dim lIdx As Int32 = -1
        Dim lDisplayIdx As Int32 = -1
        For Each oItem As SystemItemData In mcolData
            lIdx += 1
            If lIdx >= vscrScroll.Value Then
                lDisplayIdx += 1

                If moSystemItem(lDisplayIdx).Visible = False Then moSystemItem(lDisplayIdx).Visible = True
                If moSystemItem(lDisplayIdx).GetCurrentID <> oItem.lID Then moSystemItem(lDisplayIdx).SetFromSystemData(oItem)
                moSystemItem(lDisplayIdx).UpdateSelectState(oItem.lID = mlSelectedID)
                If lDisplayIdx = moSystemItem.GetUpperBound(0) Then Exit For
            End If
        Next
    End Sub

    Private Sub frmSelectRespawn_OnRender() Handles Me.OnRender

    End Sub

    Private Sub frmSelectRespawn_OnRenderEnd() Handles Me.OnRenderEnd


        'Now, render our icons on the controls
        Dim oGuildTex As Texture = goResMgr.GetTexture("GuildIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "gi.pak")
        Dim oHullTex As Texture = goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")

        If moAgentIcons Is Nothing OrElse moAgentIcons.Disposed = True Then
            moAgentIcons = goResMgr.LoadScratchTexture("AgentIcons.dds", "Interface.pak")
        End If

        Dim rcWHSrc As Rectangle = New Rectangle(0, 16, 16, 16)
        Dim rcWarSrc As Rectangle = New Rectangle(32, 48, 16, 16)
        Dim rcMinCntSrc As Rectangle = New Rectangle(32, 0, 16, 16)
        Dim clrWH As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 192, 64, 192)

        'Dim bAlpha As Boolean = GFXEngine.moDevice.RenderState.AlphaBlendEnable
        'GFXEngine.moDevice.RenderState.AlphaBlendEnable = False
        'moAgentIcons = goResMgr.LoadScratchTexture("AgentIcons.dds", "Interface.pak")
        For X As Int32 = 0 To moSystemItem.GetUpperBound(0)
            If moSystemItem(X) Is Nothing = False AndAlso moSystemItem(X).Visible = True Then
                Dim destRect As Rectangle = mrcWH
                destRect.X += moSystemItem(X).Left : destRect.Y += moSystemItem(X).Top
                BPSprite.Draw2DOnce(GFXEngine.moDevice, oGuildTex, rcWHSrc, destRect, clrWH, 64, 64)

                destRect = mrcWars
                destRect.X += moSystemItem(X).Left : destRect.Y += moSystemItem(X).Top
                BPSprite.Draw2DOnce(GFXEngine.moDevice, moAgentIcons, rcWarSrc, destRect, Color.White, 128, 128)

                destRect = mrcMinConc
                destRect.X += moSystemItem(X).Left : destRect.Y += moSystemItem(X).Top
                BPSprite.Draw2DOnce(GFXEngine.moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eQuickbar_AvailResources), destRect, Color.White, 256, 256)

                destRect = mrcMinCount
                destRect.X += moSystemItem(X).Left : destRect.Y += moSystemItem(X).Top
                BPSprite.Draw2DOnce(GFXEngine.moDevice, oHullTex, rcMinCntSrc, destRect, Color.White, 64, 64)

                destRect = mrcTotalMins
                destRect.X += moSystemItem(X).Left : destRect.Y += moSystemItem(X).Top
                BPSprite.Draw2DOnce(GFXEngine.moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eQuickbar_Mining), destRect, Color.White, 256, 256)
            End If
        Next X
        'GFXEngine.moDevice.RenderState.AlphaBlendEnable = bAlpha
    End Sub

    Public Sub HandleAddSystemItem(ByVal yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Now, get our add our new item
        Dim sKey As String = "Sys" & lID.ToString
        Dim bNew As Boolean = True
        Dim oNew As SystemItemData = Nothing

        If mcolData Is Nothing Then mcolData = New Collection
        If mcolData.Contains(sKey) = True Then
            bNew = False
            oNew = CType(mcolData(sKey), SystemItemData)
        End If
        If oNew Is Nothing Then
            bNew = True
            oNew = New SystemItemData()
        End If

        With oNew
            .lID = lID
            .sName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            .lPlanetCount = yData(lPos) : lPos += 1
            .lActiveWars = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .lMinCacheCount = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .blTotalMineralQty = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .lAvgMinCacheConc = yData(lPos) : lPos += 1

            .lMagistrateCnt = yData(lPos) : lPos += 1
            .lGovernorCnt = yData(lPos) : lPos += 1
            .lOverseerCnt = yData(lPos) : lPos += 1
            .lDukeCnt = yData(lPos) : lPos += 1
            .lBaronCnt = yData(lPos) : lPos += 1
            .lKingCnt = yData(lPos) : lPos += 1
            .lEmperorCnt = yData(lPos) : lPos += 1

            .lWHLinks = yData(lPos) : lPos += 1

            Dim lPlanetOwnerCnt As Int32 = yData(lPos) : lPos += 1
            Dim lPlayerID(lPlanetOwnerCnt - 1) As Int32
            Dim yOwnedCnt(lPlanetOwnerCnt - 1) As Byte

            For X As Int32 = 0 To lPlanetOwnerCnt - 1
                lPlayerID(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                yOwnedCnt(X) = yData(lPos) : lPos += 1
            Next X

            .lPlanetOwnerUB = -1
            .yPlanetOwnerCnt = yOwnedCnt
            .lPlanetOwnerID = lPlayerID
            .lPlanetOwnerUB = lPlanetOwnerCnt - 1
        End With

        If bNew = True Then
            mcolData.Add(oNew, sKey)
        End If
        SortCollection(mcolData, "sName", True)

    End Sub

    Public Sub DoTest()
        Dim oRand As New Random()
        If goGalaxy Is Nothing = False Then
            For X As Int32 = 0 To goGalaxy.mlSystemUB
                If goGalaxy.moSystems(X) Is Nothing = False Then
                    Dim oSys As SolarSystem = goGalaxy.moSystems(X)
                    Dim oNew As New SystemItemData()
                    With oNew
                        .lID = oSys.ObjectID
                        .sName = oSys.SystemName
                        .lPlanetCount = oSys.PlanetUB + 1
                        .lActiveWars = oRand.Next(0, 1000)
                        .lMinCacheCount = oRand.Next(0, 1000000)
                        .blTotalMineralQty = oRand.Next(0, 1000000000)
                        .lAvgMinCacheConc = oRand.Next(0, 255)

                        .lMagistrateCnt = oRand.Next(0, 10)
                        .lGovernorCnt = oRand.Next(0, 10)
                        .lOverseerCnt = oRand.Next(0, 10)
                        .lDukeCnt = oRand.Next(0, 10)
                        .lBaronCnt = oRand.Next(0, 10)
                        .lKingCnt = oRand.Next(0, 10)
                        .lEmperorCnt = oRand.Next(0, 10)

                        .lWHLinks = oRand.Next(0, 3)

                        Dim lPlanetOwnerCnt As Int32 = oSys.lVoterUB + 1
                        Dim lPlayerID(lPlanetOwnerCnt - 1) As Int32
                        Dim yOwnedCnt(lPlanetOwnerCnt - 1) As Byte

                        For Y As Int32 = 0 To lPlanetOwnerCnt - 1
                            lPlayerID(Y) = oSys.lVoterIDs(Y)
                            yOwnedCnt(Y) = oSys.yVoterCnts(Y)
                        Next Y

                        .lPlanetOwnerUB = -1
                        .yPlanetOwnerCnt = yOwnedCnt
                        .lPlanetOwnerID = lPlayerID
                        .lPlanetOwnerUB = lPlanetOwnerCnt - 1
                    End With

                    mcolData.Add(oNew, "Sys" & oNew.lID)
                End If
            Next X
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If moAgentIcons Is Nothing = False Then
            moAgentIcons.Dispose()
        End If
        moAgentIcons = Nothing

        MyBase.Finalize()
    End Sub

    Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        Dim ofrm As frmLoginDlg = CType(goUILib.GetWindow("frmLoginDlg"), frmLoginDlg)
        If ofrm Is Nothing = False Then
            ofrm.ReEnableStuff()
        End If
    End Sub

    Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click

        If mlSelectedID < 1 OrElse mlSelectedID = Int32.MaxValue Then
            MyBase.moUILib.AddNotification("You must select a system to respawn in.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'Now, send the msg to the server
        Dim yMsg(5) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRespawnSelection).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(mlSelectedID).CopyTo(yMsg, lPos) : lPos += 4
        MyBase.moUILib.SendMsgToPrimary(yMsg)

        MyBase.moUILib.RemoveWindow(Me.ControlName)

        MyBase.moUILib.AddNotification("Respawn request submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub moSystemItem_OnSelect(ByVal lSysID As Int32)
        mlSelectedID = lSysID
        'For x As Int32 = 0 To goGalaxy.mlSystemUB
        '    If goGalaxy.moSystems(x).ObjectID = lSysID Then
        '        Debug.Print(lSysID.ToString & " " & goGalaxy.moSystems(x).ObjectID.ToString & " " & goGalaxy.moSystems(x).SystemName)
        '        Exit For
        '    End If
        'Next

        Me.IsDirty = True
    End Sub

#Region "  ctlSystemItem  "
    'Interface created from Interface Builder
    Private Class ctlSystemItem
        Inherits UIWindow

        Private lblName As UILabel
        Private lblWHCnt As UILabel
        Private lblWars As UILabel
        Private lblMinConc As UILabel
        Private lblMinCount As UILabel
        Private lblTotalMins As UILabel

        Private WithEvents lstOwnerships As UIListBox
        Private WithEvents lstPlayerCounts As UIListBox

        Private mbHasUnknowns As Boolean = True
        Private moData As SystemItemData = Nothing

        Private mbSelected As Boolean = False

        Public Event OnSelect(ByVal lSysID As Int32)

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'ctlSystemItem initial props
            With Me
                .ControlName = "ctlSystemItem"
                .Left = 0
                .Top = 0
                .Width = 475
                .Height = 96
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .Resizable = False
            End With

            'lblName initial props
            lblName = New UILabel(oUILib)
            With lblName
                .ControlName = "lblName"
                .Left = 5
                .Top = 3
                .Width = 134
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "System Name"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblName, UIControl))

            'lstOwnerships initial props
            lstOwnerships = New UIListBox(oUILib)
            With lstOwnerships
                .ControlName = "lstOwnerships"
                .Left = 310
                .Top = 5
                .Width = 160
                .Height = 86
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .ToolTipText = "Players who own planets in this system"
            End With
            Me.AddChild(CType(lstOwnerships, UIControl))

            'lblWHCnt initial props
            lblWHCnt = New UILabel(oUILib)
            With lblWHCnt
                .ControlName = "lblWHCnt"
                .Left = mrcWH.X + mrcWH.Width + 5
                .Top = 30
                .Width = 36
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "0"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
                .ToolTipText = "Number of unlocked and active wormholes"
            End With
            Me.AddChild(CType(lblWHCnt, UIControl))

            'lstPlayerCounts initial props
            lstPlayerCounts = New UIListBox(oUILib)
            With lstPlayerCounts
                .ControlName = "lstPlayerCounts"
                .Left = 145
                .Top = 5
                .Width = 160
                .Height = 86
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .ToolTipText = "Number of players in the system by senate title"
            End With
            Me.AddChild(CType(lstPlayerCounts, UIControl))

            'lblWars initial props
            lblWars = New UILabel(oUILib)
            With lblWars
                .ControlName = "lblWars"
                .Left = mrcWars.X + mrcWars.Width + 5
                .Top = 30
                .Width = 46
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "0"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
                .ToolTipText = "Number of active wars in system"
            End With
            Me.AddChild(CType(lblWars, UIControl))


            'lblMinConc initial props
            lblMinConc = New UILabel(oUILib)
            With lblMinConc
                .ControlName = "lblMinConc"
                .Left = lblWHCnt.Left
                .Top = 50
                .Width = 36
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "0"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
                .ToolTipText = "Average mineral concentration of mineable mineral caches in the system"
            End With
            Me.AddChild(CType(lblMinConc, UIControl))

            'lblMinCount initial props
            lblMinCount = New UILabel(oUILib)
            With lblMinCount
                .ControlName = "lblMinCount"
                .Left = lblWars.Left
                .Top = 50
                .Width = 46
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "0"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
                .ToolTipText = "Number of mineable mineral caches in the system"
            End With
            Me.AddChild(CType(lblMinCount, UIControl))

            'lblTotalMins initial props
            lblTotalMins = New UILabel(oUILib)
            With lblTotalMins
                .ControlName = "lblTotalMins"
                .Left = lblWHCnt.Left '+ 10
                .Top = 70
                .Width = 85
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "0"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
                .ToolTipText = "Sum of the quantity of all mineable mineral caches in the system"
            End With
            Me.AddChild(CType(lblTotalMins, UIControl))
        End Sub

        Public Function GetCurrentID() As Int32
            If moData Is Nothing = False Then Return moData.lID
        End Function

        Public Sub SetFromSystemData(ByRef oData As SystemItemData)
            moData = oData
            'Handle if modata is nothing
            With oData
                lblName.Caption = oData.sName
                lblWHCnt.Caption = oData.lWHLinks.ToString("#,##0")
                lblWars.Caption = oData.lActiveWars.ToString("#,##0")
                lblMinConc.Caption = oData.lAvgMinCacheConc.ToString("#,##0")
                lblMinCount.Caption = oData.lMinCacheCount.ToString("#,##0")
                lblTotalMins.Caption = oData.blTotalMineralQty.ToString("#,##0")

                lstOwnerships.Clear()
                lstPlayerCounts.Clear()

                lstOwnerships.AddItem("Total Planets: " & .lPlanetCount.ToString(), False)
                lstOwnerships.ItemData(lstOwnerships.NewIndex) = -1

                If .lEmperorCnt > 0 Then lstPlayerCounts.AddItem("Emperors: " & .lEmperorCnt.ToString())
                If .lKingCnt > 0 Then lstPlayerCounts.AddItem("Kings: " & .lKingCnt.ToString())
                If .lBaronCnt > 0 Then lstPlayerCounts.AddItem("Barons: " & .lBaronCnt.ToString())
                If .lDukeCnt > 0 Then lstPlayerCounts.AddItem("Dukes: " & .lDukeCnt.ToString())
                If .lOverseerCnt > 0 Then lstPlayerCounts.AddItem("Overseers: " & .lOverseerCnt.ToString())
                If .lGovernorCnt > 0 Then lstPlayerCounts.AddItem("Governors: " & .lGovernorCnt.ToString())
                If .lMagistrateCnt > 0 Then lstPlayerCounts.AddItem("Magistrates: " & .lMagistrateCnt.ToString())

                mbHasUnknowns = True
            End With
        End Sub

        Public Sub Process_OnNewFrame()
            If mbHasUnknowns = True Then
                If moData Is Nothing = False Then
                    If lstOwnerships.ListCount <> moData.lPlanetOwnerUB + 2 Then    '1 for cnt, 1 for planet cnt
                        lstOwnerships.Clear()
                        lstOwnerships.AddItem("Total Planets: " & moData.lPlanetCount.ToString(), False)
                        lstOwnerships.ItemData(lstOwnerships.NewIndex) = -1
                        For X As Int32 = 0 To moData.lPlanetOwnerUB
                            Dim sName As String = GetCacheObjectValue(moData.lPlanetOwnerID(X), ObjectType.ePlayer) & " (" & moData.yPlanetOwnerCnt(X).ToString & ")"
                            lstOwnerships.AddItem(sName, False)
                            lstOwnerships.ItemData(lstOwnerships.NewIndex) = moData.lPlanetOwnerID(X)
                            lstOwnerships.ItemData2(lstOwnerships.NewIndex) = moData.yPlanetOwnerCnt(X)
                        Next X
                    Else
                        mbHasUnknowns = False
                        For X As Int32 = 0 To lstOwnerships.ListCount
                            If lstOwnerships.ItemData(X) > 0 Then
                                Dim sName As String = GetCacheObjectValue(lstOwnerships.ItemData(X), ObjectType.ePlayer)
                                If sName.ToUpper = "UNKNOWN" Then mbHasUnknowns = True
                                sName &= " (" & lstOwnerships.ItemData2(X).ToString & ")"
                                If lstOwnerships.List(X) <> sName Then
                                    lstOwnerships.List(X) = sName
                                    lstOwnerships.IsDirty = True
                                End If
                            End If
                        Next X
                    End If
                End If
            End If
        End Sub

        Private Sub ctlSystemItem_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
            If moData Is Nothing = False Then RaiseEvent OnSelect(moData.lID)
            Me.IsDirty = True
        End Sub

        Private Sub ctlSystemItem_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
            Dim pt As Point = GetAbsolutePosition()
            Dim lX As Int32 = lMouseX - pt.X
            Dim lY As Int32 = lMouseY - pt.Y

            Dim sTT As String = ""

            If mrcWH.Contains(lX, lY) Then
                sTT = lblWHCnt.ToolTipText
            ElseIf mrcWars.Contains(lX, lY) Then
                sTT = lblWars.ToolTipText
            ElseIf mrcMinConc.Contains(lX, lY) Then
                sTT = lblMinConc.ToolTipText
            ElseIf mrcMinCount.Contains(lX, lY) Then
                sTT = lblMinCount.ToolTipText
            ElseIf mrcTotalMins.Contains(lX, lY) Then
                sTT = lblTotalMins.ToolTipText
            End If

            Me.ToolTipText = sTT

        End Sub

        Public Sub UpdateSelectState(ByVal bVal As Boolean)
            'If bVal <> mbSelected Then
            mbSelected = bVal
            Dim iFill As System.Drawing.Color
            If mbSelected = True Then

                iFill = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
            Else
                iFill = muSettings.InterfaceFillColor
                'End If
            End If
            If Me.FillColor <> iFill Then Me.FillColor = iFill
        End Sub
    End Class
#End Region

End Class