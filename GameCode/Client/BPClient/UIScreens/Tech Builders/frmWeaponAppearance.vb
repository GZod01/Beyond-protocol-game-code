Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmWeaponAppearance
    Inherits UIWindow

    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Private WithEvents lblTitle As UILabel
    Private WithEvents lscAppearance As UILabelScroller
    Private WithEvents chkAudio As UICheckBox
    
    Private mbLoading As Boolean = True


    Private mbMissiles As Boolean = False
    Private mlMissileIdx As Int32 = -1

    Private moMissileMgr As MissileMgr = Nothing
    Private moWpnFX As WpnFXManager = Nothing
    Private moTex As Texture = Nothing

    Private mlROF As Int32 = 50
    Private mlLastShot As Int32

    Private mlSolidBeamType As Int32 = 0

    Public Sub SetROF(ByVal lVal As Int32)
        mlROF = lVal
    End Sub

    Public Sub New(ByRef oUILib As UILib, ByVal lLeft As Int32, ByVal lTop As Int32)
        MyBase.New(oUILib)

        'frmWeaponAppearance initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eWeaponAppearance
            .ControlName = "frmWeaponAppearance"
            .Left = lLeft
            .Top = lTop
            .Width = 180
            .Height = 70 '45
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor 'System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .FullScreen = True
            .Moveable = True
            '.FillWindow = False
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 0
            .Width = 153
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Weapon Appearance"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lscAppearance initial props
        lscAppearance = New UILabelScroller(oUILib)
        With lscAppearance
            .ControlName = "lscAppearance"
            .Left = 5
            .Top = 20
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
        End With
        Me.AddChild(CType(lscAppearance, UIControl))

        'chkAudio initial props
        chkAudio = New UICheckBox(oUILib)
        With chkAudio
            .ControlName = "chkAudio"
            .Left = 10
            .Top = 45
            .Width = 123
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Audio Enabled"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkAudio, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        mbLoading = False
    End Sub

    Public Sub SetFromWeaponClassType(ByVal yWeaponClassType As WeaponClassType)
        mbLoading = True

        mbMissiles = False
        moMissileMgr = Nothing
        moWpnFX = Nothing

        With lscAppearance
            .Clear()

            'TODO: Flesh these out a bit more

            Select Case yWeaponClassType
                Case WeaponClassType.eEnergyBeam
                    mlSolidBeamType = 0
                    .AddItem(WeaponType.eFlickerGreenBeam, "Flicker Green")
                    .AddItem(WeaponType.eFlickerPurpleBeam, "Flicker Purple")
                    .AddItem(WeaponType.eFlickerRedBeam, "Flicker Red")
                    .AddItem(WeaponType.eFlickerTealBeam, "Flicker Teal")
                    .AddItem(WeaponType.eFlickerYellowBeam, "Flicker Yellow")
                Case WeaponClassType.eEnergyPulse
                    .AddItem(WeaponType.eShortGreenPulse, "Green Pulse")
                    .AddItem(WeaponType.eShortPurplePulse, "Purple Pulse")
                    .AddItem(WeaponType.eShortRedPulse, "Red Pulse")
                    .AddItem(WeaponType.eShortTealPulse, "Teal Pulse")
                    .AddItem(WeaponType.eShortYellowPulse, "Yellow Pulse")
                Case WeaponClassType.eMissile
                    mbMissiles = True
                    .AddItem(WeaponType.eMissile_Color_1, "Light Blue")
                    .AddItem(WeaponType.eMissile_Color_2, "Green")
                    .AddItem(WeaponType.eMissile_Color_3, "Orange")
                    .AddItem(WeaponType.eMissile_Color_4, "Blue")
                    .AddItem(WeaponType.eMissile_Color_5, "Purple")
                    .AddItem(WeaponType.eMissile_Color_6, "Dark Blue")
                    .AddItem(WeaponType.eMissile_Color_7, "Red")
                    .AddItem(WeaponType.eMissile_Color_8, "Yellow")
                    .AddItem(WeaponType.eMissile_Color_9, "Dark Green")
                Case WeaponClassType.eProjectile
                    .AddItem(WeaponType.eMetallicProjectile_Bronze, "Orange")
                    .AddItem(WeaponType.eMetallicProjectile_Copper, "Red")
                    .AddItem(WeaponType.eMetallicProjectile_Gold, "Yellow")
                    .AddItem(WeaponType.eMetallicProjectile_Lead, "Gray")
                    .AddItem(WeaponType.eMetallicProjectile_Silver, "White")
                Case WeaponClassType.eBomb
                    .AddItem(WeaponType.eBomb_Gray, "Gray")
                    .AddItem(WeaponType.eBomb_Green, "Green")
                    .AddItem(WeaponType.eBomb_Purple, "Purple")
                    .AddItem(WeaponType.eBomb_Red, "Red")
                    .AddItem(WeaponType.eBomb_Teal, "Teal")
                    .AddItem(WeaponType.eBomb_Yellow, "Yellow")
            End Select
        End With
        mbLoading = False

        lscAppearance.ResetCurrentIndex()
    End Sub

    Public Sub SetSolidBeamType(ByVal lType As Int32)
        mlSolidBeamType = lType
        mbLoading = True
        With lscAppearance
            .Clear()
            If mlSolidBeamType = 0 Then
                .AddItem(WeaponType.eFlickerGreenBeam, "Flicker Green")
                .AddItem(WeaponType.eFlickerPurpleBeam, "Flicker Purple")
                .AddItem(WeaponType.eFlickerRedBeam, "Flicker Red")
                .AddItem(WeaponType.eFlickerTealBeam, "Flicker Teal")
                .AddItem(WeaponType.eFlickerYellowBeam, "Flicker Yellow")
            Else
                .AddItem(WeaponType.eSolidGreenBeam, "Solid Green")
                .AddItem(WeaponType.eSolidPurpleBeam, "Solid Purple")
                .AddItem(WeaponType.eSolidRedBeam, "Solid Red")
                .AddItem(WeaponType.eSolidTealBeam, "Solid Teal")
                .AddItem(WeaponType.eSolidYellowBeam, "Solid Yellow")
            End If
        End With
        mbLoading = False
        lscAppearance.ResetCurrentIndex()
    End Sub

    Private Sub lscAppearance_ItemChanged(ByVal lID As Integer, ByVal sDisplay As String) Handles lscAppearance.ItemChanged
        If mbLoading = True Then Return

        If mbMissiles = True Then
            If moMissileMgr Is Nothing Then
                moMissileMgr = New MissileMgr()
            End If
            moMissileMgr.KillAll()
            If lID > -1 Then mlMissileIdx = moMissileMgr.AddWpnBuilderMissile(New Vector3(0, 0, 0), New Vector3(0, 0, 10000000), 30, 50, CInt(lID) - WeaponType.eMissile_Color_1, 2)
            'moMissileMgr.MissileImpact(mlMissileIdx, 1)
        Else
            If moWpnFX Is Nothing Then
                moWpnFX = New WpnFXManager()
                AddHandler moWpnFX.WpnEnd_WpnBldrOnly, AddressOf WpnFX_WeaponEnd
                moWpnFX.GenerateSoundFX = chkAudio.Value
            End If
            If lID > -1 Then moWpnFX.AddNewEffect(0, 600, 0, 0, 0, 0, CByte(lID), True, 0)
        End If
        Dim oFrm As frmWeaponBuilder = CType(MyBase.moUILib.GetWindow("frmWeaponBuilder"), frmWeaponBuilder)
        If oFrm Is Nothing = False Then
            oFrm.IsDirty = True
        End If
    End Sub

    Private Sub HandleInstantiation()
        If mbMissiles = True Then
            If moMissileMgr Is Nothing Then
                moMissileMgr = New MissileMgr()
            End If
            'If (timeGetTime - mlLastShot) \ 30 > mlROF Then
            If lscAppearance.SelectedItemID > -1 AndAlso mlMissileIdx = -1 Then
                mlMissileIdx = moMissileMgr.AddWpnBuilderMissile(New Vector3(0, 0, 0), New Vector3(0, 0, 10000000), 30, 50, CInt(lscAppearance.SelectedItemID) - WeaponType.eMissile_Color_1, 2)
                'moMissileMgr.MissileImpact(mlMissileIdx, 1)
                'mlLastShot = timeGetTime
            End If
            'End If
        Else
            If moWpnFX Is Nothing Then
                moWpnFX = New WpnFXManager()
                AddHandler moWpnFX.WpnEnd_WpnBldrOnly, AddressOf WpnFX_WeaponEnd
                moWpnFX.GenerateSoundFX = chkAudio.Value
            End If

            If (timeGetTime - mlLastShot) \ 30I > mlROF Then
                If lscAppearance.SelectedItemID > -1 Then moWpnFX.AddNewEffect(0, 600, 0, 0, 0, 0, CByte(lscAppearance.SelectedItemID), True, 0)
                mlLastShot = timeGetTime
            End If
        End If
    End Sub

    Private Sub frmWeaponAppearance_OnNewFrameEnd() Handles Me.OnNewFrameEnd
        If mbLoading = True Then Return

        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then
            If moTex Is Nothing = False Then moTex.Dispose()
            moTex = Nothing
            Return
        End If

        HandleInstantiation()

        Dim lCX As Int32 = goCamera.mlCameraX
        Dim lCY As Int32 = goCamera.mlCameraY
        Dim lCZ As Int32 = goCamera.mlCameraZ
        Dim lCAtX As Int32 = goCamera.mlCameraAtX
        Dim lCAtY As Int32 = goCamera.mlCameraAtY
        Dim lCAtZ As Int32 = goCamera.mlCameraAtZ


        With goCamera
            .mlCameraX = 0
            .mlCameraY = 500
            .mlCameraZ = -200
            .mlCameraAtX = 0
            .mlCameraAtY = 0
            .mlCameraAtZ = 100
            If mbMissiles = False Then
                If goSound Is Nothing = False Then goSound.UpdateListenerLoc(New Vector3(.mlCameraX, .mlCameraY, .mlCameraZ), Vector3.Normalize(Vector3.Subtract(New Vector3(.mlCameraAtX, .mlCameraAtY, .mlCameraAtZ), New Vector3(.mlCameraX, .mlCameraY, .mlCameraZ))), New Vector3(0, 1, 0))
            Else
                If moMissileMgr Is Nothing = False AndAlso mlMissileIdx <> -1 Then
                    Dim vecTemp As Vector3 = moMissileMgr.GetMissileLoc(mlMissileIdx)
                    .mlCameraY = 100
                    .mlCameraAtZ = CInt(vecTemp.Z - 50)
                    .mlCameraZ = CInt(vecTemp.Z - 200)
                End If
            End If
        End With

        'Set no events
        Device.IsUsingEventHandlers = False

        DoRendering()

        'turn events back on
        Device.IsUsingEventHandlers = True

        With goCamera
            .mlCameraX = lCX
            .mlCameraY = lCY
            .mlCameraZ = lCZ
            .mlCameraAtX = lCAtX
            .mlCameraAtY = lCAtY
            .mlCameraAtZ = lCAtZ
        End With

        MyBase.RenderBorder(0, 128)
    End Sub

    Private Sub DoRendering()

        Dim matView As Matrix
        Dim matProj As Matrix
        Dim oLoc As System.Drawing.Point
        Dim oOriginal As Surface
        Dim oScene As Surface = Nothing

        'Create our texture...
        If moTex Is Nothing Then moTex = New Texture(MyBase.moUILib.oDevice, 256, 256, 1, Usage.RenderTarget, Format.R5G6B5, Pool.Default)
        With MyBase.moUILib.oDevice

            'Store our matrices beforehand...
            .RenderState.ZBufferEnable = False
            matView = MyBase.moUILib.oDevice.Transform.View
            matProj = MyBase.moUILib.oDevice.Transform.Projection

            'Ok, store our original surface
            oOriginal = .GetRenderTarget(0)

            'Get our surface to render to
            oScene = moTex.GetSurfaceLevel(0)

            'Now, set our render target to the texture's surface
            .SetRenderTarget(0, oScene)

            'Clear out our surface
            .Clear(ClearFlags.Target, System.Drawing.Color.FromArgb(255, 0, 0, 0), 1.0F, 0)

            'Set up our matrices
            If mbMissiles = True Then
                .Transform.View = Matrix.LookAtLH(New Vector3(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ), New Vector3(goCamera.mlCameraAtX, goCamera.mlCameraAtY, goCamera.mlCameraAtZ), New Vector3(0.0F, 1.0F, 0.0F))
                .Transform.Projection = Matrix.PerspectiveFovLH(0.7853982F, 2.0F, muSettings.TerrainNearClippingPlane, muSettings.FarClippingPlane)
            Else
                .Transform.View = Matrix.LookAtLH(New Vector3(0, 500, -200), New Vector3(0, 0, 100), New Vector3(0.0F, 1.0F, 0.0F))
                .Transform.Projection = Matrix.PerspectiveFovLH(0.7853982F, 2.0F, muSettings.TerrainNearClippingPlane, muSettings.FarClippingPlane)
            End If

            '.RenderState.ZBufferWriteEnable = True

            If mbMissiles = True Then
                moMissileMgr.Render(False)
            Else : moWpnFX.RenderFX(False)
            End If

            'Now, restore our original surface to the device
            .SetRenderTarget(0, oOriginal)

            'now, re-enable our Z Buffer
            '.RenderState.ZBufferWriteEnable = True

            'restore our matrices
            .Transform.View = matView
            .Transform.Projection = matProj
            .Transform.World = Matrix.Identity

            'Release all our objects
            If oScene Is Nothing = False Then oScene.Dispose()
            If oOriginal Is Nothing = False Then oOriginal.Dispose()
            oScene = Nothing
            oOriginal = Nothing

            'Now... create a sprite
            Using oTmpSprite As Sprite = New Sprite(MyBase.moUILib.oDevice)
                oTmpSprite.Begin(SpriteFlags.AlphaBlend)

                oLoc = Me.GetAbsolutePosition
                oLoc.Y += Me.Height
                oLoc.X += 26
                Dim fX As Single = oLoc.X '+ 5
                Dim fY As Single = oLoc.Y '+ 5

                Dim rcDest As Rectangle = Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + 128, oLoc.Y + 128)
                Dim rcSrc As Rectangle = Rectangle.FromLTRB(0, 0, 256, 256)

                fX = (oLoc.X) * CSng(rcSrc.Width / rcDest.Width)
                fY = (oLoc.Y) * CSng(rcSrc.Height / rcDest.Height)

                oTmpSprite.Draw2D(moTex, rcSrc, rcDest, Point.Empty, 0, New Point(CInt(fX), CInt(fY)), System.Drawing.Color.White)

                oTmpSprite.End()
                oTmpSprite.Dispose()
            End Using
            .RenderState.ZBufferEnable = True
        End With
    End Sub

    Private Sub WpnFX_WeaponEnd(ByVal yWeaponType As Byte)
        Dim lType As Int32 = lscAppearance.SelectedItemID
        'If lType = yWeaponType Then moWpnFX.AddNewEffect(0, 600, 0, 0, 0, 0, CByte(lType), True)
    End Sub

    Public Function GetWeaponTypeID() As Byte
        Dim lID As Int32 = lscAppearance.SelectedItemID
        If lID > -1 Then Return CByte(lID)
    End Function

    Public Sub SetWeaponTypeID(ByVal yVal As Byte)
        lscAppearance.SelectByID(yVal)
    End Sub

    Public Sub DisposeMe()
        Debug.Print("DisposeMe Called")
        Me.Visible = False
        moWpnFX = Nothing
        If moMissileMgr Is Nothing = False Then moMissileMgr.KillAll()
        moMissileMgr = Nothing
        If moTex Is Nothing = False Then moTex.Dispose()
        moTex = Nothing
    End Sub

    Protected Overrides Sub Finalize()
		Debug.Print(Me.ControlName & " finalized")
		MyBase.Finalize()
    End Sub

    Private Sub chkAudio_Click() Handles chkAudio.Click
        If moWpnFX Is Nothing = False Then moWpnFX.GenerateSoundFX = chkAudio.Value
    End Sub
End Class