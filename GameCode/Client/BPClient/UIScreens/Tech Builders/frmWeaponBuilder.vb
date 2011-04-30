Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmWeaponBuilder
    Inherits frmTechBuilder
    'Inherits UIWindow

    Private WithEvents fraCurrent As fraWeaponBase

    'Form's Global controls...
    Private lblWeaponName As UILabel
    Private WithEvents txtWeaponName As UITextBox
    Private lblWeaponType As UILabel
    Private WithEvents cboWeaponType As UIComboBox
    Private WithEvents btnSubmit As UIButton
    Private WithEvents btnCancel As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnRename As UIButton
    Private WithEvents btnAutoFill As UIButton

    Private WithEvents cboHullType As UIComboBox
    Private lblHullType As UILabel


    Private lblDesignFlaw As UILabel

    Private moWpnApp As frmWeaponAppearance = Nothing

    Private moTech As WeaponTech = Nothing
    Private mfrmResCost As frmProdCost = Nothing
    Private mfrmProdCost As frmProdCost = Nothing
    Private mfrmWpnDef As frmProdCost = Nothing

    Private mbDirtyMeAgain As Boolean = False
    Private mbWpnFraChanged As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmWeaponBuilder initial props
        With Me
            .ControlName = "frmWeaponBuilder"
            .Left = 10
            .Top = 10
            .Width = 490 '420
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = False
            .mbAcceptReprocessEvents = True
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'lblWeaponName initial props
        lblWeaponName = New UILabel(oUILib)
        With lblWeaponName
            .ControlName = "lblWeaponName"
            .Left = 15
            .Top = 10
            .Width = 110
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Weapon Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblWeaponName, UIControl))

        'txtWeaponName initial props
        txtWeaponName = New UITextBox(oUILib)
        With txtWeaponName
            .ControlName = "txtWeaponName"
            .Left = 135
            .Top = 10
            .Width = 201
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtWeaponName, UIControl))

        'lblWeaponType initial props
        lblWeaponType = New UILabel(oUILib)
        With lblWeaponType
            .ControlName = "lblWeaponType"
            .Left = 15
            .Top = 35
            .Width = 110
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Weapon Type:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblWeaponType, UIControl))

        'lblDesignFlaw initial props
        lblDesignFlaw = New UILabel(oUILib)
        With lblDesignFlaw
            .ControlName = "lblDesignFlaw"
            .Left = 15
            .Top = 435 '60
            .Width = Me.Width - (.Left * 2)
            .Height = 36 '18
            .Enabled = True
            .Visible = False
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(lblDesignFlaw, UIControl))

        'btnSubmit initial props
        btnSubmit = New UIButton(oUILib)
        With btnSubmit
            .ControlName = "btnSubmit"
            .Left = (Me.Width \ 2) - 50 '114
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
            .Left = Me.Width - 110 '247
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

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Left = 10 '15
            .Top = btnCancel.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = False
            .Caption = "Delete Design"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnDelete, UIControl))

        'btnAutoFill initial props
        btnAutoFill = New UIButton(oUILib)
        With btnAutoFill
            .ControlName = "btnAutoFill"
            .Left = 10
            .Top = btnCancel.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Auto Balance"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnAutoFill, UIControl))

        'btnRename initial props
        btnRename = New UIButton(oUILib)
        With btnRename
            .ControlName = "btnRename"
            .Left = txtWeaponName.Left + txtWeaponName.Width + 5
            .Top = txtWeaponName.Top - 1
            .Width = 100
            .Height = 24 'txtWeaponName.Height
            .Enabled = True
            .Visible = False
            .Caption = "Rename"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnRename, UIControl))

        'lblHullType initial props
        lblHullType = New UILabel(oUILib)
        With lblHullType
            .ControlName = "lblHullType"
            .Left = 15
            .Top = 60
            .Width = 110
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hull Type:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHullType, UIControl))

        '=============================== ALL COMBO BOXES ============================

        'cboHullType initial props
        cboHullType = New UIComboBox(oUILib)
        With cboHullType
            .ControlName = "cboHullType"
            .Left = 135
            .Top = 60
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
        End With
        Me.AddChild(CType(cboHullType, UIControl))


        'cboWeaponType initial props
        cboWeaponType = New UIComboBox(oUILib)
        With cboWeaponType
            .ControlName = "cboWeaponType"
            .Left = 135
            .Top = 35
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
        End With
        Me.AddChild(CType(cboWeaponType, UIControl))

        FillValues()

        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
        ofrm.ShowMineralDetail(Me.Left + Me.Width + 5, Me.Top, Me.Height, -1)

        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
        goFullScreenBackground = goResMgr.GetTexture("weapons.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "TechBacks.pak")

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mfrmBuilderCost.Visible = False
        glCurrentEnvirView = CurrentView.eWeaponBuilder
    End Sub

    Private Sub FillValues()
        cboWeaponType.Clear()
        'Fill in our weapontype... (in alphabetical)
        cboWeaponType.AddItem("Bomb")
        cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eBomb
        cboWeaponType.AddItem("Energy/Beam")
        cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eEnergyBeam
        cboWeaponType.AddItem("Energy/Pulse")
        cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eEnergyPulse
        'cboWeaponType.AddItem("Mine")
        'cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eMine
        cboWeaponType.AddItem("Missile")
        cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eMissile
        cboWeaponType.AddItem("Projectile")
        cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eProjectile


        Dim lMaxPowerThrust As Int32 = 10000
        If goCurrentPlayer Is Nothing = False Then lMaxPowerThrust = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)
        cboHullType.Clear()
        If lMaxPowerThrust > 110000 Then
            cboHullType.AddItem("Battlecruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.BattleCruiser
        End If
        If lMaxPowerThrust > 400000 Then
            cboHullType.AddItem("Battleship") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Battleship
        End If
        cboHullType.AddItem("Corvette") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Corvette
        If lMaxPowerThrust > 57000 Then
            cboHullType.AddItem("Cruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Cruiser
        End If
        If lMaxPowerThrust > 32000 Then
            cboHullType.AddItem("Destroyer") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Destroyer
        End If
        cboHullType.AddItem("Escort") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Escort
        cboHullType.AddItem("Facility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Facility
        cboHullType.AddItem("Fighter (Light)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.LightFighter
        cboHullType.AddItem("Fighter (Medium)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.MediumFighter
        cboHullType.AddItem("Fighter (Heavy)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.HeavyFighter
        cboHullType.AddItem("Frigate") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Frigate
        Dim bHasNavalUnit As Boolean = False
        If goCurrentPlayer Is Nothing = False Then bHasNavalUnit = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 0
        If bHasNavalUnit = True Then
            If lMaxPowerThrust > 181000 Then
                cboHullType.AddItem("Naval Battleship") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalBattleship
                cboHullType.AddItem("Naval Carrier") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalCarrier
            End If
            If lMaxPowerThrust > 83000 Then
                cboHullType.AddItem("Naval Cruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalCruiser
            End If
            If lMaxPowerThrust > 31000 Then
                cboHullType.AddItem("Naval Destroyer") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalDestroyer
            End If
            If lMaxPowerThrust > 15000 Then
                cboHullType.AddItem("Naval Frigate") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalFrigate
                If lMaxPowerThrust > 42000 Then
                    cboHullType.AddItem("Naval Submarine") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalSub
                End If
            End If
            cboHullType.AddItem("Naval Utility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Utility
        End If
        cboHullType.AddItem("Small Vehicle") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.SmallVehicle
        cboHullType.AddItem("Space Station") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.SpaceStation
        cboHullType.AddItem("Tank") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Tank
        cboHullType.AddItem("Utility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Utility

    End Sub

    Private Sub cboWeaponType_ItemSelected(ByVal lItemIndex As Integer) Handles cboWeaponType.ItemSelected

        If lItemIndex > -1 Then
            If cboWeaponType.ItemData(lItemIndex) = WeaponClassType.eBomb Then
                If cboHullType.ListIndex > -1 Then
                    If cboHullType.ItemData(cboHullType.ListIndex) <> Base_Tech.eyHullType.Frigate Then
                        cboHullType.FindComboItemData(Base_Tech.eyHullType.Frigate)
                        cboHullType_ItemSelected(cboHullType.ListIndex)
                    End If
                Else
                    cboHullType.FindComboItemData(Base_Tech.eyHullType.Frigate)
                    cboHullType_ItemSelected(cboHullType.ListIndex)
                End If
                cboHullType.Enabled = False
            Else
                cboHullType.Enabled = True
            End If
        End If

        If fraCurrent Is Nothing = False Then fraCurrent.CloseFrame()

        If lItemIndex = -1 OrElse cboHullType.ListIndex < 0 Then
            If fraCurrent Is Nothing = False Then fraCurrent.Visible = False
            If moWpnApp Is Nothing = False Then moWpnApp.Visible = False
        Else
            If moWpnApp Is Nothing = False Then
                moWpnApp.Visible = True
                moWpnApp.SetFromWeaponClassType(CType(cboWeaponType.ItemData(lItemIndex), WeaponClassType))
            End If

            If fraCurrent Is Nothing = False Then
                For X As Int32 = 0 To Me.ChildrenUB
                    If Me.moChildren(X).ControlName = fraCurrent.ControlName Then
                        Me.RemoveChild(X)
                        Exit For
                    End If
                Next X
                fraCurrent = Nothing
            End If

            'Remove cboWeaponType...
            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X).ControlName = cboWeaponType.ControlName Then
                    Me.RemoveChild(X)
                    Exit For
                End If
            Next X
            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X).ControlName = cboHullType.ControlName Then
                    Me.RemoveChild(X)
                    Exit For
                End If
            Next X

            Select Case cboWeaponType.ItemData(lItemIndex)
                Case WeaponClassType.eBomb
                    fraCurrent = New fraBomb(MyBase.moUILib)
                Case WeaponClassType.eEnergyBeam
                    fraCurrent = New fraSolidBeam(MyBase.moUILib)
                Case WeaponClassType.eMine
                    MyBase.moUILib.AddNotification("Mine builder is not yet implemented.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Case WeaponClassType.eMissile
                    fraCurrent = New fraMissiles(MyBase.moUILib)
                Case WeaponClassType.eProjectile
                    fraCurrent = New fraProjectile(MyBase.moUILib)
                Case WeaponClassType.eEnergyPulse
                    fraCurrent = New fraPulse(MyBase.moUILib)
                    'MyBase.moUILib.AddNotification("Pulse builder is currently disabled.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End Select

            If fraCurrent Is Nothing = False Then
                With fraCurrent
                    .Left = 10
                    .Top = 95 '75
                End With
                Me.AddChild(CType(fraCurrent, UIControl))
            End If

            MyBase.moUILib.RemoveWindow(Me.ControlName)

            If moWpnApp Is Nothing Then moWpnApp = New frmWeaponAppearance(MyBase.moUILib, Me.Left + Me.Width + 5, Me.Top + Me.Height + 5)
            If moWpnApp Is Nothing = False Then
                moWpnApp.Visible = True
                moWpnApp.SetFromWeaponClassType(CType(cboWeaponType.ItemData(lItemIndex), WeaponClassType))
            End If

            Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
            If ofrm Is Nothing = False Then
                ofrm.Top = Me.Top
                ofrm.Height = Me.Height
                ofrm.Left = Me.Left + Me.Width + 5
                'If moWpnApp.Visible = True Then
                '    ofrm.Left = moWpnApp.Left + moWpnApp.Width + 5
                'Else : ofrm.Left = moWpnApp.Left + Me.Width + 5
                'End If
            End If

            MyBase.moUILib.AddWindow(Me)
            MyBase.ReshowTechBuilderCost()

            'Now, readd cboWeaponType
            cboWeaponType.HasFocus = True
            cboWeaponType.HasFocus = False
            Me.AddChild(CType(cboHullType, UIControl))
            Me.AddChild(CType(cboWeaponType, UIControl))

            If cboWeaponType.ItemData(lItemIndex) = WeaponClassType.eBomb Then
                cboHullType.FindComboItemData(Base_Tech.eyHullType.Frigate)
                cboHullType_ItemSelected(cboHullType.ListIndex)
                cboHullType.Enabled = False
                Me.IsDirty = True
            Else
                cboHullType.Enabled = True
            End If

            If cboHullType.ListIndex > -1 Then
                cboHullType_ItemSelected(cboHullType.ListIndex)
            End If
        End If
    End Sub

    Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click
        'we drop focus for our controls just incase some invalid value causes recalculating
        If goUILib.FocusedControl Is Nothing = False Then
            goUILib.FocusedControl.HasFocus = False
            goUILib.FocusedControl = Nothing
        End If

        If fraCurrent Is Nothing = False Then
            If txtWeaponName.Caption.Trim = "" Then
                MyBase.moUILib.AddNotification("You must enter a name for this weapon!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            ElseIf txtWeaponName.Caption.Length > 20 Then
                MyBase.moUILib.AddNotification("Name must be less than 20 characters long!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            If cboWeaponType.ListIndex > -1 Then
                Dim yTypeID As Byte = 0
                If moWpnApp Is Nothing = False Then yTypeID = moWpnApp.GetWeaponTypeID()

                If moTech Is Nothing = False AndAlso btnSubmit.Caption.ToLower = "research" Then
                    'Ok, simply submit this for research now
                    Dim yData(13) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)

                    Dim lEntityIndex As Int32 = -1
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected Then
                            lEntityIndex = X
                            Exit For
                        End If
                    Next X
                    If lEntityIndex = -1 Then
                        MyBase.moUILib.AddNotification("You must select a research facility to research this.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If

                    goCurrentEnvir.oEntity(lEntityIndex).GetGUIDAsString.CopyTo(yData, 2)
                    moTech.GetGUIDAsString.CopyTo(yData, 8)
                    MyBase.moUILib.GetMsgSys.SendToPrimary(yData)

                    If NewTutorialManager.TutorialOn = True Then
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBuildOrderSent, moTech.ObjectID, moTech.ObjTypeID, 1, "")
                    End If

                    CloseMe()
                    Return
                End If

                If fraCurrent.ValidateData = True AndAlso fraCurrent.SubmitDesign(txtWeaponName.Caption, yTypeID) = True Then
                    If NewTutorialManager.TutorialOn = True Then
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSubmitDesignClick, ObjectType.eWeaponTech, -1, -1, "")
                    End If
                    CloseMe()
                End If
            End If
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
        CloseMe()
    End Sub

    Private Sub CloseMe()
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.RemoveWindow("frmMinDetail")
        If mfrmProdCost Is Nothing = False Then MyBase.moUILib.RemoveWindow(mfrmProdCost.ControlName)
        If mfrmResCost Is Nothing = False Then MyBase.moUILib.RemoveWindow(mfrmResCost.ControlName)
        If mfrmWpnDef Is Nothing = False Then MyBase.moUILib.RemoveWindow(mfrmWpnDef.ControlName)
        If moWpnApp Is Nothing = False Then
            moWpnApp.DisposeMe()
            MyBase.moUILib.RemoveWindow(moWpnApp.ControlName)
        End If
        ReturnToPreviousView()

        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
    End Sub

#Region "   Base Frame Class for the Weapon Builder... all child frames must inherit this...   "
    Private MustInherit Class fraWeaponBase
        Inherits UIWindow

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)
        End Sub

        Public MustOverride Function ValidateData() As Boolean
        Public MustOverride Function SubmitDesign(ByVal sName As String, ByVal yWeaponTypeID As Byte) As Boolean
        Public MustOverride Sub ViewResults(ByRef oTech As WeaponTech, ByVal lProdFactor As Int32)
        Public MustOverride Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
        Public MustOverride Sub CloseFrame()
        Public MustOverride Sub CheckForDARequest()

        Public lHullTypeID As Int32 = -1

        Protected mbImpossibleDesign As Boolean = False

        Protected mbIgnoreValueChange As Boolean = False

        Public Function mfrmBuilderCost() As frmTechBuilderCost
            Return CType(Me.ParentControl, frmWeaponBuilder).mfrmBuilderCost
        End Function

        Protected Sub SetDesignFlaw(ByVal sText As String)
            If Me.ParentControl Is Nothing Then Return
            With CType(Me.ParentControl, frmWeaponBuilder).lblDesignFlaw
                .Caption = sText
                .Visible = True
            End With
        End Sub
    End Class
#End Region

    Protected Overrides Sub Finalize()
        Debug.Write(Me.ControlName & " finalized" & vbCrLf)
        MyBase.Finalize()
    End Sub

    Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
        If moTech Is Nothing Then Return
        If btnDelete.Caption = "Delete Design" Then
            btnDelete.Caption = "CONFIRM"
        Else
            If moTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                'Delete the design - permanently
                Dim yMsg(8) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetArchiveState).CopyTo(yMsg, lPos) : lPos += 2
                moTech.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                yMsg(8) = 255
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            Else
                'Delete the design
                Dim yMsg(7) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yMsg, 0)
                moTech.GetGUIDAsString.CopyTo(yMsg, 2)
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
            CloseMe()
            ''Delete the design
            'Dim yMsg(7) As Byte
            'System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yMsg, 0)
            'moTech.GetGUIDAsString.CopyTo(yMsg, 2)
            'MyBase.moUILib.SendMsgToPrimary(yMsg)
            'CloseMe()
        End If
    End Sub

    Public Sub SetDataChanged(ByVal bValue As Boolean)
        'If bValue = True Then
        '    If btnSubmit.Caption <> "Redesign" Then btnSubmit.Caption = "Redesign"
        '    If btnSubmit.Enabled = False Then btnSubmit.Enabled = True
        'Else
        '    If btnSubmit.Caption <> "Research" Then btnSubmit.Caption = "Research"
        '    If moTech Is Nothing = False AndAlso moTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso btnSubmit.Enabled = True Then btnSubmit.Enabled = False
        'End If
        mbWpnFraChanged = bValue
        mbDirtyMeAgain = True
    End Sub

    Private Sub frmWeaponBuilder_OnNewFrame() Handles Me.OnNewFrame
        If mbDirtyMeAgain = True Then
            mbDirtyMeAgain = False
            Me.IsDirty = True
        End If
    End Sub

    Private Sub frmWeaponBuilder_OnRender() Handles Me.OnRender
        If moTech Is Nothing = False Then
            With moTech
                Dim bChanged As Boolean = mbWpnFraChanged

                If bChanged = False Then
                    If txtWeaponName.Caption <> .WeaponName AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        bChanged = True
                    End If
                    If bChanged = False Then
                        If moWpnApp Is Nothing = False AndAlso moWpnApp.GetWeaponTypeID <> .WeaponTypeID Then
                            bChanged = True
                        End If
                    End If
                End If
                If bChanged = True Then
                    If btnSubmit.Caption <> "Redesign" Then btnSubmit.Caption = "Redesign"
                    If btnSubmit.Enabled = False Then btnSubmit.Enabled = True
                Else
                    If btnSubmit.Caption <> "Research" Then btnSubmit.Caption = "Research"
                    If moTech Is Nothing = False AndAlso moTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso btnSubmit.Enabled = True Then btnSubmit.Enabled = False
                End If
            End With
        End If
    End Sub

    Protected Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
        If fraCurrent Is Nothing = False Then fraCurrent.BuilderCostValueChange(bAutoFill)
    End Sub

    Public Overloads Overrides Sub ViewResults(ByRef oTech As Base_Tech, ByVal lProdFactor As Integer)
        If oTech Is Nothing Then Return

        moTech = CType(oTech, WeaponTech)

        Me.btnSubmit.Enabled = False

        With moTech
            cboHullType.FindComboItemData(.HullTypeID)
            cboHullType_ItemSelected(cboHullType.ListIndex)
            cboWeaponType.FindComboItemData(.WeaponClassTypeID)
            cboWeaponType_ItemSelected(cboWeaponType.ListIndex)
            txtWeaponName.Caption = .WeaponName

            If fraCurrent Is Nothing = False Then fraCurrent.ViewResults(moTech, lProdFactor)
            If moWpnApp Is Nothing Then moWpnApp = New frmWeaponAppearance(MyBase.moUILib, Me.Left + Me.Width + 5, Me.Top + Me.Height + 5)
            If moWpnApp Is Nothing = False Then
                moWpnApp.Visible = True
                moWpnApp.SetFromWeaponClassType(CType(.WeaponClassTypeID, WeaponClassType))
                If .WeaponTypeID >= WeaponType.eSolidGreenBeam AndAlso .WeaponTypeID <= WeaponType.eSolidPurpleBeam Then
                    moWpnApp.SetSolidBeamType(1)
                End If
                moWpnApp.SetWeaponTypeID(.WeaponTypeID)
            End If


            lblDesignFlaw.Caption = .GetDesignFlawText()
            lblDesignFlaw.Visible = True

        End With

        If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentResearching AndAlso HasAliasedRights(AliasingRights.eAddResearch) = True Then
            btnSubmit.Caption = "Research"
            btnSubmit.Enabled = True
        End If
        If gbAliased = False Then
            btnDelete.Visible = True 'oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched
            btnRename.Visible = oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched
            'If btnDelete.Visible = True Then
            'btnSubmit.Left = Me.Width \ 2 - btnSubmit.Width \ 2
            'btnCancel.Left = Me.Width - btnDelete.Left - btnCancel.Width
            'End If

            btnDelete.Left = 10
            btnCancel.Left = Me.Width - btnCancel.Width - 10
            Dim lNewW As Int32 = (btnCancel.Left - (btnDelete.Left + btnDelete.Width))
            Dim lGapW As Int32 = lNewW - (btnSubmit.Width + btnAutoFill.Width)
            lGapW \= 3  'for 3 gaps
            btnSubmit.Left = btnCancel.Left - btnSubmit.Width - lGapW
            btnAutoFill.Left = btnDelete.Left + btnDelete.Width + lGapW
        End If

        'Now... what state is the tech in?
        'we don't care, disable weapon type to avoid bugs
        cboWeaponType.Enabled = False
        If oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
            cboHullType.Enabled = False
        End If
        If oTech.ComponentDevelopmentPhase > Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then
            'Ok, show it's research and production cost
            Dim lLeft As Int32 = Me.Left + Me.Width + 5
            If moWpnApp Is Nothing = False Then
                If MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth < 1270 Then
                    moWpnApp.Visible = False
                Else
                    lLeft = moWpnApp.Left + moWpnApp.Width + 5
                End If
            End If

            If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                If mfrmWpnDef Is Nothing Then mfrmWpnDef = New frmProdCost(MyBase.moUILib, "frmWpnDef")
                With mfrmWpnDef
                    .Visible = True
                    .Left = lLeft
                    .Top = MyBase.mfrmBuilderCost.Top
                    .SetFromWeaponDefResult(moTech.oResultWeaponDef, moTech.WeaponClassTypeID)
                    lLeft += .Width + 5
                End With
            ElseIf Not mfrmWpnDef Is Nothing Then : lLeft += mfrmWpnDef.Width + 5
            End If

            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
            With mfrmResCost
                .Visible = True
                .Left = lLeft
                .Top = MyBase.mfrmBuilderCost.Top
                .SetFromProdCost(oTech.oResearchCost, lProdFactor, True, 0, 0)
            End With

            If mfrmProdCost Is Nothing Then mfrmProdCost = New frmProdCost(MyBase.moUILib, "frmProdCost")
            With mfrmProdCost
                .Visible = True
                If MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth >= 1270 Then
                    .Left = mfrmResCost.Left + mfrmResCost.Width + 5
                    .Top = mfrmResCost.Top '+ mfrmResCost.Height + 5
                Else
                    .Left = mfrmResCost.Left '+ mfrmResCost.Width + 5
                    .Top = mfrmResCost.Top + mfrmResCost.Height + 5
                End If


                .SetFromProdCost(oTech.oProductionCost, 1000, False, moTech.HullRequired, moTech.PowerRequired)
            End With

        ElseIf oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
            With mfrmResCost
                .Visible = True
                Dim lLeft As Int32 = Me.Left + Me.Width + 5
                If moWpnApp Is Nothing = False Then
                    lLeft = moWpnApp.Left + moWpnApp.Width + 5
                End If
                If mfrmWpnDef Is Nothing = False Then
                    lLeft += mfrmWpnDef.Width + 5

                End If
                .Left = lLeft
                .Top = Me.Top + Me.Height + 5
                .SetFromFailureCode(oTech.ErrorCodeReason)
            End With
        End If
    End Sub

    Public Sub SetWeaponDetailedStats(ByVal lMinBurn As Int32, ByVal lMaxBurn As Int32, ByVal lMinBeam As Int32, ByVal lMaxBeam As Int32, ByVal lMinChem As Int32, ByVal lMaxChem As Int32, ByVal lMinECM As Int32, ByVal lMaxECM As Int32, ByVal lMinImpact As Int32, ByVal lMaxImpact As Int32, ByVal lMinPiercing As Int32, ByVal lMaxPiercing As Int32, ByVal fROF As Single)
        If mfrmWpnDef Is Nothing Then mfrmWpnDef = New frmProdCost(MyBase.moUILib, "frmWpnDef")
        With mfrmWpnDef
            .Visible = True
            If moWpnApp Is Nothing = False Then
                .Left = moWpnApp.Left + moWpnApp.Width + 5
                .Top = moWpnApp.Top
            Else
                .Left = MyBase.mfrmBuilderCost.Left + MyBase.mfrmBuilderCost.Width + 5
                .Top = MyBase.mfrmBuilderCost.Top
            End If
            If MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth < 1281 Then
                mfrmWpnDef.Width = 200
            End If

            Dim oSB As New System.Text.StringBuilder
            oSB.AppendLine("Rate of Fire: " & fROF.ToString("##0.##"))

            Dim lMaxDmg As Int32 = lMaxBurn + lMaxBeam + lMaxChem + lMaxImpact + lMaxECM + lMaxPiercing
            Dim lMinDmg As Int32 = lMinBurn + lMinBeam + lMinChem + lMinImpact + lMinECM + lMinPiercing
            Dim lFirePowerRating As Int32 = CInt(CSng((lMaxDmg * 4) + (lMinDmg * 8)) / Math.Max(1, fROF))
            oSB.AppendLine("Firepower Rating: " & lFirePowerRating)

            Dim fDPS As Single = (lMaxDmg / fROF)
            oSB.AppendLine()
            oSB.AppendLine("Damage Per Second: " & fDPS.ToString("#,##0.###"))
            oSB.AppendLine()

            oSB.AppendLine("DAMAGE")
            If lMaxBeam <> 0 OrElse lMinBeam <> 0 Then oSB.AppendLine("  Beam: " & lMinBeam & " - " & lMaxBeam)
            If lMinChem <> 0 OrElse lMaxChem <> 0 Then oSB.AppendLine("  Chemical: " & lMinChem & " - " & lMaxChem)
            If lMinBurn <> 0 OrElse lMaxBurn <> 0 Then oSB.AppendLine("  Flame: " & lMinBurn & " - " & lMaxBurn)
            If lMinImpact <> 0 OrElse lMaxImpact <> 0 Then oSB.AppendLine("  Impact: " & lMinImpact & " - " & lMaxImpact)
            If lMaxECM <> 0 OrElse lMinECM <> 0 Then oSB.AppendLine("  Magnetic: " & lMinECM & " - " & lMaxECM)
            If lMinPiercing <> 0 OrElse lMaxPiercing <> 0 Then oSB.AppendLine("  Pierce: " & lMinPiercing & " - " & lMaxPiercing)
            .SetTextDirect(oSB.ToString, "Expected Statistics")
        End With

    End Sub

    Private Sub cboHullType_ItemSelected(ByVal lItemIndex As Integer) Handles cboHullType.ItemSelected

        If fraCurrent Is Nothing Then
            If cboWeaponType.ListIndex > -1 Then
                cboWeaponType_ItemSelected(cboWeaponType.ListIndex)
            End If
        End If

        If fraCurrent Is Nothing = False Then
            If cboHullType.ListIndex > -1 Then
                Dim lID As Int32 = cboHullType.ItemData(cboHullType.ListIndex)
                fraCurrent.lHullTypeID = lID
                'fraCurrent.BuilderCostValueChange(False)
                fraCurrent.CheckForDARequest()
            End If
        End If
    End Sub

    Private msRenameVal As String = ""
    Private Sub btnRename_Click(ByVal sName As String) Handles btnRename.Click
        If moTech Is Nothing = False Then
            If moTech.OwnerID = glPlayerID Then
                If goCurrentPlayer.blCredits < 10000000 Then
                    MyBase.moUILib.AddNotification("You require 10,000,000 credits to rename a tech.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
                Dim sVal As String = txtWeaponName.Caption.Trim
                If sVal.Length > 20 Then
                    MyBase.moUILib.AddNotification("The new name is too long!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                ElseIf sVal.Length = 0 Then
                    MyBase.moUILib.AddNotification("You must enter a name for this tech to be renamed to.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
                If sVal.ToUpper <> moTech.GetComponentName().ToUpper Then
                    msRenameVal = sVal
                    Dim oFrm As New frmMsgBox(MyBase.moUILib, "Renaming a tech design costs 10,000,000 credits in order to have the GTC update all of the registries. Do you wish to proceed?", MsgBoxStyle.YesNo, "Confirm Rename")
                    oFrm.Visible = True
                    AddHandler oFrm.DialogClosed, AddressOf RenameMsgBoxResult
                Else
                    MyBase.moUILib.AddNotification("The new name is not different enough from the old name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            End If
        End If
    End Sub
    Private Sub RenameMsgBoxResult(ByVal lResult As MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            Dim yMsg(27) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityName).CopyTo(yMsg, lPos) : lPos += 2
            moTech.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            System.Text.ASCIIEncoding.ASCII.GetBytes(msRenameVal).CopyTo(yMsg, lPos) : lPos += 20
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            moTech.SetComponentName(msRenameVal)
        End If
    End Sub

    Private Sub btnAutoFill_Click(ByVal sName As String) Handles btnAutoFill.Click
        'we drop focus for our controls just incase some invalid value causes recalculating
        If goUILib.FocusedControl Is Nothing = False Then
            goUILib.FocusedControl.HasFocus = False
            goUILib.FocusedControl = Nothing
        End If
        BuilderCostValueChange(True)
    End Sub
End Class