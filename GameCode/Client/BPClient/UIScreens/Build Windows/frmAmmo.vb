Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAmmo
    Inherits UIWindow

    Private Structure WeaponAmmoDetail
        Public lEDW_ID As Int32     'entity def weapon id (UDW_ID for units or SDW_ID for structures)
        Public lMaxAmmo As Int32
        Public lCurrentAmmo As Int32
        Public lWeaponTechID As Int32
        Public yArc As Byte

        Public bHasUnknown As Boolean

        Public Function GetListBoxText() As String
            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lWeaponTechID, ObjectType.eWeaponTech)
            Dim sName As String
            If oTech Is Nothing = False Then
                sName = oTech.GetComponentName
            Else : sName = GetCacheObjectValue(lWeaponTechID, ObjectType.eWeaponTech)
            End If
            oTech = Nothing
            If sName.ToUpper = "UNKNOWN" Then bHasUnknown = True Else bHasUnknown = False

            Dim sArc As String
            Select Case yArc
                Case UnitArcs.eBackArc
                    sArc = "Rear"
                Case UnitArcs.eForwardArc
                    sArc = "Front"
                Case UnitArcs.eLeftArc
                    sArc = "Left"
                Case UnitArcs.eRightArc
                    sArc = "Right"
                Case Else
                    sArc = "All"
            End Select

            Dim sAmmoStr As String
            If lCurrentAmmo = -1 OrElse lMaxAmmo = -1 Then
                sAmmoStr = "Does Not Use Ammo"
            Else : sAmmoStr = lCurrentAmmo.ToString("#,##0") & " / " & lMaxAmmo.ToString("#,##0")
            End If

            Return sName.PadRight(23, " "c) & sArc.PadRight(9, " "c) & sAmmoStr
        End Function
    End Structure
    Private Structure CargoAmmoDetail
        Public lWeaponTechID As Int32
        Public Quantity As Int32
        Public fAmmoSize As Single

        Public bHasUnknown As Boolean

        Public Function GetListBoxText() As String
            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lWeaponTechID, ObjectType.eWeaponTech)
            Dim sName As String

            If oTech Is Nothing = False Then
                sName = oTech.GetComponentName
            Else : sName = GetCacheObjectValue(lWeaponTechID, ObjectType.eWeaponTech)
            End If
            If sName.ToUpper = "UNKNOWN" Then bHasUnknown = True Else bHasUnknown = False

            Dim lSpace As Int32 = CInt(Math.Ceiling(fAmmoSize * Quantity))

            Return sName.PadRight(23, " "c) & Quantity.ToString("#,##0").PadRight(12, " "c) & lSpace.ToString("#,##0")
        End Function
    End Structure

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblWeapons As UILabel
    Private lblArc As UILabel
    Private lblAmmoCaps As UILabel
    Private lblQty1 As UILabel
    Private lblWeaponType As UILabel
    Private lblCargoQty As UILabel
    Private lblCargoSpace As UILabel
    Private lblCargo As UILabel
    Private lblCapacity As UILabel
    Private lnDiv3 As UILine
    Private lblAmmoSize As UILabel
    Private lnDiv2 As UILine
    Private lblQty2 As UILabel
    Private txtQuantity1 As UITextBox
    Private txtQuantity2 As UITextBox

    Private WithEvents lstWeapons As UIListBox
    Private WithEvents btnClose As UIButton
    Private WithEvents btnReloadAll As UIButton
    Private WithEvents btnLoadAmmo As UIButton
    Private WithEvents lstCargo As UIListBox
    Private WithEvents cboWeaponType As UIComboBox
    Private WithEvents btnLoadCargo As UIButton

    Private mlParentID As Int32
    Private miParentTypeID As Int16
    Private mlEntityID As Int32
    Private miEntityTypeID As Int16

    Private mlParentEntityIdx As Int32 = -1

    'The Selection's Details...
    Private mlTotalCargoCap As Int32 = 0
    Private mlAvailCargoCap As Int32 = 0
    Private muAmmoDetails() As WeaponAmmoDetail
    Private mlAmmoDetailUB As Int32 = -1
    Private muCargoAmmo() As CargoAmmoDetail
    Private mlCargoAmmoUB As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmAmmo initial props
        With Me
            .ControlName = "frmAmmo"
            .Left = 235
            .Top = 56
            .Width = 512
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 2
        End With

        'lstWeapons initial props
        lstWeapons = New UIListBox(oUILib)
        With lstWeapons
            .ControlName = "lstWeapons"
            .Left = 5
            .Top = 50
            .Width = 502
            .Height = 150
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstWeapons, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 4
            .Width = 400
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
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
            .Top = 26
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

        'lblWeapons initial props
        lblWeapons = New UILabel(oUILib)
        With lblWeapons
            .ControlName = "lblWeapons"
            .Left = 5
            .Top = 30
            .Width = 159
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Weapon"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblWeapons, UIControl))

        'txtQuantity1 initial props
        txtQuantity1 = New UITextBox(oUILib)
        With txtQuantity1
            .ControlName = "txtQuantity1"
            .Left = 305
            .Top = 205
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
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtQuantity1, UIControl))

        'btnLoadAmmo initial props
        btnLoadAmmo = New UIButton(oUILib)
        With btnLoadAmmo
            .ControlName = "btnLoadAmmo"
            .Left = 410
            .Top = 205
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Load Ammo"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to load the quantity of ammo into the selected weapon." & vbCrLf & "Any overflow will be placed in the cargo bay (if possible)."
        End With
        Me.AddChild(CType(btnLoadAmmo, UIControl))

        'btnReloadAll initial props
        btnReloadAll = New UIButton(oUILib)
        With btnReloadAll
            .ControlName = "btnReloadAll"
            .Left = 5
            .Top = btnLoadAmmo.Top
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Reload All"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to load all weapons on this entity to their max."
        End With
        Me.AddChild(CType(btnReloadAll, UIControl))

        'lblArc initial props
        lblArc = New UILabel(oUILib)
        With lblArc
            .ControlName = "lblArc"
            .Left = 195
            .Top = 30
            .Width = 26
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Arc"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblArc, UIControl))

        'lblAmmoCaps initial props
        lblAmmoCaps = New UILabel(oUILib)
        With lblAmmoCaps
            .ControlName = "lblAmmoCaps"
            .Left = 265
            .Top = 30
            .Width = 180
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Ammo / Ammo Capacity"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAmmoCaps, UIControl))

        'lblQty1 initial props
        lblQty1 = New UILabel(oUILib)
        With lblQty1
            .ControlName = "lblQty1"
            .Left = 235
            .Top = 205
            .Width = 64
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Quantity:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
        End With
        Me.AddChild(CType(lblQty1, UIControl))

        'lstCargo initial props
        lstCargo = New UIListBox(oUILib)
        With lstCargo
            .ControlName = "lstCargo"
            .Left = 5
            .Top = 340 '300
            .Width = 502
            .Height = 150 ' 160
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstCargo, UIControl))

        'lblWeaponType initial props
        lblWeaponType = New UILabel(oUILib)
        With lblWeaponType
            .ControlName = "lblWeaponType"
            .Left = 5
            .Top = 320 '280
            .Width = 159
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Weapon Type"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblWeaponType, UIControl))

        'lblCargoQty initial props
        lblCargoQty = New UILabel(oUILib)
        With lblCargoQty
            .ControlName = "lblCargoQty"
            .Left = 195
            .Top = 320 '280
            .Width = 64
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Quantity"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCargoQty, UIControl))

        'lblCargoSpace initial props
        lblCargoSpace = New UILabel(oUILib)
        With lblCargoSpace
            .ControlName = "lblCargoSpace"
            .Left = 290
            .Top = 320 '280
            .Width = 145
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Cargo Space Used"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCargoSpace, UIControl))

        'lblCargo initial props
        lblCargo = New UILabel(oUILib)
        With lblCargo
            .ControlName = "lblCargo"
            .Left = 5
            .Top = 255
            .Width = 180
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Ammunition in Cargo Bay"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCargo, UIControl))

        'lblCapacity initial props
        lblCapacity = New UILabel(oUILib)
        With lblCapacity
            .ControlName = "lblCapacity"
            .Left = 255
            .Top = 255
            .Width = 250
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Capacity: 123456789 / 123456789"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCapacity, UIControl))

        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = 1
            .Top = 275
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'lblAmmoSize initial props
        lblAmmoSize = New UILabel(oUILib)
        With lblAmmoSize
            .ControlName = "lblAmmoSize"
            .Left = 5
            .Top = 300 '475
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Size: 0.0 Hull Each"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAmmoSize, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 1
            .Top = 245
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'lblQty2 initial props
        lblQty2 = New UILabel(oUILib)
        With lblQty2
            .ControlName = "lblQty2"
            .Left = 235
            .Top = 280 '455
            .Width = 64
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Quantity:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
        End With
        Me.AddChild(CType(lblQty2, UIControl))

        'txtQuantity2 initial props
        txtQuantity2 = New UITextBox(oUILib)
        With txtQuantity2
            .ControlName = "txtQuantity2"
            .Left = 305
            .Top = 280 '455
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
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtQuantity2, UIControl))

        'btnLoadCargo initial props
        btnLoadCargo = New UIButton(oUILib)
        With btnLoadCargo
            .ControlName = "btnLoadCargo"
            .Left = 410
            .Top = 280 '455
            .Width = 100
            .Height = 18
            .Enabled = False
            .Visible = True
            .Caption = "Load Ammo"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to load the quantity of ammo of the selected weapon type into the cargo bay."
        End With
        Me.AddChild(CType(btnLoadCargo, UIControl))

        'cboWeaponType initial props
        cboWeaponType = New UIComboBox(oUILib)
        With cboWeaponType
            .ControlName = "cboWeaponType"
            .Left = 5
            .Top = 280 '455
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
            .ToolTipText = "Select a weapon type to load ammo into the cargo bay."
        End With
        Me.AddChild(CType(cboWeaponType, UIControl))

        FillWeaponTypes()

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub FillWeaponTypes()
        cboWeaponType.Clear()

        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        For X As Int32 = 0 To goCurrentPlayer.mlTechUB
            If goCurrentPlayer.moTechs(X) Is Nothing = False Then
                If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eWeaponTech AndAlso goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then

                    Dim yType As WeaponClassType = CType(goCurrentPlayer.moTechs(X), WeaponTech).WeaponClassTypeID
                    If yType <> WeaponClassType.eBomb AndAlso yType <> WeaponClassType.eMine AndAlso yType <> WeaponClassType.eMissile AndAlso yType <> WeaponClassType.eProjectile Then Continue For
                    'If (CType(goCurrentPlayer.moTechs(X), WeaponTech).GetAmmoSize > 0.0F) = False Then Continue For

                    Dim lIdx As Int32 = -1

                    Dim sName As String = goCurrentPlayer.moTechs(X).GetComponentName

                    For Y As Int32 = 0 To lSortedUB
                        If goCurrentPlayer.moTechs(lSorted(Y)).GetComponentName > sName Then
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
                End If
            End If
        Next X

        For X As Int32 = 0 To lSortedUB
            With goCurrentPlayer.moTechs(lSorted(X))
                cboWeaponType.AddItem(.GetComponentName)
                cboWeaponType.ItemData(cboWeaponType.NewIndex) = .ObjectID
            End With
        Next X
    End Sub

    Public Sub SetFromEntity(ByVal lParentID As Int32, ByVal iParentTypeID As Int16, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal sName As String)

        mlParentEntityIdx = -1

        mlParentID = lParentID
        miParentTypeID = iParentTypeID
        mlEntityID = lEntityID
        miEntityTypeID = iEntityTypeID

        If goCurrentEnvir Is Nothing = False Then
            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) = mlParentID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iParentTypeID Then
                    mlParentEntityIdx = X
                    Exit For
                End If
            Next X
        End If

        Me.lblTitle.Caption = "Ammo for " & sName

        Dim yMsg(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityAmmo).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(lEntityID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(iEntityTypeID).CopyTo(yMsg, 6)
        MyBase.moUILib.SendMsgToPrimary(yMsg)

    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnLoadAmmo_Click(ByVal sName As String) Handles btnLoadAmmo.Click
		If Val(txtQuantity1.Caption) > Int32.MaxValue OrElse Val(txtQuantity1.Caption) < Int32.MinValue Then
			MyBase.moUILib.AddNotification("Invalid entry in the Quantity box.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		If lstWeapons.ListIndex = -1 Then
			MyBase.moUILib.AddNotification("Select an item in the list first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim lQty As Int32 = CInt(Val(txtQuantity1.Caption))
		Dim lEDW_ID As Int32 = lstWeapons.ItemData(lstWeapons.ListIndex)

		CreateAndSendReloadRequest(lEDW_ID, lQty, 0)

		MyBase.moUILib.AddNotification("Load Ammo request submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
	End Sub

    Private Sub lstWeapons_ItemClick(ByVal lIndex As Integer) Handles lstWeapons.ItemClick
        If lIndex > -1 Then
            Dim lID As Int32 = lstWeapons.ItemData(lIndex)
            For X As Int32 = 0 To mlAmmoDetailUB
                If muAmmoDetails(X).lEDW_ID = lID Then
                    Dim lQty As Int32 = muAmmoDetails(X).lMaxAmmo - muAmmoDetails(X).lCurrentAmmo
                    If lQty > 0 Then
                        txtQuantity1.Caption = lQty.ToString
                    Else : txtQuantity1.Caption = "0"
                    End If
                    Exit For
                End If
            Next X
        End If
    End Sub

	Private Sub btnLoadCargo_Click(ByVal sName As String) Handles btnLoadCargo.Click
		If Val(txtQuantity2.Caption) > Int32.MaxValue OrElse Val(txtQuantity2.Caption) < Int32.MinValue Then
			MyBase.moUILib.AddNotification("Invalid entry in the Quantity box.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		If cboWeaponType.ListIndex = -1 Then
			MyBase.moUILib.AddNotification("Select an item in the list first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim lQty As Int32 = CInt(Val(txtQuantity2.Caption))
		Dim lWpn_ID As Int32 = cboWeaponType.ItemData(cboWeaponType.ListIndex)

		CreateAndSendReloadRequest(lWpn_ID, lQty, 1)

		MyBase.moUILib.AddNotification("Load Ammo request submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
	End Sub

    Private Sub CreateAndSendReloadRequest(ByVal lWpnID As Int32, ByVal lQty As Int32, ByVal yType As Byte)
        Dim yMsg(16) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestLoadAmmo).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(mlEntityID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(miEntityTypeID).CopyTo(yMsg, 6)
        System.BitConverter.GetBytes(lWpnID).CopyTo(yMsg, 8)
        System.BitConverter.GetBytes(lQty).CopyTo(yMsg, 12)
        yMsg(16) = yType
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Private Sub cboWeaponType_ItemSelected(ByVal lItemIndex As Integer) Handles cboWeaponType.ItemSelected
        If lItemIndex <> -1 Then
            Dim lID As Int32 = cboWeaponType.ItemData(lItemIndex)
            Dim oWpn As WeaponTech = CType(goCurrentPlayer.GetTech(lID, ObjectType.eWeaponTech), WeaponTech)
            If oWpn Is Nothing = False Then
                'lblAmmoSize.Caption = "Size: " & oWpn.GetAmmoSize.ToString("#,##0.0###") & " Hull Each"
                ' btnLoadCargo.Enabled = oWpn.GetAmmoSize > 0.0F
            End If

        End If
    End Sub

    Private Sub lstCargo_ItemClick(ByVal lIndex As Integer) Handles lstCargo.ItemClick
        If lIndex > -1 Then
            Dim lID As Int32 = lstCargo.ItemData(lIndex)
            cboWeaponType.ListIndex = -1
            cboWeaponType.FindComboItemData(lID)
        End If
    End Sub

    Public Sub HandleRequestAmmo(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If lID = mlEntityID AndAlso iTypeID = miEntityTypeID Then
            mlTotalCargoCap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            mlAvailCargoCap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim lCnt As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            mlAmmoDetailUB = lCnt - 1
            ReDim muAmmoDetails(mlAmmoDetailUB)
            For X As Int32 = 0 To mlAmmoDetailUB
                With muAmmoDetails(X)
                    .lEDW_ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lWeaponTechID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .yArc = yData(lPos) : lPos += 1
                    .lMaxAmmo = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .lCurrentAmmo = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                End With
            Next X

            lCnt = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            mlCargoAmmoUB = lCnt - 1
            ReDim muCargoAmmo(mlCargoAmmoUB)
            For X As Int32 = 0 To mlCargoAmmoUB
                With muCargoAmmo(X)
                    .lWeaponTechID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Quantity = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .fAmmoSize = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                End With
            Next X
        End If

        Me.IsDirty = True
    End Sub

    Private Sub frmAmmo_OnRender() Handles Me.OnRender
        Try
            'Check our lists... first check for adds/updates
            For X As Int32 = 0 To mlAmmoDetailUB
                Dim bFound As Boolean = False
                For Y As Int32 = 0 To lstWeapons.ListCount - 1
                    If lstWeapons.ItemData(Y) = muAmmoDetails(X).lEDW_ID Then
                        'Ok, found it
                        bFound = True
                        Dim sTemp As String = muAmmoDetails(X).GetListBoxText()
                        If lstWeapons.List(Y) <> sTemp Then lstWeapons.List(Y) = sTemp
                        Exit For
                    End If
                Next Y

                If bFound = False Then
                    lstWeapons.AddItem(muAmmoDetails(X).GetListBoxText())
                    lstWeapons.ItemData(lstWeapons.NewIndex) = muAmmoDetails(X).lEDW_ID
                End If
            Next X

            'Now, check for removals
            Dim bDone As Boolean = False
            While bDone = False
                bDone = True
                For X As Int32 = 0 To lstWeapons.ListCount - 1
                    Dim bFound As Boolean = False
                    For Y As Int32 = 0 To mlAmmoDetailUB
                        If muAmmoDetails(Y).lEDW_ID = lstWeapons.ItemData(X) Then
                            bFound = True
                            Exit For
                        End If
                    Next Y
                    If bFound = False Then
                        lstWeapons.RemoveItem(X)
                        bDone = False
                    End If
                Next X
            End While

            'Check our Cargo List... first check for adds/updates
            For X As Int32 = 0 To mlCargoAmmoUB
                Dim bFound As Boolean = False
                For Y As Int32 = 0 To lstCargo.ListCount - 1
                    If lstCargo.ItemData(Y) = muCargoAmmo(X).lWeaponTechID Then
                        'Ok, found it
                        bFound = True
                        Dim sTemp As String = muCargoAmmo(X).GetListBoxText()
                        If lstCargo.List(Y) <> sTemp Then lstCargo.List(Y) = sTemp
                        Exit For
                    End If
                Next Y

                If bFound = False Then
                    lstCargo.AddItem(muCargoAmmo(X).GetListBoxText())
                    lstCargo.ItemData(lstCargo.NewIndex) = muCargoAmmo(X).lWeaponTechID
                End If
            Next X

            'Now, check for removals
            bDone = False
            While bDone = False
                bDone = True
                For X As Int32 = 0 To lstCargo.ListCount - 1
                    Dim bFound As Boolean = False
                    For Y As Int32 = 0 To mlCargoAmmoUB
                        If muCargoAmmo(Y).lWeaponTechID = lstCargo.ItemData(X) Then
                            bFound = True
                            Exit For
                        End If
                    Next Y
                    If bFound = False Then
                        lstCargo.RemoveItem(X)
                        bDone = False
                    End If
                Next X
            End While
        Catch
            'Do nothing, likely a race condition
        End Try

        Dim sCargo As String = "Capacity: " & (mlTotalCargoCap - mlAvailCargoCap).ToString("#,##0") & " / " & mlTotalCargoCap.ToString("#,##0")
        If lblCapacity.Caption <> sCargo Then lblCapacity.Caption = sCargo
    End Sub

	Private Sub btnReloadAll_Click(ByVal sName As String) Handles btnReloadAll.Click
		CreateAndSendReloadRequest(-1, -1, 2)
		MyBase.moUILib.AddNotification("Load Ammo request submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
	End Sub

    Public Sub CheckForAmmoProdComplete(ByVal lCreatorID As Int32, ByVal iCreatorTypeID As Int16)
        If lCreatorID = mlParentID AndAlso iCreatorTypeID = miParentTypeID Then
            Dim yMsg(7) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityAmmo).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(mlEntityID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(miEntityTypeID).CopyTo(yMsg, 6)
            MyBase.moUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub
End Class