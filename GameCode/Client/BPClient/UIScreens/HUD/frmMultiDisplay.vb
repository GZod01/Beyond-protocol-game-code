Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Multiple Selected display
Public Class frmMultiDisplay
    Inherits UIWindow

	Private WithEvents lstData As UIListBox
	Private lblFormation As UILabel
	Private WithEvents cboFormation As UIComboBox
    Private lblTotalSelected As UILabel
    Private WithEvents btnUnitGoto As UIButton
    'Private WithEvents btnDamaged As UIButton
    'Private WithEvents btnUnDamaged As UIButton
    Private WithEvents btnFilterSel As UIButton
    Private WithEvents txtRename As UITextBox
    Private lblName As UILabel

    Private rcSendToOrbit As Rectangle
    Private rcMoveThenForm As Rectangle

    Private mbPowerDisplay As Boolean = False
    Private rcPowerOn As Rectangle
    Private rcPowerOff As Rectangle

    Private mbHasUnknowns As Boolean
    Private mbLoading As Boolean = True
    Private mbDidNameCheck As Boolean = False

    Private mlLastSetEntityCycle As Int32 = -1
    Private mlLastCheckStoreCycle As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmMultiDisplay initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eMultiDisplay
            .ControlName = "frmMultiDisplay"
            '.Left = muSettings.MiniMapLocX + muSettings.MiniMapWidthHeight
            '.Top = (muSettings.MiniMapLocY + muSettings.MiniMapWidthHeight) - 86
            .Left = 0
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - 86

            If muSettings.MultiSelectLocX <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Left = muSettings.MultiSelectLocX
            If muSettings.MultiSelectLocY <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Top = muSettings.MultiSelectLocY
            If muSettings.MultiSelectWidth <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Width = muSettings.MultiSelectWidth Else .Width = 190
            If muSettings.MultiSelectHeight <> -1 AndAlso NewTutorialManager.TutorialOn = False Then .Height = muSettings.MultiSelectHeight Else .Height = 128

            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(-16768960)
            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
            .Resizable = True
            .ResizeLimits = New System.Drawing.Rectangle(190, 128, 512, 512)
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'txtRename initial props
        txtRename = New UITextBox(oUILib)
        With txtRename
            .ControlName = "txtRename"
            .Left = 47 '3
            .Top = 3
            .Width = Me.Width - 50
            .Height = 16
            .Enabled = Not gbAliased
            .Visible = True
            .Caption = "ENTITY NAME"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .BackColorEnabled = Me.FillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
            .DoNotRender = UITextBox.DoNotRenderSetting.eFillColor
        End With
        Me.AddChild(CType(txtRename, UIControl))

        'lstData initial props
        lstData = New UIListBox(oUILib)
        With lstData
            .ControlName = "lstData"
            .Left = 2
            .Top = 65 '45
            .Width = Me.Width - 4 '186
            .Height = Me.Height - .Top
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-1)
            '.FillColor = System.Drawing.Color.FromArgb(-16768960)
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            '.HighlightColor = System.Drawing.Color.FromArgb(-14663588)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
        End With
		Me.AddChild(CType(lstData, UIControl))

        'lblName initial props
        lblName = New UILabel(oUILib)
        With lblName
            .ControlName = "lblName"
            .Left = 3
            .Top = 3
            .Width = 45
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Name: "
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblName, UIControl))

		'lblTotalSelected initial props
		lblTotalSelected = New UILabel(oUILib)
		With lblTotalSelected
			.ControlName = "lblTotalSelected"
			.Left = 3
            .Top = 45 ' 25
			.Width = Me.Width - 6
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Selected: 0"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTotalSelected, UIControl))

		'lblFormation initial props
		lblFormation = New UILabel(oUILib)
		With lblFormation
			.ControlName = "lblFormation"
			.Left = 3
            .Top = 25 ' 5
			.Width = 65
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Formation:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblFormation, UIControl))

        'btnUnitGoto initial props
        btnUnitGoto = New UIButton(oUILib)
        With btnUnitGoto
            .ControlName = "btnUnitGoto"
            .Left = Me.Width - 120
            .Top = 46 ' 26
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)

            '.ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eDirectionalArrow)
            .ControlImageRect_Disabled = .ControlImageRect
            .ControlImageRect_Normal = .ControlImageRect
            .ControlImageRect_Pressed = .ControlImageRect
            .ToolTipText = "Click to send the selected units to a desired location."
        End With
        Me.AddChild(CType(btnUnitGoto, UIControl))

        ''Two buttons before the formation dropdown.
        ''btnDamaged initial props
        'btnDamaged = New UIButton(oUILib)
        'With btnDamaged
        '    .ControlName = "btnDamaged"
        '    .Left = Me.Width - 96
        '    .Top = 46 ' 26
        '    .Width = 16
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "D"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = CType(5, DrawTextFormat)
        '    .ControlImageRect = New Rectangle(0, 0, 120, 32)
        '    .ToolTipText = "Click to deselect damaged units."
        'End With
        'Me.AddChild(CType(btnDamaged, UIControl))

        ''btnUnDamaged initial props
        'btnUnDamaged = New UIButton(oUILib)
        'With btnUnDamaged
        '    .ControlName = "btnUnDamaged"
        '    .Left = Me.Width - 72
        '    .Top = 46 ' 26
        '    .Width = 16
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "U"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = CType(5, DrawTextFormat)
        '    .ControlImageRect = New Rectangle(0, 0, 120, 32)
        '    .ToolTipText = "Click to deselect undamaged units."
        'End With
        'Me.AddChild(CType(btnUnDamaged, UIControl))

        'btnFilterSel initial props
        btnFilterSel = New UIButton(oUILib)
        With btnFilterSel
            .ControlName = "btnFilterSel"
            .Left = Me.Width - 120
            .Top = 46 ' 26
            .Width = 16
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "F"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to filter selected units or facilities."
        End With
        Me.AddChild(CType(btnFilterSel, UIControl))

		'cboFormations initial props
		cboFormation = New UIComboBox(oUILib)
		With cboFormation
			.ControlName = "cboFormation"
			.Left = 70
            .Top = 25 ' 5
            .Width = Me.Width - 72 '118
			.Height = 18
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .l_ListBoxHeight = Me.Height - .Top - .Height - 2
		End With
        Me.AddChild(CType(cboFormation, UIControl))

        rcSendToOrbit = New Rectangle(Me.Width - 16 - 4, lblTotalSelected.Top, 16, 16)
        rcMoveThenForm = New Rectangle(Me.Width - 16 - 16 - 4, lblTotalSelected.Top, 16, 16)
        'btnUnDamaged.Left = rcMoveThenForm.Left - btnUnDamaged.Width - 1
        'btnDamaged.Left = btnUnDamaged.Left - btnDamaged.Width - 1
        btnFilterSel.Left = rcMoveThenForm.Left - btnFilterSel.Width - 1
        btnUnitGoto.Left = btnFilterSel.Left - btnUnitGoto.Width - 1
        rcPowerOn = New Rectangle(Me.Width - 16 - 16 - 4, lblTotalSelected.Top + 1, 16, 16)
        rcPowerOff = New Rectangle(Me.Width - 16 - 4, lblTotalSelected.Top + 1, 16, 16)

        mbLoading = False
    End Sub

    Public Sub SetFromEntity(ByVal lEntityIndex As Int32)
        If goCurrentEnvir Is Nothing Then Exit Sub
        If lEntityIndex < 0 OrElse lEntityIndex > goCurrentEnvir.lEntityUB Then Exit Sub
        If goCurrentEnvir.oEntity(lEntityIndex) Is Nothing Then Exit Sub
        'If goCurrentEnvir.oEntity(lEntityIndex).oUnitDef Is Nothing Then Exit Sub

        Dim X As Int32
        Dim sName As String = goCurrentEnvir.oEntity(lEntityIndex).EntityName
        If sName Is Nothing OrElse sName = "" Then
            sName = "Unknown"
        End If

        If goCurrentEnvir.oEntity(lEntityIndex).oMesh.bLandBased = True Then
            If btnUnitGoto.Visible = True Then btnUnitGoto.Visible = False
            cboFormation.Enabled = False
            cboFormation.ListIndex = -1
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oEntity(lEntityIndex).ObjTypeID = ObjectType.eFacility Then
                mbPowerDisplay = True
            End If
        Else
            If btnUnitGoto.Visible = False Then btnUnitGoto.Visible = True
        End If


        'always set this so that on the next check, we clear it
        mbHasUnknowns = True
        mbDidNameCheck = False

        For X = 0 To lstData.ListCount - 1
            If lstData.ItemData(X) = lEntityIndex Then
                Exit Sub
            End If
        Next X

        'Ok, we are here... need to add the item
        lstData.AddItem(sName)
        lstData.ItemData(lstData.NewIndex) = lEntityIndex


        mlLastSetEntityCycle = glCurrentCycle
    End Sub

    Public Sub ClearItems()
		lstData.Clear()
		lstData.ListIndex = -1
		cboFormation.ListIndex = 0
        cboFormation.Enabled = True
        mbPowerDisplay = False
    End Sub

    Protected Overrides Sub Finalize()
        Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
        MyBase.Finalize()
    End Sub

    Private Function CheckSendToOrbitRegion(ByVal lX As Int32, ByVal lY As Int32) As Boolean
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return False
        If oEnvir.ObjTypeID <> ObjectType.ePlanet Then Return False
        If cboFormation.Enabled = False Then Return False

        If NewTutorialManager.TutorialOn = True Then
            If goCurrentPlayer.yPlayerPhase = 0 Then Return False
            Dim sParms() As String = {"SendToOrbit"}
            If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return False
        End If

        Return rcSendToOrbit.Contains(lX, lY)
    End Function
    Private Function CheckMoveThenForm(ByVal lX As Int32, ByVal lY As Int32) As Boolean
        If cboFormation.Enabled = False Then Return False
        Return rcMoveThenForm.Contains(lX, lY)
    End Function
    Private Function CheckPowerOn(ByVal lX As Int32, ByVal ly As Int32) As Boolean
        If mbPowerDisplay = False Then Return False
        Return rcPowerOn.Contains(lX, ly)
    End Function
    Private Function CheckPowerOff(ByVal lX As Int32, ByVal ly As Int32) As Boolean
        If mbPowerDisplay = False Then Return False
        Return rcPowerOff.Contains(lX, ly)
    End Function

    Private Sub frmMultiDisplay_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        Try
            Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
            Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y
            If CheckSendToOrbitRegion(lX, lY) = True Then
                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                If oEnvir Is Nothing Then Return
                Dim lDX As Int32
                Dim lDZ As Int32
                Dim lDID As Int32
                Dim iDTypeID As Int16
                With CType(oEnvir.oGeoObject, Planet)
                    lDX = .LocX
                    lDZ = .LocZ
                    If .ParentSystem Is Nothing Then Return
                    lDID = .ParentSystem.ObjectID
                    iDTypeID = .ParentSystem.ObjTypeID
                End With
                MyBase.moUILib.GetMsgSys.SendMoveRequestMsgEx(lDX, lDZ, 0, False, False, lDID, iDTypeID)

                BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eGotoOrbitClick, 1, 0)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            ElseIf CheckMoveThenForm(lX, lY) = True Then
                muSettings.bFormationMoveThenForm = Not muSettings.bFormationMoveThenForm
                Me.IsDirty = True
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            ElseIf CheckPowerOn(lX, lY) = True Then
                Dim x As Int32 = 0
                If TogglePower(True) = False Then Return
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            ElseIf CheckPowerOff(lX, lY) = True Then
                Dim y As Int32 = 0
                If TogglePower(False) = False Then Return
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            End If
        Catch
        End Try
    End Sub
    Private Function TogglePower(ByVal bSetState As Boolean) As Boolean
        If goCurrentEnvir Is Nothing Then Exit Function
        If goCurrentEnvir.ObjTypeID <> ObjectType.ePlanet Then Exit Function

        If HasAliasedRights(AliasingRights.eAlterAutoLaunchPower) = False Then
            MyBase.moUILib.AddNotification("You lack rights to alter auto-launch and power settings.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return False
        End If

        Dim lCnt As Int32 = 0
        Dim lPos As Int32 = 0
        Dim iLen As Int16 = 12
        Dim lStatus As Int32
        If bSetState = True Then
            lStatus = elUnitStatus.eFacilityPowered
        Else
            lStatus = -elUnitStatus.eFacilityPowered
        End If

        For X As Int32 = 0 To lstData.ListCount - 1
            Dim lIdx As Int32 = lstData.ItemData(X)
            If goCurrentEnvir.lEntityIdx(lIdx) > -1 Then
                Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(lIdx)
                If oEntity Is Nothing = False Then
                    If oEntity.ObjectID <> -1 AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.bSelected = True AndAlso oEntity.oMesh.bLandBased = True AndAlso oEntity.ObjTypeID = ObjectType.eFacility Then
                        If (bSetState = True AndAlso (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0) Then
                            lCnt += 1
                        ElseIf (bSetState = False AndAlso (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0) Then
                            lCnt += 1
                        End If
                        'lCnt += 1
                    End If
                End If
            End If
        Next X
        If lCnt > 0 Then
            Dim yMsg(((iLen + 2) * lCnt) - 1) As Byte

            For X As Int32 = 0 To lstData.ListCount - 1
                Dim lIdx As Int32 = lstData.ItemData(X)
                If goCurrentEnvir.lEntityIdx(lIdx) > -1 Then
                    Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(lIdx)
                    If oEntity Is Nothing = False Then
                        If oEntity.ObjectID <> -1 AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.bSelected = True AndAlso oEntity.oMesh.bLandBased = True AndAlso oEntity.ObjTypeID = ObjectType.eFacility Then
                            'Create Packet 
                            If (bSetState = True AndAlso (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0) OrElse (bSetState = False AndAlso (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0) Then
                                System.BitConverter.GetBytes(iLen).CopyTo(yMsg, lPos) : lPos += 2
                                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yMsg, lPos) : lPos += 2
                                oEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                                System.BitConverter.GetBytes(lStatus).CopyTo(yMsg, lPos) : lPos += 4
                            End If
                        End If
                    End If
                End If
            Next
            'Dim ofrmAdv As frmAdvanceDisplay = CType(goUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
            'If ofrmAdv Is Nothing = False Then
            '    ofrmAdv.mbCurrentActive = bSetState
            'End If
            If yMsg Is Nothing = False Then
                MyBase.moUILib.SendLenAppendedMsgToPrimary(yMsg)
                If bSetState Then
                    goUILib.AddNotification(lCnt.ToString & " facilities requesting poweron.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Else
                    goUILib.AddNotification(lCnt.ToString & " facilities requesting power shutdown.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            End If
        End If
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
    End Function

    Private Sub frmMultiDisplay_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Try
            Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
            Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y
            If CheckSendToOrbitRegion(lX, lY) = True Then
                MyBase.moUILib.SetToolTip("Click to send the selected items into orbit.", lMouseX, lMouseY)
            ElseIf CheckMoveThenForm(lX, lY) = True Then
                MyBase.moUILib.SetToolTip("Click to toggle whether formation orders are Move Then Form (Green) or Form Then Move (Red).", lMouseX, lMouseY)
            ElseIf CheckPowerOn(lX, lY) = True Then
                MyBase.moUILib.SetToolTip("Click to power on all selected facilities.", lMouseX, lMouseY)
            ElseIf CheckPowerOff(lX, lY) = True Then
                MyBase.moUILib.SetToolTip("Click to power off all selected facilities.", lMouseX, lMouseY)
            End If
        Catch
        End Try
    End Sub

    Private Sub frmMultiDisplay_OnNewFrame() Handles Me.OnNewFrame

        If mlLastSetEntityCycle > mlLastCheckStoreCycle Then
            CheckMultiSelectLookup()
            mlLastCheckStoreCycle = mlLastSetEntityCycle
        End If

        If mbHasUnknowns = True Then
            mbHasUnknowns = False
            If goCurrentEnvir Is Nothing = False Then
                For X As Int32 = 0 To lstData.ListCount - 1
                    If lstData.List(X) = "Unknown" Then
                        If goCurrentEnvir.lEntityIdx(lstData.ItemData(X)) <> -1 AndAlso goCurrentEnvir.oEntity(lstData.ItemData(X)) Is Nothing = False Then
                            With goCurrentEnvir.oEntity(lstData.ItemData(X))
                                If .OwnerID <> glPlayerID Then Continue For
                                .EntityName = GetCacheObjectValue(.ObjectID, .ObjTypeID)
                                If .EntityName <> "Unknown" Then lstData.List(X) = .EntityName
                            End With
                            mbHasUnknowns = True
                        End If
                    End If
                Next X
            End If
        End If

        If mbHasUnknowns = False AndAlso mbDidNameCheck = False AndAlso lstData.ListCount > 0 Then
            If goCurrentEnvir Is Nothing = False Then
                Dim sName As String = goCurrentEnvir.oEntity(lstData.ItemData(0)).EntityName
                Dim bPass As Boolean = True

                ' check if all listed units have the same name
                For X As Int32 = 1 To lstData.ListCount - 1
                    Dim lIdx As Int32 = lstData.ItemData(X)
                    With goCurrentEnvir.oEntity(lIdx)
                        If .EntityName = "Unknown" OrElse .EntityName <> sName Then
                            bPass = False
                            Exit For
                        End If
                    End With
                Next

                ' update txtRename to show then common name, else nothing
                If bPass = False Then sName = ""
                If txtRename.Caption <> sName Then txtRename.Caption = sName

                mbDidNameCheck = True
            End If
        End If

        If goCurrentPlayer Is Nothing = False Then
            Dim lCnt As Int32 = 0
            For X As Int32 = 0 To goCurrentPlayer.lFormationUB
                If goCurrentPlayer.lFormationIdx(X) <> -1 Then
                    lCnt += 1
                End If
            Next X
            If lCnt + 1 <> cboFormation.ListCount Then
                FillFormations()
            End If
        End If

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            Try
                If oEnvir.lEntityIdx Is Nothing = False Then
                    Dim lCurUB As Int32 = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))

                    Dim clrBlack As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 32, 255)
                    Dim clrRed As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    Dim clrOrange As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 128, 64)
                    Dim clrYellow As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 0)
                    Dim clrGreen As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                    Dim clrUnknown As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)
                    Dim clrVal As System.Drawing.Color = clrGreen
                    For X As Int32 = 0 To lstData.ListCount - 1
                        Dim lIdx As Int32 = lstData.ItemData(X)

                        Dim bGood As Boolean = False
                        If lIdx > -1 AndAlso lIdx <= oEnvir.lEntityUB Then
                            If oEnvir.lEntityIdx(lIdx) <> -1 Then
                                Dim oEntity As BaseEntity = oEnvir.oEntity(lIdx)
                                If oEntity Is Nothing = False Then
                                    If oEntity.bSelected = True Then
                                        bGood = True
                                        If oEntity.bRequestedDetails = False Then
                                            Dim yMsg(13) As Byte
                                            System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yMsg, 0)
                                            oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
                                            goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, 8)
                                            MyBase.moUILib.SendMsgToRegion(yMsg)
                                            oEntity.bRequestedDetails = True
                                        End If
                                        If oEntity.lLastHPUpdate <= 0 Then
                                            clrVal = clrUnknown
                                            If oEntity.lLastHPUpdate = Int32.MinValue Then
                                                oEntity.lLastHPUpdate = Int32.MinValue + 1
                                                Dim yData(7) As Byte
                                                System.BitConverter.GetBytes(GlobalMessageCode.eUnitHPUpdate).CopyTo(yData, 0)
                                                System.BitConverter.GetBytes(oEntity.ObjectID).CopyTo(yData, 2)
                                                System.BitConverter.GetBytes(oEntity.ObjTypeID).CopyTo(yData, 6)
                                                MyBase.moUILib.SendMsgToRegion(yData)
                                            End If
                                        Else
                                            'Ok, let's get its status...
                                            '  Green is undamaged
                                            clrVal = clrGreen
                                            If (oEntity.CritList And (elUnitStatus.eEngineOperational Or elUnitStatus.eRadarOperational)) <> 0 Then
                                                '  Black is combat in-capable (if the unit was capable of combat that is)
                                                clrVal = clrBlack
                                            ElseIf oEntity.yStructureHP < 75 Then
                                                '  Red is 75% or less structure remains
                                                clrVal = clrRed
                                            Else
                                                For Z As Int32 = 0 To oEntity.yArmorHP.GetUpperBound(0)
                                                    If oEntity.yArmorHP(Z) = 0 Then
                                                        '  Orange is at least one side at 0
                                                        clrVal = clrOrange
                                                        Exit For
                                                    ElseIf oEntity.yArmorHP(Z) <> 100 AndAlso oEntity.yArmorHP(Z) <> 255 Then
                                                        '  yellow is damaged
                                                        clrVal = clrYellow
                                                    End If
                                                Next Z
                                            End If
                                        End If
                                        If lstData.ItemCustomColor(X) <> clrVal Then
                                            lstData.ItemCustomColor(X) = clrVal
                                        End If

                                        If muSettings.ShowMultiHealthBars = True AndAlso oEntity.yStructureHP < 100 Then
                                            lstData.ItemCustomBackColor(X) = Color.White
                                            lstData.ItemCustomBackPercent(X) = oEntity.yStructureHP
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If bGood = False Then
                            lstData.RemoveItem(X)
                            Exit For
                        End If

                    Next X

                    Dim sText As String = "Selected: " & lstData.ListCount
                    If lblTotalSelected.Caption <> sText Then lblTotalSelected.Caption = sText
                End If
            Catch
                'do nothing, we'll get it next time
            End Try
        End If

        If lstData.ListCount < 2 Then
            Dim lTempIdx As Int32 = -1
            If lstData.ListCount > 0 Then
                Try
                    lTempIdx = lstData.ItemData(0)
                Catch
                End Try
            End If
            MyBase.moUILib.ClearSelection()
            If lTempIdx <> -1 Then MyBase.moUILib.AddSelection(lTempIdx)
        End If
    End Sub

	Private Sub FillFormations()
		Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.oFormations, goCurrentPlayer.lFormationIdx, GetSortedIndexArrayType.eFormationName)
		If lSorted Is Nothing = False Then
			cboFormation.Clear()
			cboFormation.AddItem("No Formation")
			cboFormation.ItemData(cboFormation.NewIndex) = -1
			For X As Int32 = 0 To lSorted.GetUpperBound(0)
				cboFormation.AddItem(goCurrentPlayer.oFormations(lSorted(X)).sName)
				cboFormation.ItemData(cboFormation.NewIndex) = goCurrentPlayer.oFormations(lSorted(X)).FormationID
			Next X
		End If
	End Sub

    Private Sub lstData_ItemClick(ByVal lIndex As Integer) Handles lstData.ItemClick
        Dim lIdx As Int32 = lstData.ItemData(lIndex)
        If frmMain.mbShiftKeyDown = True Then
            If goCurrentEnvir Is Nothing = False Then
                If goCurrentEnvir.lEntityIdx(lIdx) <> -1 Then
                    goCurrentEnvir.oEntity(lIdx).bSelected = False
                    goUILib.RemoveSelection(lIdx)
                End If
            End If
        Else
            goUILib.ClearSelection()
            If goCurrentEnvir Is Nothing = False Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                        goCurrentEnvir.oEntity(X).bSelected = False
                    End If
                Next X
                If goCurrentEnvir.lEntityIdx(lIdx) <> -1 Then
                    goCurrentEnvir.oEntity(lIdx).bSelected = True
                    goUILib.AddSelection(lIdx)
                End If
            End If
        End If

        If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
        lstData.HasFocus = False
        Me.HasFocus = False
        goUILib.FocusedControl = Nothing
    End Sub

    Private Sub frmMultiDisplay_OnRenderEnd() Handles Me.OnRenderEnd
        If cboFormation.IsExpanded = True Then Return
        Using oSprite As New Sprite(MyBase.moUILib.oDevice)
            oSprite.Begin(SpriteFlags.AlphaBlend)
            Try
                If cboFormation.Enabled = True Then
                    If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                        oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.ePlanetOrbit), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcSendToOrbit.Location, System.Drawing.Color.White)
                    End If
                    Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    If muSettings.bFormationMoveThenForm = True Then
                        clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                    End If
                    oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcMoveThenForm.Location, clrVal)
                ElseIf mbPowerDisplay = True Then
                    'Render Power On / Off buttons
                    oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eSphere), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcPowerOn.Location, System.Drawing.Color.FromArgb(255, 0, 255, 0))
                    oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eSphere), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcPowerOff.Location, System.Drawing.Color.FromArgb(255, 255, 0, 0))
                End If

            Catch
            End Try

            oSprite.End()
        End Using
    End Sub

    Private Sub frmMultiDisplay_OnResize() Handles Me.OnResize
        If mbLoading = True Then Return

        muSettings.MultiSelectHeight = Me.Height
        muSettings.MultiSelectWidth = Me.Width

        If lstData Is Nothing = False AndAlso lblFormation Is Nothing = False AndAlso cboFormation Is Nothing = False AndAlso lblTotalSelected Is Nothing = False AndAlso txtRename Is Nothing = False Then
            cboFormation.Width = Me.Width - 72
            cboFormation.l_ListBoxHeight = Me.Height - cboFormation.Top - cboFormation.Height - 2
            lstData.Width = Me.Width - 4
            lstData.Height = Me.Height - lstData.Top
            lblTotalSelected.Width = Me.Width - 6
            txtRename.Width = Me.Width - 50

            'rcSendToOrbit = New Rectangle(Me.Width - 16 - 4, lblTotalSelected.Top, 16, 16)
            'rcMoveThenForm = New Rectangle(Me.Width - 16 - 16 - 4, lblTotalSelected.Top, 16, 16)
            'btnUnDamaged.Left = rcMoveThenForm.Left - btnUnDamaged.Width - 1
            'btnDamaged.Left = btnUnDamaged.Left - btnUnDamaged.Width - 1
            'btnUnitGoto.Left = btnDamaged.Left - btnUnitGoto.Width - 1
            'rcPowerOn = New Rectangle(Me.Width - 16 - 16 - 4, lblTotalSelected.Top + 2, 16, 16)
            'rcPowerOff = New Rectangle(Me.Width - 16 - 4, lblTotalSelected.Top + 2, 16, 16)

            rcSendToOrbit = New Rectangle(Me.Width - 16 - 4, lblTotalSelected.Top, 16, 16)
            rcMoveThenForm = New Rectangle(Me.Width - 16 - 16 - 4, lblTotalSelected.Top, 16, 16)
            btnFilterSel.Left = rcMoveThenForm.Left - btnFilterSel.Width - 1
            btnUnitGoto.Left = btnFilterSel.Left - btnUnitGoto.Width - 1
            rcPowerOn = New Rectangle(Me.Width - 16 - 16 - 4, lblTotalSelected.Top + 1, 16, 16)
            rcPowerOff = New Rectangle(Me.Width - 16 - 4, lblTotalSelected.Top + 1, 16, 16)

            Dim oFilter As UIWindow = MyBase.moUILib.GetWindow("frmUnitSelectFilter")
            If oFilter Is Nothing = False Then
                oFilter.Top = Me.Top
                oFilter.Left = Me.Left + Me.Width
            End If

            Dim oWin As UIWindow = MyBase.moUILib.GetWindow("frmBehavior")
            If oWin Is Nothing = False Then
                Dim rc As Rectangle = New Rectangle(Me.Left, Me.Top, Me.Width + 50, Me.Height)
                If rc.Contains(oWin.Left, oWin.Top) = True Then
                    oWin.Left = Me.Left + Me.Width
                End If
            End If
        End If
    End Sub

    Private Sub frmMultiDisplay_WindowMoved() Handles Me.WindowMoved
        muSettings.MultiSelectLocX = Me.Left
        muSettings.MultiSelectLocY = Me.Top

        Dim oFilter As UIWindow = MyBase.moUILib.GetWindow("frmUnitSelectFilter")
        If oFilter Is Nothing = False Then
            oFilter.Top = Me.Top
            oFilter.Left = Me.Left + Me.Width
        End If
	End Sub

	Public Function GetFormationID() As Int32
		If cboFormation Is Nothing = False AndAlso cboFormation.ListIndex > -1 Then
			Return cboFormation.ItemData(cboFormation.ListIndex)
		End If
		Return -1
	End Function

	Private Sub cboFormation_ItemSelected(ByVal lItemIndex As Integer) Handles cboFormation.ItemSelected
		'ok, remove focus
		If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
		cboFormation.HasFocus = False
		Me.HasFocus = False
		goUILib.FocusedControl = Nothing
    End Sub

    Public Sub SetFormationID(ByVal lID As Int32)
        If lID > 0 Then
            If cboFormation.FindComboItemData(lID) = True Then
                cboFormation_ItemSelected(cboFormation.ListIndex)
            End If
        Else : cboFormation.ListIndex = -1
        End If
    End Sub

    'Private Sub btnDamaged_Click(ByVal sName As String) Handles btnDamaged.Click
    '    Dim clrGreen As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
    '    For X As Int32 = 0 To lstData.ListCount - 1
    '        If lstData.ItemCustomColor(X) <> clrGreen Then
    '            Dim lIdx As Int32 = lstData.ItemData(X)
    '            If goCurrentEnvir Is Nothing = False Then
    '                If goCurrentEnvir.lEntityIdx(lIdx) <> -1 Then
    '                    goCurrentEnvir.oEntity(lIdx).bSelected = False
    '                    goUILib.RemoveSelection(lIdx)
    '                End If
    '            End If
    '        End If
    '    Next X
    'End Sub

    'Private Sub btnUnDamaged_Click(ByVal sName As String) Handles btnUnDamaged.Click
    '    Dim clrGreen As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
    '    For X As Int32 = 0 To lstData.ListCount - 1
    '        If lstData.ItemCustomColor(X) = clrGreen Then
    '            Dim lIdx As Int32 = lstData.ItemData(X)
    '            If goCurrentEnvir Is Nothing = False Then
    '                If goCurrentEnvir.lEntityIdx(lIdx) <> -1 Then
    '                    goCurrentEnvir.oEntity(lIdx).bSelected = False
    '                    goUILib.RemoveSelection(lIdx)
    '                End If
    '            End If
    '        End If
    '    Next X
    'End Sub

    Public Sub HandleSelectFilter(ByVal bUndamaged As Boolean, ByVal bDamaged As Boolean, ByVal bCritical As Boolean, ByVal bDisabled As Boolean, ByVal bUnits As Boolean, ByVal bFacilities As Boolean, ByVal bSelect As Boolean)
        'Disabled is when any component is offline
        'Critical when any part is lost, or structure is < 25%
        'Damaged is armor < 100%, or structure < 100%

        'Final check weeds out units - facilities.

        If goCurrentEnvir Is Nothing = True Then Return
        For X As Int32 = 0 To lstData.ListCount - 1
            Dim lIdx As Int32 = lstData.ItemData(X)
            If goCurrentEnvir.lEntityIdx(lIdx) <> -1 Then
                Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(lIdx)
                Dim bMatch As Boolean = False
                Dim bIsCritical As Boolean = False
                If bDisabled = False AndAlso bCritical = False AndAlso bDamaged = False AndAlso bUndamaged = False Then
                    bMatch = True 'Set to true.  will reset later if it doesnt match unit-fac filter
                Else
                    If ((oEntity.CritList And elUnitStatus.eEngineOperational) <> 0) OrElse _
                        ((oEntity.CritList And elUnitStatus.eRadarOperational) <> 0) OrElse _
                        ((oEntity.CritList And elUnitStatus.eForwardWeapon1) <> 0) OrElse _
                        ((oEntity.CritList And elUnitStatus.eForwardWeapon2) <> 0) OrElse _
                        ((oEntity.CritList And elUnitStatus.eLeftWeapon1) <> 0) OrElse _
                        ((oEntity.CritList And elUnitStatus.eLeftWeapon2) <> 0) OrElse _
                        ((oEntity.CritList And elUnitStatus.eAftWeapon1) <> 0) OrElse _
                        ((oEntity.CritList And elUnitStatus.eAftWeapon2) <> 0) OrElse _
                        ((oEntity.CritList And elUnitStatus.eRightWeapon1) <> 0) OrElse _
                        ((oEntity.CritList And elUnitStatus.eRightWeapon2) <> 0) Then
                        bIsCritical = True
                    End If
                    If bDisabled = True AndAlso bIsCritical = True Then
                        bMatch = True
                    Else
                        Dim yArmorHP As Byte = 255
                        For Z As Int32 = 0 To oEntity.yArmorHP.GetUpperBound(0)
                            If oEntity.yArmorHP(Z) < yArmorHP Then yArmorHP = oEntity.yArmorHP(Z)
                        Next Z
                        If bUndamaged = True AndAlso yArmorHP >= 100 AndAlso oEntity.yStructureHP = 100 AndAlso bIsCritical = False Then
                            bMatch = True
                        ElseIf bCritical = True AndAlso (bIsCritical = True OrElse oEntity.yStructureHP < 75) Then
                            bMatch = True
                        ElseIf bDamaged = True AndAlso (yArmorHP < 100 OrElse oEntity.yStructureHP < 100 OrElse bIsCritical = True) Then
                            bMatch = True
                        End If
                    End If
                End If

                'Secondary pass for Unit Type
                If (bFacilities = True AndAlso oEntity.ObjTypeID <> ObjectType.eFacility) OrElse (bUnits = True AndAlso oEntity.ObjTypeID <> ObjectType.eUnit) Then
                    bMatch = False
                End If

                If bMatch <> bSelect Then
                    goCurrentEnvir.oEntity(lIdx).bSelected = False
                    goUILib.RemoveSelection(lIdx)
                End If
            End If
        Next

    End Sub

#Region "  Formation Memory  "
    Private Structure uMultiSelectLookup
        Public lFormationID As Int32
        Public bMoveThenForm As Boolean

        Public lIDList() As Int32
        Public lHitMatch As Int32
    End Structure
    Private Shared muMultiSelectLookup() As uMultiSelectLookup
    Private Shared mlMultiSelectLookupUB As Int32 = -1
    Private Shared mlLastStoreLookup As Int32 = -1

    Private Sub CheckMultiSelectLookup()
        Try
            For X As Int32 = 0 To mlMultiSelectLookupUB
                muMultiSelectLookup(X).lHitMatch = 0
            Next X

            If lstData Is Nothing = False AndAlso goCurrentEnvir Is Nothing = False Then
                For X As Int32 = 0 To lstData.ListCount - 1
                    Dim lID As Int32 = lstData.ItemData(X)
                    If lID > -1 AndAlso lID <= goCurrentEnvir.lEntityUB Then
                        lID = goCurrentEnvir.lEntityIdx(lID)

                        For Y As Int32 = 0 To mlMultiSelectLookupUB
                            With muMultiSelectLookup(Y)
                                If .lIDList Is Nothing = False Then
                                    For Z As Int32 = 0 To .lIDList.GetUpperBound(0)
                                        If .lIDList(Z) = lID Then
                                            .lHitMatch += 1
                                            Exit For
                                        End If
                                    Next Z
                                End If
                            End With
                        Next Y
                    End If
                Next X

                'Now, our hitmatch list is populated
                Dim fBestRatio As Single = 0.0F
                Dim lBestMatchIdx As Int32 = -1
                For X As Int32 = 0 To mlMultiSelectLookupUB
                    If muMultiSelectLookup(X).lHitMatch > 0 AndAlso muMultiSelectLookup(X).lIDList Is Nothing = False Then
                        Dim fVal As Single = CSng(muMultiSelectLookup(X).lHitMatch / muMultiSelectLookup(X).lIDList.Length)
                        If fVal > fBestRatio Then
                            fBestRatio = fVal
                            lBestMatchIdx = X
                        End If
                    End If
                Next X

                If lBestMatchIdx <> -1 Then
                    cboFormation.FindComboItemData(muMultiSelectLookup(lBestMatchIdx).lFormationID)
                    muSettings.bFormationMoveThenForm = muMultiSelectLookup(lBestMatchIdx).bMoveThenForm
                End If
            End If
        Catch
        End Try
    End Sub

    Public Sub StoreMultiSelectLookup()
        Try
            If goCurrentEnvir Is Nothing Then Return

            Dim bMoveThenForm As Boolean = muSettings.bFormationMoveThenForm
            Dim lFormationID As Int32 = -1
            If cboFormation Is Nothing = False AndAlso cboFormation.ListIndex > -1 Then lFormationID = cboFormation.ItemData(cboFormation.ListIndex)

            Dim lIDs(lstData.ListCount - 1) As Int32
            For X As Int32 = 0 To lstData.ListCount - 1
                lIDs(X) = lstData.ItemData(X)
                If lIDs(X) > -1 AndAlso lIDs(X) <= goCurrentEnvir.lEntityUB Then
                    lIDs(X) = goCurrentEnvir.lEntityIdx(lIDs(X))
                Else : lIDs(X) = -1
                End If
            Next X
            'Now, create our final id list
            Dim lIDList(-1) As Int32
            For X As Int32 = 0 To lIDs.GetUpperBound(0)
                If lIDs(X) > 0 Then
                    ReDim Preserve lIDList(lIDList.GetUpperBound(0) + 1)
                    lIDList(lIDList.GetUpperBound(0)) = lIDs(X)
                End If
            Next X
            If lIDList.GetUpperBound(0) = -1 Then Return

            If mlLastStoreLookup <> -1 Then
                With muMultiSelectLookup(mlLastStoreLookup)
                    If .lIDList Is Nothing = False AndAlso .lIDList.GetUpperBound(0) = lIDList.GetUpperBound(0) Then
                        Dim bGood As Boolean = True
                        For Y As Int32 = 0 To lIDList.GetUpperBound(0)
                            Dim bFound As Boolean = False
                            For Z As Int32 = 0 To .lIDList.GetUpperBound(0)
                                If .lIDList(Z) = lIDList(Y) Then
                                    bFound = True
                                    Exit For
                                End If
                            Next Z
                            If bFound = False Then
                                bGood = False
                                Exit For
                            End If
                        Next Y

                        If bGood = True Then
                            .lFormationID = lFormationID
                            .bMoveThenForm = bMoveThenForm
                            Return
                        End If
                    End If
                End With
            End If
            mlLastStoreLookup = -1

            'Now, see if any of our current lookups are a 100% match
            Dim bAlreadyThere As Boolean = False
            For X As Int32 = 0 To mlMultiSelectLookupUB
                If muMultiSelectLookup(X).lIDList Is Nothing = False AndAlso muMultiSelectLookup(X).lIDList.GetUpperBound(0) = lIDList.GetUpperBound(0) Then
                    Dim bGood As Boolean = True
                    For Y As Int32 = 0 To lIDList.GetUpperBound(0)
                        Dim bFound As Boolean = False
                        For Z As Int32 = 0 To muMultiSelectLookup(X).lIDList.GetUpperBound(0)
                            If muMultiSelectLookup(X).lIDList(Z) = lIDList(Y) Then
                                bFound = True
                                Exit For
                            End If
                        Next Z
                        If bFound = False Then
                            bGood = False
                            Exit For
                        End If
                    Next Y

                    If bGood = True Then
                        muMultiSelectLookup(X).lFormationID = lFormationID
                        muMultiSelectLookup(X).bMoveThenForm = bMoveThenForm
                        mlLastStoreLookup = X
                        bAlreadyThere = True
                        Exit For
                    End If
                End If
            Next X
            If bAlreadyThere = False Then
                Dim uTemp As uMultiSelectLookup
                With uTemp
                    .bMoveThenForm = bMoveThenForm
                    .lFormationID = lFormationID
                    .lIDList = lIDList
                End With
                ReDim Preserve muMultiSelectLookup(mlMultiSelectLookupUB + 1)
                muMultiSelectLookup(mlMultiSelectLookupUB + 1) = uTemp
                mlMultiSelectLookupUB += 1
                mlLastStoreLookup = mlMultiSelectLookupUB
            End If

        Catch
        End Try
    End Sub
#End Region

    Private Sub txtRename_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRename.OnKeyDown
        If e.KeyCode = Keys.Enter AndAlso goCurrentEnvir Is Nothing = False AndAlso lstData.ListCount > 0 Then
            Dim sName As String = txtRename.Caption
            Dim oFrm As New frmMsgBox(goUILib, "This will rename the currently selected items to """ & sName & """. Are you sure you wish to do so?", MsgBoxStyle.YesNo, "Confirm Rename")
            oFrm.Visible = True
            AddHandler oFrm.DialogClosed, AddressOf RenameSelectedEntities
            e.Handled = True
        End If
    End Sub

    Private Sub RenameSelectedEntities(ByVal lResult As Microsoft.VisualBasic.MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            Try
                Dim lCnt As Int32 = 0
                Dim lPos As Int32 = 0
                Dim iLen As Int16 = 28
                Dim sName As String = txtRename.Caption

                If goCurrentEnvir Is Nothing = False Then
                    For X As Int32 = 0 To lstData.ListCount - 1
                        Dim lIdx As Int32 = lstData.ItemData(X)
                        If goCurrentEnvir.lEntityIdx(lIdx) > -1 Then
                            Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(lIdx)
                            If oEntity Is Nothing = False Then
                                If oEntity.ObjectID <> -1 AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.bSelected = True Then
                                    lCnt += 1
                                End If
                            End If
                        End If
                    Next X
                    If lCnt = 0 Then Exit Sub
                    Dim yMsg(((iLen + 2) * lCnt) - 1) As Byte

                    For X As Int32 = 0 To lstData.ListCount - 1
                        Dim lIdx As Int32 = lstData.ItemData(X)
                        If goCurrentEnvir.lEntityIdx(lIdx) > -1 Then
                            Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(lIdx)
                            If oEntity Is Nothing = False Then
                                If oEntity.ObjectID <> -1 AndAlso oEntity.OwnerID = glPlayerID AndAlso oEntity.bSelected = True Then
                                    'Create Packet 
                                    System.BitConverter.GetBytes(iLen).CopyTo(yMsg, lPos) : lPos += 2
                                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityName).CopyTo(yMsg, lPos) : lPos += 2
                                    oEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                                    System.Text.ASCIIEncoding.ASCII.GetBytes(Mid$(sName, 1, 20)).CopyTo(yMsg, lPos) : lPos += 20

                                    'Update local client 
                                    SetCacheObjectValue(oEntity.ObjectID, oEntity.ObjTypeID, sName)
                                    oEntity.EntityName = sName
                                    lstData.List(X) = sName
                                End If
                            End If
                        End If
                    Next

                    'If yMsg Is Nothing = False Then MyBase.moUILib.SendLenAppendedMsgToRegion(yMsg)
                    If yMsg Is Nothing = False Then MyBase.moUILib.SendLenAppendedMsgToPrimary(yMsg)
                End If

            Catch
            End Try
            If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
            txtRename.HasFocus = False
            Me.HasFocus = False
            goUILib.FocusedControl = Nothing
        End If
    End Sub

    Private Sub btnUnitGoto_Click(ByVal sName As String) Handles btnUnitGoto.Click
        HandleUnitGotoClick()
    End Sub

    Public Function HandleUnitGotoClick() As Boolean
        If goCurrentEnvir Is Nothing Then Return False
        If HasAliasedRights(AliasingRights.eMoveUnits) = False Then
            goUILib.AddNotification("You lack rights to change unit environments.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If


        Dim ofrmUnitGoto As frmUnitGoto = CType(goUILib.GetWindow("frmUnitGoto"), frmUnitGoto)
        If ofrmUnitGoto Is Nothing Then
            ofrmUnitGoto = New frmUnitGoto(goUILib)
        ElseIf ofrmUnitGoto.Visible = False Then
        Else
            goUILib.RemoveWindow("frmUnitGoto")
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            Return False
        End If

        ofrmUnitGoto.Visible = True
        ofrmUnitGoto.FillList()
        ofrmUnitGoto = Nothing
        Return True

    End Function

    Private Sub btnFilterSel_Click(ByVal sName As String) Handles btnFilterSel.Click
        HandleFilterClick()
    End Sub

    Public Sub HandleFilterClick()
        Dim ofrm As frmUnitSelectFilter = CType(goUILib.GetWindow("frmUnitSelectFilter"), frmUnitSelectFilter)
        If ofrm Is Nothing = True Then
            ofrm = New frmUnitSelectFilter(goUILib)
            ofrm.Visible = True
        Else
            goUILib.RemoveWindow("frmUnitSelectFilter")
        End If
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
        ofrm = Nothing
    End Sub

    Private Sub lstData_ItemMouseOver(ByVal lIndex As Integer, ByVal lX As Integer, ByVal lY As Integer) Handles lstData.ItemMouseOver
        If lIndex > -1 Then
            Dim lEntityIndex As Int32 = lstData.ItemData(lIndex)
            If goCurrentEnvir.oEntity(lEntityIndex) Is Nothing Then Exit Sub
            Dim sTemp As String = ""
            If goCurrentEnvir.oEntity(lEntityIndex).ObjTypeID <> ObjectType.eFacility Then
                Dim lXPRank As Int32 = Math.Abs((CInt(goCurrentEnvir.oEntity(lEntityIndex).Exp_Level) - 1) \ 25)
                sTemp = BaseEntity.GetExperienceLevelNameAndBenefits(goCurrentEnvir.oEntity(lEntityIndex).Exp_Level)
                sTemp &= vbCrLf & "CP Usage: " & Math.Max((10 - lXPRank), 1)
                'Dim lWarpointUpkeep As Int32 = GetUnitWarpointUpkeep(goCurrentEnvir.oEntity(lEntityIndex).ObjectID)
                'If lWarpointUpkeep > -1 Then sTemp &= vbCrLf & "WP Upkeep Cost: " & lWarpointUpkeep.ToString("#,##0")
            End If
            If sTemp <> "" Then MyBase.moUILib.SetToolTip(sTemp, lX, lY)
        End If
    End Sub

End Class
