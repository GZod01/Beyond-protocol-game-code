Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

''Defense Status form (hp display)
'Public Class frmDefenseDisplay_old
'    Inherits UIWindow

'    Private lblOwner As UILabel
'    Private lblShieldHP As UILabel
'    Private lblForeArmorHP As UILabel
'    Private lblLeftArmorHP As UILabel
'    Private lblRearArmorHP As UILabel
'    Private lblRightArmorHP As UILabel
'    Private lblStructureHP As UILabel
'    Private lblShieldPerc As UILabel
'    Private lblForeArmorPerc As UILabel
'    Private lblLeftArmorPerc As UILabel
'    Private lblRearArmorPerc As UILabel
'    Private lblRightArmorPerc As UILabel
'    Private lblStructurePerc As UILabel
'    Private vln2 As UILine

'    Private WithEvents btnBuild As UIButton

'    Private moLowColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
'    Private moMidColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 0)
'    Private moHiColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)

'    Private mlOwnerID As Int32

'    Private mlEntityIndex As Int32 = -1

'    Private mbBuildShowing As Boolean = False
'    Private mbRefreshBuildScreen As Boolean

'    Private mlLastHPUpdate As Int32 = 0

'    Private Function getColorFromPerc(ByVal yPerc As Byte) As System.Drawing.Color
'        If yPerc < 30 Then
'            Return moLowColor
'        ElseIf yPerc < 70 Then
'            Return moMidColor
'        Else : Return moHiColor
'        End If
'    End Function

'    Public Sub New(ByRef oUILib As UILib)
'        MyBase.New(oUILib)

'        'frmDefenseDisplay initial props
'        With Me
'            .ControlName = "frmDefenseDisplay"
'            '.Left = muSettings.MiniMapLocX + muSettings.MiniMapWidthHeight + 190
'            .Left = 190
'            '.Top = (muSettings.MiniMapLocY + muSettings.MiniMapWidthHeight) - 86
'            .Top = moUILib.oDevice.PresentationParameters.BackBufferHeight - 86
'            .Width = 126    '108
'            .Height = 86
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'            '.FillColor = System.Drawing.Color.FromArgb(-16768960)
'            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = False
'            .BorderLineWidth = 2
'        End With
'        Debug.Write(Me.ControlName & " Newed" & vbCrLf)
'        'lblOwner initial props
'        lblOwner = New UILabel(oUILib)
'        With lblOwner
'            .ControlName = "lblOwner"
'            .Left = 5
'            .Top = 2
'            .Width = 103
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "Unknown"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = 0
'        End With
'        Me.AddChild(CType(lblOwner, UIControl))

'        'lblShieldHP initial props
'        lblShieldHP = New UILabel(oUILib)
'        With lblShieldHP
'            .ControlName = "lblShieldHP"
'            .Left = 4
'            .Top = 24
'            .Width = 45
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "SHIELDS:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblShieldHP, UIControl))

'        'lblForeArmorHP initial props
'        lblForeArmorHP = New UILabel(oUILib)
'        With lblForeArmorHP
'            .ControlName = "lblForeArmorHP"
'            .Left = 4
'            .Top = 34
'            .Width = 67
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "FORE ARMOR:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblForeArmorHP, UIControl))

'        'lblLeftArmorHP initial props
'        lblLeftArmorHP = New UILabel(oUILib)
'        With lblLeftArmorHP
'            .ControlName = "lblLeftArmorHP"
'            .Left = 4
'            .Top = 44
'            .Width = 66
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "LEFT ARMOR:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblLeftArmorHP, UIControl))

'        'lblRearArmorHP initial props
'        lblRearArmorHP = New UILabel(oUILib)
'        With lblRearArmorHP
'            .ControlName = "lblRearArmorHP"
'            .Left = 4
'            .Top = 54
'            .Width = 67
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "REAR ARMOR:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRearArmorHP, UIControl))

'        'lblRightArmorHP initial props
'        lblRightArmorHP = New UILabel(oUILib)
'        With lblRightArmorHP
'            .ControlName = "lblRightArmorHP"
'            .Left = 4
'            .Top = 64
'            .Width = 72
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "RIGHT ARMOR:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRightArmorHP, UIControl))

'        'lblStructureHP initial props
'        lblStructureHP = New UILabel(oUILib)
'        With lblStructureHP
'            .ControlName = "lblStructureHP"
'            .Left = 4
'            .Top = 74
'            .Width = 62
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "STRUCTURE:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblStructureHP, UIControl))

'        'lblShieldPerc initial props
'        lblShieldPerc = New UILabel(oUILib)
'        With lblShieldPerc
'            .ControlName = "lblShieldPerc"
'            .Left = 49
'            .Top = 24
'            .Width = 24
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "100%"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblShieldPerc, UIControl))

'        'lblForeArmorPerc initial props
'        lblForeArmorPerc = New UILabel(oUILib)
'        With lblForeArmorPerc
'            .ControlName = "lblForeArmorPerc"
'            .Left = 73
'            .Top = 34
'            .Width = 24
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "100%"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblForeArmorPerc, UIControl))

'        'lblLeftArmorPerc initial props
'        lblLeftArmorPerc = New UILabel(oUILib)
'        With lblLeftArmorPerc
'            .ControlName = "lblLeftArmorPerc"
'            .Left = 70
'            .Top = 44
'            .Width = 24
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "100%"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblLeftArmorPerc, UIControl))

'        'lblRearArmorPerc initial props
'        lblRearArmorPerc = New UILabel(oUILib)
'        With lblRearArmorPerc
'            .ControlName = "lblRearArmorPerc"
'            .Left = 74
'            .Top = 54
'            .Width = 24
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "100%"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRearArmorPerc, UIControl))

'        'lblRightArmorPerc initial props
'        lblRightArmorPerc = New UILabel(oUILib)
'        With lblRightArmorPerc
'            .ControlName = "lblRightArmorPerc"
'            .Left = 77
'            .Top = 64
'            .Width = 24
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "100%"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblRightArmorPerc, UIControl))

'        'lblStructurePerc initial props
'        lblStructurePerc = New UILabel(oUILib)
'        With lblStructurePerc
'            .ControlName = "lblStructurePerc"
'            .Left = 68
'            .Top = 74
'            .Width = 24
'            .Height = 10
'            .Enabled = True
'            .Visible = True
'            .Caption = "100%"
'            .ForeColor = System.Drawing.Color.FromArgb(-16711936)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblStructurePerc, UIControl))

'        'vln2 initial props
'        vln2 = New UILine(oUILib)
'        With vln2
'            .ControlName = "vln2"
'            .Left = 1
'            .Top = 19
'            .Width = Me.Width - 1
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(vln2, UIControl))

'        'btnBuild initial props
'        btnBuild = New UIButton(oUILib)
'        With btnBuild
'            .ControlImageRect_Normal = Rectangle.FromLTRB(72, 144, 90, 215)
'            .ControlImageRect_Pressed = Rectangle.FromLTRB(72, 144, 90, 215)
'            .ControlImageRect_Disabled = Rectangle.FromLTRB(72, 144, 90, 215)
'            .ControlName = "btnBuild"
'            .Left = 107
'            .Top = 19
'            .Width = 18
'            .Height = 67
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = System.Drawing.Color.FromArgb(-1)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(btnBuild, UIControl))
'        mbBuildShowing = False
'    End Sub

'    Public Sub SetFromEntity(ByVal lEntityIndex As Int32)
'        'Verify our objects
'        If goCurrentEnvir Is Nothing Then Exit Sub
'        If lEntityIndex < 0 OrElse lEntityIndex > goCurrentEnvir.lEntityUB Then Exit Sub
'        If goCurrentEnvir.oEntity(lEntityIndex) Is Nothing Then Exit Sub
'        If goCurrentEnvir.oEntity(lEntityIndex).oUnitDef Is Nothing Then Exit Sub

'        mlEntityIndex = lEntityIndex
'        mbRefreshBuildScreen = True
'        lblOwner.Caption = "Unknown"

'        Dim oTmpWin As frmBuildWindow
'        oTmpWin = CType(MyBase.moUILib.GetWindow("frmBuildWindow"), frmBuildWindow)
'        If oTmpWin Is Nothing = False AndAlso oTmpWin.Visible = True Then
'            mbBuildShowing = True
'        Else : mbBuildShowing = False
'        End If
'        oTmpWin = Nothing

'    End Sub

'    Private Sub UpdateDisplay()
'        Dim yRel As Byte

'        If goCurrentEnvir.oEntity(mlEntityIndex).lLastHPUpdate <> mlLastHPUpdate OrElse mbRefreshBuildScreen = True Then
'            mlLastHPUpdate = goCurrentEnvir.oEntity(mlEntityIndex).lLastHPUpdate

'            With goCurrentEnvir.oEntity(mlEntityIndex)
'                lblShieldPerc.Caption = .yShieldHP & "%"
'                lblShieldPerc.ForeColor = getColorFromPerc(.yShieldHP)
'                lblForeArmorPerc.Caption = .yArmorHP(UnitArcs.eForwardArc) & "%"
'                lblForeArmorPerc.ForeColor = getColorFromPerc(.yArmorHP(UnitArcs.eForwardArc))
'                lblLeftArmorPerc.Caption = .yArmorHP(UnitArcs.eLeftArc) & "%"
'                lblLeftArmorPerc.ForeColor = getColorFromPerc(.yArmorHP(UnitArcs.eLeftArc))
'                lblRearArmorPerc.Caption = .yArmorHP(UnitArcs.eBackArc) & "%"
'                lblRearArmorPerc.ForeColor = getColorFromPerc(.yArmorHP(UnitArcs.eBackArc))
'                lblRightArmorPerc.Caption = .yArmorHP(UnitArcs.eRightArc) & "%"
'                lblRightArmorPerc.ForeColor = getColorFromPerc(.yArmorHP(UnitArcs.eRightArc))
'                lblStructurePerc.Caption = .yStructureHP & "%"
'                lblStructurePerc.ForeColor = getColorFromPerc(.yStructureHP)

'                If .OwnerID <> glPlayerID Then

'                    yRel = .yRelID

'                    If yRel <= elRelTypes.eWar Then
'                        lblOwner.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'enemy
'                    ElseIf yRel <= elRelTypes.eNeutral Then
'                        lblOwner.ForeColor = System.Drawing.Color.FromArgb(255, 192, 192, 192)  'neutral
'                    Else
'                        lblOwner.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)    'ally
'                    End If
'                Else
'                    lblOwner.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)          'player's
'                End If

'                mlOwnerID = .OwnerID
'                If lblOwner.Caption = "Unknown" OrElse lblOwner.Caption = "" Then
'                    lblOwner.Caption = GetCacheObjectValue(.OwnerID, ObjectType.ePlayer)
'                End If
'            End With

'            'Reset our build window?
'            If mbRefreshBuildScreen = True OrElse (mlOwnerID <> glPlayerID) Then
'                mbBuildShowing = False

'                MyBase.moUILib.RemoveWindow("frmBuildWindow")
'                MyBase.moUILib.RemoveWindow("frmResearchMain")
'                MyBase.moUILib.RemoveWindow("frmSelectFac")

'                btnBuild.ControlImageRect_Normal = Rectangle.FromLTRB(72, 144, 90, 215)
'                btnBuild.ControlImageRect_Disabled = Rectangle.FromLTRB(72, 144, 90, 215)
'                btnBuild.ControlImageRect_Pressed = Rectangle.FromLTRB(72, 144, 90, 215)
'                mbRefreshBuildScreen = False
'            End If
'            btnBuild.Visible = (mlOwnerID = glPlayerID)
'        End If

'        If lblOwner.Caption = "Unknown" OrElse lblOwner.Caption = "" Then
'            lblOwner.Caption = GetCacheObjectValue(goCurrentEnvir.oEntity(mlEntityIndex).OwnerID, ObjectType.ePlayer)
'        End If

'    End Sub

'    Private Sub UpdateBuildButtonStatus()
'        Dim lX As Int32

'        If mbBuildShowing = True Then
'            lX = 90
'        Else : lX = 72
'        End If

'        btnBuild.ControlImageRect_Normal = Rectangle.FromLTRB(lX, 144, lX + 18, 215)
'        btnBuild.ControlImageRect_Disabled = Rectangle.FromLTRB(lX, 144, lX + 18, 215)
'        btnBuild.ControlImageRect_Pressed = Rectangle.FromLTRB(lX, 144, lX + 18, 215)
'    End Sub

'    Private Sub btnBuild_Click(ByVal sName As String) Handles btnBuild.Click
'        mbBuildShowing = Not mbBuildShowing

'        Dim oTmp As frmBuildWindow
'        Dim bUnitBuild As Boolean
'        Dim sFrmName As String

'        If mbBuildShowing = True Then
'            If mlEntityIndex <> -1 AndAlso goCurrentEnvir Is Nothing = False Then
'                If goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eResearch Then
'                    Dim ofrmResMain As frmResearchMain = New frmResearchMain(MyBase.moUILib)
'                    ofrmResMain = Nothing
'                Else
'                    'Units cannot have a build queue...
'                    bUnitBuild = (goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eUnit)
'                    oTmp = New frmBuildWindow(MyBase.moUILib, bUnitBuild)
'                    sFrmName = oTmp.ControlName
'                    If oTmp.UpdateFromEntity(mlEntityIndex) = False Then
'                        oTmp = Nothing
'                        MyBase.moUILib.RemoveWindow(sFrmName)
'                        mbBuildShowing = False
'                    End If
'                    oTmp = Nothing

'                    'Check if this is a spacestation special
'                    If goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eSpaceStationSpecial Then
'                        'Yes... we need to acquire the space station colony's Child list
'                        Dim yMsg(11) As Byte
'                        System.BitConverter.GetBytes(EpicaMessageCode.eGetColonyChildList).CopyTo(yMsg, 0)
'                        goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
'                        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 8)
'                        MyBase.moUILib.SendMsgToPrimary(yMsg)

'                        'Now... we should already be set to the space station's entity index... so nothing needs doing there

'                        'show frmSelectFac... it will fill with what we know so far... (or as of last update)
'                        Dim ofrmSelectFac As frmSelectFac = New frmSelectFac(goUILib)
'                        ofrmSelectFac.SetFromEntityIndex(mlEntityIndex)
'                        ofrmSelectFac.Visible = True
'                        ofrmSelectFac = Nothing
'                    End If
'                End If
'            Else
'                mbBuildShowing = False
'            End If
'        Else
'            'need to hide the build screen
'            MyBase.moUILib.RemoveWindow("frmBuildWindow")
'            MyBase.moUILib.RemoveWindow("frmResearchMain")
'            MyBase.moUILib.RemoveWindow("frmSelectFac")
'        End If

'        UpdateBuildButtonStatus()

'    End Sub

'    Protected Overrides Sub Finalize()
'        Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
'        MyBase.Finalize()
'    End Sub

'    Private Sub frmDefenseDisplay_OnNewFrame() Handles MyBase.OnNewFrame
'        If goCurrentEnvir Is Nothing = False AndAlso mlEntityIndex > -1 AndAlso goCurrentEnvir.lEntityIdx(mlEntityIndex) > -1 Then
'            UpdateDisplay()
'        End If
'    End Sub

'    Public Sub ResetBuildButton()
'        mbBuildShowing = False
'        UpdateBuildButtonStatus()
'    End Sub
'End Class

Public Class frmAdvanceDisplay
    Inherits UIWindow

    Private mbContentsDisplayed As Boolean = False
    Private mbBuildDisplayed As Boolean = False
    Private mbOrdersDisplayed As Boolean = False
    Private mbAutoLaunchDisplayed As Boolean = False

    Private mlEntityIndex As Int32 = -1 

    Private mlChildID As Int32 = -1
    Private miChildTypeID As Int16 = -1
    Private mbUseChild As Boolean = False

    Private mbCurrentActive As Boolean = False
    Private mbCurrentAutoLaunch As Boolean = False

    Private mbFirstTime As Boolean = True

    Private rcContents As Rectangle = New Rectangle(3, 3, 73, 21)
    Private rcBuild As Rectangle = New Rectangle(78, 3, 50, 21)
    Private rcOrders As Rectangle = New Rectangle(130, 3, 59, 21)

    Private rcActive As Rectangle = New Rectangle(233, 5, 16, 16)
    Private rcAutoLaunch As Rectangle = New Rectangle(215, 5, 16, 16)
    Private rcRepair As Rectangle = New Rectangle(195, 5, 16, 16)
    Private rcUnitGoto As Rectangle = New Rectangle(253, 5, 16, 16)

    Private myRepairButtonVisible As Byte = 0      '0 not visible, 1 = repair, 2 = dismantle

    Private Const ml_CRITS_WE_CARE_ABOUT As Int32 = elUnitStatus.eAftWeapon1 Or elUnitStatus.eAftWeapon2 Or elUnitStatus.eCargoBayOperational Or elUnitStatus.eEngineOperational Or elUnitStatus.eForwardWeapon1 Or elUnitStatus.eForwardWeapon2 Or elUnitStatus.eFuelBayOperational Or elUnitStatus.eHangarOperational Or elUnitStatus.eLeftWeapon1 Or elUnitStatus.eLeftWeapon2 Or elUnitStatus.eRadarOperational Or elUnitStatus.eRightWeapon1 Or elUnitStatus.eRightWeapon2 Or elUnitStatus.eShieldOperational

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmAdvancedDisplay initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAdvanceDisplay
            .ControlName = "frmAdvanceDisplay"
            .Width = 256
            .Height = 27

            If muSettings.AdvanceDisplayLocX = -1 OrElse NewTutorialManager.TutorialOn = True Then .Left = 128 Else .Left = muSettings.AdvanceDisplayLocX
            If muSettings.AdvanceDisplayLocY = -1 OrElse NewTutorialManager.TutorialOn = True Then .Top = moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height Else .Top = muSettings.AdvanceDisplayLocY

            If .Top + .Height > moUILib.oDevice.PresentationParameters.BackBufferHeight Then
                .Top = moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            End If
            If .Left + .Width > moUILib.oDevice.PresentationParameters.BackBufferWidth Then
                .Left = moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            End If

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        Me.IsDirty = True
        mbFirstTime = True
    End Sub

    Private Sub frmAdvanceDisplay_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown

        Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
		Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y

		Dim oEntity As BaseEntity = Nothing
		Dim oEnvir As BaseEnvironment = goCurrentEnvir
		If oEnvir Is Nothing Then Return

		If mlEntityIndex < 0 Then Return
		If oEnvir.lEntityUB < mlEntityIndex Then Return
		If oEnvir.lEntityIdx(mlEntityIndex) < 0 Then Return
		oEntity = oEnvir.oEntity(mlEntityIndex)
		If oEntity Is Nothing Then Return

        Dim bIgnoreMouseMove As Boolean = True
        'Now, check for what the player is hovering over
        If rcContents.Contains(lX, lY) = True Then

            If NewTutorialManager.TutorialOn = True Then
                Dim sParms() As String = {"Contents"}
                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
            End If

            'contents button
            OpenContentsWindow()
            'mbContentsDisplayed = Not mbContentsDisplayed
            'If mbContentsDisplayed = True Then

            '	If oEntity.ObjTypeID = ObjectType.eFacility AndAlso (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 Then
            '		MyBase.moUILib.AddNotification("Cannot open contents interface while facility is unpowered.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '		If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowPower.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
            '		mbContentsDisplayed = False
            '		Return
            '	End If

            '	Dim oContents As frmContents = CType(MyBase.moUILib.GetWindow("frmContents"), frmContents)
            '	If oContents Is Nothing Then oContents = New frmContents(MyBase.moUILib)
            '	oContents.Visible = True
            '	oContents.SetEntityRef(mlEntityIndex)
            'Else
            '	MyBase.moUILib.RemoveWindow("frmContents")
            'End If

            'If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
        ElseIf rcBuild.Contains(lX, lY) = True Then
            'Build button

            If NewTutorialManager.TutorialOn = True Then
                Dim sParms() As String = {"Build"}
                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
            End If

            mbBuildDisplayed = Not mbBuildDisplayed
            HandleBuildButtonClick()
            If mbBuildDisplayed = True AndAlso goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
        ElseIf rcOrders.Contains(lX, lY) = True Then
            'Orders button
            OpenOrdersWindows()
        ElseIf oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEntity.ObjTypeID = ObjectType.eUnit AndAlso oEntity.oMesh.bLandBased = False AndAlso rcActive.Contains(lX, lY) = True Then
            If NewTutorialManager.TutorialOn = True Then
                If goCurrentPlayer.yPlayerPhase = 0 Then Return
                Dim sParms() As String = {"SendToOrbit"}
                If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
            End If
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
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eGotoOrbitClick, 2, 0)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
        ElseIf oEntity.ObjTypeID = ObjectType.eUnit AndAlso oEntity.oMesh.bLandBased = False AndAlso rcUnitGoto.Contains(lX, lY) = True Then
            HandleUnitGotoClick()
        ElseIf oEntity.ObjTypeID = ObjectType.eFacility Then
            If rcActive.Contains(lX, lY) = True Then
                'Power button
                If NewTutorialManager.TutorialOn = True Then
                    If goCurrentPlayer.yPlayerPhase = 0 AndAlso mbCurrentActive = True AndAlso NewTutorialManager.GetTutorialStepID() <> 206 Then Return
                    Dim sParms() As String = {"Power"}
                    If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eAdvanceDisplayPowerClick, -1, -1, -1, "")
                End If
                If HandlePowerButtonClick() = True Then
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                End If
            ElseIf mbAutoLaunchDisplayed = True AndAlso rcAutoLaunch.Contains(lX, lY) = True Then
                'Auto-Launch button
                If NewTutorialManager.TutorialOn = True Then
                    Dim sParms() As String = {"AutoLaunch"}
                    If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                End If
                If HandleAutoLaunchButtonClick() = True Then
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                End If
            ElseIf myRepairButtonVisible <> 0 AndAlso rcRepair.Contains(lX, lY) = True Then
                If NewTutorialManager.TutorialOn = True Then
                    Dim sParms() As String = {"Repair"}
                    If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                End If
                HandleFacilityRepairClick()
            Else
                bIgnoreMouseMove = False
            End If
        ElseIf oEntity.yProductionType = ProductionType.eFacility Then
            If rcRepair.Contains(lX, lY) = True Then
                If lButton = MouseButtons.Left Then
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                    MyBase.moUILib.lUISelectState = UILib.eSelectState.eRepairTarget
                    MyBase.moUILib.AddNotification("Right-Click the target object to repair...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            ElseIf rcAutoLaunch.Contains(lX, lY) = True Then
                If NewTutorialManager.TutorialOn = True Then
                    Dim sParms() As String = {"RepairDismantle"}
                    If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
                End If
                If lButton = MouseButtons.Left Then
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                    MyBase.moUILib.lUISelectState = UILib.eSelectState.eDismantleTarget
                    MyBase.moUILib.AddNotification("Right-Click the target object to dismantle...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            Else
                bIgnoreMouseMove = False
            End If
        Else
            Me.IsDirty = True
            Exit Sub
        End If
        If bIgnoreMouseMove = True Then frmMain.mbIgnoreMouseMove = True
        Me.IsDirty = True
    End Sub

    Private Sub frmAdvanceDisplay_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
		MyBase.moUILib.SetToolTip(False)

		Try

			Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
			Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y

			'Now, check for what the player is hovering over
			If rcContents.Contains(lX, lY) = True Then
				'contents button
				MyBase.moUILib.SetToolTip("Click to show/hide the contents window.", lMouseX, lMouseY)
			ElseIf rcBuild.Contains(lX, lY) = True Then
				MyBase.moUILib.SetToolTip("Click to show/hide the Build/Research window.", lMouseX, lMouseY)
			ElseIf rcOrders.Contains(lX, lY) = True Then
                MyBase.moUILib.SetToolTip("Click to show/hide the Orders window.", lMouseX, lMouseY)
            ElseIf goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eUnit AndAlso rcActive.Contains(lX, lY) AndAlso goCurrentEnvir.oEntity(mlEntityIndex).oMesh.bLandBased = False Then
                MyBase.moUILib.SetToolTip("Click to send the unit into orbit.", lMouseX, lMouseY)
            ElseIf goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eUnit AndAlso rcUnitGoto.Contains(lX, lY) AndAlso goCurrentEnvir.oEntity(mlEntityIndex).oMesh.bLandBased = False Then
                MyBase.moUILib.SetToolTip("Click to send the unit to a desired location.", lMouseX, lMouseY)
            ElseIf goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eFacility Then
                If New Rectangle(Me.Width - 23, 5, 16, 16).Contains(lX, lY) = True Then
                    'Power button

                    If mbUseChild = True Then
                        With goCurrentEnvir.oEntity(mlEntityIndex)
                            Dim oChild As StationChild = Nothing
                            For X As Int32 = 0 To .lChildUB
                                If .lChildIdx(X) = mlChildID Then
                                    If .oChild(X).iChildTypeID = miChildTypeID Then
                                        oChild = .oChild(X)
                                        Exit For
                                    End If
                                End If
                            Next X

                            If oChild Is Nothing = False Then
                                With oChild
                                    Dim sTemp As String = ""
                                    If (.lChildStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                                        If .oChildDef Is Nothing = False Then
                                            If .oChildDef.PowerFactor < 0 Then
                                                sTemp = "This facility currently generates " & Math.Abs(.oChildDef.PowerFactor).ToString & " power for the colony." & vbCrLf
                                            Else : sTemp = "This facility currently consumes " & Math.Abs(.oChildDef.PowerFactor).ToString & " power from the colony." & vbCrLf
                                            End If
                                            If .oChildDef.WorkerFactor <> 0 Then
                                                sTemp &= .oChildDef.WorkerFactor.ToString & " jobs are being created by this facility." & vbCrLf
                                            End If
                                            If .oChildDef.ProductionTypeID = ProductionType.eColonists Then
                                                sTemp &= "This facility provides residential space for " & .oChildDef.ProdFactor.ToString & " citizens." & vbCrLf
                                            End If
                                        End If
                                    Else : sTemp = "This facility is currently unpowered." & vbCrLf
                                    End If
                                    sTemp &= "Click this to toggle the facility's power status."
                                    MyBase.moUILib.SetToolTip(sTemp, lMouseX, lMouseY)
                                End With
                            End If
                        End With
                    Else
                        With goCurrentEnvir.oEntity(mlEntityIndex)
                            Dim sTemp As String = ""
                            If (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then

                                If .PowerConsumption < 0 Then
                                    sTemp = "This facility currently generates " & Math.Abs(.PowerConsumption).ToString & " power for the colony." & vbCrLf
                                Else : sTemp = "This facility currently consumes " & Math.Abs(.PowerConsumption).ToString & " power from the colony." & vbCrLf
                                End If
                                If .MaxWorkers <> 0 Then
                                    sTemp &= .MaxWorkers.ToString & " jobs are being created by this facility." & vbCrLf
                                End If
                                If .MaxColonists <> 0 Then
                                    sTemp &= "This facility provides residential space for " & .MaxColonists & " citizens." & vbCrLf
                                End If
                            Else : sTemp = "This facility is currently unpowered." & vbCrLf
                            End If
                            sTemp &= "Click this to toggle the facility's power status."
                            MyBase.moUILib.SetToolTip(sTemp, lMouseX, lMouseY)
                        End With
                    End If

                ElseIf mbAutoLaunchDisplayed = True AndAlso New Rectangle(Me.Width - 41, 5, 16, 16).Contains(lX, lY) = True Then
                    If goCurrentEnvir.oEntity(mlEntityIndex).AutoLaunch = True Then
                        MyBase.moUILib.SetToolTip("Auto-launch is currently on." & vbCrLf & "Units will automatically undock when capable." & vbCrLf & "Click this to toggle Auto-Launch.", lMouseX, lMouseY)
                    Else
                        MyBase.moUILib.SetToolTip("Auto-launch is currently off." & vbCrLf & "Units will stay docked." & vbCrLf & "Click this to toggle Auto-Launch.", lMouseX, lMouseY)
                    End If
                ElseIf myRepairButtonVisible <> 0 AndAlso rcRepair.Contains(lX, lY) = True Then
                    MyBase.moUILib.SetToolTip("Click to open the Repair window for repairing this facility.", lMouseX, lMouseY)
                End If
            ElseIf goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eFacility Then
                If rcRepair.Contains(lX, lY) = True Then
                    MyBase.moUILib.SetToolTip("Left-Click to order the selection to" & vbCrLf & "repair a selected unit or facility.", lMouseX, lMouseY)
                ElseIf rcAutoLaunch.Contains(lX, lY) = True Then
                    MyBase.moUILib.SetToolTip("Left-Click to order the selection to" & vbCrLf & "dismantle a selected unit or facility.", lMouseX, lMouseY)
                End If
            End If
		Catch
		End Try
	End Sub

	Private Sub frmAdvanceDisplay_OnNewFrame() Handles Me.OnNewFrame

		Try
			Dim oEnvir As BaseEnvironment = goCurrentEnvir

			If oEnvir Is Nothing OrElse mlEntityIndex = -1 OrElse oEnvir.lEntityIdx(mlEntityIndex) = -1 OrElse oEnvir.oEntity(mlEntityIndex) Is Nothing OrElse oEnvir.oEntity(mlEntityIndex).bSelected = False Then
				MyBase.moUILib.RemoveSelection(mlEntityIndex)
				Me.Visible = False
				MyBase.moUILib.RemoveWindow(Me.ControlName)
				MyBase.moUILib.RemoveWindow("frmResearchMain")
				MyBase.moUILib.RemoveWindow("frmBuildWindow")
				Return
			End If
			Me.IsDirty = Me.IsDirty OrElse mbCurrentAutoLaunch <> oEnvir.oEntity(mlEntityIndex).AutoLaunch

			'Verify all of the data...
			If mbBuildDisplayed = True Then
				Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmBuildWindow")
                If ofrm Is Nothing Then ofrm = MyBase.moUILib.GetWindow("frmResearchMain")
                If ofrm Is Nothing Then ofrm = MyBase.moUILib.GetWindow("frmTradeMain")
                If ofrm Is Nothing Then ofrm = MyBase.moUILib.GetWindow("frmMining")
				If ofrm Is Nothing OrElse ofrm.Visible = False Then
					mbBuildDisplayed = False
					Me.IsDirty = True
                End If
			End If
			If mbContentsDisplayed = True Then
				Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmContents")
				If ofrm Is Nothing OrElse ofrm.Visible = False Then
					mbContentsDisplayed = False
					Me.IsDirty = True
				End If
			End If
			If mbOrdersDisplayed = True Then
				Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmBehavior")
				If ofrm Is Nothing OrElse ofrm.Visible = False Then
					mbOrdersDisplayed = False
					Me.IsDirty = True
				End If
			End If

			'Now, check the status
			Dim lStatus As Int32 = GetSelectionCurrentStatus()
            If (lStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                If mbCurrentActive = False Then Me.IsDirty = True
            ElseIf mbCurrentActive = True Then
                Me.IsDirty = True
            ElseIf oEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eFacility Then
                If mbBuildDisplayed = True Then
                    MyBase.moUILib.RemoveWindow("frmBuildWindow")
                    MyBase.moUILib.RemoveWindow("frmResearchMain")
                End If
                If mbContentsDisplayed = True Then MyBase.moUILib.RemoveWindow("frmContents")
                If mbOrdersDisplayed = True Then
                    MyBase.moUILib.RemoveWindow("frmBehavior")
                    MyBase.moUILib.RemoveWindow("frmRouteConfig")
                End If
            End If
		Catch
			'Do Nothing
		End Try

	End Sub

    Private Sub frmAdvanceDisplay_OnRenderEnd() Handles Me.OnRenderEnd
        If goCurrentEnvir Is Nothing OrElse mlEntityIndex = -1 OrElse mlEntityIndex > goCurrentEnvir.lEntityUB OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then
            MyBase.moUILib.RemoveWindow(Me.ControlName)
            Return
        End If
        Dim oEntity As BaseEntity = Nothing 'goCurrentEnvir.oEntity(mlEntityIndex)
        Try
            oEntity = goCurrentEnvir.oEntity(mlEntityIndex)
        Catch
        End Try
        If oEntity Is Nothing Then Return

        Dim yPrevRepairButtonState As Byte = myRepairButtonVisible
        myRepairButtonVisible = 0

        'Render our special state buttons... first, do any color fills from the button being selected
        mbCurrentActive = (GetSelectionCurrentStatus() And elUnitStatus.eFacilityPowered) <> 0
        mbCurrentAutoLaunch = oEntity.AutoLaunch

        Dim yProdType As Byte = oEntity.yProductionType

        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

        If mbBuildDisplayed = True OrElse mbContentsDisplayed = True OrElse mbOrdersDisplayed = True OrElse oEntity.ObjTypeID = ObjectType.eFacility OrElse oEntity.yProductionType = ProductionType.eFacility OrElse (goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEntity.oMesh.bLandBased = False) Then
            Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                Dim oSelColor As System.Drawing.Color
                oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)

                oSprite.Begin(SpriteFlags.AlphaBlend)

                Try
                    If mbContentsDisplayed = True Then
                        DoMultiColorFill(rcContents, oSelColor, rcContents.Location, oSprite)
                    End If
                    If mbBuildDisplayed = True Then
                        DoMultiColorFill(rcBuild, oSelColor, rcBuild.Location, oSprite)
                    End If
                    If mbOrdersDisplayed = True Then
                        DoMultiColorFill(rcOrders, oSelColor, rcOrders.Location, oSprite)
                    End If

                    'While we have a sprite created, let's render our two buttons (Active and Auto-Launch)
                    If oEntity.ObjTypeID = ObjectType.eFacility Then
                        Dim clrVal As System.Drawing.Color

                        If (GetSelectionCurrentStatus() And elUnitStatus.eFacilityPowered) <> 0 Then
                            clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                            mbCurrentActive = True
                        Else : clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                        End If
                        oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eSphere), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcActive.Location, clrVal)

                        If (oEntity.yProductionType And ProductionType.eProduction) <> 0 OrElse (oEntity.yProductionType And ProductionType.eFacility) <> 0 OrElse (oEntity.CritList And elUnitStatus.eHangarOperational) <> 0 OrElse (oEntity.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then
                            mbAutoLaunchDisplayed = True
                            If oEntity.AutoLaunch = True Then
                                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                                mbCurrentAutoLaunch = True
                            Else : clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                            End If
                            oSprite.Draw2D(goResMgr.GetTexture("HullBuilderIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.Pak"), New Rectangle(16, 32, 16, 16), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcAutoLaunch.Location, clrVal)
                        Else
                            mbAutoLaunchDisplayed = False
                        End If


                        'Ok, facilities ALWAYS have the repair visible (if they are damaged) hehe
                        With oEntity
                            Dim bDamaged As Boolean = False
                            For X As Int32 = 0 To .yArmorHP.GetUpperBound(0)
                                If .yArmorHP(X) <> 100 AndAlso .yArmorHP(X) <> 255 Then
                                    bDamaged = True
                                End If
                            Next X
                            bDamaged = bDamaged OrElse .yStructureHP <> 100 OrElse (.CritList And ml_CRITS_WE_CARE_ABOUT) <> 0
                            If bDamaged = True Then
                                oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eWrench), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcRepair.Location, Color.White)
                                myRepairButtonVisible = 1
                            End If
                        End With
                    ElseIf oEntity.yProductionType = ProductionType.eFacility Then
                        'Engineers always have a repair button visible
                        If yPrevRepairButtonState = 0 Then myRepairButtonVisible = 1 Else myRepairButtonVisible = yPrevRepairButtonState
                        'If myRepairButtonVisible = 1 Then
                        oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eWrench), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcRepair.Location, Color.White)
                        'Else
                        '	oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, New Rectangle(192, 80, 16, 16), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcRepair.Location, Color.White)
                        'End If
                        oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDemolish), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcAutoLaunch.Location, Color.White)
                    End If
                    If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEntity.ObjTypeID = ObjectType.eUnit AndAlso oEntity.oMesh.bLandBased = False Then
                        oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.ePlanetOrbit), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcActive.Location, System.Drawing.Color.White)
                    End If
                    If oEntity.ObjTypeID = ObjectType.eUnit AndAlso oEntity.oMesh.bLandBased = False Then
                        oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcUnitGoto.Location, System.Drawing.Color.FromArgb(255, 192, 220, 255))
                    End If

                Catch
                End Try

                oSprite.End()
            End Using
        ElseIf oEntity.ObjTypeID = ObjectType.eUnit AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso oEntity.oMesh.bLandBased = False Then
            Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                oSprite.Begin(SpriteFlags.AlphaBlend)
                If oEntity.ObjTypeID = ObjectType.eUnit AndAlso oEntity.oMesh.bLandBased = False Then
                    oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcUnitGoto.Location, System.Drawing.Color.FromArgb(255, 192, 220, 255))
                End If

                oSprite.End()
            End Using

        End If

        'Now, draw our borders around the buttons 
        MyBase.RenderRoundedBorder(rcContents, 1, muSettings.InterfaceBorderColor)
        MyBase.RenderRoundedBorder(rcBuild, 1, muSettings.InterfaceBorderColor)
        MyBase.RenderRoundedBorder(rcOrders, 1, muSettings.InterfaceBorderColor)

        'Now, render our text...
        Dim clrTemp As System.Drawing.Color
        Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                oTextSpr.Begin(SpriteFlags.AlphaBlend)
                Try
                    clrTemp = muSettings.InterfaceBorderColor
                    If Not ( _
                        (((oEntity.CritList And elUnitStatus.eCargoBayOperational) <> 0 OrElse (oEntity.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0) _
                        OrElse (oEntity.CritList And elUnitStatus.eHangarOperational) <> 0 OrElse (oEntity.CurrentStatus And elUnitStatus.eHangarOperational) <> 0) _
                        ) Then

                        'If oEntity.oUnitDef Is Nothing OrElse oNode.lItemData2 <> ObjectType.eUnit Then
                        clrTemp = System.Drawing.Color.FromArgb(255, clrTemp.R \ 2, clrTemp.G \ 2, clrTemp.B \ 2)
                    End If

                    oFont.DrawText(oTextSpr, "Contents", rcContents, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrTemp)

                    If yProdType = ProductionType.eTradePost Then
                        oFont.DrawText(oTextSpr, "Trade", rcBuild, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                    ElseIf yProdType = ProductionType.eResearch Then
                        oFont.DrawText(oTextSpr, "Techs", rcBuild, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                    Else
                        clrTemp = muSettings.InterfaceBorderColor
                        If oEntity.yProductionType = 0 OrElse (oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.yProductionType = ProductionType.ePersonnel) Then ' oEntity.ObjTypeID <> ObjectType.eFacility Then ' OrElse (oEntity.ObjTypeID <> ObjectType.eFacility AndAlso oEntity.yProductionType = ProductionType.eCommandCenterSpecial) Then
                            clrTemp = System.Drawing.Color.FromArgb(255, clrTemp.R \ 2, clrTemp.G \ 2, clrTemp.B \ 2)
                            'ElseIf oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.yProductionType = 0 Then
                            'Stop
                        End If
                        oFont.DrawText(oTextSpr, "Build", rcBuild, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrTemp)
                    End If

                    oFont.DrawText(oTextSpr, "Orders", rcOrders, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                Catch
                End Try
                oTextSpr.End()
                oTextSpr.Dispose()
            End Using
            oFont.Dispose()
        End Using


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

    Public Sub SetFromEntity(ByVal lEntityIndex As Int32)
        'Verify our objects
        If goCurrentEnvir Is Nothing Then Exit Sub
        If lEntityIndex < 0 OrElse lEntityIndex > goCurrentEnvir.lEntityUB Then Exit Sub
        If goCurrentEnvir.oEntity(lEntityIndex) Is Nothing Then Exit Sub
        If goCurrentEnvir.oEntity(lEntityIndex).oUnitDef Is Nothing Then Exit Sub

        mlEntityIndex = lEntityIndex
        If goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eUnit AndAlso goCurrentEnvir.oEntity(mlEntityIndex).oMesh.bLandBased = False Then
            Me.Width = 276
        ElseIf goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eUnit AndAlso goCurrentEnvir.oEntity(mlEntityIndex).oMesh.bLandBased = True Then
            Me.Width = 234
        Else
            Me.Width = 256
        End If
        mbUseChild = False
        mlChildID = -1
        miChildTypeID = -1

        With goCurrentEnvir.oEntity(lEntityIndex)
            If .ObjTypeID = ObjectType.eFacility AndAlso .MaxColonists = 0 AndAlso .MaxWorkers = 0 AndAlso .PowerConsumption = 0 Then
                Dim yMsg(8) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityPersonnel).CopyTo(yMsg, 0)
                .GetGUIDAsString.CopyTo(yMsg, 2)
                yMsg(8) = 2     'pass two to indicate a simple request
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
        End With

        Dim oTmpWin As UIWindow
        oTmpWin = MyBase.moUILib.GetWindow("frmBuildWindow")
        If oTmpWin Is Nothing Then oTmpWin = MyBase.moUILib.GetWindow("frmMining")
        If oTmpWin Is Nothing = False AndAlso oTmpWin.Visible = True Then
            mbBuildDisplayed = True
        Else : mbBuildDisplayed = False
        End If
        oTmpWin = Nothing

    End Sub

    Public Sub ResetBuildButton()
        mbBuildDisplayed = False
        'HandleBuildButtonClick()
        Me.IsDirty = True
    End Sub

    Private Sub HandleBuildButtonClick()
        If mbBuildDisplayed = True Then
            If mlEntityIndex <> -1 AndAlso goCurrentEnvir Is Nothing = False Then

                If goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eFacility AndAlso (Not goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eTradePost) AndAlso (goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus And elUnitStatus.eFacilityPowered) = 0 Then
                    MyBase.moUILib.AddNotification("Cannot open build interface while facility is unpowered.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowPower.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                    mbBuildDisplayed = False
                    Return
                End If

                If goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eResearch Then
                    Dim ofrmResMain As frmResearchMain = New frmResearchMain(MyBase.moUILib)
                    ofrmResMain = Nothing
                ElseIf goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eTradePost Then
                    Dim ofrmTrade As frmTradeMain = CType(goUILib.GetWindow("frmTradeMain"), frmTradeMain)
                    If ofrmTrade Is Nothing Then
                        ofrmTrade = New frmTradeMain(goUILib, goCurrentEnvir.oEntity(mlEntityIndex).ObjectID)
                        'Else : ofrmTrade.SetTradePostID(goCurrentEnvir.oEntity(mlEntityIndex).ObjectID)
                        ofrmTrade.Visible = True
                    Else
                        MyBase.moUILib.RemoveWindow("frmTradeMain")
                        Return
                    End If
                    ofrmTrade = Nothing
                ElseIf goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eMining Then
                    Dim oFrm As frmMining = CType(MyBase.moUILib.GetWindow("frmMining"), frmMining)
                    If oFrm Is Nothing Then oFrm = New frmMining(MyBase.moUILib)
                    oFrm.Visible = True
                    oFrm.FindAndSelectFacility(mlEntityIndex)
                ElseIf goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = 0 OrElse goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.ePersonnel Then
                    mbBuildDisplayed = False
                Else

                    'Units cannot have a build queue...
                    Dim bUnitBuild As Boolean = (goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = ObjectType.eUnit)
                    Dim oTmp As frmBuildWindow = New frmBuildWindow(MyBase.moUILib, bUnitBuild)

                    If NewTutorialManager.TutorialOn = True Then
                        If bUnitBuild = True Then
                            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eEngineerBuildWindowOpen, -1, -1, -1, "")
                        Else
                            With goCurrentEnvir.oEntity(mlEntityIndex)
                                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eFacilityBuildWindowOpen, .ObjectID, .ObjTypeID, .yProductionType, "")
                            End With
                        End If
                    End If

                    Dim sFrmName As String = oTmp.ControlName
                    If oTmp.UpdateFromEntity(mlEntityIndex) = False Then
                        oTmp = Nothing
                        MyBase.moUILib.RemoveWindow(sFrmName)
                        mbBuildDisplayed = False
                    End If
                    oTmp = Nothing

                    'Check if this is a spacestation special
                    If goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eSpaceStationSpecial Then
                        'Yes... we need to acquire the space station colony's Child list
                        Dim yMsg(11) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eGetColonyChildList).CopyTo(yMsg, 0)
                        goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
                        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 8)
                        MyBase.moUILib.SendMsgToPrimary(yMsg)

                        'Now... we should already be set to the space station's entity index... so nothing needs doing there

                        'show frmSelectFac... it will fill with what we know so far... (or as of last update)
                        Dim ofrmSelectFac As frmSelectFac = New frmSelectFac(goUILib)
                        ofrmSelectFac.SetFromEntityIndex(mlEntityIndex)
                        ofrmSelectFac.Visible = True
                        ofrmSelectFac = Nothing
                    End If
                End If
            Else
                mbBuildDisplayed = False
            End If
        Else
            'need to hide the build screen
            MyBase.moUILib.RemoveWindow("frmBuildWindow")
            MyBase.moUILib.RemoveWindow("frmResearchMain")
            MyBase.moUILib.RemoveWindow("frmSelectFac")
            MyBase.moUILib.RemoveWindow("frmMining")
            Dim oFrm As frmTradeMain = CType(goUILib.GetWindow("frmTradeMain"), frmTradeMain)
            If oFrm Is Nothing = False AndAlso oFrm.Visible = True Then
                MyBase.moUILib.RemoveWindow("frmTradeMain")
                ReturnToPreviousView()
            End If
        End If
        'Return True
    End Sub

    Public Function HandlePowerButtonClick() As Boolean

        If goCurrentEnvir Is Nothing Then Exit Function
        If mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 OrElse goCurrentEnvir.oEntity(mlEntityIndex) Is Nothing Then Exit Function

        If HasAliasedRights(AliasingRights.eAlterAutoLaunchPower) = False Then Return False

        mbCurrentActive = Not mbCurrentActive

        Dim yMsg(11) As Byte
        Dim lStatus As Int32

        Dim lEntityStatus As Int32 = goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus
        Dim oChild As StationChild = Nothing

        If mbUseChild = True Then
            oChild = goCurrentEnvir.oEntity(mlEntityIndex).GetChild(mlChildID, miChildTypeID)
            If oChild Is Nothing = False Then
                lEntityStatus = oChild.lChildStatus
            End If
        End If

        If mbUseChild = False OrElse oChild Is Nothing = False Then
            If (lEntityStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                lStatus = -elUnitStatus.eFacilityPowered
                If mbUseChild = True Then
                    oChild.lChildStatus -= elUnitStatus.eFacilityPowered
                Else : goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus -= elUnitStatus.eFacilityPowered
                End If
            Else
                lStatus = elUnitStatus.eFacilityPowered
                If mbUseChild = True Then
                    oChild.lChildStatus = oChild.lChildStatus Or lStatus
                Else : goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus = goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus Or lStatus
                End If
            End If

            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yMsg, 0)
            If mbUseChild = True Then
                System.BitConverter.GetBytes(mlChildID).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(miChildTypeID).CopyTo(yMsg, 6)
            Else : goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
            End If
            System.BitConverter.GetBytes(lStatus).CopyTo(yMsg, 8)

            MyBase.moUILib.SendMsgToPrimary(yMsg)

            ReDim yMsg(8)
            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityPersonnel).CopyTo(yMsg, 0)
            If mbUseChild = True Then
                System.BitConverter.GetBytes(mlChildID).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(miChildTypeID).CopyTo(yMsg, 6)
            Else : goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
            End If
            yMsg(8) = 2     'pass two to indicate a simple request
            MyBase.moUILib.SendMsgToPrimary(yMsg)
        End If
        Return True
    End Function

    Public Function HandleAutoLaunchButtonClick() As Boolean
        If goCurrentEnvir Is Nothing Then Return False
        If mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return False

        Dim oChild As StationChild = Nothing

        If HasAliasedRights(AliasingRights.eAlterAutoLaunchPower) = False Then Return False
        If goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID <> ObjectType.eFacility Then Return False

        If Not ((goCurrentEnvir.oEntity(mlEntityIndex).yProductionType And ProductionType.eProduction) <> 0 OrElse (goCurrentEnvir.oEntity(mlEntityIndex).yProductionType And ProductionType.eFacility) <> 0 OrElse (goCurrentEnvir.oEntity(mlEntityIndex).CritList And elUnitStatus.eHangarOperational) <> 0 OrElse (goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus And elUnitStatus.eHangarOperational) <> 0) Then
            Return False
        End If

        Dim yMsg(8) As Byte
        Dim lPos As Int32 = 0

        Dim bCurrentVal As Boolean

        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityPersonnel).CopyTo(yMsg, lPos) : lPos += 2

        If mbUseChild = True Then
            oChild = goCurrentEnvir.oEntity(mlEntityIndex).GetChild(mlChildID, miChildTypeID)
            If oChild Is Nothing Then Return False
            System.BitConverter.GetBytes(oChild.lChildID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(oChild.iChildTypeID).CopyTo(yMsg, lPos) : lPos += 2

            'TODO: Figure out the current value... for now, return
            Return False
        Else
            goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            bCurrentVal = goCurrentEnvir.oEntity(mlEntityIndex).AutoLaunch
        End If

        'Now, toggle our currentval...
        bCurrentVal = Not bCurrentVal

        If bCurrentVal = True Then
            yMsg(lPos) = 1
        Else : yMsg(lPos) = 0
        End If

        MyBase.moUILib.SendMsgToPrimary(yMsg)
        Return True
    End Function

    Private Sub HandleFacilityRepairClick()
        If HasAliasedRights(AliasingRights.eAddProduction) = False Then Return
        Try
            Dim oRepair As frmRepair = CType(MyBase.moUILib.GetWindow("frmRepair"), frmRepair)
            If oRepair Is Nothing Then oRepair = New frmRepair(MyBase.moUILib)
            oRepair.Visible = True
            With goCurrentEnvir.oEntity(mlEntityIndex)
                oRepair.SetFromEntity(.ObjectID, .ObjTypeID, .ObjectID, .ObjTypeID)
            End With
        Catch ex As Exception
            MyBase.moUILib.RemoveWindow("frmRepair")
        End Try
    End Sub

    Public Function SetEntityStatusMsgRcvd(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lStatus As Int32) As Boolean
        If mlEntityIndex = -1 Then Return False
        If goCurrentEnvir Is Nothing Then Return False
        If goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return False

        If mbUseChild = False Then
            If goCurrentEnvir.lEntityIdx(mlEntityIndex) = lID AndAlso goCurrentEnvir.oEntity(mlEntityIndex).ObjTypeID = iTypeID Then
                If lStatus = -elUnitStatus.eFacilityPowered Then
                    If mbCurrentActive = True Then
                        'We were expecting it to be powered, so alert the player to low power
                        goUILib.AddNotification("Not enough power!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowPower.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                    End If
                Else
                    If mbCurrentActive = False Then
                        goUILib.AddNotification("Facilities require this facility to be active!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowPower.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                    End If
                End If
            Else : Return False
            End If
        Else
            Dim oChild As StationChild = goCurrentEnvir.oEntity(mlEntityIndex).GetChild(mlChildID, miChildTypeID)
            If oChild Is Nothing = False Then

                If lStatus = -elUnitStatus.eFacilityPowered Then
                    If mbCurrentActive = True Then
                        'We were expecting it to be powered, so alert the player to low power
                        goUILib.AddNotification("Not enough power!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowPower.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                    End If
                Else
                    If mbCurrentActive = False Then
                        goUILib.AddNotification("Facilities require this facility to be active!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowPower.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                    End If
                End If
            Else : Return False
            End If
        End If
        Me.IsDirty = True
        Return True
    End Function

    Public Sub SetFromEntityWithChild(ByVal lEntityIndex As Int32, ByVal lChildID As Int32, ByVal iChildTypeID As Int16)
        'Verify our objects
        If goCurrentEnvir Is Nothing Then Return
        If lEntityIndex < 0 OrElse lEntityIndex > goCurrentEnvir.lEntityUB Then Return
        If goCurrentEnvir.oEntity(lEntityIndex) Is Nothing Then Return
        If goCurrentEnvir.oEntity(lEntityIndex).oUnitDef Is Nothing Then Return

        mlEntityIndex = lEntityIndex

        Dim oTmpWin As frmBuildWindow
        oTmpWin = CType(MyBase.moUILib.GetWindow("frmBuildWindow"), frmBuildWindow)
        If oTmpWin Is Nothing = False AndAlso oTmpWin.Visible = True Then
            mbBuildDisplayed = True
        Else : mbBuildDisplayed = False
        End If
        oTmpWin = Nothing

        mbUseChild = False
        mlChildID = lChildID
        miChildTypeID = iChildTypeID
        If mlChildID > 0 AndAlso miChildTypeID > 0 Then
            Dim oChild As StationChild = goCurrentEnvir.oEntity(mlEntityIndex).GetChild(mlChildID, miChildTypeID)
            If oChild Is Nothing = False Then
                mbUseChild = True
            End If
        End If

        Me.IsDirty = True
    End Sub

    Private Function GetSelectionCurrentStatus() As Int32
		Try
			If goCurrentEnvir Is Nothing Then Return 0
			If mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return 0

			Dim lEntityStatus As Int32 = goCurrentEnvir.oEntity(mlEntityIndex).CurrentStatus

            Dim oChild As StationChild = Nothing

            If mbUseChild = True Then
                oChild = goCurrentEnvir.oEntity(mlEntityIndex).GetChild(mlChildID, miChildTypeID)
                If oChild Is Nothing = False Then
                    lEntityStatus = oChild.lChildStatus
                End If
            End If
            Return lEntityStatus
        Catch

        End Try
		Return 0
    End Function

    Public Sub OpenOrdersWindows()
        If NewTutorialManager.TutorialOn = True Then
            Dim sParms() As String = {"Orders"}
            If MyBase.moUILib.CommandAllowedWithParms(True, Me.ControlName, sParms, False) = False Then Return
        End If
        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenOrdersWindow)

        mbOrdersDisplayed = Not mbOrdersDisplayed
        If mbOrdersDisplayed = True Then
            If HasAliasedRights(AliasingRights.eChangeBehavior) = True Then
                Dim oOrders As frmBehavior = CType(MyBase.moUILib.GetWindow("frmBehavior"), frmBehavior)
                If oOrders Is Nothing Then oOrders = New frmBehavior(MyBase.moUILib)
                oOrders.MultiSelect = False
                oOrders.SetFromEntity(mlEntityIndex)
                oOrders.btnMinMax_Click("")      'ensures it is maximized
            Else
                mbOrdersDisplayed = False
            End If
        Else
            MyBase.moUILib.RemoveWindow("frmBehavior")
            MyBase.moUILib.RemoveWindow("frmRouteConfig")
        End If

        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
    End Sub

    Public Sub OpenContentsWindow()
        mbContentsDisplayed = Not mbContentsDisplayed
        If mbContentsDisplayed = True Then
            If mlEntityIndex <> -1 AndAlso goCurrentEnvir Is Nothing = False Then
                Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(mlEntityIndex)
                If oEntity Is Nothing = False AndAlso _
                    ( _
                        (((oEntity.CritList And elUnitStatus.eCargoBayOperational) <> 0 OrElse (oEntity.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0) _
                        OrElse (oEntity.CritList And elUnitStatus.eHangarOperational) <> 0 OrElse (oEntity.CurrentStatus And elUnitStatus.eHangarOperational) <> 0) _
                    ) Then

                    If oEntity.ObjTypeID = ObjectType.eFacility AndAlso (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 Then
                        MyBase.moUILib.AddNotification("Cannot open contents interface while facility is unpowered.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowPower.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                        mbContentsDisplayed = False
                        Return
                    End If

                    Dim oContents As frmContents = CType(MyBase.moUILib.GetWindow("frmContents"), frmContents)
                    If oContents Is Nothing Then oContents = New frmContents(MyBase.moUILib)
                    oContents.Visible = True
                    oContents.SetEntityRef(mlEntityIndex)
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                    Me.IsDirty = True
                Else
                    mbContentsDisplayed = False
                    Return
                End If
            End If
        Else
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            Me.IsDirty = True
            MyBase.moUILib.RemoveWindow("frmContents")
        End If

    End Sub

    Public Sub OpenBuildWindow()
        mbBuildDisplayed = Not mbBuildDisplayed
        HandleBuildButtonClick()
        If mbBuildDisplayed = True AndAlso goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
        Me.IsDirty = True
    End Sub

    Private Sub frmAdvanceDisplay_WindowMoved() Handles Me.WindowMoved
        muSettings.AdvanceDisplayLocX = Me.Left
        muSettings.AdvanceDisplayLocY = Me.Top
    End Sub

    Public Function HandleUnitGotoClick() As Boolean

        If goCurrentEnvir Is Nothing Then Return False
        If mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 OrElse goCurrentEnvir.oEntity(mlEntityIndex) Is Nothing Then Return False

        If HasAliasedRights(AliasingRights.eMoveUnits) = False Then
            goUILib.AddNotification("You lack rights to change unit environments.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        Dim oEntity As BaseEntity = Nothing
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return False

        If mlEntityIndex < 0 Then Return False
        If oEnvir.lEntityUB < mlEntityIndex Then Return False
        If oEnvir.lEntityIdx(mlEntityIndex) < 0 Then Return False
        oEntity = oEnvir.oEntity(mlEntityIndex)
        If oEntity Is Nothing Then Return False

        If Not (oEntity.ObjTypeID = ObjectType.eUnit AndAlso oEntity.oMesh.bLandBased = False) Then Return False

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
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
        Return True
    End Function

End Class