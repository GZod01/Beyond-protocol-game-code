Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmGalaxyControl
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents chkHideWormholes As UICheckBox
    Private WithEvents chkHideFleetMovement As UICheckBox
    Private WithEvents chkStarLabelFalloff As UICheckBox

    Private lnDiv1 As UILine
    Private lnDiv2 As UILine

    Private lblStar As UILabel
    Private WithEvents txtStarName As UITextBox
    Private WithEvents btnStarFind As UIButton
    Private WithEvents btnStarGoto As UIButton
    Private lblFoundStar As UILabel

    Private mlFoundIdx As Int32 = 0
    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'Dim lTempTop As Int32
        'Dim oWin As UIWindow = oUILib.GetWindow("frmEnvirDisplay")
        '
        'If oWin Is Nothing = False Then
        'lTempTop = oWin.Top + oWin.Height + 1
        'Else : lTempTop = CInt(oUILib.oDevice.PresentationParameters.BackBufferHeight / 2) - 85
        'End If
        'oWin = Nothing
        ' initial props
        With Me
            .ControlName = "frmGalaxyControl"

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1


            lLeft = muSettings.GalaxyControlX
            lTop = muSettings.GalaxyControlY

            If lLeft < 0 Then lLeft = 0
            If lTop < 0 Then lTop = 0
            If lLeft + 201 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 201
            If lTop + 200 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 200


            '.Top = lTempTop
            '.Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 203
            .Top = lTop
            .Left = lLeft
            .Width = 201
            .Height = 200
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With


        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 2
            .Width = 147
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Galaxy Control"
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
            .Top = 25
            .Width = Me.Width - 1
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        '=================================== Hide Wormholes ====================================
        'chkHideWormholes initial props
        chkHideWormholes = New UICheckBox(oUILib)
        With chkHideWormholes
            .ControlName = "chkHideWormholes"
            .Left = 5
            .Top = lnDiv1.Top + lnDiv1.Height + 5
            .Width = 112
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hide Wormholes"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbGalaxyControlHideWormholes
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkHideWormholes, UIControl))

        '=================================== Hide Fleets aka Battlegroups ====================================
        'chkHideFleetMovement initial props
        chkHideFleetMovement = New UICheckBox(oUILib)
        With chkHideFleetMovement
            .ControlName = "chkHideFleetMovement"
            .Left = 5
            .Top = chkHideWormholes.Top + chkHideWormholes.Height + 1
            .Width = 120
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hide Battlegroups"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbGalaxyControlHideFleetMovement
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkHideFleetMovement, UIControl))

        '=================================== Star Label Falloff - by distance ====================================
        'chkStarLabelFalloff initial props
        chkStarLabelFalloff = New UICheckBox(oUILib)
        With chkStarLabelFalloff
            .ControlName = "chkStarLabelFalloff"
            .Left = 5
            .Top = chkHideFleetMovement.Top + chkHideFleetMovement.Height + 1
            .Width = 112
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Star Label Falloff"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbGalaxyControlStarLabelFalloff
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkStarLabelFalloff, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 1
            .Top = chkStarLabelFalloff.Top + chkStarLabelFalloff.Height + 1
            .Width = Me.Width - 1
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        '=================================== Find / Goto Star ====================================

        'lblStar initial props
        lblStar = New UILabel(oUILib)
        With lblStar
            .ControlName = "lblStar"
            .Left = 5
            .Top = lnDiv2.Top + lnDiv2.Height + 10
            .Width = 77
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Star Name :"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblStar, UIControl))


        'txtStarName initial props
        txtStarName = New UITextBox(oUILib)
        With txtStarName
            .ControlName = "txtStarName"
            .Left = lblStar.Left + lblStar.Width + 10
            .Top = lblStar.Top
            .Width = Me.Width - .Left - 5
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
        Me.AddChild(CType(txtStarName, UIControl))

        'btnStarFind initial props
        btnStarFind = New UIButton(oUILib)
        With btnStarFind
            .ControlName = "btnSet"
            .Left = txtStarName.Left
            .Top = txtStarName.Top + txtStarName.Height + 5
            .Width = 50
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Center"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnStarFind, UIControl))

        'btnStarGoto initial props
        btnStarGoto = New UIButton(oUILib)
        With btnStarGoto
            .ControlName = "btnStarGoto"
            .Left = btnStarFind.Left + btnStarFind.Width + 5
            .Top = btnStarFind.Top
            .Width = 50
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Goto"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnStarGoto, UIControl))

        'lblFoundStar initial props
        lblFoundStar = New UILabel(oUILib)
        With lblFoundStar
            .ControlName = "lblFoundStar"
            .Left = 5
            .Top = btnStarGoto.Top + btnStarGoto.Height + 5
            .Width = Me.Width - 5
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Found Star :"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblFoundStar, UIControl))

        Me.Height = lblFoundStar.Top + lblFoundStar.Height + 1

        'goUILib.RemoveWindow("frmColonyStats")
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
    End Sub

    Private Sub frmGalaxyControl_OnNewFrame() Handles Me.OnNewFrame
        'If glCurrentEnvirView <> CurrentView.eGalaxyMapView AndAlso (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255) Then
        If glCurrentEnvirView <> CurrentView.eGalaxyMapView AndAlso (goCurrentPlayer Is Nothing = False) Then
            Me.Visible = False
            MyBase.moUILib.RemoveWindow(Me.ControlName)
        End If
    End Sub

    Private Sub frmGalaxyControl_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.GalaxyControlX = Me.Left
            muSettings.GalaxyControlY = Me.Top
        End If
    End Sub

    Private Sub chkHideWormholes_Click() Handles chkHideWormholes.Click
        muSettings.gbGalaxyControlHideWormholes = chkHideWormholes.Value
    End Sub

    Private Sub chkHideFleetMovement_Click() Handles chkHideFleetMovement.Click
        muSettings.gbGalaxyControlHideFleetMovement = chkHideFleetMovement.Value
    End Sub

    Private Sub chkStarLabelFalloff_Click() Handles chkStarLabelFalloff.Click
        muSettings.gbGalaxyControlStarLabelFalloff = chkStarLabelFalloff.Value
    End Sub

    Private Sub btnStarFind_Click(ByVal sName As String) Handles btnStarFind.Click
        'Center camera on the star
        If txtStarName.Caption.Trim <> "" Then
            lblFoundStar.Caption = "Found Star : "
            mlFoundIdx = goGalaxy.CenterOnStar(txtStarName.Caption, mlFoundIdx)
            If mlFoundIdx = 0 Then Exit Sub
            If goGalaxy.moSystems(mlFoundIdx) Is Nothing Then Exit Sub
            'goUILib.AddNotification("Found Star : " & goGalaxy.moSystems(mlFoundIdx).SystemName & " " & goGalaxy.moSystems(mlFoundIdx).ObjectID.ToString & " (" & goGalaxy.moSystems(mlFoundIdx).LocX.ToString & "," & goGalaxy.moSystems(mlFoundIdx).LocY.ToString & "," & goGalaxy.moSystems(mlFoundIdx).LocZ.ToString & ")", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            lblFoundStar.Caption = "Found Star : " & goGalaxy.moSystems(mlFoundIdx).SystemName
        End If
    End Sub

    Private Sub btnStarGoto_Click(ByVal sName As String) Handles btnStarGoto.Click
        'Change environments to the star
        If mlFoundIdx > 0 AndAlso Not goGalaxy.moSystems(mlFoundIdx) Is Nothing Then
            goGalaxy.GotoEnvironment(goGalaxy.moSystems(mlFoundIdx).ObjectID)
        ElseIf txtStarName.Caption <> "" Then
            goGalaxy.GotoStar(txtStarName.Caption)
        End If
    End Sub

    Private Sub txtStarName_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStarName.OnKeyPress
        If mlFoundIdx > 0 Then mlFoundIdx = 0
    End Sub

    Private Sub txtStarName_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStarName.OnKeyDown
        If e.KeyCode = Keys.Enter Then
            btnStarGoto_Click(Nothing)
        End If
    End Sub
End Class
