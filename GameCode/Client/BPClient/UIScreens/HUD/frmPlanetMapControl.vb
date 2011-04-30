Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D


Public Class frmStrategicMapControl
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine

    Private lblHide As UILabel
    Private WithEvents btnRenderTerrain As UIButton
    Private WithEvents btnClearTacticalData As UIButton
    Private lnDiv2 As UILine

    Private WithEvents chkBlinkUnits As UICheckBox
    Private lnDiv3 As UILine

    Private WithEvents chkDontShowMyUnits As UICheckBox
    Private WithEvents chkDontShowGuilds As UICheckBox
    Private WithEvents chkDontShowAllied As UICheckBox
    Private WithEvents chkDontShowNeutral As UICheckBox
    Private WithEvents chkDontShowEnemy As UICheckBox
    Private WithEvents chkDontShowUnknown As UICheckBox


    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
        With Me
            .ControlName = "frmStrategicMapControl"
            .Width = 158
            .Height = 210

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1
            If lLeft < 0 Then lLeft = 0
            If lTop < 0 Then lTop = 0

            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height

            .Top = lTop
            .Left = lLeft

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
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
            .Caption = "Strategic Control"
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
            .Left = 2
            .Top = 25
            .Width = Me.Width - 3
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'btnRenderTerrain initial props
        btnRenderTerrain = New UIButton(oUILib)
        With btnRenderTerrain
            .ControlName = "btnRenderTerrain"
            .Left = 5
            .Top = lnDiv1.Top + lnDiv1.Height + 5
            .Width = 150
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Hide Terrain"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
            .ToolTipText = "This button will toggle rendering of the planet terrain."
        End With
        Me.AddChild(CType(btnRenderTerrain, UIControl))

        'btnClearTacticalData initial props
        btnClearTacticalData = New UIButton(oUILib)
        With btnClearTacticalData
            .ControlName = "btnClearTacticalData"
            .Left = 5
            .Top = btnRenderTerrain.Top + btnRenderTerrain.Height + 1
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Clear Tactical Data"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnClearTacticalData, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 2
            .Top = btnClearTacticalData.Top + btnClearTacticalData.Height + 1
            .Width = Me.Width - 3
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'chkBlinkUnits initial props
        chkBlinkUnits = New UICheckBox(oUILib)
        With chkBlinkUnits
            .ControlName = "chkBlinkUnits"
            .Left = 5
            .Top = lnDiv2.Top + lnDiv2.Height + 1
            .Width = 82
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Blink Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbPlanetMapBlinkUnits
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkBlinkUnits, UIControl))

        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = 2
            .Top = chkBlinkUnits.Top + chkBlinkUnits.Height + 1
            .Width = Me.Width - 3
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'chkDontShowMyUnits initial props
        chkDontShowMyUnits = New UICheckBox(oUILib)
        With chkDontShowMyUnits
            .ControlName = "chkDontShowMyUnits"
            .Left = 5
            .Top = lnDiv3.Top + lnDiv3.Height + 1
            .Width = 103
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hide My Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbPlanetMapDontShowMyUnits
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkDontShowMyUnits, UIControl))

        'chkDontShowGuilds initial props
        chkDontShowGuilds = New UICheckBox(oUILib)
        With chkDontShowGuilds
            .ControlName = "chkDontShowGuilds"
            .Left = 5
            .Top = chkDontShowMyUnits.Top + chkDontShowMyUnits.Height + 1
            .Width = 116
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hide Guild Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbPlanetMapDontShowGuilds
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkDontShowGuilds, UIControl))

        'chkDontShowAllied initial props
        chkDontShowAllied = New UICheckBox(oUILib)
        With chkDontShowAllied
            .ControlName = "chkDontShowAllied"
            .Left = 5
            .Top = chkDontShowGuilds.Top + chkDontShowGuilds.Height + 1
            .Width = 119
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hide Allied Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbPlanetMapDontShowAllied
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkDontShowAllied, UIControl))

        'chkDontShowNeutral initial props
        chkDontShowNeutral = New UICheckBox(oUILib)
        With chkDontShowNeutral
            .ControlName = "chkDontShowNeutral"
            .Left = 5
            .Top = chkDontShowAllied.Top + chkDontShowAllied.Height + 1
            .Width = 128
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hide Neutral Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbPlanetMapDontShowNeutral
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkDontShowNeutral, UIControl))

        'chkDontShowEnemy initial props
        chkDontShowEnemy = New UICheckBox(oUILib)
        With chkDontShowEnemy
            .ControlName = "chkDontShowEnemy"
            .Left = 5
            .Top = chkDontShowNeutral.Top + chkDontShowNeutral.Height + 1
            .Width = 127
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hide Enemy Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbPlanetMapDontShowEnemy
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkDontShowEnemy, UIControl))

        'chkDontShowUnknown initial props
        chkDontShowUnknown = New UICheckBox(oUILib)
        With chkDontShowUnknown
            .ControlName = "chkDontShowUnknown"
            .Left = 5
            .Top = chkDontShowEnemy.Top + chkDontShowEnemy.Height + 1
            .Width = 140
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hide Unknown Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = muSettings.gbPlanetMapDontShowUnknown
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkDontShowUnknown, UIControl))

        oUILib.RemoveWindow(Me.ControlName)
        oUILib.AddWindow(Me)
    End Sub

    Private Sub frmStrategicMapControl_OnNewFrame() Handles Me.OnNewFrame
        Try
            If goCurrentEnvir Is Nothing Then Return
            'want on 4 = CurrentView.ePlanetMapView
            'want on 1 = CurrentView.eSystemMapView1
            '        If Not (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255) Then Return
            Dim bTerrain As Boolean
            Dim bEnabled As Boolean
            If glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                bTerrain = False
            ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                bTerrain = False
                bEnabled = True
            Else
                bTerrain = True
                bEnabled = True
            End If
            If btnRenderTerrain.Enabled <> bTerrain Then btnRenderTerrain.Enabled = bTerrain
            If btnClearTacticalData.Enabled <> bTerrain Then btnClearTacticalData.Enabled = bTerrain
            If chkDontShowAllied.Enabled <> bEnabled Then chkDontShowAllied.Enabled = bEnabled
            If chkDontShowNeutral.Enabled <> bEnabled Then chkDontShowNeutral.Enabled = bEnabled
            If chkDontShowEnemy.Enabled <> bEnabled Then chkDontShowEnemy.Enabled = bEnabled
            If chkDontShowUnknown.Enabled <> bEnabled Then chkDontShowUnknown.Enabled = bEnabled
            'lEnvirType = CurrentView.eSystemView OrElse lEnvirType = CurrentView.eSystemMapView2
            Dim sCaption As String = ""
            If muSettings.gbPlanetMapDontRenderTerrain = False Then
                sCaption = "Hide Terrain"
            Else
                sCaption = "Show Terrain"
            End If
            If btnRenderTerrain.Caption <> sCaption Then btnRenderTerrain.Caption = sCaption
            If (glCurrentEnvirView = CurrentView.ePlanetMapView AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet) OrElse (goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso (glCurrentEnvirView = CurrentView.eSystemMapView1) OrElse (glCurrentEnvirView = CurrentView.eSystemMapView2)) Then
                Return
            Else
                Me.Visible = False
                'MyBase.moUILib.RemoveWindow(Me.ControlName)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnRenderTerrain_Click(ByVal sName As String) Handles btnRenderTerrain.Click
        If muSettings.gbPlanetMapDontRenderTerrain = True Then
            muSettings.gbPlanetMapDontRenderTerrain = False
            btnRenderTerrain.Caption = "Hide Terrain"
        Else
            muSettings.gbPlanetMapDontRenderTerrain = True
            btnRenderTerrain.Caption = "Show Terrain"
        End If
        Me.IsDirty = True
    End Sub

    Private Sub frmPlanetMapControl_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.OnKeyDown
        If e.KeyCode = Keys.T Then
            e.Handled = True
            btnRenderTerrain_Click(Nothing)
        End If
    End Sub

    Private Sub chkBlinkUnits_Click() Handles chkBlinkUnits.Click
        muSettings.gbPlanetMapBlinkUnits = chkBlinkUnits.Value
    End Sub

    Private Sub chkDontShowMyUnits_Click() Handles chkDontShowMyUnits.Click
        muSettings.gbPlanetMapDontShowMyUnits = chkDontShowMyUnits.Value
    End Sub

    Private Sub chkDontShowGuilds_Click() Handles chkDontShowGuilds.Click
        muSettings.gbPlanetMapDontShowGuilds = chkDontShowGuilds.Value
    End Sub

    Private Sub chkDontShowAllied_Click() Handles chkDontShowAllied.Click
        muSettings.gbPlanetMapDontShowAllied = chkDontShowAllied.Value
    End Sub

    Private Sub btnClearTacticalData_Click(ByVal sName As String) Handles btnClearTacticalData.Click
        'Clear Tactical

        If goCurrentEnvir Is Nothing = False Then
            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                'Is the entity valid?
                If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                    'With block for slight optimization
                    With goCurrentEnvir.oEntity(X)
                        If .yVisibility = eVisibilityType.FacilityIntel Then
                            .yVisibility = eVisibilityType.NoVisibility
                        End If
                    End With
                End If
            Next
            goCurrentEnvir.SaveEnvironmentTacticalData()
        End If
    End Sub

    Private Sub chkDontShowNeutral_Click() Handles chkDontShowNeutral.Click
        muSettings.gbPlanetMapDontShowNeutral = chkDontShowNeutral.Value
    End Sub

    Private Sub chkDontShowEnemy_Click() Handles chkDontShowEnemy.Click
        muSettings.gbPlanetMapDontShowEnemy = chkDontShowEnemy.Value
    End Sub

    Private Sub chkDontShowUnknown_Click() Handles chkDontShowUnknown.Click
        muSettings.gbPlanetMapDontShowUnknown = chkDontShowUnknown.Value
    End Sub



End Class

