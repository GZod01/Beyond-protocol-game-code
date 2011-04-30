Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmMinimapZoom
    Inherits UIWindow

    Private WithEvents btnZoomIn As UIButton
    Private lblZoom As UILabel
    Private WithEvents btnZoomOut As UIButton

    Private lblSize As UILabel
    Private WithEvents btnSizeUp As UIButton
    Private WithEvents btnSizeDown As UIButton

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmMinimapZoom initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eMinimapZoom
            .ControlName = "frmMinimapZoom"
            .Left = muSettings.MiniMapLocX
            .Top = muSettings.MiniMapLocY + muSettings.MiniMapWidthHeight
            .Width = 120
            .Height = 32
            If NewTutorialManager.TutorialOn Then .Height = 16
            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 1
            .Moveable = False
        End With

        'btnZoomIn initial props
        btnZoomIn = New UIButton(oUILib)
        With btnZoomIn
            .ControlName = "btnZoomIn"
            .Left = 50
            .Top = 1
            .Width = 32
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = "In"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnZoomIn, UIControl))

        'lblZoom initial props
        lblZoom = New UILabel(oUILib)
        With lblZoom
            .ControlName = "lblZoom"
            .Left = 2
            .Top = 0
            .Width = 100
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Zoom:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblZoom, UIControl))

        'btnZoomOut initial props
        btnZoomOut = New UIButton(oUILib)
        With btnZoomOut
            .ControlName = "btnZoomOut"
            .Left = 85
            .Top = 1
            .Width = 32
            .Height = 14
            .Enabled = True
            .Visible = True
            .Caption = "Out"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnZoomOut, UIControl))

        'lblSize initial props
        lblSize = New UILabel(oUILib)
        With lblSize
            .ControlName = "lblSize"
            .Left = 2
            .Top = 17
            .Width = 100
            .Height = 16
            .Enabled = True
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = "Size:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSize, UIControl))

        'btnSizeDown initial props
        btnSizeDown = New UIButton(oUILib)
        With btnSizeDown
            .ControlName = "btnSizeDown"
            .Left = 50
            .Top = 18
            .Width = 32
            .Height = 14
            .Enabled = Not (muSettings.MiniMapWidthHeight <= 120)
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = "-"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSizeDown, UIControl))

        'btnSizeUp initial props
        btnSizeUp = New UIButton(oUILib)
        With btnSizeUp
            .ControlName = "btnSizeUp"
            .Left = 85
            .Top = 18
            .Width = 32
            .Height = 14
            .Enabled = Not (muSettings.MiniMapWidthHeight >= 300)
            .Visible = Not NewTutorialManager.TutorialOn
            .Caption = "+"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSizeUp, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

	Private Sub btnZoomIn_Click(ByVal sName As String) Handles btnZoomIn.Click
		muSettings.PlanetMinimapZoomLevel += CByte(1)
		If muSettings.PlanetMinimapZoomLevel > 3 Then muSettings.PlanetMinimapZoomLevel = 3
	End Sub

	Private Sub btnZoomOut_Click(ByVal sName As String) Handles btnZoomOut.Click
        If muSettings.PlanetMinimapZoomLevel <> 0 Then muSettings.PlanetMinimapZoomLevel -= CByte(1)
	End Sub

    Private Sub frmMinimapZoom_OnNewFrame() Handles Me.OnNewFrame
        If glCurrentEnvirView <> CurrentView.ePlanetView OrElse (muSettings.ShowMiniMap = False AndAlso (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 255)) Then
            Me.Visible = False
            MyBase.moUILib.RemoveWindow(Me.ControlName)
        End If
    End Sub

    Private Sub btnSizeUp_Click(ByVal sName As String) Handles btnSizeUp.Click
        muSettings.MiniMapWidthHeight += 60
        btnSizeDown.Enabled = True
        If muSettings.MiniMapWidthHeight >= 300 Then btnSizeUp.Enabled = False
        Me.Top = muSettings.MiniMapLocY + muSettings.MiniMapWidthHeight
    End Sub

    Private Sub btnSizeDown_Click(ByVal sName As String) Handles btnSizeDown.Click
        muSettings.MiniMapWidthHeight -= 60
        btnSizeUp.Enabled = True
        If muSettings.MiniMapWidthHeight <= 120 Then btnSizeDown.Enabled = False
        Me.Top = muSettings.MiniMapLocY + muSettings.MiniMapWidthHeight
    End Sub
End Class