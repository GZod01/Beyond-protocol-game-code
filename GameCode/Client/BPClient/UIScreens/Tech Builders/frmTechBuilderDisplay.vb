Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmTechBuilderDisplay
    Inherits UIWindow

    Private WithEvents lblDisplay As UILabel
    Private WithEvents btnClose As UIButton
    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmTechBuilderDisplay initial props
        With Me
            .ControlName = "frmTechBuilderDisplay"
            .Left = 4
            .Top = 11
            .Width = 256
            .Height = 760
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 2
        End With

        'lblDisplay initial props
        lblDisplay = New UILabel(oUILib)
        With lblDisplay
            .ControlName = "lblDisplay"
            .Left = 5
            .Top = 5
            .Width = 246
            .Height = 750
            .Enabled = True
            .Visible = True
            .Caption = "Display"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDisplay, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
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

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Public Sub SetValues(ByVal sString As String)
        lblDisplay.Caption = sString
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub frmTechBuilderDisplay_OnNewFrame() Handles Me.OnNewFrame

        Dim sNames() As String = {"frmAlloyBuilder", "frmArmorBuilder", "frmEngineBuilder", "frmHullBuilder", "frmPrototypeBuilder", "frmRadarBuilder", "frmShieldBuilder", "frmWeaponBuilder"}
        For X As Int32 = 0 To sNames.GetUpperBound(0)
            Dim oWin As UIWindow = MyBase.moUILib.GetWindow(sNames(X))
            If oWin Is Nothing = False AndAlso oWin.Visible = True Then Return
        Next X
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub
End Class