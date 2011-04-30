Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class ctlSetTechMineral
    Inherits UIControl

    Private btnSet As UIButton
    Private lblVal As UILabel

    Public mlMineralIndex As Int32 = -1

    Public Event SetButtonClicked(ByVal lIdx As Int32)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        With Me
            .bAcceptEvents = True
            .ControlName = "ctlSetTechMineral"
            .Enabled = True
            .Height = 18
            .Left = 0
            .Top = 0
            .Visible = True
            .Width = 175
        End With

        'lblVal initial props
        lblVal = New UILabel(oUILib)
        With lblVal
            .ControlName = "lblCaseMatItem"
            .Left = 0 - 5
            .Top = 0
            .Width = Me.Width - 38
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Unselected"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        End With
        Me.AddChild(CType(lblVal, UIControl))

        'btnSet initial props
        btnSet = New UIButton(oUILib)
        With btnSet
            .ControlName = "btnSet"
            .Left = Me.Width - 42
            .Top = 2
            .Width = 45
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Set"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnSet, UIControl))

        AddHandler btnSet.Click, AddressOf ButtonClicked
    End Sub

    Private Sub ButtonClicked(ByVal sName As String)
        RaiseEvent SetButtonClicked(mlMineralIndex)
    End Sub

    Public Sub SetMineralName(ByVal sName As String)
        lblVal.Caption = sName
    End Sub

    Private Sub ctlSetTechMineral_OnResize() Handles Me.OnResize
        If btnSet Is Nothing = False Then
            btnSet.Left = Me.Width - 42
            btnSet.Height = Me.Height
        End If
        If lblVal Is Nothing = False Then
            lblVal.Width = Me.Width - 38
            lblVal.Height = Me.Height
        End If
    End Sub
End Class
