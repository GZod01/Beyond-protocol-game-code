Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmSkillDetail
    Inherits UIWindow

    Private lblSkillName As UILabel
    Private lnDiv1 As UILine
    Private txtDesc As UITextBox
    Private lblSkillProf As UILabel

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmSkillDetail initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eSkillDetailWindow
            .ControlName = "frmSkillDetail"
            .Left = 249
            .Top = 269
            .Width = 256
            .Height = 192
            .Enabled = False
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
            .FullScreen = True
            .Moveable = True
        End With

        'lblSkillName initial props
        lblSkillName = New UILabel(oUILib)
        With lblSkillName
            .ControlName = "lblSkillName"
            .Left = 5
            .Top = 5
            .Width = 195
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "SkillName"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSkillName, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 25
            .Width = 255
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'txtDesc initial props
        txtDesc = New UITextBox(oUILib)
        With txtDesc
            .ControlName = "txtDesc"
            .Left = 5
            .Top = 30
            .Width = 246
            .Height = Me.Height - 36
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
            .Locked = True
        End With
        Me.AddChild(CType(txtDesc, UIControl))

        'lblSkillProf initial props
        lblSkillProf = New UILabel(oUILib)
        With lblSkillProf
            .ControlName = "lblSkillProf"
            .Left = 200
            .Top = 5
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "100/100"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSkillProf, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Public Sub SetFromSkill(ByRef oSkill As Skill, ByVal lProf As Int32)
        lblSkillName.Caption = oSkill.SkillName
        txtDesc.Caption = oSkill.SkillDesc
        lblSkillProf.Caption = lProf.ToString & "/" & oSkill.MaxVal.ToString
    End Sub
End Class