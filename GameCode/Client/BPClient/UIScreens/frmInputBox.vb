Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmInputBox
    Inherits UIWindow

    Public lblEnter As UILabel
    Public WithEvents txtValue As UITextBox
    Private WithEvents btnOK As UIButton
    Private WithEvents btnCancel As UIButton

    Public Event InputBoxClosed(ByVal bCancelled As Boolean, ByVal sValue As String)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmInputBox initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eInputBox
            .ControlName = "frmInputBox"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 120
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 35
            .Width = 240
            .Height = 70
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

        'lblEnter initial props
        lblEnter = New UILabel(oUILib)
        With lblEnter
            .ControlName = "lblEnter"
            .Left = 10
            .Top = 10
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Enter new value:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblEnter, UIControl))

        'txtValue initial props
        txtValue = New UITextBox(oUILib)
        With txtValue
            .ControlName = "txtValue"
            .Left = 110
            .Top = 10
            .Width = 120
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtValue, UIControl))

        'btnOK initial props
        btnOK = New UIButton(oUILib)
        With btnOK
            .ControlName = "btnOK"
            .Left = 10
            .Top = 40
            .Width = 80
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "OK"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnOK, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = 150
            .Top = 40
            .Width = 80
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

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

	Private Sub btnOK_Click(ByVal sName As String) Handles btnOK.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		RaiseEvent InputBoxClosed(False, txtValue.Caption)
	End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		RaiseEvent InputBoxClosed(True, "")
	End Sub

    Public Sub SetText(ByVal sName As String)
        If txtValue.Caption <> sName Then txtValue.Caption = sName
        If txtValue.HasFocus = False Then txtValue.HasFocus = True
    End Sub
End Class