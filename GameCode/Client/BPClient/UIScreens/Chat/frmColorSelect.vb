Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmColorSelect
    Inherits UIWindow

    Private lblRed As UILabel
    Private lblGreen As UILabel
    Private lblBlue As UILabel
    Private WithEvents hscrRed As UIScrollBar
    Private WithEvents hscrGreen As UIScrollBar
    Private WithEvents hscrBlue As UIScrollBar
    Private fraResult As UIWindow
    Private WithEvents btnOK As UIButton
    Private WithEvents btnCancel As UIButton
    Private WithEvents txtRed As UITextBox
    Private WithEvents txtGreen As UITextBox
    Private WithEvents txtBlue As UITextBox

    Private mbLoading As Boolean = True

    Public Event FormClosed(ByVal clrResult As System.Drawing.Color)

    Private mclrOriginal As System.Drawing.Color

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        ' initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eColorSelect
            .ControlName = "frmColorSelect"
            .Left = 341
            .Top = 374
            .Width = 275
            .Height = 100
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 1
            .Moveable = True
        End With

        'lblRed initial props
        lblRed = New UILabel(oUILib)
        With lblRed
            .ControlName = "lblRed"
            .Left = 5
            .Top = 5
            .Width = 38
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Red:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblRed, UIControl))

        'lblGreen initial props
        lblGreen = New UILabel(oUILib)
        With lblGreen
            .ControlName = "lblGreen"
            .Left = 5
            .Top = 25
            .Width = 41
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Green:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblGreen, UIControl))

        'lblBlue initial props
        lblBlue = New UILabel(oUILib)
        With lblBlue
            .ControlName = "lblBlue"
            .Left = 5
            .Top = 45
            .Width = 38
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Blue:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblBlue, UIControl))

        'hscrRed initial props
        hscrRed = New UIScrollBar(oUILib, False)
        With hscrRed
            .ControlName = "hscrRed"
            .Left = 60
            .Top = 6
            .Width = 100
            .Height = 16
            .Enabled = True
            .Visible = True
            .Value = 0
            .MaxValue = 255
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 4
            .ReverseDirection = False
        End With
        Me.AddChild(CType(hscrRed, UIControl))

        'hscrGreen initial props
        hscrGreen = New UIScrollBar(oUILib, False)
        With hscrGreen
            .ControlName = "hscrGreen"
            .Left = 60
            .Top = 26
            .Width = 100
            .Height = 16
            .Enabled = True
            .Visible = True
            .Value = 0
            .MaxValue = 255
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 4
            .ReverseDirection = False
        End With
        Me.AddChild(CType(hscrGreen, UIControl))

        'hscrBlue initial props
        hscrBlue = New UIScrollBar(oUILib, False)
        With hscrBlue
            .ControlName = "hscrBlue"
            .Left = 60
            .Top = 46
            .Width = 100
            .Height = 16
            .Enabled = True
            .Visible = True
            .Value = 0
            .MaxValue = 255
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 4
            .ReverseDirection = False
        End With
        Me.AddChild(CType(hscrBlue, UIControl))

        'fraResult initial props
        fraResult = New UIWindow(oUILib)
        With fraResult
            .ControlName = "fraResult"
            .Left = 210
            .Top = 6
            .Width = 60
            .Height = 55
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With
        Me.AddChild(CType(fraResult, UIControl))

        'btnOK initial props
        btnOK = New UIButton(oUILib)
        With btnOK
            .ControlName = "btnOK"
            .Left = 10
            .Top = 70
            .Width = 100
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
            .Left = 165
            .Top = 70
            .Width = 100
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

        'txtRed initial props
        txtRed = New UITextBox(oUILib)
        With txtRed
            .ControlName = "txtRed"
            .Left = 165
            .Top = 5
            .Width = 40
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
            .MaxLength = 3
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtRed, UIControl))

        'txtGreen initial props
        txtGreen = New UITextBox(oUILib)
        With txtGreen
            .ControlName = "txtGreen"
            .Left = 165
            .Top = 25
            .Width = 40
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
            .MaxLength = 3
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtGreen, UIControl))

        'txtBlue initial props
        txtBlue = New UITextBox(oUILib)
        With txtBlue
            .ControlName = "txtBlue"
            .Left = 165
            .Top = 45
            .Width = 40
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
            .MaxLength = 3
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtBlue, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        mbLoading = False
    End Sub

    Public Sub ShowForm(ByVal clrVal As System.Drawing.Color, ByVal lLeft As Int32, ByVal lTop As Int32)
        mbLoading = True
        mclrOriginal = clrVal
        hscrRed.Value = clrVal.R
        hscrGreen.Value = clrVal.G
        hscrBlue.Value = clrVal.B
        mbLoading = False
        RefreshAll()
    End Sub

    Private Sub ColorValueChange() Handles hscrBlue.ValueChange, hscrGreen.ValueChange, hscrRed.ValueChange
        If mbLoading = True Then Return
        mbLoading = True
        'we set textboxes to scrollbars
        txtBlue.Caption = hscrBlue.Value.ToString
        txtGreen.Caption = hscrGreen.Value.ToString
        txtRed.Caption = hscrRed.Value.ToString
        mbLoading = False

        RefreshAll()
    End Sub

    Private Sub TextboxValueChange() Handles txtBlue.TextChanged, txtGreen.TextChanged, txtRed.TextChanged
        If mbLoading = True Then Return
        mbLoading = True
        'we set scroll bars to textboxes
        Dim lVal As Int32 = CInt(Val(txtBlue.Caption))
        hscrBlue.Value = Math.Min(Math.Max(0, lVal), 255)
        lVal = CInt(Val(txtGreen.Caption))
        hscrGreen.Value = Math.Min(Math.Max(0, lVal), 255)
        lVal = CInt(Val(txtRed.Caption))
        hscrRed.Value = Math.Min(Math.Max(0, lVal), 255)
        mbLoading = False

        RefreshAll()
    End Sub

    Private Sub RefreshAll()
        If mbLoading = True Then Return
        fraResult.FillColor = System.Drawing.Color.FromArgb(255, hscrRed.Value, hscrGreen.Value, hscrBlue.Value)
    End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		RaiseEvent FormClosed(mclrOriginal)
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnOK_Click(ByVal sName As String) Handles btnOK.Click
		RaiseEvent FormClosed(fraResult.FillColor)
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub
End Class