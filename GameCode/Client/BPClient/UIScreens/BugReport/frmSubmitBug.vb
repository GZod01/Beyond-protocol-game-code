Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmSubmitBug
    Inherits UIWindow

    Private WithEvents btnClose As UIButton
    Private WithEvents btnSubmit As UIButton
    Private txtDesc As UITextBox
    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblDetails As UILabel
    Private txtDetail As UITextBox

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmSubmitBug initial props
        With Me
            .ControlName = "frmSubmitBug"
            .Left = 285
            .Top = 60
            .Width = 512
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 487
            .Top = 2
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

        'btnSubmit initial props
        btnSubmit = New UIButton(oUILib)
        With btnSubmit
            .ControlName = "btnSubmit"
            .Left = 405
            .Top = 480
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Submit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSubmit, UIControl))

        'txtDesc initial props
        txtDesc = New UITextBox(oUILib)
        With txtDesc
            .ControlName = "txtDesc"
            .Left = 5
            .Top = 30
            .Width = 501
            .Height = 130
            .Enabled = True
            .Visible = True
            .Caption = "At DSE, we are always working to provide our customers a better gaming experience. We are sorry if you have encountered a bug that may be impacting that experience. " & _
              "Please be as detailed as possible with your explanation of the bug or exploit. If it is reproduceable, include steps to reproduce it." & vbCrLf & vbCrLf & _
              "Please realize that you are not guaranteed a response from support. This is intended for bug reporting. If we need further information, we will contact you."
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
            .BackColorEnabled = Me.FillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
            .Locked = True
        End With
        Me.AddChild(CType(txtDesc, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 166
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Bug/Exploit Submission"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'NewControl5 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 25
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblDetails initial props
        lblDetails = New UILabel(oUILib)
        With lblDetails
            .ControlName = "lblDetails"
            .Left = 5
            .Top = txtDesc.Top + txtDesc.Height + 5
            .Width = 211
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Please Describe the Problem:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDetails, UIControl))

        'txtDetail initial props
        txtDetail = New UITextBox(oUILib)
        With txtDetail
            .ControlName = "txtDetail"
            .Left = 5
            .Top = lblDetails.Top + lblDetails.Height + 5
            .Width = 501
            .Height = (btnSubmit.Top - 5) - .Top
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
        End With
        Me.AddChild(CType(txtDetail, UIControl))

        ''chkResponse initial props
        'chkResponse = New UICheckBox(oUILib)
        'With chkResponse
        '    .ControlName = "chkResponse"
        '    .Left = 15
        '    .Top = 482
        '    .Width = 220
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Request a response from Support"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        'End With
        'Me.AddChild(CType(chkResponse, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click
        Dim sBody As String = txtDetail.Caption


    End Sub
End Class