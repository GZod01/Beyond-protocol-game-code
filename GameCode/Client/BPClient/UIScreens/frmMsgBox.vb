Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmMsgBox
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private txtBody As UITextBox
    Private btnCmd() As UIButton

    Private lblSignature As UILabel
    Private txtSignature As UITextBox
    Private mbRequiresSignature As Boolean = False

    Public Event DialogClosed(ByVal lResult As MsgBoxResult)

    Public Sub New(ByRef oUILib As UILib, ByVal sMsg As String, ByVal eMsgBoxStyle As MsgBoxStyle, ByVal sTitle As String)
        MyBase.New(oUILib)

        mbRequiresSignature = (eMsgBoxStyle And MsgBoxStyle.Critical) <> 0
        If mbRequiresSignature = True Then
            eMsgBoxStyle = eMsgBoxStyle Xor MsgBoxStyle.Critical
        End If

        Dim bBiggieSize As Boolean = (eMsgBoxStyle And MsgBoxStyle.AbortRetryIgnore) <> 0
        If bBiggieSize = True Then
            eMsgBoxStyle = eMsgBoxStyle Xor MsgBoxStyle.AbortRetryIgnore
        End If

        'frmMsgBox initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eMsgBox
            .ControlName = "frmMsgBox"

            If bBiggieSize = True Then
                .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
                .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
                .Width = 512
                .Height = 256 '148
            Else
                .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
                .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 64
                .Width = 256
                .Height = 128
            End If

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = Me.Width - 10
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = sTitle
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
            .Width = Me.Width - 2
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'txtBody initial props
        txtBody = New UITextBox(oUILib)
        With txtBody
            .ControlName = "txtBody"
            .Left = 5
            .Top = 30
            .Width = Me.Width - 10
            If bBiggieSize = True Then
                .Height = 191 '65
            Else
                .Height = 65
            End If
            .Enabled = True
            .Visible = True
            .Caption = sMsg
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .BorderColor = muSettings.InterfaceBorderColor
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .MultiLine = True
            .Locked = True
        End With
        Me.AddChild(CType(txtBody, UIControl))

        'We'll accept a few of these for now, but eventually we should expand this...
        If eMsgBoxStyle = MsgBoxStyle.OkOnly Then
            ReDim btnCmd(0)
            btnCmd(0) = New UIButton(oUILib)
            With btnCmd(0)
                .ControlName = "btnOK"
                .Left = (Me.Width \ 2) - 100
                .Top = txtBody.Top + txtBody.Height + 5
                .Width = 100
                .Height = 24
                .Visible = True
                .Enabled = True
                .Caption = "OK"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnCmd(0), UIControl))
        ElseIf eMsgBoxStyle = MsgBoxStyle.OkCancel Then
            ReDim btnCmd(1)
            btnCmd(0) = New UIButton(oUILib)
            With btnCmd(0)
                .ControlName = "btnOK"
                .Left = (Me.Width \ 2) - 100
                .Top = txtBody.Top + txtBody.Height + 5
                .Width = 100
                .Height = 24
                .Visible = True
                .Enabled = True
                .Caption = "OK"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnCmd(0), UIControl))

            btnCmd(1) = New UIButton(oUILib)
            With btnCmd(1)
                .ControlName = "btnCancel"
                .Left = (Me.Width \ 2) '- 100
                .Top = txtBody.Top + txtBody.Height + 5
                .Width = 100
                .Height = 24
                .Visible = True
                .Enabled = True
                .Caption = "Cancel"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnCmd(1), UIControl))

        ElseIf eMsgBoxStyle = MsgBoxStyle.YesNo Then
            ReDim btnCmd(1)
            btnCmd(0) = New UIButton(oUILib)
            With btnCmd(0)
                .ControlName = "btnYes"
                .Left = (Me.Width \ 2) - 100
                .Top = txtBody.Top + txtBody.Height + 5
                .Width = 100
                .Height = 24
                .Visible = True
                .Enabled = Not mbRequiresSignature
                .Caption = "Yes"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnCmd(0), UIControl))

            btnCmd(1) = New UIButton(oUILib)
            With btnCmd(1)
                .ControlName = "btnNo"
                .Left = (Me.Width \ 2) '+ 100
                .Top = txtBody.Top + txtBody.Height + 5
                .Width = 100
                .Height = 24
                .Visible = True
                .Enabled = True
                .Caption = "No"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnCmd(1), UIControl))
        End If

        If btnCmd Is Nothing = False Then
            For X As Int32 = 0 To btnCmd.GetUpperBound(0)
                AddHandler btnCmd(X).Click, AddressOf BtnClick
            Next X
        End If

        If mbRequiresSignature = True Then
            If bBiggieSize = True Then
                txtBody.Height = 160
            Else
                txtBody.Height = 35
            End If

            'lblSignature initial props
            lblSignature = New UILabel(oUILib)
            With lblSignature
                .ControlName = "lblSignature"
                .Left = txtBody.Left
                .Top = txtBody.Top + txtBody.Height + 5
                .Width = 300
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Type your player name here to confirm:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblSignature, UIControl))

            txtSignature = New UITextBox(oUILib)
            With txtSignature
                .ControlName = "txtSignature"
                .Left = lblSignature.Left + lblSignature.Width + 5
                .Top = txtBody.Top + txtBody.Height + 5
                .Width = (Me.Width - 5) - .Left
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .BorderColor = muSettings.InterfaceBorderColor
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(0, DrawTextFormat)
            End With
            Me.AddChild(CType(txtSignature, UIControl))


        End If

        'MyBase.moUILib.RemoveWindow(Me.ControlName)
        'MyBase.moUILib.AddWindow(Me)
        MyBase.moUILib.SetMsgBox(Me)
    End Sub

    Private Sub BtnClick(ByVal sName As String)
        Select Case sName.ToUpper
            Case "BTNNO"
                RaiseEvent DialogClosed(MsgBoxResult.No)
            Case "BTNYES"
                RaiseEvent DialogClosed(MsgBoxResult.Yes)
            Case "BTNOK"
                RaiseEvent DialogClosed(MsgBoxResult.Ok)
            Case "BTNCANCEL"
                RaiseEvent DialogClosed(MsgBoxResult.Cancel)
        End Select
        'MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.SetMsgBox(Nothing)
    End Sub

    Private Sub frmMsgBox_OnNewFrame() Handles Me.OnNewFrame
        Dim bButton1Enabled As Boolean = True

        If mbRequiresSignature = True Then
            bButton1Enabled = False
            If txtSignature Is Nothing = False Then
                Dim sVal As String = txtSignature.Caption
                If sVal Is Nothing = False Then
                    sVal = sVal.ToUpper
                    If goCurrentPlayer Is Nothing = False Then
                        If sVal = goCurrentPlayer.PlayerName.ToUpper Then
                            bButton1Enabled = True
                        End If
                    End If
                End If
            End If
        End If

        If btnCmd Is Nothing = False AndAlso btnCmd.GetUpperBound(0) > -1 Then
            If btnCmd(0) Is Nothing = False Then
                If btnCmd(0).Enabled <> bButton1Enabled Then btnCmd(0).Enabled = bButton1Enabled
            End If
        End If
    End Sub
End Class