Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmSpawnChoice
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private txtDisplay As UITextBox
    Private WithEvents btnAccept As UIButton
    Private WithEvents btnReject As UIButton

    Private myChoice As Byte

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmSpawnChoice initial props
        With Me
            .ControlName = "frmSpawnChoice"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            .Width = 255
            .Height = 255
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = False
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 0
            .Top = 3
            .Width = 255
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "READ CAREFULLY"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 22
            .Width = 255
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'txtDisplay initial props
        txtDisplay = New UITextBox(oUILib)
        With txtDisplay
            .ControlName = "txtDisplay"
            .Left = 5
            .Top = 27
            .Width = 245
            .Height = 190
            .Enabled = True
            .Visible = True
            .Caption = "You have recently upgraded your account to a full subscription. This gives you many new features such as special project research. It also grants you this single opportunity to spawn into a Spawn System (S). You only receive this option this one time. It is an important decision." & vbCrLf & vbCrLf & _
                "If you click Accept, all of your units, facilities, death budget and political relationships will be reset. You will be placed into a Spawn system with other players who have recently joined the subscription community. You will receive some protection from the veteran players." & vbCrLf & vbCrLf & _
                "If you click Refuse, you will continue playing just like you are now and nothing will change."
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .Locked = True
            .MultiLine = True
        End With
        Me.AddChild(CType(txtDisplay, UIControl))

        'btnAccept initial props
        btnAccept = New UIButton(oUILib)
        With btnAccept
            .ControlName = "btnAccept"
            .Left = 5
            .Top = 225
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Accept"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAccept, UIControl))

        'btnReject initial props
        btnReject = New UIButton(oUILib)
        With btnReject
            .ControlName = "btnReject"
            .Left = 150
            .Top = 225
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Reject"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnReject, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        goUILib.RetainTooltip = False
    End Sub

    Private Sub btnAccept_Click(ByVal sName As String) Handles btnAccept.Click
        myChoice = 1
        Me.Enabled = False

        Dim ofrmMsgBox As New frmMsgBox(goUILib, "You have chosen to spawn into a Spawn System. By clicking OK you acknowledge that you understand that all of your assets will be reset in order to provide a fair balance within the Spawn System.", MsgBoxStyle.OkCancel, "Confirm Choice")
        AddHandler ofrmMsgBox.DialogClosed, AddressOf MsgBox_Closed
    End Sub

    Private Sub btnReject_Click(ByVal sName As String) Handles btnReject.Click
        myChoice = 2
        Me.Enabled = False

        Dim ofrmMsgBox As New frmMsgBox(goUILib, "You have chosen to refuse spawning into a Spawn System. By clicking OK you acknowledge that you understand this option will not be given again.", MsgBoxStyle.OkCancel, "Confirm Choice")
        AddHandler ofrmMsgBox.DialogClosed, AddressOf MsgBox_Closed
    End Sub

    Private Sub MsgBox_Closed(ByVal yRes As MsgBoxResult)
        If yRes = MsgBoxResult.Ok Then
            'Ok, send it
            Dim yMsg(6) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRespawnSelection).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yMsg, 2)
            yMsg(6) = myChoice
            goUILib.SendMsgToPrimary(yMsg)

            goUILib.AddNotification("You will need to log in once more.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            gfrmMain.CancelLogin()

            MyBase.moUILib.RemoveWindow(Me.ControlName)
            Return
        Else
            myChoice = 0
        End If
        Me.Enabled = True
    End Sub
End Class