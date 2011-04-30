Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmBudget
    'Interface created from Interface Builder
    Public Class fraDeathBudget
        Inherits UIWindow

        Private lblBalance As UILabel
        Public lblMax As UILabel
		Public txtAmount As UITextBox
        Public WithEvents btnDeposit As UIButton

        Private mlMaxValue As Int32

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraDeathBudget initial props
            With Me
                .ControlName = "fraDeathBudget"
                .Left = 409
                .Top = 130
                .Width = 215
                .Height = 82 '64
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .Caption = "Death Budget"
            End With

            'lblBalance initial props
            lblBalance = New UILabel(oUILib)
            With lblBalance
                .ControlName = "lblBalance"
                .Left = 5
                .Top = 10
                .Width = 163
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Balance: 0,000,000,000"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "This is an account of funds that your advisors" & vbCrLf & "will use to ensure our civilization never" & vbCrLf & "ceases to exist. Funds allocated to this account will" & vbCrLf & "be available to use as a contingency plan. When used, " & vbCrLf & "anything built with the funds will automatically complete."
            End With
            Me.AddChild(CType(lblBalance, UIControl))

            'lblMax initial props
            lblMax = New UILabel(oUILib)
            With lblMax
                .ControlName = "lblMax"
                .Left = 5
                .Top = 35
                .Width = 163
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Maximum: 0"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "This is an account of funds that your advisors" & vbCrLf & "will use to ensure our civilization never" & vbCrLf & "ceases to exist. Funds allocated to this account will" & vbCrLf & "be available to use as a contingency plan. When used, " & vbCrLf & "anything built with the funds will automatically complete."
            End With
            Me.AddChild(CType(lblMax, UIControl))

            'txtAmount initial props
            txtAmount = New UITextBox(oUILib)
            With txtAmount
                .ControlName = "txtAmount"
                .Left = 5
                .Top = 60 '35
                .Width = 100
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
                .MaxLength = 10
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(txtAmount, UIControl))

            'btnDeposit initial props
            btnDeposit = New UIButton(oUILib)
            With btnDeposit
                .ControlName = "btnDeposit"
                .Left = 110
                .Top = 57 '32
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Deposit"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnDeposit, UIControl))
        End Sub

        Public Sub SetDeathBalance(ByVal lValue As Int32)
            lblBalance.Caption = "Balance: " & lValue.ToString("#,##0")
        End Sub

        Public Sub SetMaximum(ByVal lMaxValue As Int32)
            mlMaxValue = lMaxValue
            lblMax.Caption = "Maximum: " & lMaxValue.ToString("#,##0")
            lblMax.ToolTipText = "This is based off of the total colonists within your empire." & vbCrLf & "For every colonist in your empire, you can put 100 credits into the budget"
        End Sub

        Private Sub btnDeposit_Click(ByVal sName As String) Handles btnDeposit.Click
            If gbAliased = True Then
                MyBase.moUILib.AddNotification("Unable to alter death budget as an aliased player.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            If goCurrentPlayer Is Nothing = False Then
                Dim blVal As Int64 = CLng(Val(txtAmount.Caption))
                Dim lVal As Int32 = 0
                If goCurrentPlayer.yPlayerPhase <> 0 Then
                    If blVal > mlMaxValue Then
                        MyBase.moUILib.AddNotification("That exceeds the maximum value for the death budget account.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                    lVal = CInt(Val(txtAmount.Caption))
                    If lVal < 1 Then
                        MyBase.moUILib.AddNotification("You must enter an amount to deposit.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                    If lVal > goCurrentPlayer.blCredits Then
                        MyBase.moUILib.AddNotification("You do not have enough credits to deposit that amount.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                ElseIf blVal < Int32.MaxValue Then
                    lVal = CInt(blVal)
                End If
                If lVal < 1 Then Return

                Dim yMsg(5) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eDeathBudgetDeposit).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(lVal).CopyTo(yMsg, lPos) : lPos += 4
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                MyBase.moUILib.AddNotification("Deposit request submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                txtAmount.Caption = "0"
            End If
        End Sub
    End Class
End Class