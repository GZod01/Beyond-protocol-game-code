Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmBudget
    'Interface created from Interface Builder
    Public Class fraBudgetSummary
        Inherits UIWindow

        Private lblRevenue As UILabel
        Private lblExpense As UILabel
        Private lblCashFlow As UILabel
        Private lblChanged As UILabel
        Private lnDiv1 As UILine

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraBudgetSummary initial props
            With Me
                .ControlName = "fraBudgetSummary"
                .Left = 232
                .Top = 71
                .Width = 310
                .Height = 80
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 1
                .Moveable = False
                .Caption = "Budget Summary"
            End With

            'lblRevenue initial props
            lblRevenue = New UILabel(oUILib)
            With lblRevenue
                .ControlName = "lblRevenue"
                .Left = 10
                .Top = 11 ' old 15
                .Width = 300
                .Height = 12
                .Enabled = True
                .Visible = True
                .Caption = "Revenue:  00,000,000,000,000,000,000"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblRevenue, UIControl))

            'lblExpense initial props
            lblExpense = New UILabel(oUILib)
            With lblExpense
                .ControlName = "lblExpense"
                .Left = 10
                .Top = 24 ' old 30
                .Width = 300
                .Height = 12
                .Enabled = True
                .Visible = True
                .Caption = "Expense:  00,000,000,000,000,000,000"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblExpense, UIControl))

            'lblCashFlow initial props
            lblCashFlow = New UILabel(oUILib)
            With lblCashFlow
                .ControlName = "lblCashFlow"
                .Left = 10
                .Top = 48 ' old 60
                .Width = 300
                .Height = 12
                .Enabled = True
                .Visible = True
                .Caption = "Cashflow: -0,000,000"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCashFlow, UIControl))

            'lblChanged
            lblChanged = New UILabel(oUILib)
            With lblChanged
                .ControlName = "lblChanged"
                .Left = 10
                .Top = 61
                .Width = 300
                .Height = 12
                .Enabled = True
                .Visible = True
                .Caption = "Changed:  -0,000,000"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblChanged, UIControl))

            'lnDiv1 initial props
            lnDiv1 = New UILine(oUILib)
            With lnDiv1
                .ControlName = "lnDiv1"
                .Left = 0
                .Top = 42 ' old 50
                .Width = 310
                .Height = 0
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(lnDiv1, UIControl))
        End Sub

		Public Sub SetValues(ByVal blRev As Int64, ByVal blExp As Int64, ByVal blAgentMaint As Int64)
			blExp += blAgentMaint
            Dim blCash As Int64 = blRev - blExp '- blAgentMaint
            Static xblCashLast As Int64 = blCash
            Dim blChanged As Int64 = blCash - xblCashLast
            xblCashLast = blCash

			lblRevenue.Caption = "Revenue:  " & blRev.ToString("#,###").PadLeft(26, " "c)
			lblExpense.Caption = "Expense:  " & blExp.ToString("#,###").PadLeft(26, " "c)
            lblCashFlow.Caption = "Cashflow: " & blCash.ToString("#,###").PadLeft(26, " "c)
            lblChanged.Caption = "Changed:  " & blChanged.ToString("#,###").PadLeft(26, " "c)

            If blCash < 0 Then
                lblCashFlow.ForeColor = Color.FromArgb(255, 255, 0, 0)
            Else : lblCashFlow.ForeColor = Color.FromArgb(255, 0, 255, 0)
            End If

            If blChanged > 0 Then
                lblChanged.ForeColor = Color.FromArgb(255, 0, 255, 0)
            ElseIf blChanged < 0 Then
                lblChanged.ForeColor = Color.FromArgb(255, 255, 0, 0)
            Else : lblChanged.ForeColor = muSettings.InterfaceBorderColor
            End If
		End Sub
    End Class
End Class