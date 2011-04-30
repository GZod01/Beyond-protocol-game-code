
Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Partial Class frmArenaConfig
    Private Class fraDMArena
        Inherits fraArenaSubFrame

        Private lblRounds As UILabel
        Private lblKillGoal As UILabel
        Private lblTimeLimit As UILabel
        Private lblRespawns As UILabel
        Private lblRespawnDelay As UILabel
        Private lblSeconds As UILabel

        Private WithEvents txtRounds As UITextBox
        Private WithEvents txtKillGoal As UITextBox
        Private WithEvents txtTimeLimit As UITextBox
        Private txtRespawnDelay As UITextBox

        Private WithEvents hscrRounds As UIScrollBar
        Private WithEvents hscrKillGoal As UIScrollBar
        Private WithEvents hscrTimeLimit As UIScrollBar
        Private WithEvents hscrRespawns As UIScrollBar

        Private mbLoading As Boolean = True

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraDMArena initial props
            With Me
                .ControlName = "fraDMArena"
                .Left = 53
                .Top = 96
                .Width = 500
                .Height = 115
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .BorderLineWidth = 2
                .Caption = "Advanced Configuration"
            End With

            'lblRounds initial props
            lblRounds = New UILabel(oUILib)
            With lblRounds
                .ControlName = "lblRounds"
                .Left = 10
                .Top = 15
                .Width = 49
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Rounds:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblRounds, UIControl))

            'txtRounds initial props
            txtRounds = New UITextBox(oUILib)
            With txtRounds
                .ControlName = "txtRounds"
                .Left = 100
                .Top = 15
                .Width = 40
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "1"
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 2
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
            End With
            Me.AddChild(CType(txtRounds, UIControl))

            'hscrRounds initial props
            hscrRounds = New UIScrollBar(oUILib, False)
            With hscrRounds
                .ControlName = "hscrRounds"
                .Left = 150
                .Top = 16
                .Width = 100
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 1
                .MaxValue = 99
                .MinValue = 1
                .SmallChange = 1
                .LargeChange = 5
                .ReverseDirection = False
            End With
            Me.AddChild(CType(hscrRounds, UIControl))

            'lblKillGoal initial props
            lblKillGoal = New UILabel(oUILib)
            With lblKillGoal
                .ControlName = "lblKillGoal"
                .Left = 10
                .Top = 45
                .Width = 52
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Kill Goal:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblKillGoal, UIControl))

            'txtKillGoal initial props
            txtKillGoal = New UITextBox(oUILib)
            With txtKillGoal
                .ControlName = "txtKillGoal"
                .Left = 100
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
                .bNumericOnly = True
            End With
            Me.AddChild(CType(txtKillGoal, UIControl))

            'hscrKillGoal initial props
            hscrKillGoal = New UIScrollBar(oUILib, False)
            With hscrKillGoal
                .ControlName = "hscrKillGoal"
                .Left = 150
                .Top = 46
                .Width = 100
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 1
                .MaxValue = 999
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 5
                .ReverseDirection = False
            End With
            Me.AddChild(CType(hscrKillGoal, UIControl))

            'lblTimeLimit initial props
            lblTimeLimit = New UILabel(oUILib)
            With lblTimeLimit
                .ControlName = "lblTimeLimit"
                .Left = 10
                .Top = 75
                .Width = 65
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Time Limit:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTimeLimit, UIControl))

            'txtTimeLimit initial props
            txtTimeLimit = New UITextBox(oUILib)
            With txtTimeLimit
                .ControlName = "txtTimeLimit"
                .Left = 100
                .Top = 75
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
                .bNumericOnly = True
            End With
            Me.AddChild(CType(txtTimeLimit, UIControl))

            'hscrTimeLimit initial props
            hscrTimeLimit = New UIScrollBar(oUILib, False)
            With hscrTimeLimit
                .ControlName = "hscrTimeLimit"
                .Left = 150
                .Top = 76
                .Width = 100
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 1
                .MaxValue = 999
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 5
                .ReverseDirection = False
            End With
            Me.AddChild(CType(hscrTimeLimit, UIControl))

            'lblRespawns initial props
            lblRespawns = New UILabel(oUILib)
            With lblRespawns
                .ControlName = "lblRespawns"
                .Left = 265
                .Top = 15
                .Width = 113
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Respawns: None"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblRespawns, UIControl))

            'hscrRespawns initial props
            hscrRespawns = New UIScrollBar(oUILib, False)
            With hscrRespawns
                .ControlName = "hscrRespawns"
                .Left = 385
                .Top = 16
                .Width = 100
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 0
                .MaxValue = 99
                .MinValue = -1
                .SmallChange = 1
                .LargeChange = 5
                .ReverseDirection = False
            End With
            Me.AddChild(CType(hscrRespawns, UIControl))

            'lblRespawnDelay initial props
            lblRespawnDelay = New UILabel(oUILib)
            With lblRespawnDelay
                .ControlName = "lblRespawnDelay"
                .Left = 265
                .Top = 45
                .Width = 113
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Respawn Delay:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblRespawnDelay, UIControl))

            'txtRespawnDelay initial props
            txtRespawnDelay = New UITextBox(oUILib)
            With txtRespawnDelay
                .ControlName = "txtRespawnDelay"
                .Left = 385
                .Top = 45
                .Width = 47
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
                .MaxLength = 0
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
            End With
            Me.AddChild(CType(txtRespawnDelay, UIControl))

            'lblSeconds initial props
            lblSeconds = New UILabel(oUILib)
            With lblSeconds
                .ControlName = "lblSeconds"
                .Left = 440
                .Top = 45
                .Width = 52
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "seconds"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblSeconds, UIControl))

            mbLoading = False
        End Sub

        Public Overrides Sub SetFromArena(ByRef oArena As Arena)
            'Set our values from the arena object
            If oArena.yGameMode = eyGameMode.eDeathMatch Then
                With CType(oArena, DM_Arena)
                    hscrRounds.Value = .lRounds
                    hscrKillGoal.Value = .lKillGoal
                    hscrTimeLimit.Value = .lRoundTimeLimit
                    hscrRespawns.Value = .lRespawnLimit
                    txtRespawnDelay.Caption = .lRespawnDelay.ToString
                End With
            End If
        End Sub

        Private Sub hscrKillGoal_ValueChange() Handles hscrKillGoal.ValueChange
            If mbLoading = True Then Return
            mbLoading = True
            txtKillGoal.Caption = hscrKillGoal.Value.ToString
            mbLoading = False
        End Sub

        Private Sub hscrRespawns_ValueChange() Handles hscrRespawns.ValueChange
            If hscrRespawns.Value = -1 Then
                lblRespawns.Caption = "Respawns: Infinite"
            ElseIf hscrRespawns.Value = 0 Then
                lblRespawns.Caption = "Respawns: None"
            Else
                lblRespawns.Caption = "Respawns: " & hscrRespawns.Value.ToString
            End If
        End Sub

        Private Sub hscrRounds_ValueChange() Handles hscrRounds.ValueChange
            If mbLoading = True Then Return
            mbLoading = True
            txtRounds.Caption = hscrRounds.Value.ToString
            mbLoading = False
        End Sub

        Private Sub hscrTimeLimit_ValueChange() Handles hscrTimeLimit.ValueChange
            If mbLoading = True Then Return
            mbLoading = True
            txtTimeLimit.Caption = hscrTimeLimit.Value.ToString
            mbLoading = False
        End Sub

        Private Sub txtKillGoal_TextChanged() Handles txtKillGoal.TextChanged
            If mbLoading = True Then Return
            mbLoading = True
            Dim lVal As Int32 = CInt(Val(txtKillGoal.Caption))
            With hscrKillGoal
                lVal = Math.Min(Math.Max(.MinValue, lVal), .MaxValue)
                .Value = lVal
            End With
            mbLoading = False
        End Sub

        Private Sub txtRounds_TextChanged() Handles txtRounds.TextChanged
            If mbLoading = True Then Return
            mbLoading = True
            Dim lVal As Int32 = CInt(Val(txtRounds.Caption))
            With hscrRounds
                lVal = Math.Min(Math.Max(.MinValue, lVal), .MaxValue)
                .Value = lVal
            End With
            mbLoading = False
        End Sub

        Private Sub txtTimeLimit_TextChanged() Handles txtTimeLimit.TextChanged
            If mbLoading = True Then Return
            mbLoading = True
            Dim lVal As Int32 = CInt(Val(txtTimeLimit.Caption))
            With hscrTimeLimit
                lVal = Math.Min(Math.Max(.MinValue, lVal), .MaxValue)
                .Value = lVal
            End With
            mbLoading = False
        End Sub

        Public Overrides Function AdvancedConfigMsgAppend() As Byte()
            Dim yMsg(AdvancedConfigMsgLen() - 1) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(hscrRounds.Value).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(hscrKillGoal.Value).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(hscrTimeLimit.Value).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(hscrRespawns.Value).CopyTo(yMsg, lPos) : lPos += 4
            Dim lVal As Int32 = CInt(Val(txtRespawnDelay.Caption))
            System.BitConverter.GetBytes(lVal).CopyTo(yMsg, lPos) : lPos += 4
            Return yMsg
        End Function

        Public Overrides Function AdvancedConfigMsgLen() As Integer
            Return 20
        End Function

        Public Overrides Function AdvancedConfigValid() As Boolean
            If txtRespawnDelay.Caption Is Nothing OrElse IsNumeric(txtRespawnDelay.Caption) = False Then
                MyBase.moUILib.AddNotification("Invalid Respawn Delay.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lVal As Int32 = CInt(Val(txtRespawnDelay.Caption))
            If lVal < 0 Then
                MyBase.moUILib.AddNotification("Invalid Respawn Delay.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Return True
        End Function
    End Class
End Class
