Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmGuildMain
    'Interface created from Interface Builder
    Public Class fra_Finances
        Inherits guildframe

        Private lblBalance As UILabel
        Private lblPrevTaxIncome As UILabel
        Private lblTransLog As UILabel
        Private lstTransLog As UIListBox

        Private Structure GuildTransLog
            Public lPlayerID As Int32
            Public sDateTime As String
            Public blAmount As Int64
            Public blBalance As Int64

            Public Function GetTransLogString() As String
                Return sDateTime.PadRight(20, " "c) & GetCacheObjectValue(lPlayerID, ObjectType.ePlayer).PadRight(20, " "c) & blAmount.ToString("#,##0").PadLeft(20, " "c) & blBalance.ToString("#,##0").PadLeft(20, " "c)
            End Function
        End Structure
        Private Shared muTransLog() As GuildTransLog
        Private Shared mlTransLogUB As Int32 = -1
        Private Shared mblLastPrevTaxIncome As Int64 = 0

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fra_Finances initial props
            With Me
                .lWindowMetricID = BPMetrics.eWindow.eGuildMain_Finances
                .ControlName = "fra_Finances"
                .Left = 15
                .Top = ml_CONTENTS_TOP
                .Width = 128
                .Height = 128
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 2
                .Moveable = False
            End With

            'lblBalance initial props
            lblBalance = New UILabel(oUILib)
            With lblBalance
                .ControlName = "lblBalance"
                .Left = 10
                .Top = 10
                .Width = 300
                .Height = 16
                .Enabled = True
                .Visible = True
                .Caption = "Treasury Balance: 0"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBalance, UIControl))

            'lblPrevTaxIncome initial props
            lblPrevTaxIncome = New UILabel(oUILib)
            With lblPrevTaxIncome
                .ControlName = "lblPrevTaxIncome"
                .Left = 10
                .Top = 30
                .Width = 300
                .Height = 16
                .Enabled = True
                .Visible = True
                .Caption = "Previous Tax Income: 0"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPrevTaxIncome, UIControl))

            'lblTransLog initial props
            lblTransLog = New UILabel(oUILib)
            With lblTransLog
                .ControlName = "lblTransLog"
                .Left = 10
                .Top = 70
                .Width = 122
                .Height = 16
                .Enabled = True
                .Visible = True
                .Caption = "Transaction Log:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTransLog, UIControl))

            'lstTransLog initial props
            lstTransLog = New UIListBox(oUILib)
            With lstTransLog
                .ControlName = "lstTransLog"
                .Left = 10
                .Top = 90
                .Width = 775
                .Height = 441
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Courier New", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            End With
            Me.AddChild(CType(lstTransLog, UIControl))

            Dim yMsg(1) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildTreasury).CopyTo(yMsg, 0)
            MyBase.moUILib.SendMsgToPrimary(yMsg)
        End Sub

        Public Overrides Sub NewFrame()
            Try
                If mlTransLogUB <> lstTransLog.ListCount - 1 Then
                    FillWholeList()
                Else
                    For X As Int32 = 0 To mlTransLogUB
                        Dim sVal As String = muTransLog(X).GetTransLogString()
                        If lstTransLog.List(X) <> sVal Then lstTransLog.List(X) = sVal
                    Next X
                End If

                Dim sTemp As String = "Treasury Balance: " & goCurrentPlayer.oGuild.blTreasury.ToString("#,##0")
                If lblBalance.Caption <> sTemp Then lblBalance.Caption = sTemp
                sTemp = "Previous Tax Income: " & mblLastPrevTaxIncome.ToString("#,##0")
                If lblPrevTaxIncome.Caption <> sTemp Then lblPrevTaxIncome.Caption = sTemp
            Catch
            End Try
        End Sub

        Private Sub FillWholeList()
            Try
                lstTransLog.Clear()
                lstTransLog.sHeaderRow = "Transaction Date".PadRight(20, " "c) & "Player".PadRight(20, " "c) & "Amount".PadLeft(20, " "c) & "Balance".PadLeft(20, " "c)

                For X As Int32 = 0 To mlTransLogUB
                    lstTransLog.AddItem(muTransLog(X).GetTransLogString(), muTransLog(X).lPlayerID = glPlayerID)
                Next X
            Catch
            End Try
        End Sub

        Public Overrides Sub RenderEnd()
            '
        End Sub

        Public Shared Sub HandleUpdateGuildTreasury(ByVal yData() As Byte)
            Dim lPos As Int32 = 2
            Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            mblLastPrevTaxIncome = System.BitConverter.ToInt64(yData, lPos) : lPos += 8

            Dim lTmpUB As Int32 = lCnt - 1
            Dim uTmp(lTmpUB) As GuildTransLog

            For X As Int32 = 0 To lTmpUB
                With uTmp(X)
                    .lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim lTransDate As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    'Ok, convert our date
                    Dim dtVal As Date = Date.SpecifyKind(GetDateFromNumber(lTransDate), DateTimeKind.Utc).ToLocalTime
                    .sDateTime = dtVal.ToString("MM/dd/yyyy HH:mm")

                    .blAmount = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                    .blBalance = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                End With
            Next X

            mlTransLogUB = -1
            muTransLog = uTmp
            mlTransLogUB = lTmpUB
        End Sub
    End Class
End Class
