Public Class frmColors

    Public Sub DoInitialColors()
        btnAlias.BackColor = Color.Black : btnAlias.ForeColor = AliasChatColor
        btnChannel.BackColor = Color.Black : btnChannel.ForeColor = ChannelChatColor
        btnGuild.BackColor = Color.Black : btnGuild.ForeColor = GuildChatColor
        btnLocal.BackColor = Color.Black : btnLocal.ForeColor = LocalChatColor
        btnPM.BackColor = Color.Black : btnPM.ForeColor = PMChatColor
        btnSenate.BackColor = Color.Black : btnSenate.ForeColor = SenateChatColor
        btnSysAdm.BackColor = Color.Black : btnSysAdm.ForeColor = AlertChatColor
        btnBackground.BackColor = TextBoxBackColor : btnBackground.ForeColor = Color.White
    End Sub

    Public bDoSaveColors As Boolean = False

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        bDoSaveColors = True

        AliasChatColor = btnAlias.ForeColor
        ChannelChatColor = btnChannel.ForeColor
        GuildChatColor = btnGuild.ForeColor
        LocalChatColor = btnLocal.ForeColor
        PMChatColor = btnPM.ForeColor
        SenateChatColor = btnSenate.ForeColor
        AlertChatColor = btnSysAdm.ForeColor
        TextBoxBackColor = btnBackground.BackColor
        SaveSettings()
        Me.Hide()
    End Sub
    Private Sub btnAlias_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAlias.Click
        cdMain.Color = btnAlias.ForeColor
        If cdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            btnAlias.ForeColor = cdMain.Color
        End If
    End Sub
    Private Sub btnChannel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChannel.Click
        cdMain.Color = ChannelChatColor
        If cdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            ChannelChatColor = cdMain.Color
            btnChannel.ForeColor = cdMain.Color
        End If
    End Sub
    Private Sub btnGuild_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGuild.Click
        cdMain.Color = btnGuild.ForeColor
        If cdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            btnGuild.ForeColor = cdMain.Color
        End If
    End Sub
    Private Sub btnLocal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLocal.Click
        cdMain.Color = btnLocal.ForeColor
        If cdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            btnLocal.ForeColor = cdMain.Color
        End If
    End Sub
    Private Sub btnPM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPM.Click
        cdMain.Color = btnPM.ForeColor
        If cdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            btnPM.ForeColor = cdMain.Color
        End If
    End Sub
    Private Sub btnSenate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSenate.Click
        cdMain.Color = btnSenate.ForeColor
        If cdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            btnSenate.ForeColor = cdMain.Color
        End If
    End Sub
    Private Sub btnSysAdm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSysAdm.Click
        cdMain.Color = btnSysAdm.ForeColor
        If cdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            btnSysAdm.ForeColor = cdMain.Color
        End If
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Hide()
    End Sub
    Private Sub btnBackground_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackground.Click
        cdMain.Color = btnBackground.BackColor
        If cdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            btnBackground.BackColor = cdMain.Color
        End If
    End Sub
End Class