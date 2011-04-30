Public Class frmChatTabProps

    Private moTab As ChatTab = Nothing
    Private moFrmMain As frmMain = Nothing

    Public Sub SetReturn(ByRef frm As frmMain)
        moFrmMain = frm
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        With moTab
            .sTabName = txtTabName.Text
            .sChannel = txtChannel.Text
            .sMessagePrefix = txtPrefix.Text

            Dim lFilter As Int32 = 0
            If chkLocal.Checked = True Then lFilter = lFilter Or ChatTab.ChatFilter.eLocalMessages
            If chkSysAdm.Checked = True Then lFilter = lFilter Or ChatTab.ChatFilter.eSysAdminMessages
            If chkChannel.Checked = True Then lFilter = lFilter Or ChatTab.ChatFilter.eChannelMessages
            If chkGuild.Checked = True Then lFilter = lFilter Or ChatTab.ChatFilter.eAllianceMessages
            If chkAlias.Checked = True Then lFilter = lFilter Or ChatTab.ChatFilter.eAliasChatMessage
            If chkPM.Checked = True Then lFilter = lFilter Or ChatTab.ChatFilter.ePMs
            If chkNotification.Checked = True Then lFilter = lFilter Or ChatTab.ChatFilter.eNotificationMessages

            .lFilter = lFilter
            If .oTab Is Nothing = False Then
                .oTab.Text = .sTabName
            End If
        End With
        If moFrmMain Is Nothing = False Then moFrmMain.ForceSaveTabs()
        Me.Close()
    End Sub

    Public Sub SetTab(ByRef oTab As ChatTab)
        moTab = oTab
        With moTab
            txtTabName.Text = .sTabName
            chkLocal.Checked = (.lFilter And ChatTab.ChatFilter.eLocalMessages) <> 0
            chkSysAdm.Checked = (.lFilter And ChatTab.ChatFilter.eSysAdminMessages) <> 0
            chkChannel.Checked = (.lFilter And ChatTab.ChatFilter.eChannelMessages) <> 0
            txtChannel.Text = .sChannel
            chkGuild.Checked = (.lFilter And ChatTab.ChatFilter.eAllianceMessages) <> 0
            chkAlias.Checked = (.lFilter And ChatTab.ChatFilter.eAliasChatMessage) <> 0
            chkPM.Checked = (.lFilter And ChatTab.ChatFilter.ePMs) <> 0
            chkNotification.Checked = (.lFilter And ChatTab.ChatFilter.eNotificationMessages) <> 0
            txtPrefix.Text = .sMessagePrefix
        End With
    End Sub
End Class