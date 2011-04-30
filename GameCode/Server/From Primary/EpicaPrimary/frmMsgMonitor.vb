Public Class frmMsgMonitor

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        goMsgSys.moMonitor.FillListView(tvwData)
    End Sub
End Class