Public Class frmLogin

    Public Sub DoEnableControls(ByVal bVal As Boolean)
        btnCancel.Enabled = bVal
        btnOK.Enabled = bVal
        txtUsername.Enabled = bVal
        txtPassword.Enabled = bVal
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        DoEnableControls(False)

        gsUserName = txtUsername.Text
        gsPassword = txtPassword.Text

        If gsUserName.Trim = "" OrElse gsPassword.Trim = "" Then
            MsgBox("Enter a valid username and password.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Invalid Credentials")
            DoEnableControls(True)
            Return
        End If

        gsUserName = gsUserName.ToUpper
        gsPassword = gsPassword.ToUpper
        gfrmChat.LoadChatTabs()
        Me.Visible = False
        gfrmChat.ReConnect.Visible = True
        If gfrmChat.DoLogin() = False Then
            Me.Visible = True
            DoEnableControls(True)
            Me.Focus()
        End If

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        Process.GetCurrentProcess.Kill()
    End Sub

    Private Sub frmLogin_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        gfrmChat = New frmMain()
        gfrmChat.Show()
        gfrmChat.Panel1.Enabled = False
        AddHandler gfrmChat.LoginAttemptResult, AddressOf gfrmChat_LoginAttemptResult

        Me.Focus()
        Me.Left = (gfrmChat.Width \ 2) + gfrmChat.Left - (Me.Width \ 2)
        Me.Top = (gfrmChat.Height \ 2) + gfrmChat.Top - (Me.Height \ 2)
    End Sub

    Private Sub gfrmChat_LoginAttemptResult(ByVal bSuccess As Boolean)
        If bSuccess = True Then
            Me.Close()
            gfrmChat.Panel1.Enabled = True
            gfrmChat.ReConnect.Visible = False
        Else
            Me.Visible = True
            DoEnableControls(True)
            Me.Focus()
        End If
    End Sub
End Class
