Public Class frmMain

    Private moSB As System.Text.StringBuilder

    Public Sub AddEventLine(ByVal sLine As String)
        If moSB Is Nothing Then moSB = New System.Text.StringBuilder

        If moSB.Length > 1000000 Then moSB.Remove(500000, moSB.Length - 500000)
        moSB.Insert(0, sLine & vbCrLf, 1)

        If txtEvents.InvokeRequired() = True Then
            txtEvents.Invoke(New doUpdate(AddressOf delegateDoUpdate), moSB.ToString)
        Else
            txtEvents.Text = moSB.ToString
        End If
    End Sub

    Private Delegate Sub doUpdate(ByVal sEvent As String)
    Private Sub delegateDoUpdate(ByVal sText As String)
        txtEvents.Text = sText
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        If btnClose.Text.ToLower = "confirm" Then
            'close me down
            gbRunning = False
        Else : btnClose.Text = "Confirm"
        End If
    End Sub

    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        CloseConn()
    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Form.CheckForIllegalCrossThreadCalls = False
        gfrmMain = Me
    End Sub

    Private Sub frmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If ParseCommandLine() = False Then
            Me.AddEventLine("Unable to parse command line, exiting...")
            Return
        End If

        Dim oThread As Threading.Thread = New Threading.Thread(AddressOf MainLoop)
        oThread.Start()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        gbDisabled = CheckBox1.Checked
    End Sub

    Private Function ParseCommandLine() As Boolean
        Try
            Dim sValue As String = My.Application.CommandLineArgs(0)
            If IsNumeric(sValue) = False Then
                MsgBox("Unable to determine Box Operator ID in Command Line")
                Return False
            End If
            glBoxOperatorID = CInt(Val(sValue))

            gsOperatorIP = My.Application.CommandLineArgs(1)
            glOperatorPort = CInt(Val(My.Application.CommandLineArgs(2)))
            Return True
        Catch ex As Exception
            MsgBox("ParseCommandLine: " & ex.Message)
            Return False
        End Try
    End Function

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		Try
			GNSMgr.GetGNSMgr().ForceReloadTemplates()
			Me.AddEventLine("Force reload finished")
		Catch ex As Exception
			Me.AddEventLine("ERROR: " & ex.Message)
		End Try
	End Sub

    Private Sub btnResetLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetLog.Click
        ChatManager.ResetLog()
    End Sub
End Class
