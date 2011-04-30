Public Class frmSetRect
    Private mbSDown As Boolean = False
    Private mbLoading As Boolean = False

    Private mlX As Int32
    Private mlY As Int32

    Public Event SetRectLoc(ByVal lX As Int32, ByVal lY As Int32, ByVal lWidth As Int32, ByVal lHeight As Int32)

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        If mbSDown = True Then
            mlX = e.X
            mlY = e.Y
            mbLoading = True
            txtTop.Text = e.Y.ToString
            txtLeft.Text = e.X.ToString
            mbLoading = False
            RaiseEvent SetRectLoc(e.X, e.Y, CInt(Val(txtWidth.Text)), CInt(Val(txtHeight.Text)))
        End If
    End Sub

    Private Sub frmSetRect_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Control = True Then
            mbSDown = True
        End If
    End Sub

    Private Sub frmSetRect_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        If e.Control = False Then
            mbSDown = False
        End If
    End Sub

    Private Sub txtWidth_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWidth.TextChanged
        If mbLoading Then Return
        RaiseEvent SetRectLoc(mlX, mlY, CInt(Val(txtWidth.Text)), CInt(Val(txtHeight.Text)))
    End Sub

    Private Sub txtHeight_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHeight.TextChanged
        If mbLoading Then Return
        RaiseEvent SetRectLoc(mlX, mlY, CInt(Val(txtWidth.Text)), CInt(Val(txtHeight.Text)))
    End Sub

    Public Sub SetByRect(ByVal rcTemp As Rectangle)
        mbLoading = True
        txtWidth.Text = rcTemp.Width.ToString
        txtHeight.Text = rcTemp.Height.ToString
        mlX = rcTemp.X
        mlY = rcTemp.Y
        mbLoading = False
    End Sub

    Private Sub txtTop_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTop.TextChanged
        If mbLoading Then Return
        mlY = CInt(Val(txtTop.Text))
        RaiseEvent SetRectLoc(mlX, mlY, CInt(Val(txtWidth.Text)), CInt(Val(txtHeight.Text)))
    End Sub

    Private Sub txtLeft_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeft.TextChanged
        If mbLoading Then Return
        mlX = CInt(Val(txtLeft.Text))
        RaiseEvent SetRectLoc(mlX, mlY, CInt(Val(txtWidth.Text)), CInt(Val(txtHeight.Text)))
    End Sub
End Class