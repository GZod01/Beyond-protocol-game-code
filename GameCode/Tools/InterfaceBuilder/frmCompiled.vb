Public Class frmCompiled
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtResult = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtResult
        '
        Me.txtResult.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtResult.Location = New System.Drawing.Point(8, 8)
        Me.txtResult.Multiline = True
        Me.txtResult.Name = "txtResult"
        Me.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtResult.Size = New System.Drawing.Size(768, 576)
        Me.txtResult.TabIndex = 0
        Me.txtResult.Text = ""
        '
        'frmCompiled
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 589)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtResult})
        Me.Name = "frmCompiled"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmCompiled"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub SetText(ByVal sText As String)
        Me.txtResult.Text = sText
        Me.txtResult.SelectionLength = 0
    End Sub

End Class
