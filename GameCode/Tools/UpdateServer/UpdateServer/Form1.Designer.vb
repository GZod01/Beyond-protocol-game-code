<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.tmrCleanup = New System.Windows.Forms.Timer(Me.components)
        Me.Button1 = New System.Windows.Forms.Button
        Me.txtNotes = New System.Windows.Forms.TextBox
        Me.btnDownloaders = New System.Windows.Forms.Button
        Me.btnResetMsgs = New System.Windows.Forms.Button
        Me.btnDisconnectAll = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'tmrCleanup
        '
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(416, 246)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtNotes
        '
        Me.txtNotes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNotes.Location = New System.Drawing.Point(12, 12)
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtNotes.Size = New System.Drawing.Size(479, 228)
        Me.txtNotes.TabIndex = 1
        '
        'btnDownloaders
        '
        Me.btnDownloaders.Location = New System.Drawing.Point(12, 246)
        Me.btnDownloaders.Name = "btnDownloaders"
        Me.btnDownloaders.Size = New System.Drawing.Size(85, 23)
        Me.btnDownloaders.TabIndex = 2
        Me.btnDownloaders.Text = "Downloaders"
        Me.btnDownloaders.UseVisualStyleBackColor = True
        '
        'btnResetMsgs
        '
        Me.btnResetMsgs.Location = New System.Drawing.Point(103, 246)
        Me.btnResetMsgs.Name = "btnResetMsgs"
        Me.btnResetMsgs.Size = New System.Drawing.Size(97, 23)
        Me.btnResetMsgs.TabIndex = 3
        Me.btnResetMsgs.Text = "Reset All Msgs"
        Me.btnResetMsgs.UseVisualStyleBackColor = True
        '
        'btnDisconnectAll
        '
        Me.btnDisconnectAll.Location = New System.Drawing.Point(206, 246)
        Me.btnDisconnectAll.Name = "btnDisconnectAll"
        Me.btnDisconnectAll.Size = New System.Drawing.Size(97, 23)
        Me.btnDisconnectAll.TabIndex = 4
        Me.btnDisconnectAll.Text = "Disconnect All"
        Me.btnDisconnectAll.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(503, 281)
        Me.Controls.Add(Me.btnDisconnectAll)
        Me.Controls.Add(Me.btnResetMsgs)
        Me.Controls.Add(Me.btnDownloaders)
        Me.Controls.Add(Me.txtNotes)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Update Server"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tmrCleanup As System.Windows.Forms.Timer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtNotes As System.Windows.Forms.TextBox
    Friend WithEvents btnDownloaders As System.Windows.Forms.Button
    Friend WithEvents btnResetMsgs As System.Windows.Forms.Button
    Friend WithEvents btnDisconnectAll As System.Windows.Forms.Button

End Class
