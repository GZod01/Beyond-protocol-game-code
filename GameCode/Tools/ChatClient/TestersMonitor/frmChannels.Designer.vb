<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChannels
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lstChannels = New System.Windows.Forms.ListBox
        Me.btnCreate = New System.Windows.Forms.Button
        Me.btnJoinLeave = New System.Windows.Forms.Button
        Me.btnAdmin = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lstChannels
        '
        Me.lstChannels.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstChannels.FormattingEnabled = True
        Me.lstChannels.Location = New System.Drawing.Point(12, 12)
        Me.lstChannels.Name = "lstChannels"
        Me.lstChannels.Size = New System.Drawing.Size(212, 173)
        Me.lstChannels.TabIndex = 0
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(12, 202)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(101, 23)
        Me.btnCreate.TabIndex = 1
        Me.btnCreate.Text = "Create New"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'btnJoinLeave
        '
        Me.btnJoinLeave.Location = New System.Drawing.Point(119, 202)
        Me.btnJoinLeave.Name = "btnJoinLeave"
        Me.btnJoinLeave.Size = New System.Drawing.Size(101, 23)
        Me.btnJoinLeave.TabIndex = 2
        Me.btnJoinLeave.Text = "Join Channel"
        Me.btnJoinLeave.UseVisualStyleBackColor = True
        '
        'btnAdmin
        '
        Me.btnAdmin.Location = New System.Drawing.Point(65, 231)
        Me.btnAdmin.Name = "btnAdmin"
        Me.btnAdmin.Size = New System.Drawing.Size(101, 23)
        Me.btnAdmin.TabIndex = 3
        Me.btnAdmin.Text = "Admin"
        Me.btnAdmin.UseVisualStyleBackColor = True
        '
        'frmChannels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(236, 265)
        Me.Controls.Add(Me.btnAdmin)
        Me.Controls.Add(Me.btnJoinLeave)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.lstChannels)
        Me.Name = "frmChannels"
        Me.Text = "Available Channels"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lstChannels As System.Windows.Forms.ListBox
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents btnJoinLeave As System.Windows.Forms.Button
    Friend WithEvents btnAdmin As System.Windows.Forms.Button
End Class
