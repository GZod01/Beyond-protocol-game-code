<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.txtEvents = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnResetLog = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtEvents
        '
        Me.txtEvents.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEvents.Location = New System.Drawing.Point(12, 12)
        Me.txtEvents.Multiline = True
        Me.txtEvents.Name = "txtEvents"
        Me.txtEvents.ReadOnly = True
        Me.txtEvents.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtEvents.Size = New System.Drawing.Size(456, 208)
        Me.txtEvents.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(393, 226)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close Server"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(306, 230)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(67, 17)
        Me.CheckBox1.TabIndex = 3
        Me.CheckBox1.Text = "Disabled"
        Me.CheckBox1.UseVisualStyleBackColor = True
        Me.CheckBox1.Checked = gb_IS_TEST_SERVER
        If gb_IS_TEST_SERVER = True Then Me.CheckBox1.CheckState = CheckState.Checked Else Me.CheckBox1.CheckState = CheckState.Unchecked
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(225, 226)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Reload GNS"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnResetLog
        '
        Me.btnResetLog.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnResetLog.Location = New System.Drawing.Point(12, 226)
        Me.btnResetLog.Name = "btnResetLog"
        Me.btnResetLog.Size = New System.Drawing.Size(75, 23)
        Me.btnResetLog.TabIndex = 5
        Me.btnResetLog.Text = "Reset Log"
        Me.btnResetLog.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(480, 261)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnResetLog)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtEvents)
        Me.Name = "frmMain"
        Me.Text = "Email Server"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtEvents As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnResetLog As System.Windows.Forms.Button

End Class
