<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChannelConfig
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtChannelName = New System.Windows.Forms.TextBox
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkPublic = New System.Windows.Forms.CheckBox
        Me.btnUpdate = New System.Windows.Forms.Button
        Me.grpAdmin = New System.Windows.Forms.GroupBox
        Me.lstMembers = New System.Windows.Forms.ListBox
        Me.btnAdmin = New System.Windows.Forms.Button
        Me.btnKick = New System.Windows.Forms.Button
        Me.btnInvite = New System.Windows.Forms.Button
        Me.txtInvite = New System.Windows.Forms.TextBox
        Me.grpAdmin.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Channel Name:"
        '
        'txtChannelName
        '
        Me.txtChannelName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtChannelName.Location = New System.Drawing.Point(98, 6)
        Me.txtChannelName.Name = "txtChannelName"
        Me.txtChannelName.Size = New System.Drawing.Size(128, 20)
        Me.txtChannelName.TabIndex = 1
        '
        'txtPassword
        '
        Me.txtPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPassword.Location = New System.Drawing.Point(98, 32)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(128, 20)
        Me.txtPassword.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Password:"
        '
        'chkPublic
        '
        Me.chkPublic.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.chkPublic.AutoSize = True
        Me.chkPublic.Location = New System.Drawing.Point(52, 58)
        Me.chkPublic.Name = "chkPublic"
        Me.chkPublic.Size = New System.Drawing.Size(155, 17)
        Me.chkPublic.TabIndex = 4
        Me.chkPublic.Text = "Channel Viewable to Public"
        Me.chkPublic.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnUpdate.Location = New System.Drawing.Point(47, 81)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(144, 23)
        Me.btnUpdate.TabIndex = 5
        Me.btnUpdate.Text = "Update Configuration"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'grpAdmin
        '
        Me.grpAdmin.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpAdmin.Controls.Add(Me.btnKick)
        Me.grpAdmin.Controls.Add(Me.btnAdmin)
        Me.grpAdmin.Controls.Add(Me.lstMembers)
        Me.grpAdmin.Location = New System.Drawing.Point(12, 110)
        Me.grpAdmin.Name = "grpAdmin"
        Me.grpAdmin.Size = New System.Drawing.Size(214, 168)
        Me.grpAdmin.TabIndex = 6
        Me.grpAdmin.TabStop = False
        Me.grpAdmin.Text = "Channel Members"
        '
        'lstMembers
        '
        Me.lstMembers.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstMembers.FormattingEnabled = True
        Me.lstMembers.Location = New System.Drawing.Point(6, 19)
        Me.lstMembers.Name = "lstMembers"
        Me.lstMembers.Size = New System.Drawing.Size(202, 108)
        Me.lstMembers.TabIndex = 0
        '
        'btnAdmin
        '
        Me.btnAdmin.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAdmin.Location = New System.Drawing.Point(6, 133)
        Me.btnAdmin.Name = "btnAdmin"
        Me.btnAdmin.Size = New System.Drawing.Size(75, 23)
        Me.btnAdmin.TabIndex = 1
        Me.btnAdmin.Text = "Make Admin"
        Me.btnAdmin.UseVisualStyleBackColor = True
        '
        'btnKick
        '
        Me.btnKick.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnKick.Location = New System.Drawing.Point(133, 133)
        Me.btnKick.Name = "btnKick"
        Me.btnKick.Size = New System.Drawing.Size(75, 23)
        Me.btnKick.TabIndex = 2
        Me.btnKick.Text = "Kick"
        Me.btnKick.UseVisualStyleBackColor = True
        '
        'btnInvite
        '
        Me.btnInvite.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInvite.Location = New System.Drawing.Point(151, 284)
        Me.btnInvite.Name = "btnInvite"
        Me.btnInvite.Size = New System.Drawing.Size(75, 23)
        Me.btnInvite.TabIndex = 3
        Me.btnInvite.Text = "Invite"
        Me.btnInvite.UseVisualStyleBackColor = True
        '
        'txtInvite
        '
        Me.txtInvite.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInvite.Location = New System.Drawing.Point(12, 286)
        Me.txtInvite.Name = "txtInvite"
        Me.txtInvite.Size = New System.Drawing.Size(133, 20)
        Me.txtInvite.TabIndex = 7
        '
        'frmChannelConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(238, 315)
        Me.Controls.Add(Me.txtInvite)
        Me.Controls.Add(Me.btnInvite)
        Me.Controls.Add(Me.grpAdmin)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.chkPublic)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtChannelName)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmChannelConfig"
        Me.Text = "Channel Configuration"
        Me.grpAdmin.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtChannelName As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkPublic As System.Windows.Forms.CheckBox
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents grpAdmin As System.Windows.Forms.GroupBox
    Friend WithEvents btnKick As System.Windows.Forms.Button
    Friend WithEvents btnAdmin As System.Windows.Forms.Button
    Friend WithEvents lstMembers As System.Windows.Forms.ListBox
    Friend WithEvents btnInvite As System.Windows.Forms.Button
    Friend WithEvents txtInvite As System.Windows.Forms.TextBox
End Class
