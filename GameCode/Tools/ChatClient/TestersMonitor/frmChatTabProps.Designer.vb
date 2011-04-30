<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChatTabProps
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
        Me.txtTabName = New System.Windows.Forms.TextBox
        Me.grpFilter = New System.Windows.Forms.GroupBox
        Me.chkLocal = New System.Windows.Forms.CheckBox
        Me.chkSysAdm = New System.Windows.Forms.CheckBox
        Me.chkChannel = New System.Windows.Forms.CheckBox
        Me.chkGuild = New System.Windows.Forms.CheckBox
        Me.chkAlias = New System.Windows.Forms.CheckBox
        Me.chkPM = New System.Windows.Forms.CheckBox
        Me.chkNotification = New System.Windows.Forms.CheckBox
        Me.txtPrefix = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtChannel = New System.Windows.Forms.TextBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.grpFilter.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Chat Tab Name:"
        '
        'txtTabName
        '
        Me.txtTabName.Location = New System.Drawing.Point(103, 6)
        Me.txtTabName.Name = "txtTabName"
        Me.txtTabName.Size = New System.Drawing.Size(122, 20)
        Me.txtTabName.TabIndex = 1
        '
        'grpFilter
        '
        Me.grpFilter.Controls.Add(Me.txtChannel)
        Me.grpFilter.Controls.Add(Me.chkNotification)
        Me.grpFilter.Controls.Add(Me.chkPM)
        Me.grpFilter.Controls.Add(Me.chkAlias)
        Me.grpFilter.Controls.Add(Me.chkGuild)
        Me.grpFilter.Controls.Add(Me.chkChannel)
        Me.grpFilter.Controls.Add(Me.chkSysAdm)
        Me.grpFilter.Controls.Add(Me.chkLocal)
        Me.grpFilter.Location = New System.Drawing.Point(15, 32)
        Me.grpFilter.Name = "grpFilter"
        Me.grpFilter.Size = New System.Drawing.Size(210, 181)
        Me.grpFilter.TabIndex = 2
        Me.grpFilter.TabStop = False
        Me.grpFilter.Text = "Filter"
        '
        'chkLocal
        '
        Me.chkLocal.AutoSize = True
        Me.chkLocal.Location = New System.Drawing.Point(6, 19)
        Me.chkLocal.Name = "chkLocal"
        Me.chkLocal.Size = New System.Drawing.Size(103, 17)
        Me.chkLocal.TabIndex = 0
        Me.chkLocal.Text = "Local Messages"
        Me.chkLocal.UseVisualStyleBackColor = True
        '
        'chkSysAdm
        '
        Me.chkSysAdm.AutoSize = True
        Me.chkSysAdm.Location = New System.Drawing.Point(6, 42)
        Me.chkSysAdm.Name = "chkSysAdm"
        Me.chkSysAdm.Size = New System.Drawing.Size(138, 17)
        Me.chkSysAdm.TabIndex = 1
        Me.chkSysAdm.Text = "System Admin Message"
        Me.chkSysAdm.UseVisualStyleBackColor = True
        '
        'chkChannel
        '
        Me.chkChannel.AutoSize = True
        Me.chkChannel.Location = New System.Drawing.Point(6, 65)
        Me.chkChannel.Name = "chkChannel"
        Me.chkChannel.Size = New System.Drawing.Size(116, 17)
        Me.chkChannel.TabIndex = 2
        Me.chkChannel.Text = "Channel Messages"
        Me.chkChannel.UseVisualStyleBackColor = True
        '
        'chkGuild
        '
        Me.chkGuild.AutoSize = True
        Me.chkGuild.Location = New System.Drawing.Point(6, 88)
        Me.chkGuild.Name = "chkGuild"
        Me.chkGuild.Size = New System.Drawing.Size(101, 17)
        Me.chkGuild.TabIndex = 3
        Me.chkGuild.Text = "Guild Messages"
        Me.chkGuild.UseVisualStyleBackColor = True
        '
        'chkAlias
        '
        Me.chkAlias.AutoSize = True
        Me.chkAlias.Location = New System.Drawing.Point(6, 111)
        Me.chkAlias.Name = "chkAlias"
        Me.chkAlias.Size = New System.Drawing.Size(99, 17)
        Me.chkAlias.TabIndex = 4
        Me.chkAlias.Text = "Alias Messages"
        Me.chkAlias.UseVisualStyleBackColor = True
        '
        'chkPM
        '
        Me.chkPM.AutoSize = True
        Me.chkPM.Location = New System.Drawing.Point(6, 134)
        Me.chkPM.Name = "chkPM"
        Me.chkPM.Size = New System.Drawing.Size(137, 17)
        Me.chkPM.TabIndex = 5
        Me.chkPM.Text = "Private Messages (tells)"
        Me.chkPM.UseVisualStyleBackColor = True
        '
        'chkNotification
        '
        Me.chkNotification.AutoSize = True
        Me.chkNotification.Location = New System.Drawing.Point(6, 157)
        Me.chkNotification.Name = "chkNotification"
        Me.chkNotification.Size = New System.Drawing.Size(130, 17)
        Me.chkNotification.TabIndex = 6
        Me.chkNotification.Text = "Notification Messages"
        Me.chkNotification.UseVisualStyleBackColor = True
        '
        'txtPrefix
        '
        Me.txtPrefix.Location = New System.Drawing.Point(137, 219)
        Me.txtPrefix.Name = "txtPrefix"
        Me.txtPrefix.Size = New System.Drawing.Size(88, 20)
        Me.txtPrefix.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 222)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(119, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Default Message Prefix:"
        '
        'txtChannel
        '
        Me.txtChannel.Location = New System.Drawing.Point(122, 63)
        Me.txtChannel.Name = "txtChannel"
        Me.txtChannel.Size = New System.Drawing.Size(82, 20)
        Me.txtChannel.TabIndex = 5
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(122, 245)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(103, 23)
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "Save and Close"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(12, 245)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(103, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmChatTabProps
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(236, 276)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtPrefix)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.grpFilter)
        Me.Controls.Add(Me.txtTabName)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmChatTabProps"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Chat Tab Properties"
        Me.grpFilter.ResumeLayout(False)
        Me.grpFilter.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTabName As System.Windows.Forms.TextBox
    Friend WithEvents grpFilter As System.Windows.Forms.GroupBox
    Friend WithEvents chkAlias As System.Windows.Forms.CheckBox
    Friend WithEvents chkGuild As System.Windows.Forms.CheckBox
    Friend WithEvents chkChannel As System.Windows.Forms.CheckBox
    Friend WithEvents chkSysAdm As System.Windows.Forms.CheckBox
    Friend WithEvents chkLocal As System.Windows.Forms.CheckBox
    Friend WithEvents chkNotification As System.Windows.Forms.CheckBox
    Friend WithEvents chkPM As System.Windows.Forms.CheckBox
    Friend WithEvents txtChannel As System.Windows.Forms.TextBox
    Friend WithEvents txtPrefix As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
