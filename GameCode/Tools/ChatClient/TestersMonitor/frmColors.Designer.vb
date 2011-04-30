<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmColors
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
        Me.btnBackground = New System.Windows.Forms.Button
        Me.cdMain = New System.Windows.Forms.ColorDialog
        Me.btnAlias = New System.Windows.Forms.Button
        Me.btnChannel = New System.Windows.Forms.Button
        Me.btnGuild = New System.Windows.Forms.Button
        Me.btnLocal = New System.Windows.Forms.Button
        Me.btnPM = New System.Windows.Forms.Button
        Me.btnSenate = New System.Windows.Forms.Button
        Me.btnSysAdm = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btnBackground
        '
        Me.btnBackground.Location = New System.Drawing.Point(12, 12)
        Me.btnBackground.Name = "btnBackground"
        Me.btnBackground.Size = New System.Drawing.Size(164, 23)
        Me.btnBackground.TabIndex = 0
        Me.btnBackground.Text = "Background"
        Me.btnBackground.UseVisualStyleBackColor = True
        '
        'btnAlias
        '
        Me.btnAlias.Location = New System.Drawing.Point(12, 41)
        Me.btnAlias.Name = "btnAlias"
        Me.btnAlias.Size = New System.Drawing.Size(164, 23)
        Me.btnAlias.TabIndex = 1
        Me.btnAlias.Text = "Alias Chat"
        Me.btnAlias.UseVisualStyleBackColor = True
        '
        'btnChannel
        '
        Me.btnChannel.Location = New System.Drawing.Point(12, 70)
        Me.btnChannel.Name = "btnChannel"
        Me.btnChannel.Size = New System.Drawing.Size(164, 23)
        Me.btnChannel.TabIndex = 2
        Me.btnChannel.Text = "Channel Chat"
        Me.btnChannel.UseVisualStyleBackColor = True
        '
        'btnGuild
        '
        Me.btnGuild.Location = New System.Drawing.Point(12, 99)
        Me.btnGuild.Name = "btnGuild"
        Me.btnGuild.Size = New System.Drawing.Size(164, 23)
        Me.btnGuild.TabIndex = 3
        Me.btnGuild.Text = "Guild Chat"
        Me.btnGuild.UseVisualStyleBackColor = True
        '
        'btnLocal
        '
        Me.btnLocal.Location = New System.Drawing.Point(12, 128)
        Me.btnLocal.Name = "btnLocal"
        Me.btnLocal.Size = New System.Drawing.Size(164, 23)
        Me.btnLocal.TabIndex = 4
        Me.btnLocal.Text = "Local Chat"
        Me.btnLocal.UseVisualStyleBackColor = True
        '
        'btnPM
        '
        Me.btnPM.Location = New System.Drawing.Point(12, 157)
        Me.btnPM.Name = "btnPM"
        Me.btnPM.Size = New System.Drawing.Size(164, 23)
        Me.btnPM.TabIndex = 5
        Me.btnPM.Text = "Private Messages"
        Me.btnPM.UseVisualStyleBackColor = True
        '
        'btnSenate
        '
        Me.btnSenate.Location = New System.Drawing.Point(12, 186)
        Me.btnSenate.Name = "btnSenate"
        Me.btnSenate.Size = New System.Drawing.Size(164, 23)
        Me.btnSenate.TabIndex = 6
        Me.btnSenate.Text = "Senate Chat"
        Me.btnSenate.UseVisualStyleBackColor = True
        '
        'btnSysAdm
        '
        Me.btnSysAdm.Location = New System.Drawing.Point(12, 215)
        Me.btnSysAdm.Name = "btnSysAdm"
        Me.btnSysAdm.Size = New System.Drawing.Size(164, 23)
        Me.btnSysAdm.TabIndex = 8
        Me.btnSysAdm.Text = "System Admin / Alerts"
        Me.btnSysAdm.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(12, 260)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(164, 23)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "Cancel, Do Not Save"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(12, 289)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(164, 23)
        Me.btnSave.TabIndex = 10
        Me.btnSave.Text = "Save and Close"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'frmColors
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(189, 323)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSysAdm)
        Me.Controls.Add(Me.btnSenate)
        Me.Controls.Add(Me.btnPM)
        Me.Controls.Add(Me.btnLocal)
        Me.Controls.Add(Me.btnGuild)
        Me.Controls.Add(Me.btnChannel)
        Me.Controls.Add(Me.btnAlias)
        Me.Controls.Add(Me.btnBackground)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmColors"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Color Configuration"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnBackground As System.Windows.Forms.Button
    Friend WithEvents cdMain As System.Windows.Forms.ColorDialog
    Friend WithEvents btnAlias As System.Windows.Forms.Button
    Friend WithEvents btnChannel As System.Windows.Forms.Button
    Friend WithEvents btnGuild As System.Windows.Forms.Button
    Friend WithEvents btnLocal As System.Windows.Forms.Button
    Friend WithEvents btnPM As System.Windows.Forms.Button
    Friend WithEvents btnSenate As System.Windows.Forms.Button
    Friend WithEvents btnSysAdm As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
End Class
