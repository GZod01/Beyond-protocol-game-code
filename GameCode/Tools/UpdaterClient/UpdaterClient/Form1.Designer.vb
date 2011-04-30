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
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.prgOverall = New System.Windows.Forms.ProgressBar
        Me.prgCurrent = New System.Windows.Forms.ProgressBar
        Me.btnLaunch = New System.Windows.Forms.Button
        Me.grpEULA = New System.Windows.Forms.GroupBox
        Me.wbEULA = New System.Windows.Forms.WebBrowser
        Me.btnDecline = New System.Windows.Forms.Button
        Me.btnAccept = New System.Windows.Forms.Button
        Me.btnConfig = New System.Windows.Forms.Button
        Me.wbStatus = New System.Windows.Forms.WebBrowser
        Me.btnManual = New System.Windows.Forms.Button
        Me.grpEULA.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtStatus
        '
        Me.txtStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtStatus.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(92, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtStatus.ForeColor = System.Drawing.Color.Cyan
        Me.txtStatus.Location = New System.Drawing.Point(12, 516)
        Me.txtStatus.Multiline = True
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStatus.Size = New System.Drawing.Size(378, 69)
        Me.txtStatus.TabIndex = 0
        '
        'prgOverall
        '
        Me.prgOverall.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgOverall.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(92, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.prgOverall.ForeColor = System.Drawing.Color.Cyan
        Me.prgOverall.Location = New System.Drawing.Point(396, 534)
        Me.prgOverall.Name = "prgOverall"
        Me.prgOverall.Size = New System.Drawing.Size(419, 15)
        Me.prgOverall.TabIndex = 4
        Me.prgOverall.Value = 50
        '
        'prgCurrent
        '
        Me.prgCurrent.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgCurrent.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(92, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.prgCurrent.ForeColor = System.Drawing.Color.Cyan
        Me.prgCurrent.Location = New System.Drawing.Point(396, 516)
        Me.prgCurrent.Name = "prgCurrent"
        Me.prgCurrent.Size = New System.Drawing.Size(419, 15)
        Me.prgCurrent.TabIndex = 6
        Me.prgCurrent.Value = 50
        '
        'btnLaunch
        '
        Me.btnLaunch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLaunch.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnLaunch.Enabled = False
        Me.btnLaunch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLaunch.Location = New System.Drawing.Point(679, 555)
        Me.btnLaunch.Name = "btnLaunch"
        Me.btnLaunch.Size = New System.Drawing.Size(136, 30)
        Me.btnLaunch.TabIndex = 7
        Me.btnLaunch.Text = "Launch Game"
        Me.btnLaunch.UseVisualStyleBackColor = False
        '
        'grpEULA
        '
        Me.grpEULA.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpEULA.Controls.Add(Me.wbEULA)
        Me.grpEULA.Controls.Add(Me.btnDecline)
        Me.grpEULA.Controls.Add(Me.btnAccept)
        Me.grpEULA.Location = New System.Drawing.Point(0, 0)
        Me.grpEULA.Name = "grpEULA"
        Me.grpEULA.Size = New System.Drawing.Size(826, 592)
        Me.grpEULA.TabIndex = 9
        Me.grpEULA.TabStop = False
        '
        'wbEULA
        '
        Me.wbEULA.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.wbEULA.Location = New System.Drawing.Point(12, 19)
        Me.wbEULA.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbEULA.Name = "wbEULA"
        Me.wbEULA.Size = New System.Drawing.Size(803, 528)
        Me.wbEULA.TabIndex = 3
        '
        'btnDecline
        '
        Me.btnDecline.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDecline.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnDecline.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDecline.Location = New System.Drawing.Point(12, 553)
        Me.btnDecline.Name = "btnDecline"
        Me.btnDecline.Size = New System.Drawing.Size(102, 30)
        Me.btnDecline.TabIndex = 2
        Me.btnDecline.Text = "DECLINE"
        Me.btnDecline.UseVisualStyleBackColor = False
        '
        'btnAccept
        '
        Me.btnAccept.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAccept.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnAccept.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAccept.Location = New System.Drawing.Point(713, 553)
        Me.btnAccept.Name = "btnAccept"
        Me.btnAccept.Size = New System.Drawing.Size(102, 30)
        Me.btnAccept.TabIndex = 1
        Me.btnAccept.Text = "ACCEPT"
        Me.btnAccept.UseVisualStyleBackColor = False
        '
        'btnConfig
        '
        Me.btnConfig.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnConfig.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnConfig.Enabled = False
        Me.btnConfig.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConfig.Location = New System.Drawing.Point(396, 555)
        Me.btnConfig.Name = "btnConfig"
        Me.btnConfig.Size = New System.Drawing.Size(134, 30)
        Me.btnConfig.TabIndex = 10
        Me.btnConfig.Text = "Configure Client"
        Me.btnConfig.UseVisualStyleBackColor = False
        '
        'wbStatus
        '
        Me.wbStatus.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.wbStatus.Location = New System.Drawing.Point(12, 12)
        Me.wbStatus.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbStatus.Name = "wbStatus"
        Me.wbStatus.Size = New System.Drawing.Size(803, 498)
        Me.wbStatus.TabIndex = 11
        '
        'btnManual
        '
        Me.btnManual.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnManual.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnManual.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManual.Location = New System.Drawing.Point(536, 555)
        Me.btnManual.Name = "btnManual"
        Me.btnManual.Size = New System.Drawing.Size(136, 30)
        Me.btnManual.TabIndex = 12
        Me.btnManual.Text = "View Manual"
        Me.btnManual.UseVisualStyleBackColor = False
        Me.btnManual.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(827, 597)
        Me.Controls.Add(Me.btnManual)
        Me.Controls.Add(Me.grpEULA)
        Me.Controls.Add(Me.btnConfig)
        Me.Controls.Add(Me.btnLaunch)
        Me.Controls.Add(Me.prgCurrent)
        Me.Controls.Add(Me.prgOverall)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.wbStatus)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Dark Sky Entertainment Updater"
        Me.grpEULA.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents prgOverall As System.Windows.Forms.ProgressBar
    Friend WithEvents prgCurrent As System.Windows.Forms.ProgressBar
    Friend WithEvents btnLaunch As System.Windows.Forms.Button
    Friend WithEvents grpEULA As System.Windows.Forms.GroupBox
    Friend WithEvents btnDecline As System.Windows.Forms.Button
    Friend WithEvents btnAccept As System.Windows.Forms.Button
    Friend WithEvents wbEULA As System.Windows.Forms.WebBrowser
    Friend WithEvents btnConfig As System.Windows.Forms.Button
    Friend WithEvents wbStatus As System.Windows.Forms.WebBrowser
    Friend WithEvents btnManual As System.Windows.Forms.Button

End Class
