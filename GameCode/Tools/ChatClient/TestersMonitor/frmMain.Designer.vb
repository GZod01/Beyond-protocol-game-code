<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Credits = New System.Windows.Forms.Label
        Me.txtNew = New System.Windows.Forms.TextBox
        Me.btnSend = New System.Windows.Forms.Button
        Me.btnColors = New System.Windows.Forms.Button
        Me.btnChannels = New System.Windows.Forms.Button
        Me.btnAddTab = New System.Windows.Forms.Button
        Me.btnDeleteTab = New System.Windows.Forms.Button
        Me.btnOptions = New System.Windows.Forms.Button
        Me.tbMain = New System.Windows.Forms.TabControl
        Me.ReConnect = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Credits)
        Me.Panel1.Controls.Add(Me.txtNew)
        Me.Panel1.Controls.Add(Me.btnSend)
        Me.Panel1.Controls.Add(Me.btnColors)
        Me.Panel1.Controls.Add(Me.btnChannels)
        Me.Panel1.Controls.Add(Me.btnAddTab)
        Me.Panel1.Controls.Add(Me.btnDeleteTab)
        Me.Panel1.Controls.Add(Me.btnOptions)
        Me.Panel1.Controls.Add(Me.tbMain)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(986, 648)
        Me.Panel1.TabIndex = 13
        '
        'Credits
        '
        Me.Credits.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Credits.AutoSize = True
        Me.Credits.Location = New System.Drawing.Point(12, 603)
        Me.Credits.Name = "Credits"
        Me.Credits.Size = New System.Drawing.Size(80, 13)
        Me.Credits.TabIndex = 21
        Me.Credits.Text = "Credits: $0 / $0"
        '
        'txtNew
        '
        Me.txtNew.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNew.Enabled = False
        Me.txtNew.Location = New System.Drawing.Point(12, 619)
        Me.txtNew.Name = "txtNew"
        Me.txtNew.Size = New System.Drawing.Size(795, 20)
        Me.txtNew.TabIndex = 13
        '
        'btnSend
        '
        Me.btnSend.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSend.Enabled = False
        Me.btnSend.Location = New System.Drawing.Point(813, 617)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(75, 23)
        Me.btnSend.TabIndex = 14
        Me.btnSend.Text = "Send"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        'btnColors
        '
        Me.btnColors.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnColors.Enabled = False
        Me.btnColors.Location = New System.Drawing.Point(652, 8)
        Me.btnColors.Name = "btnColors"
        Me.btnColors.Size = New System.Drawing.Size(75, 23)
        Me.btnColors.TabIndex = 20
        Me.btnColors.Text = "Colors"
        Me.btnColors.UseVisualStyleBackColor = True
        '
        'btnChannels
        '
        Me.btnChannels.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnChannels.Enabled = False
        Me.btnChannels.Location = New System.Drawing.Point(733, 8)
        Me.btnChannels.Name = "btnChannels"
        Me.btnChannels.Size = New System.Drawing.Size(75, 23)
        Me.btnChannels.TabIndex = 19
        Me.btnChannels.Text = "Channels"
        Me.btnChannels.UseVisualStyleBackColor = True
        '
        'btnAddTab
        '
        Me.btnAddTab.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAddTab.Location = New System.Drawing.Point(814, 8)
        Me.btnAddTab.Name = "btnAddTab"
        Me.btnAddTab.Size = New System.Drawing.Size(75, 23)
        Me.btnAddTab.TabIndex = 18
        Me.btnAddTab.Text = "Add Tab"
        Me.btnAddTab.UseVisualStyleBackColor = True
        '
        'btnDeleteTab
        '
        Me.btnDeleteTab.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDeleteTab.Location = New System.Drawing.Point(895, 8)
        Me.btnDeleteTab.Name = "btnDeleteTab"
        Me.btnDeleteTab.Size = New System.Drawing.Size(75, 23)
        Me.btnDeleteTab.TabIndex = 15
        Me.btnDeleteTab.Text = "Delete Tab"
        Me.btnDeleteTab.UseVisualStyleBackColor = True
        '
        'btnOptions
        '
        Me.btnOptions.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOptions.Enabled = False
        Me.btnOptions.Location = New System.Drawing.Point(899, 617)
        Me.btnOptions.Name = "btnOptions"
        Me.btnOptions.Size = New System.Drawing.Size(75, 23)
        Me.btnOptions.TabIndex = 17
        Me.btnOptions.Text = "Options"
        Me.btnOptions.UseVisualStyleBackColor = True
        '
        'tbMain
        '
        Me.tbMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbMain.Location = New System.Drawing.Point(12, 15)
        Me.tbMain.Name = "tbMain"
        Me.tbMain.SelectedIndex = 0
        Me.tbMain.Size = New System.Drawing.Size(962, 585)
        Me.tbMain.TabIndex = 16
        '
        'ReConnect
        '
        Me.ReConnect.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ReConnect.Location = New System.Drawing.Point(282, 162)
        Me.ReConnect.Name = "ReConnect"
        Me.ReConnect.Size = New System.Drawing.Size(402, 280)
        Me.ReConnect.TabIndex = 21
        Me.ReConnect.Text = "Re-Connect"
        Me.ReConnect.UseVisualStyleBackColor = True
        Me.ReConnect.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 2000
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(986, 648)
        Me.Controls.Add(Me.ReConnect)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(431, 316)
        Me.Name = "frmMain"
        Me.Text = "Chat Monitor"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtNew As System.Windows.Forms.TextBox
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents btnColors As System.Windows.Forms.Button
    Friend WithEvents btnChannels As System.Windows.Forms.Button
    Friend WithEvents btnAddTab As System.Windows.Forms.Button
    Friend WithEvents btnDeleteTab As System.Windows.Forms.Button
    Friend WithEvents btnOptions As System.Windows.Forms.Button
    Friend WithEvents tbMain As System.Windows.Forms.TabControl
    Friend WithEvents ReConnect As System.Windows.Forms.Button
    Friend WithEvents Credits As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

    Private Sub tbMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbMain.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                txtNew.Focus()
                e.Handled = True
            End If
        Catch
        End Try
    End Sub
End Class
