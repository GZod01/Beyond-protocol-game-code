<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProcMon
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
        Me.lstProcs = New System.Windows.Forms.ListBox
        Me.btnMonitor = New System.Windows.Forms.Button
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtTo = New System.Windows.Forms.TextBox
        Me.txtServerName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtThreshold = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtHost = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtUN = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtPW = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'lstProcs
        '
        Me.lstProcs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstProcs.FormattingEnabled = True
        Me.lstProcs.Location = New System.Drawing.Point(12, 12)
        Me.lstProcs.Name = "lstProcs"
        Me.lstProcs.Size = New System.Drawing.Size(243, 342)
        Me.lstProcs.TabIndex = 0
        '
        'btnMonitor
        '
        Me.btnMonitor.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMonitor.Location = New System.Drawing.Point(180, 371)
        Me.btnMonitor.Name = "btnMonitor"
        Me.btnMonitor.Size = New System.Drawing.Size(75, 23)
        Me.btnMonitor.TabIndex = 1
        Me.btnMonitor.Text = "Monitor"
        Me.btnMonitor.UseVisualStyleBackColor = True
        '
        'btnRefresh
        '
        Me.btnRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRefresh.Location = New System.Drawing.Point(12, 371)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btnRefresh.TabIndex = 2
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 402)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Alert Recipients:"
        '
        'txtTo
        '
        Me.txtTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTo.Location = New System.Drawing.Point(102, 399)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(153, 20)
        Me.txtTo.TabIndex = 4
        Me.txtTo.Text = "7275345867@txt.att.net, 4073013665@messaging.sprintpcs.com, mcampbell@darkskyente" & _
            "rtainment.com, ccoey@darkskyentertainment.com, 9728246676@txt.att.net"
        '
        'txtServerName
        '
        Me.txtServerName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtServerName.Location = New System.Drawing.Point(102, 425)
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.Size = New System.Drawing.Size(153, 20)
        Me.txtServerName.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 428)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Server Name:"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 454)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(90, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Threshold (in ms):"
        '
        'txtThreshold
        '
        Me.txtThreshold.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtThreshold.Location = New System.Drawing.Point(102, 451)
        Me.txtThreshold.Name = "txtThreshold"
        Me.txtThreshold.Size = New System.Drawing.Size(153, 20)
        Me.txtThreshold.TabIndex = 7
        Me.txtThreshold.Text = "60000"
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 480)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Out Email:"
        '
        'txtHost
        '
        Me.txtHost.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHost.Location = New System.Drawing.Point(102, 477)
        Me.txtHost.Name = "txtHost"
        Me.txtHost.Size = New System.Drawing.Size(153, 20)
        Me.txtHost.TabIndex = 9
        Me.txtHost.Text = "10.70.5.51"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 506)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(61, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Email Login"
        '
        'txtUN
        '
        Me.txtUN.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUN.Location = New System.Drawing.Point(102, 503)
        Me.txtUN.Name = "txtUN"
        Me.txtUN.Size = New System.Drawing.Size(153, 20)
        Me.txtUN.TabIndex = 11
        Me.txtUN.Text = "fleetcommander"
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 532)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Email PW"
        '
        'txtPW
        '
        Me.txtPW.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPW.Location = New System.Drawing.Point(102, 529)
        Me.txtPW.Name = "txtPW"
        Me.txtPW.Size = New System.Drawing.Size(153, 20)
        Me.txtPW.TabIndex = 13
        Me.txtPW.Text = "nu4tuibezl@pspos"
        '
        'frmProcMon
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(267, 556)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtPW)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtUN)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtHost)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtThreshold)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtServerName)
        Me.Controls.Add(Me.txtTo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.btnMonitor)
        Me.Controls.Add(Me.lstProcs)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmProcMon"
        Me.Text = "Proc Mon"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstProcs As System.Windows.Forms.ListBox
    Friend WithEvents btnMonitor As System.Windows.Forms.Button
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTo As System.Windows.Forms.TextBox
    Friend WithEvents txtServerName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtThreshold As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtHost As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtUN As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPW As System.Windows.Forms.TextBox

End Class
