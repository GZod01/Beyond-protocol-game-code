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
        Catch
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
        Me.chkAcceptClients = New System.Windows.Forms.CheckBox
        Me.txtEvents = New System.Windows.Forms.TextBox
        Me.btnShutdown = New System.Windows.Forms.Button
        Me.txtExpectedBoxes = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnSet = New System.Windows.Forms.Button
        Me.btnSetExtAddy = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtExtAddy = New System.Windows.Forms.TextBox
        Me.txtUserID = New System.Windows.Forms.TextBox
        Me.cboStatus = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnLegalMins = New System.Windows.Forms.Button
        Me.btnSpawnNeighborhood = New System.Windows.Forms.Button
        Me.btnRegionDist = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'chkAcceptClients
        '
        Me.chkAcceptClients.AutoSize = True
        Me.chkAcceptClients.Location = New System.Drawing.Point(12, 12)
        Me.chkAcceptClients.Name = "chkAcceptClients"
        Me.chkAcceptClients.Size = New System.Drawing.Size(94, 17)
        Me.chkAcceptClients.TabIndex = 0
        Me.chkAcceptClients.Text = "Accept Clients"
        Me.chkAcceptClients.UseVisualStyleBackColor = True
        '
        'txtEvents
        '
        Me.txtEvents.Location = New System.Drawing.Point(12, 35)
        Me.txtEvents.Multiline = True
        Me.txtEvents.Name = "txtEvents"
        Me.txtEvents.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtEvents.Size = New System.Drawing.Size(423, 261)
        Me.txtEvents.TabIndex = 1
        '
        'btnShutdown
        '
        Me.btnShutdown.Location = New System.Drawing.Point(360, 301)
        Me.btnShutdown.Name = "btnShutdown"
        Me.btnShutdown.Size = New System.Drawing.Size(75, 23)
        Me.btnShutdown.TabIndex = 2
        Me.btnShutdown.Text = "Shutdown"
        Me.btnShutdown.UseVisualStyleBackColor = True
        '
        'txtExpectedBoxes
        '
        Me.txtExpectedBoxes.Location = New System.Drawing.Point(286, 10)
        Me.txtExpectedBoxes.Name = "txtExpectedBoxes"
        Me.txtExpectedBoxes.Size = New System.Drawing.Size(68, 20)
        Me.txtExpectedBoxes.TabIndex = 3
        Me.txtExpectedBoxes.Text = glExpectedBoxCnt.ToString
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(196, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Expected Boxes"
        '
        'btnSet
        '
        Me.btnSet.Location = New System.Drawing.Point(360, 8)
        Me.btnSet.Name = "btnSet"
        Me.btnSet.Size = New System.Drawing.Size(75, 23)
        Me.btnSet.TabIndex = 5
        Me.btnSet.Text = "Set"
        Me.btnSet.UseVisualStyleBackColor = True
        '
        'btnSetExtAddy
        '
        Me.btnSetExtAddy.Location = New System.Drawing.Point(220, 301)
        Me.btnSetExtAddy.Name = "btnSetExtAddy"
        Me.btnSetExtAddy.Size = New System.Drawing.Size(75, 23)
        Me.btnSetExtAddy.TabIndex = 8
        Me.btnSetExtAddy.Text = "Set"
        Me.btnSetExtAddy.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 306)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "External Address"
        '
        'txtExtAddy
        '
        Me.txtExtAddy.Location = New System.Drawing.Point(102, 303)
        Me.txtExtAddy.Name = "txtExtAddy"
        Me.txtExtAddy.Size = New System.Drawing.Size(112, 20)
        Me.txtExtAddy.TabIndex = 6
        Me.txtExtAddy.Text = gsExternalAddress
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(102, 332)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(121, 20)
        Me.txtUserID.TabIndex = 9
        '
        'cboStatus
        '
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Location = New System.Drawing.Point(229, 332)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(121, 21)
        Me.cboStatus.TabIndex = 10
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(360, 330)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 11
        Me.Button1.Text = "Set"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 335)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "User ID"
        '
        'btnLegalMins
        '
        Me.btnLegalMins.Location = New System.Drawing.Point(12, 358)
        Me.btnLegalMins.Name = "btnLegalMins"
        Me.btnLegalMins.Size = New System.Drawing.Size(167, 23)
        Me.btnLegalMins.TabIndex = 13
        Me.btnLegalMins.Text = "Legalize Mineral Locations"
        Me.btnLegalMins.UseVisualStyleBackColor = True
        '
        'btnSpawnNeighborhood
        '
        Me.btnSpawnNeighborhood.Location = New System.Drawing.Point(185, 358)
        Me.btnSpawnNeighborhood.Name = "btnSpawnNeighborhood"
        Me.btnSpawnNeighborhood.Size = New System.Drawing.Size(167, 23)
        Me.btnSpawnNeighborhood.TabIndex = 14
        Me.btnSpawnNeighborhood.Text = "Spawn Neighborhood"
        Me.btnSpawnNeighborhood.UseVisualStyleBackColor = True
        '
        'btnRegionDist
        '
        Me.btnRegionDist.Location = New System.Drawing.Point(12, 387)
        Me.btnRegionDist.Name = "btnRegionDist"
        Me.btnRegionDist.Size = New System.Drawing.Size(107, 23)
        Me.btnRegionDist.TabIndex = 15
        Me.btnRegionDist.Text = "Server Dist"
        Me.btnRegionDist.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(447, 419)
        Me.Controls.Add(Me.btnRegionDist)
        Me.Controls.Add(Me.btnSpawnNeighborhood)
        Me.Controls.Add(Me.btnLegalMins)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cboStatus)
        Me.Controls.Add(Me.txtUserID)
        Me.Controls.Add(Me.btnSetExtAddy)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtExtAddy)
        Me.Controls.Add(Me.btnSet)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtExpectedBoxes)
        Me.Controls.Add(Me.btnShutdown)
        Me.Controls.Add(Me.txtEvents)
        Me.Controls.Add(Me.chkAcceptClients)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.Text = "Operator"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
	Friend WithEvents chkAcceptClients As System.Windows.Forms.CheckBox
	Friend WithEvents txtEvents As System.Windows.Forms.TextBox
	Friend WithEvents btnShutdown As System.Windows.Forms.Button
	Friend WithEvents txtExpectedBoxes As System.Windows.Forms.TextBox
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents btnSet As System.Windows.Forms.Button
	Friend WithEvents btnSetExtAddy As System.Windows.Forms.Button
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents txtExtAddy As System.Windows.Forms.TextBox
	Friend WithEvents txtUserID As System.Windows.Forms.TextBox
	Friend WithEvents cboStatus As System.Windows.Forms.ComboBox
	Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnLegalMins As System.Windows.Forms.Button
    Friend WithEvents btnSpawnNeighborhood As System.Windows.Forms.Button
    Friend WithEvents btnRegionDist As System.Windows.Forms.Button

End Class
