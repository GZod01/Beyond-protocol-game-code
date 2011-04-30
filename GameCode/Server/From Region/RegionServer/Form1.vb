Option Strict On

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtObs1 As System.Windows.Forms.TextBox
    Friend WithEvents txtObs2 As System.Windows.Forms.TextBox
    Friend WithEvents txtObs3 As System.Windows.Forms.TextBox
    Friend WithEvents btnObs1 As System.Windows.Forms.Button
    Friend WithEvents btnObs2 As System.Windows.Forms.Button
    Friend WithEvents btnObs3 As System.Windows.Forms.Button
    Friend WithEvents lblObs1 As System.Windows.Forms.Label
    Friend WithEvents lblObs2 As System.Windows.Forms.Label
    Friend WithEvents lblObs3 As System.Windows.Forms.Label
    Friend WithEvents chkFastRefresh As System.Windows.Forms.CheckBox
	Friend WithEvents Button1 As System.Windows.Forms.Button
	Friend WithEvents Button2 As System.Windows.Forms.Button
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents txtEvents As System.Windows.Forms.TextBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtEvents = New System.Windows.Forms.TextBox
        Me.txtObs1 = New System.Windows.Forms.TextBox
        Me.txtObs2 = New System.Windows.Forms.TextBox
        Me.txtObs3 = New System.Windows.Forms.TextBox
        Me.btnObs1 = New System.Windows.Forms.Button
        Me.btnObs2 = New System.Windows.Forms.Button
        Me.btnObs3 = New System.Windows.Forms.Button
        Me.lblObs1 = New System.Windows.Forms.Label
        Me.lblObs2 = New System.Windows.Forms.Label
        Me.lblObs3 = New System.Windows.Forms.Label
        Me.chkFastRefresh = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(8, 264)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(151, 96)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Label1"
        '
        'txtEvents
        '
        Me.txtEvents.Location = New System.Drawing.Point(8, 8)
        Me.txtEvents.Multiline = True
        Me.txtEvents.Name = "txtEvents"
        Me.txtEvents.ReadOnly = True
        Me.txtEvents.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtEvents.Size = New System.Drawing.Size(640, 248)
        Me.txtEvents.TabIndex = 1
        Me.txtEvents.TabStop = False
        '
        'txtObs1
        '
        Me.txtObs1.Location = New System.Drawing.Point(165, 262)
        Me.txtObs1.Name = "txtObs1"
        Me.txtObs1.Size = New System.Drawing.Size(100, 20)
        Me.txtObs1.TabIndex = 2
        '
        'txtObs2
        '
        Me.txtObs2.Location = New System.Drawing.Point(165, 288)
        Me.txtObs2.Name = "txtObs2"
        Me.txtObs2.Size = New System.Drawing.Size(100, 20)
        Me.txtObs2.TabIndex = 3
        '
        'txtObs3
        '
        Me.txtObs3.Location = New System.Drawing.Point(165, 314)
        Me.txtObs3.Name = "txtObs3"
        Me.txtObs3.Size = New System.Drawing.Size(100, 20)
        Me.txtObs3.TabIndex = 4
        '
        'btnObs1
        '
        Me.btnObs1.Location = New System.Drawing.Point(271, 260)
        Me.btnObs1.Name = "btnObs1"
        Me.btnObs1.Size = New System.Drawing.Size(75, 23)
        Me.btnObs1.TabIndex = 5
        Me.btnObs1.Text = "Observe"
        Me.btnObs1.UseVisualStyleBackColor = True
        '
        'btnObs2
        '
        Me.btnObs2.Location = New System.Drawing.Point(271, 286)
        Me.btnObs2.Name = "btnObs2"
        Me.btnObs2.Size = New System.Drawing.Size(75, 23)
        Me.btnObs2.TabIndex = 6
        Me.btnObs2.Text = "Observe"
        Me.btnObs2.UseVisualStyleBackColor = True
        '
        'btnObs3
        '
        Me.btnObs3.Location = New System.Drawing.Point(271, 312)
        Me.btnObs3.Name = "btnObs3"
        Me.btnObs3.Size = New System.Drawing.Size(75, 23)
        Me.btnObs3.TabIndex = 7
        Me.btnObs3.Text = "Observe"
        Me.btnObs3.UseVisualStyleBackColor = True
        '
        'lblObs1
        '
        Me.lblObs1.AutoSize = True
        Me.lblObs1.Location = New System.Drawing.Point(654, 11)
        Me.lblObs1.Name = "lblObs1"
        Me.lblObs1.Size = New System.Drawing.Size(39, 13)
        Me.lblObs1.TabIndex = 8
        Me.lblObs1.Text = "Label2"
        '
        'lblObs2
        '
        Me.lblObs2.AutoSize = True
        Me.lblObs2.Location = New System.Drawing.Point(654, 134)
        Me.lblObs2.Name = "lblObs2"
        Me.lblObs2.Size = New System.Drawing.Size(39, 13)
        Me.lblObs2.TabIndex = 9
        Me.lblObs2.Text = "Label2"
        '
        'lblObs3
        '
        Me.lblObs3.AutoSize = True
        Me.lblObs3.Location = New System.Drawing.Point(654, 260)
        Me.lblObs3.Name = "lblObs3"
        Me.lblObs3.Size = New System.Drawing.Size(39, 13)
        Me.lblObs3.TabIndex = 10
        Me.lblObs3.Text = "Label2"
        '
        'chkFastRefresh
        '
        Me.chkFastRefresh.AutoSize = True
        Me.chkFastRefresh.Location = New System.Drawing.Point(165, 340)
        Me.chkFastRefresh.Name = "chkFastRefresh"
        Me.chkFastRefresh.Size = New System.Drawing.Size(86, 17)
        Me.chkFastRefresh.TabIndex = 11
        Me.chkFastRefresh.Text = "Fast Refresh"
        Me.chkFastRefresh.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(509, 285)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "Refresh"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(509, 310)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 13
        Me.Button2.Text = "Unused Grid"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Location = New System.Drawing.Point(352, 264)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(151, 96)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Label2"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(792, 365)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.chkFastRefresh)
        Me.Controls.Add(Me.lblObs3)
        Me.Controls.Add(Me.lblObs2)
        Me.Controls.Add(Me.lblObs1)
        Me.Controls.Add(Me.btnObs3)
        Me.Controls.Add(Me.btnObs2)
        Me.Controls.Add(Me.btnObs1)
        Me.Controls.Add(Me.txtObs3)
        Me.Controls.Add(Me.txtObs2)
        Me.Controls.Add(Me.txtObs1)
        Me.Controls.Add(Me.txtEvents)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.Text = "Running..."
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region " Diagnostics "
	Private mbObs1 As Boolean = False
	Private mbObs2 As Boolean = False
	Private mbObs3 As Boolean = False
	Private mlIdx1 As Int32 = -1
	Private mlIdx2 As Int32 = -1
	Private mlIdx3 As Int32 = -1

	Private Sub HandleDiagnostics()

		If mbObs1 = True OrElse mbObs2 = True OrElse mbObs3 = True Then
			Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder()

			If mbObs1 = True Then
                If glEntityIdx(mlIdx1) < 1 Then
                    mbObs1 = False
                Else
                    oSB.Remove(0, oSB.Length)

                    With goEntity(mlIdx1)
                        oSB.AppendLine("LocX: " & .LocX)
                        oSB.AppendLine("LocZ: " & .LocZ)
                        oSB.AppendLine("LocA: " & .LocAngle)
                        oSB.AppendLine("DestX: " & .DestX)
                        oSB.AppendLine("DestZ: " & .DestZ)
                        oSB.AppendLine("VelX: " & CInt(.VelX))
                        oSB.AppendLine("VelZ: " & CInt(.VelZ))
                        If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                            oSB.AppendLine("ENGAGED")
                        End If
                        If .bDecelerating = True Then oSB.AppendLine("DECELERATING")
                    End With
                    lblObs1.Text = oSB.ToString
                End If
			End If

			If mbObs2 = True Then
                If glEntityIdx(mlIdx2) < 1 Then
                    mbObs2 = False
                Else
                    oSB.Remove(0, oSB.Length)

                    With goEntity(mlIdx2)
                        oSB.AppendLine("LocX: " & .LocX)
                        oSB.AppendLine("LocZ: " & .LocZ)
                        oSB.AppendLine("LocA: " & .LocAngle)
                        oSB.AppendLine("DestX: " & .DestX)
                        oSB.AppendLine("DestZ: " & .DestZ)
                        oSB.AppendLine("VelX: " & CInt(.VelX))
                        oSB.AppendLine("VelZ: " & CInt(.VelZ))
                        If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                            oSB.AppendLine("ENGAGED")
                        End If
                        If .bDecelerating = True Then oSB.AppendLine("DECELERATING")
                    End With
                    lblObs2.Text = oSB.ToString
                End If
			End If

			If mbObs3 = True Then
                If glEntityIdx(mlIdx3) < 1 Then
                    mbObs3 = False
                Else
                    oSB.Remove(0, oSB.Length)

                    With goEntity(mlIdx3)
                        oSB.AppendLine("LocX: " & .LocX)
                        oSB.AppendLine("LocZ: " & .LocZ)
                        oSB.AppendLine("LocA: " & .LocAngle)
                        oSB.AppendLine("DestX: " & .DestX)
                        oSB.AppendLine("DestZ: " & .DestZ)
                        oSB.AppendLine("VelX: " & CInt(.VelX))
                        oSB.AppendLine("VelZ: " & CInt(.VelZ))
                        If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                            oSB.AppendLine("ENGAGED")
                        End If
                        If .bDecelerating = True Then oSB.AppendLine("DECELERATING")
                    End With
                    lblObs3.Text = oSB.ToString
                End If
			End If
		End If
	End Sub

	Private Sub btnObs1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnObs1.Click
		mbObs1 = Not mbObs1

		If mbObs1 = True Then
			Dim lID As Int32 = CInt(Val(txtObs1.Text))
			Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If glEntityIdx(X) = lID AndAlso goEntity(X).ObjTypeID = ObjectType.eUnit Then
					mlIdx1 = X
					Exit For
				End If
			Next X
		End If
	End Sub

	Private Sub btnObs2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnObs2.Click
		mbObs2 = Not mbObs2

		If mbObs2 = True Then
			Dim lID As Int32 = CInt(Val(txtObs2.Text))
			Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If glEntityIdx(X) = lID AndAlso goEntity(X).ObjTypeID = ObjectType.eUnit Then
					mlIdx2 = X
					Exit For
				End If
			Next X
		End If
	End Sub

	Private Sub btnObs3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnObs3.Click
		mbObs3 = Not mbObs3

		If mbObs3 = True Then
			Dim lID As Int32 = CInt(Val(txtObs3.Text))
			Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If glEntityIdx(X) = lID AndAlso goEntity(X).ObjTypeID = ObjectType.eUnit Then
					mlIdx3 = X
					Exit For
				End If
			Next X
		End If
	End Sub
#End Region

	Private oEventSB As System.Text.StringBuilder

	'#If EXTENSIVELOGGING = 1 Then
	Private mlLastEventSave As Int32 = 0
	Private moOutFile As System.IO.FileStream
	Private moWriter As System.IO.StreamWriter
	Private moFileSB As System.Text.StringBuilder
	'#End If

	Public Sub AddEventLine(ByVal sEvent As String)
		If oEventSB Is Nothing Then oEventSB = New System.Text.StringBuilder
		oEventSB.AppendLine(sEvent)

		If oEventSB.Length > 1200000 Then
			Dim sTemp As String = oEventSB.ToString.Substring(oEventSB.Length - 1000000)
			oEventSB.Length = 0
			oEventSB.Append(sTemp)
		End If

		'#If EXTENSIVELOGGING = 1 Then
		'LogToFile(sEvent)
		'#End If
	End Sub

	Public Sub LogToFile(ByVal sEvent As String)
		If moFileSB Is Nothing Then moFileSB = New System.Text.StringBuilder
		moFileSB.AppendLine(Now.ToString("HH:mm:ss") & "|" & sEvent)
		If glCurrentCycle - mlLastEventSave > 150 Then WriteToFile()
	End Sub

	'#If EXTENSIVELOGGING = 1 Then
	Private Sub WriteToFile()
		Try
			If moOutFile Is Nothing Then
				Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
				If sFile.EndsWith("\") = False Then sFile &= "\"
				sFile &= "Log.txt"
				moOutFile = New IO.FileStream(sFile, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
			End If
			If moWriter Is Nothing Then
				moWriter = New IO.StreamWriter(moOutFile)
				moWriter.AutoFlush = True
			End If

			moWriter.Write(moFileSB.ToString)
			moFileSB.Length = 0
		Catch
			'Do nothing
		End Try
	End Sub
	'#End If

	Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		gbRunning = False
	End Sub

	'#If EXTENSIVELOGGING = 1 Then
	Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		If moWriter Is Nothing = False Then
			moWriter.Close()
			moWriter.Dispose()
		End If
		moWriter = Nothing
		If moOutFile Is Nothing = False Then
			moOutFile.Close()
			moOutFile.Dispose()
		End If
		moOutFile = Nothing
	End Sub
	'#End If

	'Public Sub RefreshUnitLabel()
	'    Static xlLastTime As Int32
	'    Static xlPeakAggs As Int32

	'    Dim X As Int32
	'    Dim lCombats As Int32 = 0

	'    For X = 0 To glEntityUB
	'        If glEntityIdx(X) <> -1 Then
	'            If ((goEntity(X).CurrentStatus And elUnitStatus.eSide1HasTarget) <> 0) Or _
	'               ((goEntity(X).CurrentStatus And elUnitStatus.eSide2HasTarget) <> 0) Or _
	'               ((goEntity(X).CurrentStatus And elUnitStatus.eSide3HasTarget) <> 0) Or _
	'               ((goEntity(X).CurrentStatus And elUnitStatus.eSide4HasTarget) <> 0) Or _
	'               ((goEntity(X).CurrentStatus And elUnitStatus.eUnitEngaged) <> 0) Then
	'                lCombats += 1
	'            End If
	'        End If
	'    Next X

	'    'If timeGetTime - xlLastTime > 1000 Then
	'    '    xlLastTime = timeGetTime
	'    '    Label1.Text = "Engine Loops: " & glEngineLoops & vbCrLf & "Average Aggressions: " & (glAggressions / glEngineLoops) & vbCrLf & "Average Movements: " & (glMovements / glEngineLoops)
	'    '    glEngineLoops = 0
	'    '    glAggressions = 0
	'    '    glMovements = 0
	'    '    Me.Refresh()
	'    'End If
	'    If glAggressions > xlPeakAggs Then xlPeakAggs = glAggressions
	'    Label1.Text = "Aggressions: " & glAggressions & "(" & xlPeakAggs & ")" & vbCrLf & "Movements: " & glMovements & vbCrLf & "Engagements: " & lCombats

	'    glMovements = 0
	'    glAggressions = 0
	'    glEngineLoops = 0
	'    Me.Refresh()
	'End Sub

	'Private Sub chkFastRefresh_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFastRefresh.CheckedChanged
	'    gbFastRefreshInterval = chkFastRefresh.Checked
	'End Sub

	Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
		Dim sFinal As String

		sFinal = ""

		sFinal = "Aggressions: " & gblAggressions & vbCrLf
		sFinal = sFinal & "Cycle = " & glCurrentCycle & vbCrLf
		If goMsgSys Is Nothing = False Then sFinal &= "Clients = " & goMsgSys.GetConnectedClientCount & vbCrLf
		sFinal &= "Full Cycles = " & gbl_Full_Cycles & vbCrLf & "FCD: " & gbl_Full_Cycle_Duration

		Label1.Text = sFinal

		sFinal = "Mvmt: " & goSWMvmt.ElapsedMilliseconds.ToString & vbCrLf & "Combat: " & goSWCombat.ElapsedMilliseconds.ToString & vbCrLf
		sFinal &= "UpdateCP: " & goSWCP.ElapsedMilliseconds.ToString & vbCrLf & "Missile: " & goSWMissile.ElapsedMilliseconds.ToString & vbCrLf
		sFinal &= "Bomb: " & goSWBomb.ElapsedMilliseconds.ToString & vbCrLf

		If oEventSB Is Nothing = False Then txtEvents.Text = sFinal & oEventSB.ToString
	End Sub

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		'Form1_Paint(Nothing, Nothing)
		oEventSB.Length = 0
	End Sub

	Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim lEnvirUB As Int32 = -1
        'Dim lUnusedGrids As Int32 = 0
        'Dim lUsedGrids As Int32 = 0
        'Dim lInUseGrids As Int32 = 0
        'Dim lOldGrids As Int32 = 0
        'Dim lTotalGrids As Int32 = 0

        'If goEnvirs Is Nothing = False Then
        '	lEnvirUB = Math.Min(glEnvirUB, goEnvirs.GetUpperBound(0))
        '	For X As Int32 = 0 To lEnvirUB
        '		Dim oEnvir As Envir = goEnvirs(X)
        '		If oEnvir Is Nothing = False Then
        '			oEnvir.FillGridUsageValues(lTotalGrids, lUnusedGrids, lUsedGrids, lInUseGrids, lOldGrids)
        '		End If
        '	Next X
        'End If

        'Label2.Text = "Total Grids: " & lTotalGrids.ToString & vbCrLf & "Used: " & lUsedGrids.ToString & vbCrLf & "Unused: " & lUnusedGrids & _
        ' vbCrLf & "In Use: " & lInUseGrids.ToString & vbCrLf & "Old: " & lOldGrids.ToString

	End Sub
End Class
