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
    Friend WithEvents btnGenMapFiles As System.Windows.Forms.Button

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtEvents As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtEvents = New System.Windows.Forms.TextBox
        Me.btnGenMapFiles = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtEvents
        '
        Me.txtEvents.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEvents.Location = New System.Drawing.Point(8, 8)
        Me.txtEvents.Multiline = True
        Me.txtEvents.Name = "txtEvents"
        Me.txtEvents.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtEvents.Size = New System.Drawing.Size(288, 256)
        Me.txtEvents.TabIndex = 0
        Me.txtEvents.TabStop = False
        '
        'btnGenMapFiles
        '
        Me.btnGenMapFiles.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenMapFiles.Location = New System.Drawing.Point(84, 270)
        Me.btnGenMapFiles.Name = "btnGenMapFiles"
        Me.btnGenMapFiles.Size = New System.Drawing.Size(126, 23)
        Me.btnGenMapFiles.TabIndex = 1
        Me.btnGenMapFiles.Text = "Generate Map Files"
        Me.btnGenMapFiles.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(304, 301)
        Me.Controls.Add(Me.btnGenMapFiles)
        Me.Controls.Add(Me.txtEvents)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pathfinding Server"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

	Private moLogger As IO.FileStream = Nothing
	Private moWriter As IO.StreamWriter = Nothing

	Private Delegate Sub doUpdate(ByVal sEvent As String)

    Private Sub delegateUpdateEvents(ByVal sEvent As String)
        If txtEvents.Text Is Nothing = False AndAlso txtEvents.Text.Length > 100000 Then
            txtEvents.Text = txtEvents.Text.Substring(0, 100000)
        End If
        txtEvents.Text = sEvent & vbCrLf & txtEvents.Text
        Me.Refresh()
    End Sub

	Public Sub AddEventLine(ByVal sEvent As String)
		txtEvents.Invoke(New doUpdate(AddressOf delegateUpdateEvents), sEvent)
		'txtEvents.Text = sEvent & vbCrLf & txtEvents.Text
		'Me.Refresh()
	End Sub

	Public Sub LogSpecificEvent(ByVal sEvent As String)
		If moLogger Is Nothing Then
			Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
			If sPath.EndsWith("\") = False Then sPath &= "\"
			moLogger = New IO.FileStream(sPath & "PF.txt", IO.FileMode.Append)
		End If
		If moWriter Is Nothing Then
			moWriter = New IO.StreamWriter(moLogger)
		End If
		moWriter.WriteLine(Now.ToString("HH:mm:ss") & "|" & sEvent)
	End Sub

    Private Function ParseCommandLine() As Boolean
        Try
            Dim sValue As String = My.Application.CommandLineArgs(0)
            If IsNumeric(sValue) = False Then
                MsgBox("Unable to determine Box Operator ID in Command Line")
                Return False
            End If
            glBoxOperatorID = CInt(Val(sValue))

            gsOperatorIP = My.Application.CommandLineArgs(1)
            glOperatorPort = CInt(Val(My.Application.CommandLineArgs(2)))
            Return True
        Catch ex As Exception
            MsgBox("ParseCommandLine: " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim X As Int32
		Form.CheckForIllegalCrossThreadCalls = False

        Me.Visible = True
        gfrmDisplayForm = Me

        AddEventLine("Parsing Command Line...")
        If ParseCommandLine() = False Then Return
        AddEventLine("BoxOperator: " & glBoxOperatorID & ", OperatorIP: " & gsOperatorIP & ":" & glOperatorPort)

        AddEventLine("Initializing Message System")
        goMsgSys = New MsgSystem()

        AddEventLine("Loading Model Data")
        LoadModelVP()

        AddEventLine("Connecting to Operator...")
        If goMsgSys.ConnectToOperator() = False Then
            goMsgSys = Nothing
            AddEventLine("Aborting startup...")
            Return
        End If

        AddEventLine("Waiting for Primary Connection Info...")
        While goMsgSys.bHavePrimaryConnInfo = False
            Threading.Thread.Sleep(10)
            Application.DoEvents()
        End While

        AddEventLine("Connecting to Primary...")
        If goMsgSys.ConnectToPrimary = True Then
            'Ok, we can proceed... next, we need to...
            If goMsgSys.SendPrimaryMyInfo() = True Then
                'but right now, I just want to get this thing working... so... we'll
                '  start accepting domains
                goMsgSys.AcceptingDomains = True

                'now, we begin loading all of our data... first, request our startypes
                AddEventLine("Requesting Star Types...")
                If goMsgSys.GetStarTypes() = False Then
                    'Debug.Assert(False)
                End If

                'now, next
                AddEventLine("Requesting Galaxy and Systems...")
                If goMsgSys.GetGalaxyAndSystems() = False Then
                    'Debug.Assert(False)
                End If

                For X = 0 To goGalaxy.mlSystemUB
                    goGalaxy.CurrentSystemIdx = X
                    AddEventLine("Requesting System Details " & goGalaxy.moSystems(X).ObjectID & "...")
                    If goMsgSys.GetSystemDetails(goGalaxy.moSystems(X).ObjectID) = False Then
                        'Debug.Assert(False)
                    End If
                Next X

                'This was unnecessary since GetSystemDetails does it inline
                'Now, ensure our terrains are generated...
                'Dim lCnt As Int32 = 0
                'AddEventLine("Populating Planet Terrains...")
                'For X = 0 To goGalaxy.mlSystemUB
                '    For Y As Int32 = 0 To goGalaxy.moSystems(X).PlanetUB
                '        goGalaxy.moSystems(X).moPlanets(Y).PopulateData()
                '        lCnt += 1
                '    Next Y
                'Next X
                'AddEventLine(lCnt.ToString & " Planet Terrains Loaded!")

                AddEventLine("Requesting Entities from Primary...")
                goMsgSys.SendRequestEntities()
                AddEventLine("Indicating to Operator Server Ready")
                goMsgSys.SendReadyStateToOperator()
                AddEventLine("Pathfinding Server Initialized!")
            Else
                goMsgSys.CloseAllConnections()
                goMsgSys = Nothing
            End If
        Else
            goMsgSys = Nothing
        End If
    End Sub

	Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		If moWriter Is Nothing = False Then moWriter.Close()
		moWriter = Nothing
		If moLogger Is Nothing = False Then moLogger.Close()
		moLogger = Nothing
		If goMsgSys Is Nothing = False Then
			If goMsgSys.AcceptingDomains Then goMsgSys.AcceptingDomains = False
			goMsgSys.CloseAllConnections()
		End If
		goMsgSys = Nothing
	End Sub

    Private Sub btnGenMapFiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenMapFiles.Click
        For X As Int32 = 0 To glPlanetVPUB
            If goPlanetVP(X).lID = 358 Then goPlanetVP(X).oPlanetRef.SaveToFile()
        Next X
        AddEventLine("Saved all Planets to File.")
    End Sub
End Class
