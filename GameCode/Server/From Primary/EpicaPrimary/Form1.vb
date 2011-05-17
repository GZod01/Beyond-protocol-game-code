Option Strict On

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        Form.CheckForIllegalCrossThreadCalls = False

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
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtEvents As System.Windows.Forms.TextBox
    Friend WithEvents txtMsg As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents lblCensus As System.Windows.Forms.Label
    Friend WithEvents txtMOTD As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents btnMonitor As System.Windows.Forms.Button
    Friend WithEvents chkAcceptClients As System.Windows.Forms.CheckBox
    Friend WithEvents chkDebug As System.Windows.Forms.CheckBox
    Friend WithEvents btnBegin As System.Windows.Forms.Button
	Friend WithEvents btnFixLinks As System.Windows.Forms.Button
	Friend WithEvents txtResTimeMult As System.Windows.Forms.TextBox
    Friend WithEvents btnResTimeMult As System.Windows.Forms.Button
    Friend WithEvents btnDisconnectClients As System.Windows.Forms.Button
    Friend WithEvents btnInvalidate As System.Windows.Forms.Button
    Friend WithEvents btnFixMinLoad As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents btnPlayer As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblStatus = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.txtEvents = New System.Windows.Forms.TextBox
        Me.btnPlayer = New System.Windows.Forms.Button
        Me.txtMsg = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.lblCensus = New System.Windows.Forms.Label
        Me.txtMOTD = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.btnMonitor = New System.Windows.Forms.Button
        Me.chkAcceptClients = New System.Windows.Forms.CheckBox
        Me.chkDebug = New System.Windows.Forms.CheckBox
        Me.btnBegin = New System.Windows.Forms.Button
        Me.btnFixLinks = New System.Windows.Forms.Button
        Me.txtResTimeMult = New System.Windows.Forms.TextBox
        Me.btnResTimeMult = New System.Windows.Forms.Button
        Me.btnDisconnectClients = New System.Windows.Forms.Button
        Me.btnInvalidate = New System.Windows.Forms.Button
        Me.btnFixMinLoad = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lblStatus
        '
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(267, 261)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(205, 32)
        Me.lblStatus.TabIndex = 2
        Me.lblStatus.Text = "Status"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(339, 319)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(133, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.TabStop = False
        Me.Button1.Text = "Shutdown Server"
        '
        'txtEvents
        '
        Me.txtEvents.Location = New System.Drawing.Point(12, 8)
        Me.txtEvents.Multiline = True
        Me.txtEvents.Name = "txtEvents"
        Me.txtEvents.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtEvents.Size = New System.Drawing.Size(460, 247)
        Me.txtEvents.TabIndex = 4
        '
        'btnPlayer
        '
        Me.btnPlayer.Location = New System.Drawing.Point(186, 290)
        Me.btnPlayer.Name = "btnPlayer"
        Me.btnPlayer.Size = New System.Drawing.Size(75, 23)
        Me.btnPlayer.TabIndex = 6
        Me.btnPlayer.Text = "Player..."
        '
        'txtMsg
        '
        Me.txtMsg.Location = New System.Drawing.Point(12, 261)
        Me.txtMsg.Multiline = True
        Me.txtMsg.Name = "txtMsg"
        Me.txtMsg.Size = New System.Drawing.Size(168, 81)
        Me.txtMsg.TabIndex = 0
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(186, 319)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 1
        Me.Button3.Text = "Send Msg"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(186, 261)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 9
        Me.Button4.Text = "Who's On"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'lblCensus
        '
        Me.lblCensus.AutoSize = True
        Me.lblCensus.Location = New System.Drawing.Point(281, 295)
        Me.lblCensus.Name = "lblCensus"
        Me.lblCensus.Size = New System.Drawing.Size(87, 13)
        Me.lblCensus.TabIndex = 10
        Me.lblCensus.Text = "Galactic Census:"
        '
        'txtMOTD
        '
        Me.txtMOTD.Location = New System.Drawing.Point(12, 348)
        Me.txtMOTD.Multiline = True
        Me.txtMOTD.Name = "txtMOTD"
        Me.txtMOTD.Size = New System.Drawing.Size(168, 81)
        Me.txtMOTD.TabIndex = 2
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(186, 406)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Set MOTD"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(339, 348)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(133, 23)
        Me.Button5.TabIndex = 13
        Me.Button5.TabStop = False
        Me.Button5.Text = "Close"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'btnMonitor
        '
        Me.btnMonitor.Location = New System.Drawing.Point(339, 406)
        Me.btnMonitor.Name = "btnMonitor"
        Me.btnMonitor.Size = New System.Drawing.Size(133, 23)
        Me.btnMonitor.TabIndex = 14
        Me.btnMonitor.Text = "Message Monitor"
        Me.btnMonitor.UseVisualStyleBackColor = True
        '
        'chkAcceptClients
        '
        Me.chkAcceptClients.AutoSize = True
        Me.chkAcceptClients.Location = New System.Drawing.Point(12, 442)
        Me.chkAcceptClients.Name = "chkAcceptClients"
        Me.chkAcceptClients.Size = New System.Drawing.Size(94, 17)
        Me.chkAcceptClients.TabIndex = 15
        Me.chkAcceptClients.Text = "Accept Clients"
        Me.chkAcceptClients.UseVisualStyleBackColor = True
        '
        'chkDebug
        '
        Me.chkDebug.AutoSize = True
        Me.chkDebug.Location = New System.Drawing.Point(122, 442)
        Me.chkDebug.Name = "chkDebug"
        Me.chkDebug.Size = New System.Drawing.Size(58, 17)
        Me.chkDebug.TabIndex = 16
        Me.chkDebug.Text = "Debug"
        Me.chkDebug.UseVisualStyleBackColor = True
        '
        'btnBegin
        '
        Me.btnBegin.Location = New System.Drawing.Point(204, 438)
        Me.btnBegin.Name = "btnBegin"
        Me.btnBegin.Size = New System.Drawing.Size(94, 23)
        Me.btnBegin.TabIndex = 17
        Me.btnBegin.Text = "Begin Engine"
        Me.btnBegin.UseVisualStyleBackColor = True
        '
        'btnFixLinks
        '
        Me.btnFixLinks.Location = New System.Drawing.Point(339, 438)
        Me.btnFixLinks.Name = "btnFixLinks"
        Me.btnFixLinks.Size = New System.Drawing.Size(133, 23)
        Me.btnFixLinks.TabIndex = 18
        Me.btnFixLinks.Text = "Fix Links"
        Me.btnFixLinks.UseVisualStyleBackColor = True
        '
        'txtResTimeMult
        '
        Me.txtResTimeMult.Location = New System.Drawing.Point(284, 377)
        Me.txtResTimeMult.Name = "txtResTimeMult"
        Me.txtResTimeMult.Size = New System.Drawing.Size(50, 20)
        Me.txtResTimeMult.TabIndex = 19
        Me.txtResTimeMult.Text = "1.0"
        '
        'btnResTimeMult
        '
        Me.btnResTimeMult.Location = New System.Drawing.Point(340, 377)
        Me.btnResTimeMult.Name = "btnResTimeMult"
        Me.btnResTimeMult.Size = New System.Drawing.Size(133, 23)
        Me.btnResTimeMult.TabIndex = 20
        Me.btnResTimeMult.Text = "Set Research Multiplier"
        Me.btnResTimeMult.UseVisualStyleBackColor = True
        '
        'btnDisconnectClients
        '
        Me.btnDisconnectClients.Location = New System.Drawing.Point(12, 465)
        Me.btnDisconnectClients.Name = "btnDisconnectClients"
        Me.btnDisconnectClients.Size = New System.Drawing.Size(125, 23)
        Me.btnDisconnectClients.TabIndex = 21
        Me.btnDisconnectClients.Text = "Disconnect Clients"
        Me.btnDisconnectClients.UseVisualStyleBackColor = True
        '
        'btnInvalidate
        '
        Me.btnInvalidate.Location = New System.Drawing.Point(143, 465)
        Me.btnInvalidate.Name = "btnInvalidate"
        Me.btnInvalidate.Size = New System.Drawing.Size(125, 23)
        Me.btnInvalidate.TabIndex = 22
        Me.btnInvalidate.Text = "Invalidate Techs"
        Me.btnInvalidate.UseVisualStyleBackColor = True
        '
        'btnFixMinLoad
        '
        Me.btnFixMinLoad.Location = New System.Drawing.Point(274, 465)
        Me.btnFixMinLoad.Name = "btnFixMinLoad"
        Me.btnFixMinLoad.Size = New System.Drawing.Size(125, 23)
        Me.btnFixMinLoad.TabIndex = 23
        Me.btnFixMinLoad.Text = "Fix Mineral Load"
        Me.btnFixMinLoad.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(267, 261)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(79, 23)
        Me.Button6.TabIndex = 24
        Me.Button6.Text = "Start Routes"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(485, 500)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.btnFixMinLoad)
        Me.Controls.Add(Me.btnInvalidate)
        Me.Controls.Add(Me.btnDisconnectClients)
        Me.Controls.Add(Me.btnResTimeMult)
        Me.Controls.Add(Me.txtResTimeMult)
        Me.Controls.Add(Me.btnFixLinks)
        Me.Controls.Add(Me.btnBegin)
        Me.Controls.Add(Me.chkDebug)
        Me.Controls.Add(Me.chkAcceptClients)
        Me.Controls.Add(Me.btnMonitor)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.txtMOTD)
        Me.Controls.Add(Me.lblCensus)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.txtMsg)
        Me.Controls.Add(Me.btnPlayer)
        Me.Controls.Add(Me.txtEvents)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lblStatus)
        Me.Name = "Form1"
        Me.Text = "Epica Primary Server"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private mfrmPHack As frmPlayer
    Private msStatusText As String = ""

    Private Delegate Sub doUpdate(ByVal sEvent As String)

    Private Sub delegateUpdateEvents(ByVal sEvent As String)
        If txtEvents.Text Is Nothing = False AndAlso txtEvents.Text.Length > 300000 Then txtEvents.Text = txtEvents.Text.Substring(0, 300000)
        txtEvents.Text = sEvent & vbCrLf & txtEvents.Text
        If msStatusText <> "" Then
            lblStatus.Text = msStatusText
        End If
        Me.Refresh()
    End Sub
    Public Shared bHideEventLines As Boolean = False
    Public Sub AddEventLine(ByVal sEvent As String)
        If bHideEventLines = True Then Return
        txtEvents.Invoke(New doUpdate(AddressOf delegateUpdateEvents), sEvent)
    End Sub

    'Public Sub AddEventLine(ByVal sEvent As String)
    '	txtEvents.Text = sEvent & vbCrLf & txtEvents.Text

    '	If msStatusText <> "" Then
    '		lblStatus.Text = msStatusText
    '	End If

    '	Me.Refresh()
    'End Sub

    Public Sub SetCensus(ByVal blCensus As Int64)
        lblCensus.Text = "Galactic Census: " & blCensus.ToString
    End Sub

    Public Sub SetStatusText(ByVal sText As String)
        msStatusText = sText
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
            gsExternalIP = ParseForIpAddress(My.Application.CommandLineArgs(3))
            Return True
        Catch ex As Exception
            MsgBox("ParseCommandLine: " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Visible = True

        gbServerInitializing = True

        gfrmDisplayForm = Me
        Dim oINI As InitFile = New InitFile
        gsMOTD = oINI.GetString("SETTINGS", "MOTD", "MOTD: Servers are up!")
        txtMOTD.Text = gsMOTD

        Dim lLastAssess As Int32 = CInt(oINI.GetString("MineralAssess", "LastCalc", "0"))
        oINI = Nothing

        lblStatus.Text = "Loading Data..."

        If ParseCommandLine() = False Then Application.Exit()
        LogEvent(LogEventType.Informational, "Initiated by Box Operator ID: " & glBoxOperatorID)
        LogEvent(LogEventType.Informational, "Operator Connection Data: " & gsOperatorIP & ":" & glOperatorPort)

        'Now, do our stuff
        LogEvent(GlobalVars.LogEventType.Informational, "Creating Logging and Database Services...")
        InitializeGlobals()

        LogEvent(GlobalVars.LogEventType.Informational, "Setting Global Cycle Number...")
        glCurrentCycle = 0      'we set it to 0... now this comes at a cost... we can never allow production for longer than 800 and something days...

        LogEvent(LogEventType.Informational, "Instantiating Model Def data services...")
        goModelDefs = New ModelDefs()

        goAgentEngine = New AgentEngine()

        'MSC - 01/15/09 - moved the before the connect to operator so that all the static data is loaded, because the operator will send us data that we need
        '  to have the static data loaded for before hand. This may resolve the missing assignments bug
        If LoadAllStaticData() = False Then
            'TODO: what now?
            Return      '???
        End If

        If goMsgSys.ConnectToOperator(gsOperatorIP, glOperatorPort) = False Then
            LogEvent(LogEventType.CriticalError, "Unable to connect to Operator, exiting initialization routine")
            Return
        End If

        'LoadAllData()

        While goMsgSys.bReceivedDomains = False
            Application.DoEvents()
            Threading.Thread.Sleep(10)
        End While

        'Ok, because we only run 1 primary
        For X As Int32 = 0 To glSystemUB
            If glSystemIdx(X) > 0 Then
                Dim oSystem As SolarSystem = goSystem(X)
                If oSystem Is Nothing = False Then
                    If oSystem.InMyDomain = False Then
                        LogEvent(LogEventType.Informational, "Forcefully setting " & BytesToString(oSystem.SystemName) & " (System) to my domain")
                        oSystem.InMyDomain = True
                        oSystem.MarkChildrenInMyDomain(True)
                    End If
                    'Stop
                End If
            End If
        Next X

        'By now, operator should have indicated to us what our domains are...
        If LoadNeighborhoodData() = False Then
            'TODO: What now?
            Return      '???
        End If

        goMsgSys.AcceptingDomains = True
        goMsgSys.AcceptingPathfinding = True

        'Ok, we are ready, tell the operator that we are ready
        LogEvent(LogEventType.Informational, "Indicating server is ready to Operator...")
        goMsgSys.SendReadyStateToOperator()

        LogEvent(GlobalVars.LogEventType.Informational, "Server Initialized! Waiting for Begin Button...")

        lblStatus.Text = "Listening (Running)"

        goMsgSys.AcceptingClients = True
        chkAcceptClients.Checked = True

        gbServerInitializing = False

        Dim dtFirstAssess As Date = Now
        If lLastAssess > 0 Then dtFirstAssess = GetDateFromNumber(lLastAssess)
        Dim lSeconds As Int32 = CInt(Now.Subtract(dtFirstAssess).TotalSeconds)
        lSeconds *= 30
        'cycles that have lapsed
        lSeconds = 2592000 - lSeconds
        AddToQueue(lSeconds, QueueItemType.eReassessMinimumMineralValues, -1, -1, -1, -1, -1, -1, -1, -1)
    End Sub

    Private Delegate Sub doRenableButton1()
    Private Sub DelegateRenableButton1()
        Button1.Enabled = True
    End Sub
    Public Sub ReEnableButton1()
        If Button1.InvokeRequired = True Then
            Button1.Invoke(New doRenableButton1(AddressOf DelegateRenableButton1))
        Else
            Button1.Enabled = True
        End If
    End Sub

    Public Sub btnShutdown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Button1.Enabled = False

        If Button1.Enabled = False Then Return

        If Button1.Text.ToLower = "shutdown server" Then
            Button1.Enabled = False
            AddEventLine("Shutting down... disabling acceptance...")

            bHideEventLines = True
            'We are no longer accepting applications, thanks for your interest
            goMsgSys.AcceptingClients = False
            goMsgSys.AcceptingDomains = False
            goMsgSys.AcceptingPathfinding = False

            'Get out of the main loop
            gb_Main_Loop_Running = False
            While gb_In_Main_Loop = True
                'Application.DoEvents()
                Threading.Thread.Sleep(20)        'give everyone time to catch up...
            End While
            AddEventLine("Last main loop ran...")

            'Notify Domains that we are shutting down and notify clients...
            AddEventLine("Broadcasting Server Shutdown...")
            goMsgSys.BroadcastServerShutdownsAndDisconnectClients()

            'Now, forcefully disconnect any clients who are still connected
            AddEventLine("Dropping clients...")
            goMsgSys.ForceDisconnectClients()
            AddEventLine("Waiting for Domain Updates...")

            Button1.Text = "End Program"

            'if there are no domains...
            If goMsgSys.GetDomainUB = -1 Then
                BeginPrimarySave()
            End If
        Else
            If Button1.Enabled = True Then Application.Exit()
        End If
    End Sub

    Private Sub btnPlayer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlayer.Click

        'For X As Int32 = 0 To glUnitUB
        '    If glUnitIdx(X) > -1 Then
        '        Dim oUnit As Unit = goUnit(X)
        '        If oUnit Is Nothing = False AndAlso oUnit.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then
        '            If oUnit.EntityDef.ObjectID = 2562 Then
        '                oUnit.Q1_HP = 50
        '                oUnit.Q2_HP = 50
        '                oUnit.Q3_HP = 50
        '                oUnit.Q4_HP = 50
        '                oUnit.SaveObject()
        '                oUnit.DataChanged()

        '                Dim oSock As NetSock = Nothing
        '                If oUnit.ParentObject Is Nothing = False Then
        '                    Dim iTemp As Int16 = CType(oUnit.ParentObject, Epica_GUID).ObjTypeID
        '                    If iTemp = ObjectType.ePlanet Then
        '                        oSock = CType(oUnit.ParentObject, Planet).oDomain.DomainSocket
        '                    ElseIf iTemp = ObjectType.eSolarSystem Then
        '                        oSock = CType(oUnit.ParentObject, SolarSystem).oDomain.DomainSocket
        '                    End If
        '                End If
        '                If oSock Is Nothing Then Continue For

        '                Dim yMsg(15) As Byte
        '                System.BitConverter.GetBytes(GlobalMessageCode.eRepairCompleted).CopyTo(yMsg, 0)
        '                oUnit.GetGUIDAsString.CopyTo(yMsg, 2)
        '                System.BitConverter.GetBytes(-2I).CopyTo(yMsg, 8)
        '                System.BitConverter.GetBytes(50I).CopyTo(yMsg, 12)
        '                oSock.SendData(yMsg)
        '                System.BitConverter.GetBytes(-3I).CopyTo(yMsg, 8)
        '                oSock.SendData(yMsg)
        '                System.BitConverter.GetBytes(-4I).CopyTo(yMsg, 8)
        '                oSock.SendData(yMsg)
        '                System.BitConverter.GetBytes(-5I).CopyTo(yMsg, 8)
        '                oSock.SendData(yMsg)
        '            End If
        '        End If
        '    End If
        'Next X
        'For X As Int32 = 0 To glUnitDefUB
        '    If glUnitDefIdx(X) > -1 Then
        '        Dim oDef As Epica_Entity_Def = goUnitDef(X)
        '        If oDef Is Nothing = False AndAlso oDef.OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID Then


        '            If oDef.ObjectID = 2562 Then
        '                oDef.BeamResist = 100
        '                oDef.ChemicalResist = 100
        '                oDef.ECMResist = 100
        '                oDef.FlameResist = 100
        '                oDef.ImpactResist = 100
        '                oDef.PiercingResist = 100

        '                oDef.Q1_MaxHP = 50
        '                oDef.Q2_MaxHP = 50
        '                oDef.Q3_MaxHP = 50
        '                oDef.Q4_MaxHP = 50

        '                oDef.SaveObject()
        '                oDef.DataChanged()
        '                goMsgSys.BroadcastToDomains(goMsgSys.GetAddObjectMessage(oDef, GlobalMessageCode.eAddObjectCommand))
        '            End If


        '        End If
        '    End If
        'Next X
        '     Dim oPlayer As Player = GetEpicaPlayer(7330)
        '     Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
        '"We have found the location of the pirate factory! Your agent was also successful in causing a malfunction in the factory's machinery which resulted in an explosion doing massive damage to the factory. The waypoint is attached to this email. For the honor of your empire!", _
        '"Pirate Factory", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
        '     If oPC Is Nothing = False Then
        '         oPC.AddEmailAttachment(1, oPlayer.ObjectID + 500000000, ObjectType.ePlanet, -15000, -16600, "Factory Loc")
        '         oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
        '     End If
        '     oPC = Nothing
        mfrmPHack = New frmPlayer()
        mfrmPHack.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim X As Int32
        Dim yMsg() As Byte = MsgSystem.CreateChatMsg(-1, txtMsg.Text, ChatMessageType.eSysAdminMessage)

        For X = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).oSocket Is Nothing = False Then
                goPlayer(X).oSocket.SendData(yMsg)
            End If
        Next X

        'Dim oSys As SolarSystem = GetEpicaSystem(56)
        'For X = 0 To glSystemUB
        '    If glSystemIdx(X) > -1 AndAlso goSystem(X).SystemType <> 3 AndAlso goSystem(X).SystemType <> 255 Then
        '        Dim oWH As New Wormhole()
        '        oWH.StartCycle = 1
        '        oWH.System1 = goSystem(X)
        '        oWH.System2 = oSys
        '        oWH.ObjTypeID = ObjectType.eWormhole
        '        oWH.ObjectID = -1
        '        oWH.InMyDomain = True
        '        oWH.LocX1 = CInt(Rnd() * 10000) - 5000
        '        oWH.LocX2 = CInt(Rnd() * 10000) - 5000
        '        oWH.LocY1 = CInt(Rnd() * 10000) - 5000
        '        oWH.LocY2 = CInt(Rnd() * 10000) - 5000

        '        If oWH.SaveObject() = True Then
        '            glWormholeUB += 1
        '            ReDim Preserve goWormhole(glWormholeUB)
        '            ReDim Preserve glWormholeIdx(glWormholeUB)
        '            goWormhole(glWormholeUB) = oWH
        '            glWormholeIdx(glWormholeUB) = oWH.ObjectID
        '            If oWH.System1 Is Nothing = False Then oWH.System1.AddWormholeReference(oWH)
        '            If oWH.System2 Is Nothing = False Then oWH.System2.AddWormholeReference(oWH)

        '            Dim yWH() As Byte = goMsgSys.GetAddObjectMessage(oWH, GlobalMessageCode.eAddObjectCommand)
        '            oWH.System1.oDomain.DomainSocket.SendData(yWH)
        '            oWH.System2.oDomain.DomainSocket.SendData(yWH)
        '        End If
        '    End If
        'Next X

        'LogEvent(LogEventType.Informational, "Loading Wormholes...")
        'Dim sSQL As String = "SELECT * FROM tblWormhole where wormholeid > 97"
        'Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
        'Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader(CommandBehavior.Default)
        'While oData.Read
        '    glWormholeUB = glWormholeUB + 1
        '    ReDim Preserve goWormhole(glWormholeUB)
        '    ReDim Preserve glWormholeIdx(glWormholeUB)
        '    goWormhole(glWormholeUB) = New Wormhole()
        '    With goWormhole(glWormholeUB)
        '        .EndCycle = CInt(oData("EndCycle"))     'TODO: instead of EndCycle, let's put a stability... as it is used, stability can degrade
        '        .LocX1 = CInt(oData("LocX1"))
        '        .LocY1 = CInt(oData("LocY1"))
        '        .LocX2 = CInt(oData("LocX2"))
        '        .LocY2 = CInt(oData("LocY2"))
        '        .ObjectID = CInt(oData("WormholeID"))
        '        .ObjTypeID = ObjectType.eWormhole
        '        .StartCycle = CInt(oData("StartCycle"))
        '        .System1 = GetEpicaSystem(CInt(oData("System1ID")))
        '        .System2 = GetEpicaSystem(CInt(oData("System2ID")))
        '        glWormholeIdx(glWormholeUB) = .ObjectID

        '        If .System1 Is Nothing = False Then .System1.AddWormholeReference(goWormhole(glWormholeUB))
        '        If .System2 Is Nothing = False Then .System2.AddWormholeReference(goWormhole(glWormholeUB))
        '    End With
        'End While
        'oData.Close()
        'oData = Nothing
        'oComm = Nothing
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim X As Int32
        Dim lLen As Int32
        Dim Y As Int32
        Dim lCnt As Int32 = 0
        For X = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).oSocket Is Nothing = False Then
                lCnt += 1
                lLen = goPlayer(X).PlayerName.Length
                For Y = 0 To goPlayer(X).PlayerName.Length - 1
                    If goPlayer(X).PlayerName(Y) = 0 Then
                        lLen = Y - 1
                        Exit For
                    End If
                Next
                AddEventLine(Mid$(System.Text.ASCIIEncoding.ASCII.GetString(goPlayer(X).PlayerName), 1, lLen))
            End If
        Next X
        AddEventLine("Total Players Online: " & lCnt)

        AddEventLine("Queue Size: " & GetQueueSize())
        AddEventLine("Queue Item Count: " & GetQueueItemCount())
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        gsMOTD = txtMOTD.Text
        Dim oINI As New InitFile
        oINI.WriteString("SETTINGS", "MOTD", gsMOTD)
        oINI = Nothing
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If Button5.Text = "Close" Then
            Button5.Text = "Confirm"
        ElseIf Button5.Text = "Confirm" Then
            gb_Main_Loop_Running = False

            goMsgSys.AcceptingClients = False
            goMsgSys.AcceptingDomains = False
            goMsgSys.AcceptingPathfinding = False
            'Notify Domains that we are shutting down and notify clients...
            AddEventLine("Broadcasting Server Shutdown...")
            goMsgSys.BroadcastServerShutdownsAndDisconnectClients()

            'Get out of the main loop
            gb_Main_Loop_Running = False
            While gb_In_Main_Loop = True
                'Application.DoEvents()
                Threading.Thread.Sleep(20)        'give everyone time to catch up...
            End While
            AddEventLine("Last main loop ran...")

            'Now, forcefully disconnect any clients who are still connected
            AddEventLine("Dropping clients...")
            goMsgSys.ForceDisconnectClients()

            BeginPrimarySave()
            'Application.Exit()
            Button5.Text = "End Program"
        Else
            Application.Exit()
        End If
    End Sub

    Private Sub btnMonitor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMonitor.Click
        frmMsgMonitor.Show()
    End Sub

    Private Sub chkAcceptClients_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAcceptClients.CheckedChanged
        goMsgSys.AcceptingClients = chkAcceptClients.Checked
    End Sub

    Private Sub chkDebug_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDebug.CheckedChanged
        goMsgSys.bDebug = chkDebug.Checked
    End Sub

    Public Sub BeginServerEngine()
        btnBegin.Visible = False

        LogEvent(LogEventType.Informational, "Connecting to Email Server...")
        If goMsgSys.ConnectToEmailSrvr() = False Then
            btnBegin.Visible = True
            Exit Sub
        End If

        LogEvent(LogEventType.Informational, "Force updating CP counts...")
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 AndAlso glPlayerIdx(X) <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                If goPlayer(X).InMyDomain = True Then

                    For Y As Int32 = 0 To goPlayer(X).oSpecials.mlSpecialTechUB
                        If goPlayer(X).oSpecials.mlSpecialTechIdx(Y) > -1 Then
                            Dim oTech As PlayerSpecialTech = goPlayer(X).oSpecials.moSpecialTech(Y)
                            If oTech Is Nothing = False AndAlso oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                If oTech.oTech Is Nothing = False Then
                                    If oTech.oTech.ProgramControl = PlayerSpecialAttributeID.eCPLimit Then
                                        goPlayer(X).oSpecials.ProcessTechResearched(oTech.oTech)
                                    End If
                                End If
                            End If
                        End If
                    Next Y

                    'Now, send to our regions...
                    Dim yMsg(11) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yMsg, 2)
                    System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.eBadWarDecCPPenalty).CopyTo(yMsg, 6)
                    System.BitConverter.GetBytes(goPlayer(X).BadWarDecCPIncrease).CopyTo(yMsg, 8)
                    goMsgSys.BroadcastToDomains(yMsg)

                    'Now, verify their slots for factions
                    goPlayer(X).ReverifySlots()

                End If

            End If
        Next X

        
        LogEvent(LogEventType.Informational, "Updating Invulnerability states and WPV regeneration")
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso goPlayer(X).AccountStatus = AccountStatusType.eActiveAccount Then

                'goPlayer(X).RegenerateWPVValues()

                If (goPlayer(X).lStatusFlags And elPlayerStatusFlag.FullInvulnerabilityRaised) <> 0 Then
                    goPlayer(X).yIronCurtainState = eIronCurtainState.RaisingIronCurtainOnEverything
                Else : goPlayer(X).yIronCurtainState = eIronCurtainState.RaisingIronCurtainOnSelectedPlanet
                End If

                AddToQueue(glCurrentCycle + CInt(Rnd() * 1000), QueueItemType.eIronCurtainRaise, goPlayer(X).ObjectID, 0, 0, 0, 0, 0, 0, 0)
                If goPlayer(X).InMyDomain = True Then
                    Dim yMsg(11) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yMsg, 2)
                    System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.eBadWarDecCPPenalty).CopyTo(yMsg, 6)
                    System.BitConverter.GetBytes(goPlayer(X).BadWarDecCPIncrease).CopyTo(yMsg, 8)
                    goMsgSys.BroadcastToDomains(yMsg)
                End If
            End If
        Next X

        LogEvent(LogEventType.Informational, "Updating Planet Colony Limit Indicators...")
        For X As Int32 = 0 To glPlanetUB
            If glPlanetIdx(X) <> -1 Then
                Dim oPlanet As Planet = goPlanet(X)
                If oPlanet Is Nothing = False AndAlso oPlanet.PlanetInCorruption(-1, -1) = True AndAlso oPlanet.InMyDomain = True AndAlso oPlanet.oDomain Is Nothing = False AndAlso oPlanet.oDomain.DomainSocket Is Nothing = False Then
                    Dim yMsg(8) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eAdvancedEventConfig).CopyTo(yMsg, 0)
                    oPlanet.GetGUIDAsString.CopyTo(yMsg, 2)
                    yMsg(8) = 1

                    oPlanet.oDomain.DomainSocket.SendData(yMsg)
                End If
            End If
        Next X


        StartMainLoop()
        'chkAcceptClients.Checked = True
        'goMsgSys.AcceptingClients = True
        LogEvent(GlobalVars.LogEventType.Informational, "Server running...")
    End Sub

    Private Sub btnBegin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBegin.Click
        BeginServerEngine()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFixLinks.Click
        'For X As Int32 = 0 To glPlayerUB
        '	If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True Then
        '		If goPlayer(X).lLastViewedEnvir <> 0 AndAlso goPlayer(X).iLastViewedEnvirType <> 0 Then
        '			goPlayer(X).PerformLinkTest()
        '		End If
        '	End If
        'Next X
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True AndAlso goPlayer(X).lPlayerIcon <> 0 Then
                Dim bFound As Boolean = False
                For Y As Int32 = 0 To goPlayer(X).oSpecials.mlSpecialTechUB
                    If goPlayer(X).oSpecials.mlSpecialTechIdx(Y) <> -1 Then
                        If goPlayer(X).oSpecials.moSpecialTech(Y).bLinked = True Then
                            If goPlayer(X).oSpecials.moSpecialTech(Y).ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                bFound = True
                                Exit For
                            End If
                        End If
                    End If
                Next Y
                If bFound = False Then
                    goPlayer(X).oSpecials.PerformLinkTest()
                End If
            End If
        Next X
    End Sub

    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If msStatusText <> "" Then
            If lblStatus.Text <> msStatusText Then lblStatus.Text = msStatusText
        End If
    End Sub

    Private Sub btnResTimeMult_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResTimeMult.Click
        Dim fValue As Single = CSng(Val(txtResTimeMult.Text))
        If fValue = 0 Then fValue = 0.01F
        gfResTimeMult = fValue
    End Sub

    Private Sub btnDisconnectClients_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisconnectClients.Click
        goMsgSys.ForceDisconnectClients()
    End Sub

    Private Sub btnInvalidate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvalidate.Click
        If btnInvalidate.Text = "Confirm" Then
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) > -1 AndAlso goPlayer(X).InMyDomain = True Then

                    With goPlayer(X)
                        'Invalidate all engines, weapons, shields, radar
                        For Y As Int32 = 0 To .mlEngineUB
                            If .mlEngineIdx(Y) > -1 Then
                                If .moEngine(Y).Owner.ObjectID = .ObjectID Then
                                    .moEngine(Y).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign
                                    .moEngine(Y).SaveObject()
                                    .SendPlayerMessage(goMsgSys.GetAddObjectMessage(.moEngine(Y), GlobalMessageCode.eAddObjectCommand), False, 0)
                                End If
                            End If
                        Next Y
                        For Y As Int32 = 0 To .mlWeaponUB
                            If .mlWeaponIdx(Y) > -1 Then
                                If .moWeapon(Y).Owner.ObjectID = .ObjectID Then
                                    .moWeapon(Y).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign
                                    .moWeapon(Y).SaveObject()
                                    .SendPlayerMessage(goMsgSys.GetAddObjectMessage(.moWeapon(Y), GlobalMessageCode.eAddObjectCommand), False, 0)
                                End If
                            End If
                        Next Y
                        For Y As Int32 = 0 To .mlShieldUB
                            If .mlShieldIdx(Y) > -1 Then
                                If .moShield(Y).Owner.ObjectID = .ObjectID Then
                                    .moShield(Y).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign
                                    .moShield(Y).SaveObject()
                                    .SendPlayerMessage(goMsgSys.GetAddObjectMessage(.moShield(Y), GlobalMessageCode.eAddObjectCommand), False, 0)
                                End If
                            End If
                        Next Y
                        For Y As Int32 = 0 To .mlRadarUB
                            If .mlRadarIdx(Y) > -1 Then
                                If .moRadar(Y).Owner.ObjectID = .ObjectID Then
                                    .moRadar(Y).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign
                                    .moRadar(Y).SaveObject()
                                    .SendPlayerMessage(goMsgSys.GetAddObjectMessage(.moRadar(Y), GlobalMessageCode.eAddObjectCommand), False, 0)
                                End If
                            End If
                        Next Y

                        'Then, mark prototypes using those items with DELETED
                        For Y As Int32 = 0 To .mlPrototypeUB
                            If .mlPrototypeIdx(Y) > -1 Then
                                Dim oEngine As EngineTech = CType(.GetTech(.moPrototype(Y).lEngineTech, ObjectType.eEngineTech), EngineTech)
                                Dim oRadar As RadarTech = CType(.GetTech(.moPrototype(Y).lRadarTech, ObjectType.eRadarTech), RadarTech)
                                Dim oShield As ShieldTech = CType(.GetTech(.moPrototype(Y).lShieldTech, ObjectType.eShieldTech), ShieldTech)
                                If (oEngine Is Nothing = False AndAlso oEngine.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign) OrElse _
                                   (oRadar Is Nothing = False AndAlso oRadar.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign) OrElse _
                                   (oShield Is Nothing = False AndAlso oShield.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign) Then
                                    .moPrototype(Y).yArchived = 255
                                    .moPrototype(Y).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign

                                    Dim yResp(7) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yResp, 0)
                                    System.BitConverter.GetBytes(.moPrototype(Y).ObjectID).CopyTo(yResp, 2)
                                    System.BitConverter.GetBytes(.moPrototype(Y).ObjTypeID).CopyTo(yResp, 6)
                                    .SendPlayerMessage(yResp, False, 0)
                                Else
                                    If .moPrototype(Y).ValidateDesign() = False Then
                                        .moPrototype(Y).yArchived = 255
                                        .moPrototype(Y).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign

                                        Dim yResp(7) As Byte
                                        System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yResp, 0)
                                        System.BitConverter.GetBytes(.moPrototype(Y).ObjectID).CopyTo(yResp, 2)
                                        System.BitConverter.GetBytes(.moPrototype(Y).ObjTypeID).CopyTo(yResp, 6)
                                        .SendPlayerMessage(yResp, False, 0)
                                    End If
                                End If

                            End If
                        Next Y

                    End With

                End If
            Next X
        Else
            btnInvalidate.Text = "Confirm"
        End If
    End Sub

    Private Sub btnFixMinLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFixMinLoad.Click
        If btnFixMinLoad.Text.ToUpper = "CONFIRM" Then
            Try
                Dim oComm As New OleDb.OleDbCommand("SELECT * FROM tblMineralCache WHERE ParentTypeID = 3 AND ParentID IN ( SELECT PlanetID FROM tblPlanet WHERE ParentID > 107)", goCN)
                Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader()
                While oData.Read = True
                    Dim lCacheID As Int32 = CInt(oData("CacheID"))
                    Dim oCache As MineralCache = GetEpicaMineralCache(lCacheID)
                    If oCache Is Nothing Then
                        oCache = New MineralCache()
                        With oCache
                            .CacheTypeID = CByte(oData("CacheTypeID"))
                            .Concentration = CInt(oData("Concentration"))
                            .LocX = CInt(oData("LocX"))
                            .LocZ = CInt(oData("LocY"))
                            .ObjectID = CInt(oData("CacheID"))
                            .ObjTypeID = ObjectType.eMineralCache
                            .oMineral = GetEpicaMineral(CInt(oData("MineralID")))
                            .ParentObject = GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
                            .OriginalConcentration = CInt(oData("OriginalConcentration"))
                            .Quantity = CInt(oData("Quantity"))

                            Dim lCurrentIdx As Int32 = glMineralCacheUB + 1
                            ReDim Preserve glMineralCacheIdx(lCurrentIdx)
                            ReDim Preserve goMineralCache(lCurrentIdx)
                            .lServerIndex = lCurrentIdx
                            glMineralCacheIdx(lCurrentIdx) = .ObjectID
                            goMineralCache(lCurrentIdx) = oCache
                            glMineralCacheUB += 1
                        End With
                    End If

                    With oCache
                        If .ParentObject Is Nothing = False Then
                            Dim oPlanet As Planet = CType(.ParentObject, Planet)
                            oPlanet.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oCache, GlobalMessageCode.eAddObjectCommand))
                        End If
                    End With
                End While
                oData.Close()
                oData = Nothing
                oComm.Dispose()
                oComm = Nothing
            Catch
            End Try

            btnFixMinLoad.Text = "Fix Mineral Load"
        Else
            btnFixMinLoad.Text = "Confirm"
        End If
    End Sub

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Button6.Visible = False

        LogEvent(LogEventType.Informational, "Starting Movement Patterns...")

        bHideEventLines = True
        For X As Int32 = 0 To glUnitUB
            If glUnitIdx(X) <> -1 Then
                With goUnit(X)

                    If .Owner Is Nothing = False Then
                        If .Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID AndAlso .Owner.AccountStatus <> AccountStatusType.eActiveAccount AndAlso .Owner.AccountStatus <> AccountStatusType.eTrialAccount AndAlso .Owner.AccountStatus <> AccountStatusType.eMondelisActive Then
                            Continue For
                        End If
                    End If

                    'Let's get the parentobject as a GUID first
                    Dim oParent As Epica_GUID = CType(.ParentObject, Epica_GUID)
                    If oParent Is Nothing Then Continue For

                    'so, let's check if it has a cargo route
                    Dim bHasRoute As Boolean  '= .lRefineryIndex <> -1 AndAlso glFacilityIdx(.lRefineryIndex) <> -1
                    'bHasRoute = bHasRoute AndAlso ((.lCacheIndex <> -1 AndAlso glMineralCacheIdx(.lCacheIndex) <> -1) OrElse (.lMiningFacIndex <> -1 AndAlso glFacilityIdx(.lMiningFacIndex) <> -1))
                    bHasRoute = .lRouteUB <> -1

                    If bHasRoute = True AndAlso .bRoutePaused = False Then
                        'Ok, first, is the entity in a planet or system environment?
                        If oParent.ObjTypeID = ObjectType.ePlanet OrElse oParent.ObjTypeID = ObjectType.eSolarSystem Then
                            'Yes, ok, check if it has cargo 
                            '.CurrentRouteItemAction()
                            Dim lTmpIdx As Int32 = .lCurrentRouteIdx
                            lTmpIdx -= 1
                            If lTmpIdx < 0 Then lTmpIdx = .lRouteUB
                            .lCurrentRouteIdx = lTmpIdx
                            '.ProcessNextRouteItem()
                            AddToQueue(glCurrentCycle + CInt(Rnd() * 1500) + 30, QueueItemType.eProcessNextRouteItem, .ObjectID, -1, -1, -1, -1, -1, -1, -1)

                            'If .Cargo_Cap = 0 Then
                            '	'Ok, send it to the dropoff
                            '	'.HandleReturnToRefinery()
                            'Else
                            '	'ok, send it to the source
                            '	'.HandleReturnToMiningSite()
                            'End If
                        Else
                            'Is it inside the refinery or mining site? which MUST be a facility (for now)
                            If oParent.ObjTypeID = ObjectType.eFacility Then
                                '.CheckRouteArrival()
                                AddToQueue(glCurrentCycle + CInt(Rnd() * 1500) + 30, QueueItemType.eUndockAndReturnToRefinery_QIT, .ObjectID, .ObjTypeID, oParent.ObjectID, oParent.ObjTypeID, 0, 0, 0, 0)

                                'ok, is it in the dropoff? or the mining facility?
                                'If oParent.ObjectID = glFacilityIdx(.lRefineryIndex) Then
                                '	'yes, it is in the refinery, is the facility on Auto Launch and active?
                                '	If goFacility(.lRefineryIndex).AutoLaunch = True AndAlso goFacility(.lRefineryIndex).Active = True Then
                                '		'yes, ok, add to our queue to undock and return to the mine
                                '		AddToQueue(glCurrentCycle, EngineCode.QueueItemType.eUndockAndReturnToMine_QIT, .ObjectID, .ObjTypeID, oParent.ObjectID, oParent.ObjTypeID)
                                '	End If
                                'ElseIf .lMiningFacIndex <> -1 AndAlso oParent.ObjectID = glFacilityIdx(.lMiningFacIndex) Then
                                '	'It is in the mining facility, is the facility on Auto Launch and Active?
                                '	If goFacility(.lMiningFacIndex).AutoLaunch = True AndAlso goFacility(.lMiningFacIndex).Active = True Then
                                '		'Yes, ok, add to our queue to undock and return to the dropoff
                                '		AddToQueue(glCurrentCycle, EngineCode.QueueItemType.eUndockAndReturnToRefinery_QIT, .ObjectID, .ObjTypeID, oParent.ObjectID, oParent.ObjTypeID)
                                '	End If
                                'End If
                            End If
                        End If
                    Else
                        'Does not have a cargo route, is the dest = loc?
                        If .LocX <> .DestX OrElse .LocZ <> .DestZ OrElse .DestEnvirID <> oParent.ObjectID OrElse .DestEnvirTypeID <> oParent.ObjTypeID Then
                            'No, so let's send a move request to the region
                            'MsgCode - 2, DestX - 4, DestZ - 4, DestA - 2, DestID - 4, DestTypeID - 2, GUID List...
                            Dim yMsg(23) As Byte
                            Dim lPos As Int32 = 0
                            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMsg, lPos) : lPos += 2
                            System.BitConverter.GetBytes(.DestX).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.DestZ).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.LocAngle).CopyTo(yMsg, lPos) : lPos += 2
                            System.BitConverter.GetBytes(.DestEnvirID).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.DestEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
                            .GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6

                            If oParent.ObjTypeID = ObjectType.ePlanet Then
                                CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                            ElseIf oParent.ObjTypeID = ObjectType.eSolarSystem Then
                                CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                            End If
                        End If
                    End If

                End With
            End If
        Next X

    End Sub
End Class
