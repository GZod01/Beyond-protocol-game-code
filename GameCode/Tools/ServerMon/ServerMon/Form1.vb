Imports System.Net.Mail

Public Class frmProcMon
    Private mlMonitorID As Int32 = 0
    Private mlNoResponseCnt As Int32 = 0

    Private mlThreshold As Int32 = 0
    Private sFirstEnvironment As String = ""

    Private Sub btnMonitor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMonitor.Click

        If btnMonitor.Text.ToUpper = "MONITOR" Then
            If txtTo.Text Is Nothing OrElse txtTo.Text = "" Then
                MsgBox("Enter a recipients list!")
                Return
            End If
            If lstProcs.SelectedItem Is Nothing Then
                MsgBox("Select a process in the list to monitor.")
                Return
            End If

            mlThreshold = CInt(txtThreshold.Text)
            mlMonitorID = CType(lstProcs.SelectedItem, Process).Id
            mlNoResponseCnt = 0
            Timer1.Interval = 1000 ' mlThreshold \ 1000
            Timer1.Enabled = True
            Me.Text = "Monitor ProcID " & mlMonitorID
            btnMonitor.Text = "Stop"
        Else
            btnMonitor.Text = "Monitor"
            Timer1.Enabled = False
        End If
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        lstProcs.Items.Clear()

        For Each oProc As Process In System.Diagnostics.Process.GetProcesses
            lstProcs.Items.Add(oProc)
        Next
        lstProcs.DisplayMember = "ProcessName"
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        Dim oProc As Process = Nothing
        Try
            oProc = System.Diagnostics.Process.GetProcessById(mlMonitorID)
        Catch
        End Try
        If oProc Is Nothing Then
            'MsgBox("Monitored Process not found!")
            SendMailMsg("Monitored Process Not Found for " & txtServerName.Text & "!", "Server Issue")
            Return
        ElseIf oProc.HasExited = True Then
            'MsgBox("Monitored Process has exited!")
            SendMailMsg("Monitored Process Has Exited for " & txtServerName.Text & "!", "Server Issue")
            Return
        ElseIf oProc.Responding = False Then
            mlNoResponseCnt += 1

            If mlNoResponseCnt * 1000 >= mlThreshold Then
                'MsgBox("Monitored process has stopped responding!")
                SendMailMsg("Monitored Process Not Responding for " & txtServerName.Text & "!", "Server Issue")
                Return
            End If
            'Else : mlNoResponseCnt = 0
        Else
            If mlNoResponseCnt > 0 Then
                SendMailMsg("Monitored Process Appears Responding for " & txtServerName.Text & ".  " & mlNoResponseCnt.ToString & " failed responces", "Server Issue")
                mlNoResponseCnt = 0
            End If
        End If

        Timer1.Enabled = True
    End Sub

    Public Function SendMailMsg(ByVal sBody As String, ByVal sSubject As String) As Boolean
        Try
            'create the mail message object
            Dim sHost As String = txtHost.Text
            Dim sUN As String = txtUN.Text
            Dim sPW As String = txtPW.Text

            Dim oMailMsg As New MailMessage("support@darkskyentertainment.com", txtTo.Text)
            oMailMsg.BodyEncoding = System.Text.Encoding.Default
            oMailMsg.Subject = sSubject.Trim()
            oMailMsg.Body = sBody.Trim() & vbCrLf
            oMailMsg.Priority = MailPriority.High
            oMailMsg.IsBodyHtml = False
            oMailMsg.ReplyTo = New MailAddress("support@darkskyentertainment.com")

            'create Smtpclient to send the mail message
            Dim oSmtpMail As New SmtpClient
            oSmtpMail.Host = sHost '"10.70.5.51"
            oSmtpMail.UseDefaultCredentials = False
            oSmtpMail.Credentials = New Net.NetworkCredential(sUN, sPW) '"fleetcommander", "nu4tuibezl@pspos")

            oSmtpMail.Send(oMailMsg)

            Return True

        Catch ex As Exception
            'Message Error
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Private Sub frmProcMon_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim oINI As New InitFile

        oINI.WriteString("CONFIG", "ToBox", txtTo.Text)
        oINI.WriteString("CONFIG", "ServerName", txtServerName.Text)
        oINI.WriteString("CONFIG", "Threshold", txtThreshold.Text)
        oINI.WriteString("CONFIG", "Host", txtHost.Text)
        oINI.WriteString("CONFIG", "UN", txtUN.Text)
        oINI.WriteString("CONFIG", "PW", txtPW.Text)

        oINI = Nothing

    End Sub

    Private Sub frmProcMon_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Dim oINI As New InitFile()

        txtTo.Text = oINI.GetString("CONFIG", "ToBox", txtTo.Text)
        txtServerName.Text = oINI.GetString("CONFIG", "ServerName", txtServerName.Text)
        LookupFirstEnvironment(txtServerName.Text)
        txtThreshold.Text = oINI.GetString("CONFIG", "Threshold", txtThreshold.Text)
        txtHost.Text = oINI.GetString("CONFIG", "Host", txtHost.Text)
        txtUN.Text = oINI.GetString("CONFIG", "UN", txtUN.Text)
        txtPW.Text = oINI.GetString("CONFIG", "PW", txtPW.Text)

        oINI = Nothing
    End Sub

    Private Sub LookupFirstEnvironment(ByVal sServerName As String)
        'Send a packet to Operator and ask for an enviroment this monitored region hosts.
    End Sub
End Class


Public Class InitFile
    ' API functions
    Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
      ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function FlushPrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As Integer, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer
    Private strFilename As String

    ' Constructor, accepting a filename
    Public Sub New(Optional ByVal Filename As String = "")
        If Filename = "" Then
            'Ok, use the app.path
            strFilename = System.AppDomain.CurrentDomain.BaseDirectory()
            If Right$(strFilename, 1) <> "\" Then strFilename = strFilename & "\"
            strFilename = strFilename & Replace$(System.AppDomain.CurrentDomain.FriendlyName().ToLower, ".exe", ".ini")
        Else
            strFilename = Filename
        End If
    End Sub

    ' Read-only filename property
    ReadOnly Property FileName() As String
        Get
            Return strFilename
        End Get
    End Property

    Public Function GetString(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As String) As String
        ' Returns a string from your INI file
        Dim intCharCount As Integer
        Dim objResult As New System.Text.StringBuilder(2048)
        intCharCount = GetPrivateProfileString(Section, Key, _
           [Default], objResult, objResult.Capacity, strFilename)
        If intCharCount > 0 Then Return Left(objResult.ToString, intCharCount) Else Return ""
    End Function

    Public Function GetInteger(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Integer) As Integer
        ' Returns an integer from your INI file
        Return GetPrivateProfileInt(Section, Key, _
           [Default], strFilename)
    End Function

    Public Function GetBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Boolean) As Boolean
        ' Returns a boolean from your INI file
        Return (GetPrivateProfileInt(Section, Key, _
           CInt([Default]), strFilename) = 1)
    End Function

    Public Sub WriteString(ByVal Section As String, _
      ByVal Key As String, ByVal Value As String)
        ' Writes a string to your INI file
        WritePrivateProfileString(Section, Key, Value, strFilename)
        Flush()
    End Sub

    Public Sub WriteInteger(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Integer)
        ' Writes an integer to your INI file
        WriteString(Section, Key, CStr(Value))
        Flush()
    End Sub

    Public Sub WriteBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Boolean)
        ' Writes a boolean to your INI file
        WriteString(Section, Key, CStr(CInt(Value)))
        Flush()
    End Sub

    Private Sub Flush()
        ' Stores all the cached changes to your INI file
        FlushPrivateProfileString(0, 0, 0, strFilename)
    End Sub

End Class
