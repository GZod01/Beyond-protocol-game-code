Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmPlayerSetup
    Inherits UIWindow

    Private lblTitle As UILabel

    Private lblUserName As UILabel
    Private txtUserName As UITextBox
    Private lblPassword As UILabel
    Private txtPassword As UITextBox

    Private lblPlayerName As UILabel
    Private txtPlayerName As UITextBox
    Private lblEmpireName As UILabel
    Private lblGender As UILabel
    Private txtEmpireName As UITextBox
	Private lnDiv1 As UILine
	Private lblEmail As UILabel
	Private txtEmail As UITextBox

    Private WithEvents optMale As UIOption
    Private WithEvents optFemale As UIOption
    Private WithEvents btnSubmit As UIButton
    Private WithEvents btnCancel As UIButton

    Private fraIconChooser As frmIconChooser

    Public sUserName As String
    Public sPassword As String

    Private mbLoading As Boolean = True

    Private mlOriginalMaster As Int32 = 0

	Public Sub New(ByRef oUILib As UILib, ByVal sUserName As String, ByVal sPassword As String)
        MyBase.New(oUILib)

        If goSound Is Nothing = False Then
            mlOriginalMaster = goSound.MasterVolume
            goSound.MasterVolume = 0
            goSound.UpdateSoundVolumes()
        End If

		'frmPlayerSetup initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.ePlayerSetup
            .ControlName = "frmPlayerSetup"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 245
            .Width = 256
            .Height = 495 '470
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 5
			.Width = 179
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Player Initial Setup"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 14.25F, FontStyle.Bold Or FontStyle.Italic, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'lblUserName initial props
		lblUserName = New UILabel(oUILib)
		With lblUserName
			.ControlName = "lblUserName"
			.Left = 10
			.Top = 50
			.Width = 95
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Username: " ' "New Username:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.ToolTipText = "The Username to enter to log in to Beyond Protocol."
		End With
		Me.AddChild(CType(lblUserName, UIControl))

		'txtUserName initial props
		txtUserName = New UITextBox(oUILib)
		With txtUserName
			.ControlName = "txtUserName"
			.Left = 110
			.Top = 50
			.Width = 140
			.Height = 18
			.Enabled = False
			.Visible = True
			.Caption = sUserName
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 20
			.BorderColor = muSettings.InterfaceBorderColor
			.ToolTipText = "The Username to enter to log in to Beyond Protocol."
		End With
		Me.AddChild(CType(txtUserName, UIControl))

		'lblPassword initial props
		lblPassword = New UILabel(oUILib)
		With lblPassword
			.ControlName = "lblPassword"
			.Left = 10
			.Top = 75
			.Width = 95
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Password" ' "New Password:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.ToolTipText = "The Password to enter to log in to Beyond Protocol."
		End With
		Me.AddChild(CType(lblPassword, UIControl))

		'txtPassword initial props
		txtPassword = New UITextBox(oUILib)
		With txtPassword
			.ControlName = "txtPassword"
			.Left = 110
			.Top = 75
			.Width = 140
			.Height = 18
			.Enabled = False
			.Visible = True
			.Caption = sPassword
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 20
			.PasswordChar = "*"
			.BorderColor = muSettings.InterfaceBorderColor
			.ToolTipText = "The Password to enter to log in to Beyond Protocol."
		End With
		Me.AddChild(CType(txtPassword, UIControl))

		'lblPlayerName initial props
		lblPlayerName = New UILabel(oUILib)
		With lblPlayerName
			.ControlName = "lblPlayerName"
			.Left = 10
			.Top = 100
			.Width = 85
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Player Name:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.ToolTipText = "This is what other players will see."
		End With
		Me.AddChild(CType(lblPlayerName, UIControl))

		'txtPlayerName initial props
		txtPlayerName = New UITextBox(oUILib)
		With txtPlayerName
			.ControlName = "txtPlayerName"
			.Left = 110
			.Top = 100
			.Width = 140
			.Height = 18
            .Enabled = True
			.Visible = True
			.Caption = sUserName
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 20
			.BorderColor = muSettings.InterfaceBorderColor
			.ToolTipText = "This is what other players will see."
		End With
		Me.AddChild(CType(txtPlayerName, UIControl))

		'lblEmpireName initial props
		lblEmpireName = New UILabel(oUILib)
		With lblEmpireName
			.ControlName = "lblEmpireName"
			.Left = 10
			.Top = 125
			.Width = 85
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Empire Name:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.ToolTipText = "The name displayed for GNS and indirect references to your domain."
		End With
		Me.AddChild(CType(lblEmpireName, UIControl))

		'lblGender initial props
		lblGender = New UILabel(oUILib)
		With lblGender
			.ControlName = "lblGender"
			.Left = 10
			.Top = 150
			.Width = 85
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Gender:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblGender, UIControl))

		'lblEmail initial props
		lblEmail = New UILabel(oUILib)
		With lblEmail
			.ControlName = "lblEmail"
			.Left = 10
			.Top = 175
			.Width = 85
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Email Alerts:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblEmail, UIControl))

		'btnSubmit initial props
		btnSubmit = New UIButton(oUILib)
		With btnSubmit
			.ControlName = "btnSubmit"
			.Left = 6
			.Top = 465 '440
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Submit"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnSubmit, UIControl))

		'btnCancel initial props
		btnCancel = New UIButton(oUILib)
		With btnCancel
			.ControlName = "btnCancel"
			.Left = 150
			.Top = 465 '440
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Cancel"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnCancel, UIControl))

		'txtEmpireName initial props
		txtEmpireName = New UITextBox(oUILib)
		With txtEmpireName
			.ControlName = "txtEmpireName"
			.Left = 110
			.Top = 125
			.Width = 140
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 20
			.BorderColor = muSettings.InterfaceBorderColor
			.ToolTipText = "The name displayed for GNS and indirect references to your domain."
		End With
		Me.AddChild(CType(txtEmpireName, UIControl))

		'optMale initial props
		optMale = New UIOption(oUILib)
		With optMale
			.ControlName = "optMale"
			.Left = 80
			.Top = 150
			.Width = 50
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Male"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optMale, UIControl))

		'optFemale initial props
		optFemale = New UIOption(oUILib)
		With optFemale
			.ControlName = "optFemale"
			.Left = 175
			.Top = 150
			.Width = 66
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Female"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optFemale, UIControl))

		'txtEmail initial props
		txtEmail = New UITextBox(oUILib)
		With txtEmail
			.ControlName = "txtEmail"
			.Left = 110
			.Top = 175
			.Width = 140
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 255
			.BorderColor = muSettings.InterfaceBorderColor
			.ToolTipText = "The email address that alerts will be sent to." & vbCrLf & "This can be any email address including cell phones."
		End With
		Me.AddChild(CType(txtEmail, UIControl))

		'NewControl8 initial props
		lnDiv1 = New UILine(oUILib)
		With lnDiv1
			.ControlName = "lnDiv1"
			.Left = 1
			.Top = 35
			.Width = 255
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'fraIconChooser initial props
		fraIconChooser = New frmIconChooser(oUILib, False)
		With fraIconChooser
			.Left = 3
			.Top = 207 '182
			.Width = 250
			.Height = 250
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
		End With
		Me.AddChild(CType(fraIconChooser, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)

        mbLoading = False

        GFXEngine.bRenderInProgress = False

		LoginScreen.UpdateButNoRender = True
	End Sub

    Private Function ValidateData() As Boolean
        If optMale.Value = False AndAlso optFemale.Value = False Then
            MyBase.moUILib.AddNotification("Please select whether you are male or female.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        If txtEmpireName.Caption.Contains("*") = True Then
            MyBase.moUILib.AddNotification("Please enter a valid Empire Name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If
        If txtPlayerName.Caption.Contains("*") = True Then
            MyBase.moUILib.AddNotification("Please enter a valid Player Name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If
        If txtEmpireName.Caption.Trim.Length < 3 Then
            MyBase.moUILib.AddNotification("Please enter a valid Empire Name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If
        'If txtPlayerName.Caption.Trim.Length < 3 Then
        '    MyBase.moUILib.AddNotification("Please enter a valid Player Name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return False
        'End If
        If txtEmpireName.Caption.Trim.Length > 20 OrElse txtPlayerName.Caption.Trim.Length > 20 Then
            MyBase.moUILib.AddNotification("Max length of Empire and Player name is 20 characters.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        If txtUserName.Caption.Trim.Length = 0 Then 'OrElse txtUserName.Caption.ToUpper = sUserName.ToUpper Then
            MyBase.moUILib.AddNotification("Please enter a User Name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If
        If txtPassword.Caption.Trim.Length = 0 Then 'OrElse txtPassword.Caption.ToUpper = sPassword.ToUpper Then
            MyBase.moUILib.AddNotification("Please enter a Password.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If
        If txtUserName.Caption.Trim.Length > 20 OrElse txtPassword.Caption.Trim.Length > 20 Then
            MyBase.moUILib.AddNotification("Max length of Username and Password fields is 20 characters.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        txtEmail.Caption = txtEmail.Caption.Trim
        Dim sEmail As String = txtEmail.Caption
        If sEmail Is Nothing = False AndAlso sEmail.Trim <> "" Then
            Dim bResult As Boolean = True
            If sEmail.Contains("@") = False Then bResult = False
            If sEmail.Contains(".") = False Then bResult = False
            If sEmail.Contains("<") = True Then bResult = False
            If sEmail.Contains(">") = True Then bResult = False
            If sEmail.Contains("(") = True Then bResult = False
            If sEmail.Contains(")") = True Then bResult = False
            If sEmail.Contains("[") = True Then bResult = False
            If sEmail.Contains("]") = True Then bResult = False
            If sEmail.Contains(";") = True Then bResult = False
            If sEmail.Contains(":") = True Then bResult = False
            If sEmail.Contains(",") = True Then bResult = False
            If sEmail.Contains("\") = True Then bResult = False
            If sEmail.Contains(" ") = True Then bResult = False

            'space, pound-sign
            If bResult = False Then
                MyBase.moUILib.AddNotification("Please enter a valid email address.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            Dim lAtIdx As Int32 = sEmail.LastIndexOf("@")
            Dim lDotIdx As Int32 = sEmail.LastIndexOf(".")
            If lDotIdx < lAtIdx Then bResult = False
            Dim sChrsAtEnd As String = sEmail.Substring(lDotIdx)
            If sChrsAtEnd.Length < 3 Then bResult = False
            Dim sChrsBeforeAt As String = sEmail.Substring(0, lAtIdx)
            If sChrsBeforeAt.Length = 0 Then bResult = False
            If sChrsBeforeAt.Contains("@") = True Then bResult = False
            If "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789".Contains(sEmail.Substring(0, 1)) = False Then bResult = False

            If bResult = False Then
                MyBase.moUILib.AddNotification("Please enter a valid email address.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
        End If

        'Character checker
        Dim sTemp As String = txtPlayerName.Caption & txtEmpireName.Caption
        For X As Int32 = 0 To sTemp.Length - 1
            Dim lChrVal As Int32 = Asc(Mid$(sTemp, X + 1, 1))
            If lChrVal <> 39 AndAlso lChrVal <> 44 AndAlso lChrVal <> 45 AndAlso lChrVal <> 46 AndAlso lChrVal <> 96 AndAlso lChrVal <> 32 AndAlso lChrVal <> 95 Then
                If (lChrVal > 64 AndAlso lChrVal < 91) = False AndAlso (lChrVal > 96 AndAlso lChrVal < 123) = False AndAlso (lChrVal > 47 AndAlso lChrVal < 58) = False Then
                    MyBase.moUILib.AddNotification("Empire or Player Name contains invalid characters. Alphabetical characters only.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
            End If
        Next X
        If txtPlayerName.Caption.Contains(" ") = True Then
            MyBase.moUILib.AddNotification("Player Name cannot have a space in it.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        If fraIconChooser.GetStartingIcon() = 0 Then
            MyBase.moUILib.AddNotification("Please select a valid Icon.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        Return True
    End Function

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		Dim frmLogin As frmLoginDlg = CType(goUILib.GetWindow("frmLoginDlg"), frmLoginDlg)
		If frmLogin Is Nothing = False Then
			frmLogin.cmdExit.Enabled = True
			frmLogin.cmdLogin.Enabled = True

			LoginScreen.UpdateButNoRender = False
		End If
	End Sub

	Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click

		If ValidateData() = False Then Return

		Dim yMsg(381) As Byte '126) As Byte
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(GlobalMessageCode.ePlayerInitialSetup).CopyTo(yMsg, lPos) : lPos += 2
		System.Text.ASCIIEncoding.ASCII.GetBytes(sUserName).CopyTo(yMsg, lPos) : lPos += 20
		System.Text.ASCIIEncoding.ASCII.GetBytes(sPassword).CopyTo(yMsg, lPos) : lPos += 20
		System.Text.ASCIIEncoding.ASCII.GetBytes(txtPlayerName.Caption).CopyTo(yMsg, lPos) : lPos += 20
		System.Text.ASCIIEncoding.ASCII.GetBytes(txtEmpireName.Caption).CopyTo(yMsg, lPos) : lPos += 20
		If optMale.Value = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 2
		lPos += 1
		System.BitConverter.GetBytes(fraIconChooser.GetStartingIcon()).CopyTo(yMsg, lPos) : lPos += 4

		System.Text.ASCIIEncoding.ASCII.GetBytes(txtUserName.Caption.ToUpper).CopyTo(yMsg, lPos) : lPos += 20
		System.Text.ASCIIEncoding.ASCII.GetBytes(txtPassword.Caption.ToUpper).CopyTo(yMsg, lPos) : lPos += 20

		System.Text.ASCIIEncoding.ASCII.GetBytes(txtEmail.Caption.Trim).CopyTo(yMsg, lPos) : lPos += 255
		MyBase.moUILib.SendMsgToOperator(yMsg)

		btnSubmit.Enabled = False
		btnCancel.Enabled = False
		MyBase.moUILib.AddNotification("Submitting to server for approval... Please wait...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
	End Sub

    Private Sub optFemale_Click() Handles optFemale.Click
        If mbLoading = True Then Return
        mbLoading = True
        optMale.Value = False
        mbLoading = False
    End Sub

    Private Sub optMale_Click() Handles optMale.Click
        If mbLoading = True Then Return
        mbLoading = True
        optFemale.Value = False
        mbLoading = False
    End Sub

    Public Sub HandleSetupMsg(ByRef yData() As Byte)
        Dim yResp As Byte = yData(2)

        btnSubmit.Enabled = True
        btnCancel.Enabled = True

        If yResp = 0 Then
            MyBase.moUILib.AddNotification("Setup Properties Accepted, please log in again.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            MyBase.moUILib.RemoveWindow(Me.ControlName)

            Dim frmLogin As frmLoginDlg = CType(goUILib.GetWindow("frmLoginDlg"), frmLoginDlg)
            If frmLogin Is Nothing = False Then
                frmLogin.cmdExit.Enabled = True
                frmLogin.cmdLogin.Enabled = True

                LoginScreen.UpdateButNoRender = False
            End If
        ElseIf yResp = 1 Then
            MyBase.moUILib.AddNotification("That Player Name is already taken, please choose another.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        ElseIf yResp = 2 Then
            MyBase.moUILib.AddNotification("That Empire Name is already taken, please choose another.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        ElseIf yResp = 3 Then
            MyBase.moUILib.AddNotification("Invalid Icon Selected, please choose another.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        ElseIf yResp = 4 Then
            MyBase.moUILib.AddNotification("That Username is already taken, please choose another.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        ElseIf yResp = 5 Then
            MyBase.moUILib.AddNotification("The player name entered contains questionable wording.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        ElseIf yResp = 6 Then
            MyBase.moUILib.AddNotification("The empire name entered contains questionable wording.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            MyBase.moUILib.AddNotification("Credentials failed for unknown reason. Please try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub frmPlayerSetup_OnRenderEnd() Handles Me.OnRenderEnd
        fraIconChooser.frmIconChooser_OnRenderEnd()
    End Sub

    Private Sub frmPlayerSetup_WindowClosed() Handles Me.WindowClosed
        If goSound Is Nothing = False Then
            goSound.MasterVolume = mlOriginalMaster
            goSound.UpdateSoundVolumes()
        End If
    End Sub
End Class