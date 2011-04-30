Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Login Dialog
Public Class frmLoginDlg
    Inherits UIWindow

    Private WithEvents txtUserName As UITextBox
    Private lblUserName As UILabel
    Private lblPassword As UILabel
    Private WithEvents txtPassword As UITextBox
    Public WithEvents cmdLogin As UIButton
    Public WithEvents cmdExit As UIButton 
    Private lblSignon As UILabel

    Private WithEvents chkAliased As UICheckBox
    Private lblAliasUserName As UILabel
    Private lblAliasPassword As UILabel
	Private WithEvents txtAliasUserName As UITextBox
	Private WithEvents txtAliasPassword As UITextBox

	Private msActPassword As String

    Public Shared lStatusTop As Int32 = -1

    Public Event StartLogin(ByVal sUserName As String, ByVal sPassword As String, ByVal bAlias As Boolean, ByVal sAliasUserName As String, ByVal sAliasPassword As String, ByRef frmLogin As frmLoginDlg)
    Public Event ExitProgram()

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        Dim oINI As InitFile = New InitFile()

        'frmSignon initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eLoginScreen
            .ControlName = "frmLoginDlg"
            .Left = CInt(oUILib.oDevice.PresentationParameters.BackBufferWidth / 2) - 144
            .Top = CInt(oUILib.oDevice.PresentationParameters.BackBufferHeight / 2) - 126
            .Width = 288
            .Height = 170
            .Enabled = True
            .Visible = False
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
        End With

		lStatusTop = Me.Top + Me.Height + 5

        Debug.Write(Me.ControlName & " Newed" & vbCrLf)
        'txtUserName initial props
        txtUserName = New UITextBox(oUILib)
        With txtUserName
            .ControlName = "txtUserName"
            .Left = 102
            .Top = 42
            .Width = 171
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = oINI.GetString("SIGNON", "LastUserName", "")
            '.ForeColor = System.Drawing.Color.FromArgb(-16777216)
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Courier New", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            '.BackColorEnabled = System.Drawing.Color.FromArgb(-1)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
            .MaxLength = 20
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
            .HasFocus = True
            '.ToolTipText = "Type in your User Name Here"
            .bFilterBadWords = False
        End With
        If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
        goUILib.FocusedControl = txtUserName
        Me.AddChild(CType(txtUserName, UIControl))

        'lblUserName initial props
        lblUserName = New UILabel(oUILib)
        With lblUserName
            .ControlName = "lblUserName"
            .Left = 13
            .Top = 42
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "User Name:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblUserName, UIControl))

        'lblPassword initial props
        lblPassword = New UILabel(oUILib)
        With lblPassword
            .ControlName = "lblPassword"
            .Left = 13
            .Top = 72
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Password:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblPassword, UIControl))

        'txtPassword initial props
        txtPassword = New UITextBox(oUILib)
        With txtPassword
            .ControlName = "txtPassword"
            .Left = 102
            .Top = 72
            .Width = 171
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = ""
            '.ForeColor = System.Drawing.Color.FromArgb(-16777216)
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Courier New", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            '.BackColorEnabled = System.Drawing.Color.FromArgb(-1)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
            .MaxLength = 20
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
            .PasswordChar = "*"
            .bFilterBadWords = False
            '.ToolTipText = "Type in your Password Here"
        End With
        Me.AddChild(CType(txtPassword, UIControl))

        'cmdLogin initial props
        cmdLogin = New UIButton(oUILib)
        With cmdLogin
            .ControlName = "cmdLogin"
            .Left = 28
            .Top = 125
            .Width = 100
            .Height = 29
            .Enabled = True
            .Visible = True
            .Caption = "Login"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        AddHandler cmdLogin.Click, AddressOf LoginClick
        Me.AddChild(CType(cmdLogin, UIControl))

        'cmdExit initial props
        cmdExit = New UIButton(oUILib)
        With cmdExit
            .ControlName = "cmdExit"
            .Left = 162
            .Top = 125
            .Width = 100
            .Height = 29
            .Enabled = True
            .Visible = True
            .Caption = "Exit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        AddHandler cmdExit.Click, AddressOf ExitClick
        Me.AddChild(CType(cmdExit, UIControl))

        'lblSignon initial props
        lblSignon = New UILabel(oUILib)
        With lblSignon
            .ControlName = "lblSignon"
            .Left = 0
            .Top = 3
            .Width = 288
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Beyond Protocol Sign On"
            '.ForeColor = System.Drawing.Color.FromArgb(-16777216)
            '.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblSignon, UIControl))

        'lblAliasUserName initial props
        lblAliasUserName = New UILabel(oUILib)
        With lblAliasUserName
            .ControlName = "lblAliasUserName"
            .Left = 13
            .Top = 100
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Alias Name:"
            .ForeColor = Color.Gray  'muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblAliasUserName, UIControl))

        'txtAliasUserName initial props
        txtAliasUserName = New UITextBox(oUILib)
        With txtAliasUserName
            .ControlName = "txtAliasUserName"
            .Left = 102
            .Top = 100
            .Width = 171
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Courier New", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
            .ToolTipText = "Alias User Name that was given by the player you wish to alias. Leave this blank if you do not wish to use an alias."
            .bFilterBadWords = False
        End With
        Me.AddChild(CType(txtAliasUserName, UIControl))
        'lblAliasPassword initial props
        lblAliasPassword = New UILabel(oUILib)
        With lblAliasPassword
            .ControlName = "lblAliasPassword"
            .Left = 13
            .Top = 155
            .Width = 100
            .Height = 20
            .Enabled = True
            .Visible = False
            .Caption = "Password:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblAliasPassword, UIControl))

        'txtAliasPassword initial props
        txtAliasPassword = New UITextBox(oUILib)
        With txtAliasPassword
            .ControlName = "txtAliasPassword"
            .Left = 102
            .Top = 155
            .Width = 171
            .Height = 20
            .Enabled = True
            .Visible = False
            .Caption = "PASSWORD"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Courier New", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
            .PasswordChar = "*"
            .ToolTipText = "Alias password supplied by the player who is to be aliased."
            .bFilterBadWords = False
        End With
        Me.AddChild(CType(txtAliasPassword, UIControl))

        'chkAliased initial props
        'chkAliased = New UICheckBox(oUILib)
        'With chkAliased
        '    .ControlName = "chkAliased"
        '    .Left = 100
        '    .Top = 100
        '    .Width = 100
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Alias Player"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat) 
        '    .Value = False
        '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkAliased, UIControl))

        'now set our focus...
        If txtUserName.Caption.Length = 0 Then
            If oUILib.FocusedControl Is Nothing = False Then oUILib.FocusedControl.HasFocus = False
            oUILib.FocusedControl = txtUserName
            txtUserName.HasFocus = True
        Else
            If oUILib.FocusedControl Is Nothing = False Then oUILib.FocusedControl.HasFocus = False
            oUILib.FocusedControl = txtPassword
            txtPassword.HasFocus = True
        End If

    End Sub

	'Private Sub LoginClick()
	'	Dim oThread As New Threading.Thread(AddressOf LoginClick_ac)
	'	oThread.Start()
	'End Sub 
    Private Sub LoginClick(ByVal sName As String)
        Dim bAlias As Boolean = False
        cmdLogin.Enabled = False
        If txtUserName.Caption.Trim.Length = 0 Then
            goUILib.AddNotification("Please enter a valid UserName.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            cmdLogin.Enabled = True
            Return
        End If
        If txtPassword.Caption.Trim.Length = 0 Then
            goUILib.AddNotification("Please enter a valid Password.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            cmdLogin.Enabled = True
            Return
        End If
        If txtAliasUserName.Caption <> "" Then
            If txtAliasUserName.Caption.Trim.Length = 0 Then
                goUILib.AddNotification("Please enter a valid Alias Username.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                cmdLogin.Enabled = True
                Return
            End If
            If txtAliasPassword.Caption.Trim.Length = 0 Then
                goUILib.AddNotification("Please enter a valid Alias Password.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                cmdLogin.Enabled = True
                Return
            End If
            bAlias = True
        End If

        RaiseEvent StartLogin(txtUserName.Caption, txtPassword.Caption, bAlias, txtAliasUserName.Caption, txtAliasPassword.Caption, Me)
        'cmdLogin.Enabled = True
    End Sub

	Private Sub ExitClick(ByVal sName As String)
		Me.Enabled = False
		RaiseEvent ExitProgram()
	End Sub

    Private Sub txtUserName_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUserName.OnKeyPress
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\TextWriter.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Vector3.Empty, Vector3.Empty)
    End Sub

    Private Sub txtPassword_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPassword.OnKeyPress
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\TextWriter.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Vector3.Empty, Vector3.Empty)
    End Sub

    Private Sub txtUserName_OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserName.OnKeyUp
        If e.KeyCode = Keys.Enter Then
			If cmdLogin.Enabled = True Then LoginClick(cmdLogin.ControlName)
		ElseIf e.KeyCode = Keys.Escape Then
			Dim oOptions As frmOptions = New frmOptions(goUILib)
			oOptions.Visible = True
			oOptions = Nothing
			e.Handled = True
		End If
	End Sub

    Private Sub txtPassword_OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.OnKeyUp
        If e.KeyCode = Keys.Enter Then
			If cmdLogin.Enabled = True Then LoginClick(cmdLogin.ControlName)
        ElseIf e.KeyCode = Keys.Escape Then
            Dim oOptions As frmOptions = New frmOptions(goUILib)
            oOptions.Visible = True
            oOptions = Nothing
            e.Handled = True
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Debug.Write(Me.ControlName & " Finalized" & vbCrLf)
        MyBase.Finalize()
    End Sub

    'Private Sub chkAliased_Click() Handles chkAliased.Click
    '    'lblAliasPassword.Visible = chkAliased.Value
    '    lblAliasUserName.Visible = chkAliased.Value
    '    'txtAliasPassword.Visible = chkAliased.Value
    '    txtAliasUserName.Visible = chkAliased.Value
    '    If chkAliased.Value = True Then
    '        cmdLogin.Top = 188
    '        cmdExit.Top = 188
    '        Me.Height = 230
    '    Else
    '        cmdLogin.Top = 125
    '        cmdExit.Top = 125
    '        Me.Height = 170
    '    End If
    'End Sub

    'Private Sub txtAliasPassword_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAliasPassword.OnKeyPress
    '	If goSound Is Nothing = False Then goSound.StartSound("UserInterface\TextWriter.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Vector3.Empty, Vector3.Empty)
    'End Sub

    'Private Sub txtAliasPassword_OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAliasPassword.OnKeyUp
    '	If e.KeyCode = Keys.Enter Then
    '		LoginClick(cmdLogin.ControlName)
    '	ElseIf e.KeyCode = Keys.Escape Then
    '		Dim oOptions As frmOptions = New frmOptions(goUILib)
    '		oOptions.Visible = True
    '		oOptions = Nothing
    '		e.Handled = True
    '	End If
    'End Sub

	Private Sub txtAliasUserName_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAliasUserName.OnKeyPress
		If goSound Is Nothing = False Then goSound.StartSound("UserInterface\TextWriter.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Vector3.Empty, Vector3.Empty)
	End Sub

	Private Sub txtAliasUserName_OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAliasUserName.OnKeyUp
		If e.KeyCode = Keys.Enter Then
			LoginClick(cmdLogin.ControlName)
		ElseIf e.KeyCode = Keys.Escape Then
			Dim oOptions As frmOptions = New frmOptions(goUILib)
			oOptions.Visible = True
			oOptions = Nothing
			e.Handled = True
		End If
	End Sub

    Private Shared mlHighlightBall As Int32 = 0
    Private Shared mfActualHighlightBall As Single = 0.0F
    Private Shared mrctBalls() As Rectangle = Nothing
    Private Shared swHighlight As Stopwatch = Nothing
    Public Shared Sub RenderInProgress()
        'If cmdLogin.Enabled = False Then

        If mrctBalls Is Nothing = True Then
            swHighlight = Stopwatch.StartNew
            ReDim mrctBalls(7)
            Dim lCntrPtX As Int32 = 0 'Me.Width \ 2
            Dim lCntrPtY As Int32 = 200 'cmdLogin.Top + cmdLogin.Height \ 2
            mrctBalls(0) = New Rectangle(lCntrPtX, lCntrPtY - 12, 8, 8)      'top
            mrctBalls(1) = New Rectangle(lCntrPtX + 6, lCntrPtY - 6, 8, 8)      'top - right
            mrctBalls(2) = New Rectangle(lCntrPtX + 12, lCntrPtY, 8, 8)         'right
            mrctBalls(3) = New Rectangle(lCntrPtX + 6, lCntrPtY + 6, 8, 8)      'bottom - right
            mrctBalls(4) = New Rectangle(lCntrPtX, lCntrPtY + 12, 8, 8)      'bottom
            mrctBalls(5) = New Rectangle(lCntrPtX - 6, lCntrPtY + 6, 8, 8)      'bottom - left
            mrctBalls(6) = New Rectangle(lCntrPtX - 12, lCntrPtY, 8, 8)         'left
            mrctBalls(7) = New Rectangle(lCntrPtX - 6, lCntrPtY - 6, 8, 8)      'top - left
        End If

        If swHighlight.ElapsedMilliseconds > 60 Then
            mfActualHighlightBall += 1.0F
        Else
            mfActualHighlightBall += (swHighlight.ElapsedMilliseconds / 60.0F)
        End If
        swHighlight.Stop()
        swHighlight.Reset()
        swHighlight.Start()

        'ok, a login is in progress, we will do our render thing
        If mfActualHighlightBall > 7 Then mfActualHighlightBall = 0
        mlHighlightBall = CInt(mfActualHighlightBall)
        If mlHighlightBall > 7 Then mlHighlightBall = 0

        If goUILib.oInterfaceTexture Is Nothing Then Return
        Try
            Dim srcRect As Rectangle = grc_UI(elInterfaceRectangle.eSphere)
            'goUILib.Pen.Begin(SpriteFlags.AlphaBlend)
            goUILib.BeginPenSprite(SpriteFlags.AlphaBlend)
            For X As Int32 = mlHighlightBall To mlHighlightBall + 7
                Dim lIdx As Int32 = X
                Dim lDistFromBall As Int32 = lIdx - mlHighlightBall
                If lIdx > 7 Then lIdx -= 8
                If lIdx < 0 Then lIdx = 0
                If lIdx > 7 Then lIdx = 0

                Dim lAlpha As Int32 = CInt(((lDistFromBall / 8.0F)) * 255)
                Dim ptPos As Point = New Point(goUILib.oDevice.PresentationParameters.BackBufferWidth \ 2, goUILib.oDevice.PresentationParameters.BackBufferHeight \ 2)
                ptPos.X += mrctBalls(lIdx).Location.X
                ptPos.Y += mrctBalls(lIdx).Location.Y
                ptPos.X *= 2
                ptPos.Y *= 2
                goUILib.Pen.Draw2D(goUILib.oInterfaceTexture, srcRect, mrctBalls(lIdx), New Point(8, 8), 0, ptPos, System.Drawing.Color.FromArgb(lAlpha, 255, 255, 255))

            Next X
            'goUILib.Pen.End()
            goUILib.EndPenSprite()
        Catch
        End Try

        'End If
    End Sub
 
    Private Sub txtAliasUserName_TextChanged() Handles txtAliasUserName.TextChanged
        If txtAliasUserName.Caption = "" Then
            lblAliasUserName.ForeColor = Color.Gray
        Else
            lblAliasUserName.ForeColor = muSettings.InterfaceBorderColor
        End If
    End Sub

    Public Sub ReEnableStuff()
        cmdLogin.Enabled = True
        cmdExit.Enabled = True
        txtUserName.Enabled = True
        txtPassword.Enabled = True

    End Sub
End Class
