Option Strict On

Public Class frmPlayer
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
    Friend WithEvents txtPlayerName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtEmpireName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtRaceName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents cboList As System.Windows.Forms.ComboBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnOverride As System.Windows.Forms.Button
    Friend WithEvents cboOverride As System.Windows.Forms.ComboBox
	Friend WithEvents txtOverride As System.Windows.Forms.TextBox
	Friend WithEvents txtID As System.Windows.Forms.TextBox
	Friend WithEvents btnResetTutorial As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents grpClaimable As System.Windows.Forms.GroupBox
    Friend WithEvents txtClaimPlayer As System.Windows.Forms.TextBox
    Friend WithEvents txtClaimTypeID As System.Windows.Forms.TextBox
    Friend WithEvents txtClaimID As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtOfferCode As System.Windows.Forms.TextBox
    Friend WithEvents btnAddClaimable As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtClaimItemName As System.Windows.Forms.TextBox
    Friend WithEvents btnReset As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtPlayerName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtEmpireName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtRaceName = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.cboList = New System.Windows.Forms.ComboBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnReset = New System.Windows.Forms.Button
        Me.btnOverride = New System.Windows.Forms.Button
        Me.cboOverride = New System.Windows.Forms.ComboBox
        Me.txtOverride = New System.Windows.Forms.TextBox
        Me.txtID = New System.Windows.Forms.TextBox
        Me.btnResetTutorial = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.grpClaimable = New System.Windows.Forms.GroupBox
        Me.btnAddClaimable = New System.Windows.Forms.Button
        Me.txtOfferCode = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtClaimID = New System.Windows.Forms.TextBox
        Me.txtClaimTypeID = New System.Windows.Forms.TextBox
        Me.txtClaimPlayer = New System.Windows.Forms.TextBox
        Me.txtClaimItemName = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.grpClaimable.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtPlayerName
        '
        Me.txtPlayerName.Location = New System.Drawing.Point(92, 2)
        Me.txtPlayerName.MaxLength = 20
        Me.txtPlayerName.Name = "txtPlayerName"
        Me.txtPlayerName.Size = New System.Drawing.Size(100, 20)
        Me.txtPlayerName.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Player Name:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Empire Name:"
        '
        'txtEmpireName
        '
        Me.txtEmpireName.Location = New System.Drawing.Point(92, 26)
        Me.txtEmpireName.MaxLength = 20
        Me.txtEmpireName.Name = "txtEmpireName"
        Me.txtEmpireName.Size = New System.Drawing.Size(100, 20)
        Me.txtEmpireName.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(12, 86)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(76, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Race Name:"
        '
        'txtRaceName
        '
        Me.txtRaceName.Location = New System.Drawing.Point(92, 84)
        Me.txtRaceName.MaxLength = 20
        Me.txtRaceName.Name = "txtRaceName"
        Me.txtRaceName.Size = New System.Drawing.Size(100, 20)
        Me.txtRaceName.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(12, 130)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(76, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Username:"
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(92, 128)
        Me.txtUserName.MaxLength = 20
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(100, 20)
        Me.txtUserName.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(12, 156)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 16)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Password:"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(92, 154)
        Me.txtPassword.MaxLength = 20
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(100, 20)
        Me.txtPassword.TabIndex = 8
        '
        'cboList
        '
        Me.cboList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboList.Location = New System.Drawing.Point(12, 220)
        Me.cboList.Name = "cboList"
        Me.cboList.Size = New System.Drawing.Size(180, 21)
        Me.cboList.TabIndex = 10
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(58, 186)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(88, 23)
        Me.btnNew.TabIndex = 11
        Me.btnNew.Text = "Create New"
        '
        'btnReset
        '
        Me.btnReset.Enabled = False
        Me.btnReset.Location = New System.Drawing.Point(71, 252)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(75, 23)
        Me.btnReset.TabIndex = 12
        Me.btnReset.Text = "Reset"
        '
        'btnOverride
        '
        Me.btnOverride.Location = New System.Drawing.Point(58, 334)
        Me.btnOverride.Name = "btnOverride"
        Me.btnOverride.Size = New System.Drawing.Size(88, 23)
        Me.btnOverride.TabIndex = 13
        Me.btnOverride.Text = "Override"
        Me.btnOverride.UseVisualStyleBackColor = True
        '
        'cboOverride
        '
        Me.cboOverride.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboOverride.Items.AddRange(New Object() {"Credits", "Engineer", "Colonists", "Enlisted", "Officers"})
        Me.cboOverride.Location = New System.Drawing.Point(12, 281)
        Me.cboOverride.Name = "cboOverride"
        Me.cboOverride.Size = New System.Drawing.Size(180, 21)
        Me.cboOverride.TabIndex = 14
        '
        'txtOverride
        '
        Me.txtOverride.Location = New System.Drawing.Point(48, 308)
        Me.txtOverride.Name = "txtOverride"
        Me.txtOverride.Size = New System.Drawing.Size(108, 20)
        Me.txtOverride.TabIndex = 15
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(39, 420)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(49, 20)
        Me.txtID.TabIndex = 16
        '
        'btnResetTutorial
        '
        Me.btnResetTutorial.Location = New System.Drawing.Point(94, 420)
        Me.btnResetTutorial.Name = "btnResetTutorial"
        Me.btnResetTutorial.Size = New System.Drawing.Size(102, 23)
        Me.btnResetTutorial.TabIndex = 17
        Me.btnResetTutorial.Text = "Reset Tutorial"
        Me.btnResetTutorial.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 425)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(21, 13)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "ID:"
        '
        'grpClaimable
        '
        Me.grpClaimable.Controls.Add(Me.Label11)
        Me.grpClaimable.Controls.Add(Me.Label10)
        Me.grpClaimable.Controls.Add(Me.Label9)
        Me.grpClaimable.Controls.Add(Me.txtClaimItemName)
        Me.grpClaimable.Controls.Add(Me.txtClaimPlayer)
        Me.grpClaimable.Controls.Add(Me.txtClaimTypeID)
        Me.grpClaimable.Controls.Add(Me.txtClaimID)
        Me.grpClaimable.Controls.Add(Me.Label8)
        Me.grpClaimable.Controls.Add(Me.Label7)
        Me.grpClaimable.Controls.Add(Me.txtOfferCode)
        Me.grpClaimable.Controls.Add(Me.btnAddClaimable)
        Me.grpClaimable.Location = New System.Drawing.Point(200, 4)
        Me.grpClaimable.Name = "grpClaimable"
        Me.grpClaimable.Size = New System.Drawing.Size(200, 180)
        Me.grpClaimable.TabIndex = 19
        Me.grpClaimable.TabStop = False
        Me.grpClaimable.Text = "Claimable"
        '
        'btnAddClaimable
        '
        Me.btnAddClaimable.Location = New System.Drawing.Point(60, 148)
        Me.btnAddClaimable.Name = "btnAddClaimable"
        Me.btnAddClaimable.Size = New System.Drawing.Size(75, 23)
        Me.btnAddClaimable.TabIndex = 0
        Me.btnAddClaimable.Text = "Add"
        Me.btnAddClaimable.UseVisualStyleBackColor = True
        '
        'txtOfferCode
        '
        Me.txtOfferCode.Location = New System.Drawing.Point(94, 19)
        Me.txtOfferCode.Name = "txtOfferCode"
        Me.txtOfferCode.Size = New System.Drawing.Size(100, 20)
        Me.txtOfferCode.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 22)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(58, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "OfferCode:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(18, 13)
        Me.Label8.TabIndex = 3
        Me.Label8.Text = "ID"
        '
        'txtClaimID
        '
        Me.txtClaimID.Location = New System.Drawing.Point(94, 45)
        Me.txtClaimID.Name = "txtClaimID"
        Me.txtClaimID.Size = New System.Drawing.Size(100, 20)
        Me.txtClaimID.TabIndex = 4
        '
        'txtClaimTypeID
        '
        Me.txtClaimTypeID.Location = New System.Drawing.Point(94, 71)
        Me.txtClaimTypeID.Name = "txtClaimTypeID"
        Me.txtClaimTypeID.Size = New System.Drawing.Size(100, 20)
        Me.txtClaimTypeID.TabIndex = 5
        '
        'txtClaimPlayer
        '
        Me.txtClaimPlayer.Location = New System.Drawing.Point(94, 97)
        Me.txtClaimPlayer.Name = "txtClaimPlayer"
        Me.txtClaimPlayer.Size = New System.Drawing.Size(100, 20)
        Me.txtClaimPlayer.TabIndex = 6
        '
        'txtClaimItemName
        '
        Me.txtClaimItemName.Location = New System.Drawing.Point(94, 123)
        Me.txtClaimItemName.Name = "txtClaimItemName"
        Me.txtClaimItemName.Size = New System.Drawing.Size(100, 20)
        Me.txtClaimItemName.TabIndex = 7
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 74)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(42, 13)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "TypeID"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 100)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(47, 13)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "PlayerID"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 126)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(58, 13)
        Me.Label11.TabIndex = 10
        Me.Label11.Text = "ItemName:"
        '
        'frmPlayer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(412, 452)
        Me.Controls.Add(Me.grpClaimable)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnResetTutorial)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.txtOverride)
        Me.Controls.Add(Me.cboOverride)
        Me.Controls.Add(Me.btnOverride)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.cboList)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtUserName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtRaceName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtEmpireName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtPlayerName)
        Me.Name = "frmPlayer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Player Hack"
        Me.grpClaimable.ResumeLayout(False)
        Me.grpClaimable.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

	Private mlItemData() As Int32

	Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
		'ok, create a new player...
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand
		Dim oData As OleDb.OleDbDataReader

		sSQL = "INSERT INTO tblPlayer (PlayerName, RaceName, EmpireName, PlayerUserName, PlayerPassword, SenateID, " & _
		  "CommEncryptLevel, EmpireTaxRate, LastViewedID, LastViewedTypeID, Credits, AccountStatus) VALUES ('" & _
		  Replace$(txtPlayerName.Text, "'", "''") & "', '" & Replace$(txtRaceName.Text, "'", "''") & "', '" & _
		  Replace$(txtEmpireName.Text, "'", "''") & "', '" & Replace$(txtUserName.Text, "'", "''") & "', '" & _
		  Replace$(txtPassword.Text, "'", "''") & "', 0, 0, 0, 0, 0, 100000, 1)"
		oComm = New OleDb.OleDbCommand(sSQL, goCN)
		oComm.ExecuteNonQuery()
		oComm = Nothing

		sSQL = "SELECT * FROM tblPlayer WHERE PlayerID IN (SELECT MAX(PlayerID) FROM tblPlayer WHERE PlayerUserName = '" & Replace$(txtUserName.Text, "'", "''") & "')"
		oComm = New OleDb.OleDbCommand(sSQL, goCN)
		oData = oComm.ExecuteReader(CommandBehavior.Default)

        If oData.Read Then

            ' check the player id will be within the player redim
            Dim i As Int32 = CInt(oData("PlayerID"))
            CheckPlayerExtent(i)


            goPlayer(i) = New Player()
            With goPlayer(i)
                .ServerIndex = i
                .CommEncryptLevel = CShort(oData("CommEncryptLevel"))
                .EmpireName = StringToBytes(CStr(oData("EmpireName")))
                .EmpireTaxRate = CByte(oData("EmpireTaxRate"))
                .ObjectID = i
                .ObjTypeID = ObjectType.ePlayer
                '.oSenate = GetEpicaSenate(CInt(oData("SenateID")))
                .PlayerName = StringToBytes(CStr(oData("PlayerName")))
                StringToBytes(CStr(oData("PlayerPassword"))).CopyTo(.PlayerPassword, 0)
                StringToBytes(CStr(oData("PlayerUserName"))).CopyTo(.PlayerUserName, 0)
                .RaceName = StringToBytes(CStr(oData("RaceName")))
                .lLastViewedEnvir = CInt(oData("LastViewedID"))
                .iLastViewedEnvirType = CShort(oData("LastViewedTypeID"))
                glPlayerIdx(glPlayerUB) = .ObjectID
                .blCredits = CLng(oData("Credits"))
                .AccountStatus = CInt(oData("AccountStatus"))
            End With

            If goPlayer(i).SaveObject(False) = True Then
                Dim yMsg() As Byte = goMsgSys.GetAddObjectMessage(goPlayer(i), GlobalMessageCode.eAddObjectCommand)
                goMsgSys.BroadcastToDomains(yMsg)
            End If
        End If

        oData.Close()
		oComm.Dispose()
		oData = Nothing
		oComm = Nothing

		txtUserName.Text = ""
		txtPassword.Text = ""
		txtPlayerName.Text = ""
		txtRaceName.Text = ""
		txtEmpireName.Text = ""

		RefillCombo()

	End Sub

	Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
		If cboList.SelectedIndex > -1 Then
			Dim oPlayer As Player = GetEpicaPlayer(mlItemData(cboList.SelectedIndex))
			oPlayer.PlayerIsDead = True
		End If
	End Sub

	Private Sub frmPlayer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		RefillCombo()
	End Sub

	Private Sub RefillCombo()
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand
		Dim oData As OleDb.OleDbDataReader
		Dim lUB As Int32 = -1

		cboList.Items.Clear()

		sSQL = "SELECT PlayerID, PlayerName FROM tblPlayer ORDER BY PlayerName"

		oComm = New OleDb.OleDbCommand(sSQL, goCN)
		oData = oComm.ExecuteReader(CommandBehavior.Default)
		While oData.Read
			cboList.Items.Add(oData("PlayerName"))
			lUB += 1
			ReDim Preserve mlItemData(lUB)
			mlItemData(lUB) = CInt(oData("PlayerID"))
		End While
		oData.Close()
		oData = Nothing
		oComm.Dispose()
		oComm = Nothing
	End Sub

	Private Sub btnOverride_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOverride.Click
		If cboList.SelectedIndex < 0 Then Return

		Dim oPlayer As Player = GetEpicaPlayer(mlItemData(cboList.SelectedIndex))
		Dim X As Int32
		Dim lVal As Int64 = CLng(Val(txtOverride.Text))
		If oPlayer Is Nothing Then Return

		Select Case cboOverride.Text.ToLower
			Case "colonists"
				For X = 0 To glColonyUB
					If glColonyIdx(X) <> -1 AndAlso goColony(X).Owner.ObjectID = oPlayer.ObjectID Then
						goColony(X).Population += CInt(lVal)
					End If
				Next X
			Case "enlisted"
				For X = 0 To glColonyUB
					If glColonyIdx(X) <> -1 AndAlso goColony(X).Owner.ObjectID = oPlayer.ObjectID Then
						goColony(X).ColonyEnlisted += CInt(lVal)
					End If
				Next X
			Case "officers"
				For X = 0 To glColonyUB
					If glColonyIdx(X) <> -1 AndAlso goColony(X).Owner.ObjectID = oPlayer.ObjectID Then
						goColony(X).ColonyOfficers += CInt(lVal)
					End If
				Next X
			Case "credits"
				oPlayer.blCredits += lVal
			Case "engineer"
		End Select
	End Sub

	Private Sub btnResetTutorial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetTutorial.Click
		'Ok, call the stored procedure...
		Dim lPlayerID As Int32 = CInt(Val(txtID.Text))

		If lPlayerID < 1 OrElse lPlayerID = gl_HARDCODE_PIRATE_PLAYER_ID Then Return

		Try
			Dim oSP As New OleDb.OleDbCommand("sp_ResetPlayerTutorial", goCN)
			oSP.Parameters.Add(New OleDb.OleDbParameter("@PlayerID", lPlayerID))
			oSP.ExecuteNonQuery()
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "ResetTutorial Failed: " & ex.Message)
			Return
		End Try

		Dim lCurUB As Int32 = -1
		If glUnitDefIdx Is Nothing = False Then lCurUB = Math.Min(glUnitDefIdx.GetUpperBound(0), glUnitDefUB)
		For X As Int32 = 0 To lCurUB
			Try
				If glUnitIdx(X) > -1 Then
					Dim oDef As Epica_Entity_Def = goUnitDef(X)
					If oDef Is Nothing = False Then
						If oDef.OwnerID = lPlayerID Then
							oDef.DeleteMe()
						End If
					End If
				End If
			Catch
			End Try
		Next X

		lCurUB = -1
		If glFacilityDefIdx Is Nothing = False Then lCurUB = Math.Min(glFacilityDefIdx.GetUpperBound(0), glFacilityDefUB)
		For X As Int32 = 0 To lCurUB
			Try
				If glFacilityDefIdx(X) > -1 Then
					Dim oDef As Epica_Entity_Def = goFacilityDef(X)
					If oDef Is Nothing = False Then
						If oDef.OwnerID = lPlayerID Then
							oDef.DeleteMe()
						End If
					End If
				End If
			Catch
			End Try
		Next X

		LogEvent(LogEventType.Informational, "ResetTutorial for " & lPlayerID & " succeeded.")
	End Sub

    Private Sub btnAddClaimable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddClaimable.Click
        Try
            Dim oClaimable As New Claimable()
            With oClaimable
                .iTypeID = CShort(txtClaimTypeID.Text)
                .lID = CInt(txtClaimID.Text)
                .lOfferCode = CInt(txtOfferCode.Text)
                .lPlayerID = CInt(txtClaimPlayer.Text)
                .yClaimFlag = 0
                .yItemName = StringToBytes(txtClaimItemName.Text)
            End With
            Dim oPlayer As Player = GetEpicaPlayer(oClaimable.lPlayerID)
            If oPlayer Is Nothing = False Then
                ReDim Preserve oPlayer.Claimables(oPlayer.Claimables.GetUpperBound(0) + 1)
                oPlayer.Claimables(oPlayer.Claimables.GetUpperBound(0)) = oClaimable
                oClaimable.SaveObject()
            End If
            LogEvent(LogEventType.Informational, "Claimable added to " & oPlayer.sPlayerNameProper)
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, ex.Message)
        End Try
    End Sub
End Class
