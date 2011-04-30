Option Strict On
Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmEnvirShortcuts
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents btnClose As UIButton
    Private WithEvents lstEnvir As UIListBox
    Private WithEvents btnNew As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnRename As UIButton
    Private txtShortcutName As UITextBox
    Private Shared miLastListIndex As Int32 = -1
    Private Shared miLastScrlBarVal As Int32 = -1
    Private miIndex As Int32 = -1

    Private Structure EnvirShortcut
        Dim iIndex As Int32
        Dim iEnvirTypeID As Int16
        Dim lEnvirID As Int32
        Dim iLocX As Int32
        'Dim iLocY As Int32
        Dim iLocZ As Int32
        Dim sName As String
        Dim bNoSave As Boolean
    End Structure


    Private mbLoading As Boolean = True
    Private Shared mbShortcutsLoaded As Boolean = False
    Private moShortCuts() As EnvirShortcut

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
        Dim ofrmED As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)

        'frmEnvirList initial props
        With Me
            .ControlName = "frmEnvirShortcuts"
            .Width = 211
            .Height = 275
            .Left = ofrmED.Left - .Width
            .Top = ofrmED.Top

            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Enabled = True
            .Visible = True
            .Moveable = False
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 4
            .Top = 1
            .Width = 160
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Environment Shortcuts"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lstEnvir initial props
        lstEnvir = New UIListBox(oUILib)
        With lstEnvir
            .ControlName = "lstEnvir"
            .Left = 2
            .Top = 27
            .Width = Me.Width - 4
            .Height = Me.Height - .Top - 27
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstEnvir, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'btnNew initial props
        btnNew = New UIButton(oUILib)
        With btnNew
            .ControlName = "btnNew"
            .Width = 44
            .Height = 24
            .Left = 2
            .Top = Me.Height - .Height - 2
            .Enabled = True
            .Visible = True
            .Caption = "New"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnNew, UIControl))

        'btnRename initial props
        btnRename = New UIButton(oUILib)
        With btnRename
            .ControlName = "btnRename"
            .Width = 65
            .Height = 24
            .Left = btnNew.Left + btnNew.Width + 1
            .Top = btnNew.Top
            .Enabled = True
            .Visible = True
            .Caption = "Rename"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRename, UIControl))

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Width = 60
            .Height = 24
            .Left = btnRename.Left + btnRename.Width + 1
            .Top = btnNew.Top
            .Enabled = True
            .Visible = True
            .Caption = "Delete"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDelete, UIControl))

        ''txtName initial props
        'txtShortcutName = New UITextBox(oUILib)
        'With txtShortcutName
        '    .ControlName = "txtShortcutName"
        '    .Left = btnDelete.Left + btnDelete.Width + 1
        '    .Width = Me.Width - .Left - 2
        '    .Height = 18
        '    .Top = btnNew.Top
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(4, DrawTextFormat)
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 20
        '    .BorderColor = muSettings.InterfaceBorderColor
        'End With
        'Me.AddChild(CType(txtShortcutName, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
        LoadShortcuts()
    End Sub

    Private Sub LoadShortcuts()
        Dim ShortcutUB As Int32 = -1
        ReDim moShortCuts(ShortcutUB)
        LoadShortcutFile()

        'Popuate List
        lstEnvir.Clear()
        ShortcutUB = moShortCuts.GetUpperBound(0)
        If ShortcutUB = -1 Then Return

        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        For X As Int32 = 0 To ShortcutUB
            Dim lIdx As Int32 = -1

            Dim sName As String = moShortCuts(X).sName

            For Y As Int32 = 0 To lSortedUB
                Dim sOtherName As String = moShortCuts(lSorted(Y)).sName
                If sOtherName > sName Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedUB += 1
            ReDim Preserve lSorted(lSortedUB)
            If lIdx = -1 Then
                lSorted(lSortedUB) = X
            Else
                For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                    lSorted(Y) = lSorted(Y - 1)
                Next Y
                lSorted(lIdx) = X
            End If
        Next X

        For X As Int32 = 0 To ShortcutUB
            With moShortCuts(lSorted(X))
                Dim sText As String = .sName
                lstEnvir.AddItem(sText, False)
                .iIndex = lstEnvir.NewIndex
            End With
        Next X
        lstEnvir.RestorePosition()
        If lstEnvir.ListIndex = -1 AndAlso lstEnvir.ScrollBarValue = 0 Then
            lstEnvir.ListIndex = miLastListIndex
            lstEnvir.ScrollBarValue = miLastScrlBarVal
        End If
    End Sub

    Private Sub LoadShortcutFile()
        If goCurrentPlayer Is Nothing Then Return

        'Read from file
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= goCurrentPlayer.PlayerName & ".sht"

        If My.Computer.FileSystem.FileExists(sFile) = True Then
            Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Open)
            Dim oReader As IO.BinaryReader = New IO.BinaryReader(fsFile)
            Dim lPos As Int32 = 0
            Dim lCnt As Int32 = 0
            Try
                While fsFile.Position < fsFile.Length
                    Dim yHdr() As Byte = oReader.ReadBytes(34)
                    If yHdr Is Nothing = False Then
                        lPos = 0
                        ReDim Preserve moShortCuts(lCnt)
                        With moShortCuts(lCnt)
                            .iEnvirTypeID = System.BitConverter.ToInt16(yHdr, lPos) : lPos += 2
                            .lEnvirID = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .iLocX = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            '.iLocY = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .iLocZ = System.BitConverter.ToInt32(yHdr, lPos) : lPos += 4
                            .sName = GetStringFromBytes(yHdr, lPos, 20) : lPos += 20
                            lCnt += 1
                        End With
                    Else
                        Exit While
                    End If
                End While
            Catch
            End Try
            oReader.Close()
            oReader = Nothing
            fsFile.Close()
            fsFile.Dispose()
            fsFile = Nothing
        End If
    End Sub

    Private Sub SaveShortcuts()
        If goCurrentPlayer Is Nothing Then Return

        Dim ShortcutUB As Int32 = moShortCuts.GetUpperBound(0)

        'Write to file
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= goCurrentPlayer.PlayerName & ".sht"
        Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Create)

        Dim lPos As Int32 = 0
        For X As Int32 = 0 To ShortcutUB
            Dim yShortcutSave(33) As Byte
            lPos = 0
            With moShortCuts(X)
                If .bNoSave = False Then
                    System.BitConverter.GetBytes(.iEnvirTypeID).CopyTo(yShortcutSave, lPos) : lPos += 2
                    System.BitConverter.GetBytes(.lEnvirID).CopyTo(yShortcutSave, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.iLocX).CopyTo(yShortcutSave, lPos) : lPos += 4
                    'System.BitConverter.GetBytes(.iLocY).CopyTo(yShortcutSave, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.iLocZ).CopyTo(yShortcutSave, lPos) : lPos += 4
                    System.Text.ASCIIEncoding.ASCII.GetBytes(.sName).CopyTo(yShortcutSave, lPos) : lPos += 20
                    fsFile.Write(yShortcutSave, 0, yShortcutSave.Length)
                End If
            End With
        Next X
        fsFile.Close()
        fsFile.Dispose()
        fsFile = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        Me.Visible = False
    End Sub

    Private Sub lstEnvir_ItemClick(ByVal lIndex As Integer) Handles lstEnvir.ItemClick
        'reset stuff
        If btnDelete.Caption.ToUpper <> "DELETE" Then btnDelete.Caption = "Delete"
    End Sub

    Private Sub lstEnvir_ItemDblClick(ByVal lIndex As Integer) Handles lstEnvir.ItemDblClick
        If lIndex = -1 Then Return
        For x As Int32 = 0 To moShortCuts.GetUpperBound(0)
            If moShortCuts(x).iIndex = lstEnvir.ListIndex Then
                With moShortCuts(x)

                    If (.lEnvirID < 1 OrElse .iEnvirTypeID < 1) AndAlso goCurrentPlayer.yPlayerPhase = 255 Then Return

                    If .lEnvirID > 500000000 AndAlso .iEnvirTypeID = ObjectType.ePlanet AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then
                        If goUILib Is Nothing = False Then goUILib.AddNotification("That location no longer exists.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

                        'Remove it, Save, and Reload
                        moShortCuts(x).bNoSave = True
                        SaveShortcuts()
                        LoadShortcuts()
                    End If

                    If .iEnvirTypeID = ObjectType.ePlanet OrElse .iEnvirTypeID = ObjectType.eSolarSystem Then
                        If goCurrentEnvir Is Nothing = False Then

                            If NewTutorialManager.TutorialOn = True Then
                                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eEmailWaypointJumpedTo, -1, -1, -1, "")
                            End If

                            PlayerComm.l_JumpToX = .iLocX
                            'PlayerComm.l_JumpToY = .iLocY
                            PlayerComm.l_JumpToZ = .iLocZ

                            If goCurrentPlayer.yPlayerPhase <> 255 Then
                                PlayerComm.FinishJumpToEvent()
                                Return
                            End If

                            If goCurrentEnvir.ObjectID = .lEnvirID AndAlso goCurrentEnvir.ObjTypeID = .iEnvirTypeID Then
                                'we're already here... nothing to do...
                                PlayerComm.FinishJumpToEvent()
                            Else
                                goUILib.lUISelectState = UILib.eSelectState.eEmailAttachmentJumpTo
                                frmMain.ForceChangeEnvironment(.lEnvirID, .iEnvirTypeID)
                                btnClose_Click(Nothing)
                            End If
                        End If
                    End If
                End With
                Return
            End If
        Next x
    End Sub

    Private Sub btnNew_Click(ByVal sName As String) Handles btnNew.Click
        Dim iEnvirTypeID As Int16 = -1
        Dim lEnvirID As Int32 = -1
        Dim sShortcutName As String = ""

        If glCurrentEnvirView = CurrentView.ePlanetView Then
            If goGalaxy.CurrentSystemIdx <> -1 Then
                Dim oSystem As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                If oSystem.CurrentPlanetIdx <> -1 Then
                    With oSystem.moPlanets(oSystem.CurrentPlanetIdx)
                        lEnvirID = .ObjectID
                        iEnvirTypeID = .ObjTypeID
                        sShortcutName = .PlanetName
                    End With
                End If
            End If
        ElseIf glCurrentEnvirView = CurrentView.eSystemView Then
            If goGalaxy.CurrentSystemIdx <> -1 Then
                With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                    lEnvirID = .ObjectID
                    iEnvirTypeID = .ObjTypeID
                    sShortcutName = .SystemName
                End With
            End If
        End If

        If lEnvirID = -1 OrElse iEnvirTypeID = -1 Then
            MyBase.moUILib.AddNotification("You must be in Planet view or System view to add a shortcut.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If


        Dim ofrmInputBox As New frmInputBox(goUILib)
        'ofrmInputBox.lblEnter.Caption = "Enter Shortcut Name:"
        ofrmInputBox.txtValue.MaxLength = 20
        ofrmInputBox.SetText(sShortcutName)
        AddHandler ofrmInputBox.InputBoxClosed, AddressOf AddNewResponse

    End Sub

    Private Sub AddNewResponse(ByVal bCancelled As Boolean, ByVal sValue As String)
        If goGalaxy Is Nothing Then Return

        If bCancelled = True OrElse sValue = "" Then Return

        'Check Name
        If sValue.Length > 20 Then
            MyBase.moUILib.AddNotification("The new name is too long!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf sValue.Length = 0 Then
            MyBase.moUILib.AddNotification("You must enter a name for this shortcut.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'Verify current location is valid
        Dim iLocX As Int32 = -1
        Dim iLocY As Int32 = -1
        Dim iLocZ As Int32 = -1
        Dim iEnvirTypeID As Int16 = -1
        Dim lEnvirID As Int32 = -1

        If glCurrentEnvirView = CurrentView.ePlanetView Then
            If goGalaxy.CurrentSystemIdx <> -1 Then
                Dim oSystem As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                If oSystem.CurrentPlanetIdx <> -1 Then
                    With oSystem.moPlanets(oSystem.CurrentPlanetIdx)
                        lEnvirID = .ObjectID
                        iEnvirTypeID = .ObjTypeID
                    End With
                End If
            End If
        ElseIf glCurrentEnvirView = CurrentView.eSystemView Then
            If goGalaxy.CurrentSystemIdx <> -1 Then
                With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                    lEnvirID = .ObjectID
                    iEnvirTypeID = .ObjTypeID
                End With
            End If
        End If

        If lEnvirID = -1 OrElse iEnvirTypeID = -1 Then
            MyBase.moUILib.AddNotification("You must be in Planet view or System view to add a shortcut.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'All is well, add it
        Dim ShortcutUB As Int32 = moShortCuts.GetUpperBound(0)
        ShortcutUB += 1
        ReDim Preserve moShortCuts(ShortcutUB)
        moShortCuts(ShortcutUB).iIndex = -1
        moShortCuts(ShortcutUB).sName = sValue
        'Debug.Print("(" & goCamera.mlCameraAtX & "," & goCamera.mlCameraAtY & "," & goCamera.mlCameraAtZ & ") (" & goCamera.mlCameraX & "," & goCamera.mlCameraY & "," & goCamera.mlCameraZ & ")")
        moShortCuts(ShortcutUB).iLocX = goCamera.mlCameraAtX
        'moShortCuts(ShortcutUB).iLocY = goCamera.mlCameraY
        moShortCuts(ShortcutUB).iLocZ = goCamera.mlCameraAtZ
        moShortCuts(ShortcutUB).iEnvirTypeID = iEnvirTypeID
        moShortCuts(ShortcutUB).lEnvirID = lEnvirID

        'Save and reload/refresh
        SaveShortcuts()
        LoadShortcuts()
    End Sub

    Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
        If lstEnvir.ListIndex = -1 OrElse lstEnvir.ListCount = 0 Then
            goUILib.AddNotification("Please select a shortcut to delete.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If btnDelete.Caption.ToUpper = "CONFIRM" Then
            'Delete it
            btnDelete.Caption = "Delete"
        Else
            goUILib.AddNotification("Press Delete again to confirm Shortcut Deletion. This action cannot be undone.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            miIndex = lstEnvir.ListIndex
            btnDelete.Caption = "Confirm"
            Return
        End If
        If miIndex = -1 Then Return
        'Update it
        Dim bFound As Boolean = False
        For x As Int32 = 0 To moShortCuts.GetUpperBound(0)
            If moShortCuts(x).iIndex = miIndex Then
                moShortCuts(x).bNoSave = True
                bFound = True
                Exit For
            End If
        Next x
        If bFound = False Then Return 'Ehh shouldn't be possible
        miIndex = -1
        'Save and reload
        SaveShortcuts()
        LoadShortcuts()
        lstEnvir.ListIndex = -1

    End Sub

    Private Sub btnRename_Click(ByVal sName As String) Handles btnRename.Click
        If lstEnvir.ListIndex = -1 Then
            goUILib.AddNotification("Please select a shortcut to rename.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        miIndex = lstEnvir.ListIndex
        Dim ofrmInputBox As New frmInputBox(goUILib)
        'ofrmInputBox.lblEnter.Caption = "Enter Shortcut Name:"
        ofrmInputBox.txtValue.MaxLength = 20
        ofrmInputBox.SetText(lstEnvir.List(lstEnvir.ListIndex))
        AddHandler ofrmInputBox.InputBoxClosed, AddressOf RenameResponce

    End Sub

    Private Sub RenameResponce(ByVal bCancelled As Boolean, ByVal sValue As String)
        If bCancelled = True OrElse sValue = "" Then Return

        'Get Name
        If sValue.Length > 20 Then
            MyBase.moUILib.AddNotification("The new name is too long!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf sValue.Length = 0 Then
            MyBase.moUILib.AddNotification("You must enter a name for this shortcut to be renamed to.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'Update it
        Dim bFound As Boolean = False
        For x As Int32 = 0 To moShortCuts.GetUpperBound(0)
            If moShortCuts(x).iIndex = miIndex Then
                moShortCuts(x).sName = sValue
                bFound = True
                Exit For
            End If
        Next x
        If bFound = False Then Return 'Ehh shouldn't be possible

        'Save and reload
        miIndex = -1
        SaveShortcuts()
        LoadShortcuts()
    End Sub

    Private Sub frmEnvirShortcuts_WindowClosed() Handles Me.WindowClosed
        miLastListIndex = lstEnvir.ListIndex
        miLastScrlBarVal = lstEnvir.ScrollBarValue
    End Sub
End Class
