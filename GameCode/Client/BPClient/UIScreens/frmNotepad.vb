Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmNotePad
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private WithEvents lstNotes As UIListBox
    Private lblNoteTitle As UILabel
    Private txtNotetitle As UITextBox
    Private txtDetail As UITextBox
    Private WithEvents btnCreate As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnSave As UIButton
    Private WithEvents btnClose As UIButton

    Private Class NotepadItem
        Public sFile As String
        Public sTitle As String
        Public sBody As String
    End Class
    Private moItems() As NotepadItem
    Private mlItemUB As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmNotePad initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eNotePad
            .ControlName = "frmNotePad"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256
            .Width = 512
            .Height = 512
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
            .Top = 3
            .Width = 158
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Notes Management"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 26
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lstNotes initial props
        lstNotes = New UIListBox(oUILib)
        With lstNotes
            .ControlName = "lstNotes"
            .Left = 5
            .Top = 30
            .Width = 175
            .Height = 440
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstNotes, UIControl))

        'lblNoteTitle initial props
        lblNoteTitle = New UILabel(oUILib)
        With lblNoteTitle
            .ControlName = "lblNoteTitle"
            .Left = 190
            .Top = 35
            .Width = 50
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Title:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblNoteTitle, UIControl))

        'txtNotetitle initial props
        txtNotetitle = New UITextBox(oUILib)
        With txtNotetitle
            .ControlName = "txtNotetitle"
            .Left = 235
            .Top = 35
            .Width = 130
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Untitled"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtNotetitle, UIControl))

        'txtDetail initial props
        txtDetail = New UITextBox(oUILib)
        With txtDetail
            .ControlName = "txtDetail"
            .Left = 190
            .Top = 64
            .Width = 310
            .Height = 405
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .MultiLine = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtDetail, UIControl))

        'btnCreate initial props
        btnCreate = New UIButton(oUILib)
        With btnCreate
            .ControlName = "btnCreate"
            .Left = 402
            .Top = 35
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Begin Note"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCreate, UIControl))

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Left = 190
            .Top = 480
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Delete Note"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDelete, UIControl))

        'btnSave initial props
        btnSave = New UIButton(oUILib)
        With btnSave
            .ControlName = "btnSave"
            .Left = 401
            .Top = 480
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Save Note"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSave, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 487
            .Top = 2
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

        LoadList()

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnCreate_Click(ByVal sName As String) Handles btnCreate.Click
        lstNotes.ListIndex = -1
        txtNotetitle.Caption = ""
        txtDetail.Caption = ""
        btnDelete.Caption = "Delete Note"

        If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
        btnCreate.HasFocus = False
        goUILib.FocusedControl = txtNotetitle
        txtNotetitle.HasFocus = True
    End Sub

    Private Function GetAndEnsureNoteDirectoryExists() As String
        Dim sDirectory As String = AppDomain.CurrentDomain.BaseDirectory()
        If sDirectory.EndsWith("\") = False Then sDirectory &= "\"
        sDirectory &= "\Notes"
        If Exists(sDirectory) = False Then
            MkDir(sDirectory)
        End If
        sDirectory &= "\" & goCurrentPlayer.PlayerName
        If Exists(sDirectory) = False Then
            MkDir(sDirectory)
        End If
        Return sDirectory
    End Function

    Private Sub LoadList()
        Dim sDirectory As String = GetAndEnsureNoteDirectoryExists()
        Dim colResults As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(sDirectory, FileIO.SearchOption.SearchAllSubDirectories, "*.txt")

        mlItemUB = -1

        For Each sValue As String In colResults

            Dim sTitle As String = ""
            Dim sBody As String = ""
            If Exists(sValue) = True Then
                Dim oFS As IO.FileStream = Nothing
                Dim oReader As IO.StreamReader = Nothing
                Try
                    oFS = New IO.FileStream(sValue, IO.FileMode.Open)
                    oReader = New IO.StreamReader(oFS)

                    If oReader.EndOfStream = False Then sTitle = oReader.ReadLine
                    If oReader.EndOfStream = False Then sBody = oReader.ReadToEnd
                Catch
                Finally
                    If oReader Is Nothing = False Then oReader.Close()
                    If oFS Is Nothing = False Then oFS.Close()
                    oReader = Nothing
                    oFS = Nothing
                End Try
            Else : Continue For
            End If

            mlItemUB += 1
            ReDim Preserve moItems(mlItemUB)
            moItems(mlItemUB) = New NotepadItem

            With moItems(mlItemUB)
                .sFile = sValue
                .sTitle = sTitle
                .sBody = sBody
            End With
        Next

        lstNotes.Clear()
        For X As Int32 = 0 To mlItemUB
            lstNotes.AddItem(moItems(X).sTitle)
            lstNotes.ItemData(lstNotes.NewIndex) = X
        Next X

    End Sub

    Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
        If btnDelete.Caption.ToUpper = "CONFIRM" Then
            If lstNotes.ListIndex < 0 Then
                goUILib.AddNotification("Select a note in the list to delete first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            Dim lIdx As Int32 = lstNotes.ItemData(lstNotes.ListIndex)
            If lIdx > -1 Then
                Dim sFile As String = moItems(lIdx).sFile
                Kill(sFile)
                LoadList()
            End If
            btnDelete.Caption = "Delete Note"
            btnCreate_Click(btnCreate.ControlName)
        Else
            btnDelete.Caption = "Confirm"
        End If
        
    End Sub

    Private Sub btnSave_Click(ByVal sName As String) Handles btnSave.Click
        Dim sTitle As String = txtNotetitle.Caption
        Dim sBody As String = txtDetail.Caption

        'Verify our title is not in use yet.
        For X As Int32 = 0 To mlItemUB
            If moItems(X).sTitle = sTitle Then
                If lstNotes.ListIndex > -1 AndAlso lstNotes.ItemData(lstNotes.ListIndex) = X Then Continue For
                MyBase.moUILib.AddNotification("A Note Item already exists with that title!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
        Next X

        'now, save it... is it new?
        Dim sFile As String = ""
        If lstNotes.ListIndex > -1 Then
            'no, overwrite the current one...
            Dim lIdx As Int32 = lstNotes.ItemData(lstNotes.ListIndex)
            sFile = moItems(lIdx).sFile
        Else
            'yes
            sFile = GetAndEnsureNoteDirectoryExists()
            sFile &= "\" & Mid$(sTitle, 1, Math.Min(sTitle.Length, 8)) & ".txt"
        End If

        If sFile = "" Then Return

        Dim oFS As IO.FileStream = Nothing
        Dim oWriter As IO.StreamWriter = Nothing
        Try
            oFS = New IO.FileStream(sFile, IO.FileMode.Create)
            oWriter = New IO.StreamWriter(oFS)
            oWriter.WriteLine(sTitle)
            oWriter.WriteLine(sBody)
        Catch
        Finally
            If oWriter Is Nothing = False Then oWriter.Close()
            If oFS Is Nothing = False Then oFS.Close()
            oWriter = Nothing
            oFS = Nothing
        End Try

        If lstNotes.ListIndex < 0 Then
            'ok, add the note entry
            mlItemUB += 1
            ReDim Preserve moItems(mlItemUB)
            moItems(mlItemUB) = New NotepadItem
            With moItems(mlItemUB)
                .sTitle = sTitle
                .sFile = sFile
                .sBody = sBody
            End With

            lstNotes.AddItem(sTitle, False)
            lstNotes.ItemData(lstNotes.NewIndex) = mlItemUB
        End If
    End Sub

    Private Sub lstNotes_ItemClick(ByVal lIndex As Integer) Handles lstNotes.ItemClick
        btnDelete.Caption = "Delete Note"
        If lIndex > -1 Then
            Dim lIdx As Int32 = lstNotes.ItemData(lstNotes.ListIndex)
            If lIdx > -1 Then
                With moItems(lIdx)
                    txtNotetitle.Caption = .sTitle
                    txtDetail.Caption = .sBody
                End With
            End If
        End If
    End Sub
End Class