Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

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
    Friend WithEvents tmrUpdate As System.Windows.Forms.Timer
    Friend WithEvents btnCreatePak As System.Windows.Forms.Button
    Friend WithEvents lvwTOC As System.Windows.Forms.ListView
    Friend WithEvents colName As System.Windows.Forms.ColumnHeader
    Friend WithEvents colPos As System.Windows.Forms.ColumnHeader
    Friend WithEvents colSize As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnUnpack As System.Windows.Forms.Button
    Friend WithEvents chkSound As System.Windows.Forms.CheckBox
    Friend WithEvents btnDisplayTOC As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.tmrUpdate = New System.Windows.Forms.Timer(Me.components)
        Me.btnCreatePak = New System.Windows.Forms.Button
        Me.btnDisplayTOC = New System.Windows.Forms.Button
        Me.lvwTOC = New System.Windows.Forms.ListView
        Me.colName = New System.Windows.Forms.ColumnHeader
        Me.colPos = New System.Windows.Forms.ColumnHeader
        Me.colSize = New System.Windows.Forms.ColumnHeader
        Me.btnUnpack = New System.Windows.Forms.Button
        Me.chkSound = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'tmrUpdate
        '
        Me.tmrUpdate.Interval = 30
        '
        'btnCreatePak
        '
        Me.btnCreatePak.Location = New System.Drawing.Point(14, 12)
        Me.btnCreatePak.Name = "btnCreatePak"
        Me.btnCreatePak.Size = New System.Drawing.Size(98, 23)
        Me.btnCreatePak.TabIndex = 0
        Me.btnCreatePak.Text = "Create PAK"
        '
        'btnDisplayTOC
        '
        Me.btnDisplayTOC.Location = New System.Drawing.Point(14, 44)
        Me.btnDisplayTOC.Name = "btnDisplayTOC"
        Me.btnDisplayTOC.Size = New System.Drawing.Size(98, 23)
        Me.btnDisplayTOC.TabIndex = 1
        Me.btnDisplayTOC.Text = "Display TOC"
        '
        'lvwTOC
        '
        Me.lvwTOC.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvwTOC.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colName, Me.colPos, Me.colSize})
        Me.lvwTOC.Location = New System.Drawing.Point(118, 12)
        Me.lvwTOC.Name = "lvwTOC"
        Me.lvwTOC.Size = New System.Drawing.Size(286, 121)
        Me.lvwTOC.TabIndex = 2
        Me.lvwTOC.UseCompatibleStateImageBehavior = False
        Me.lvwTOC.View = System.Windows.Forms.View.Details
        '
        'colName
        '
        Me.colName.Text = "File Name"
        Me.colName.Width = 142
        '
        'colPos
        '
        Me.colPos.Text = "Position"
        Me.colPos.Width = 68
        '
        'colSize
        '
        Me.colSize.Text = "Size"
        Me.colSize.Width = 66
        '
        'btnUnpack
        '
        Me.btnUnpack.Location = New System.Drawing.Point(14, 73)
        Me.btnUnpack.Name = "btnUnpack"
        Me.btnUnpack.Size = New System.Drawing.Size(98, 23)
        Me.btnUnpack.TabIndex = 3
        Me.btnUnpack.Text = "Unpack"
        '
        'chkSound
        '
        Me.chkSound.AutoSize = True
        Me.chkSound.Location = New System.Drawing.Point(14, 116)
        Me.chkSound.Name = "chkSound"
        Me.chkSound.Size = New System.Drawing.Size(57, 17)
        Me.chkSound.TabIndex = 4
        Me.chkSound.Text = "Sound"
        Me.chkSound.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(416, 145)
        Me.Controls.Add(Me.chkSound)
        Me.Controls.Add(Me.btnUnpack)
        Me.Controls.Add(Me.lvwTOC)
        Me.Controls.Add(Me.btnDisplayTOC)
        Me.Controls.Add(Me.btnCreatePak)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Private Const ml_ENCRYPT_SEED As Int32 = 777

    Private Sub CreatePakOld()
        Dim oFS As IO.FileStream
        Dim oFSWriter As IO.BinaryWriter
        Dim oFSReader As IO.BinaryReader

        Dim yBody() As Byte
        Dim yTOC() As Byte

        Dim sFile As String
        Dim lPos As Int32 = 0
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        Dim X As Int32

        Dim lEntryUB As Int32 = -1
        Dim lEntryPos() As Int32 = Nothing
        Dim lEntryLen() As Int32 = Nothing
        Dim sEntry() As String = Nothing

        Dim yTemp() As Byte

        If sPath.EndsWith("\") = False Then sPath &= "\"
        sPath &= "Packer\"

        ReDim yBody(-1)

        sFile = sPath & "*.*"
        sFile = Dir$(sFile)
        While sFile <> ""
            oFS = New IO.FileStream(sPath & sFile, IO.FileMode.Open)
            oFSReader = New IO.BinaryReader(oFS)

            'Read the entire file, convert it to bytes and store in our temp array
            yTemp = oFSReader.ReadBytes(oFSReader.BaseStream.Length)

            'Do our encryption on temp array
            yTemp = EncBytes(yTemp)

            'Set up our body array
            ReDim Preserve yBody(yBody.Length + yTemp.Length - 1)

            'Copy the temp result to the body
            yTemp.CopyTo(yBody, lPos)

            'set up our TOC entry
            lEntryUB += 1
            ReDim Preserve lEntryPos(lEntryUB)
            ReDim Preserve sEntry(lEntryUB)
            ReDim Preserve lEntryLen(lEntryUB)
            sEntry(lEntryUB) = sFile
            lEntryLen(lEntryUB) = yTemp.Length - 1
            lEntryPos(lEntryUB) = lPos

            'increment lPos for the next entry
            lPos += yTemp.Length - 1

            'close the reader
            oFSReader.Close()
            oFSReader = Nothing
            'close the stream
            oFS.Close()
            oFS = Nothing

            'Get next file
            sFile = Dir()
        End While

        'Now, create our TOC
        ReDim yTOC(((lEntryUB + 1) * 28) - 1)
        'First 4 bytes need to be the Length of the TOC
        'System.BitConverter.GetBytes(yTOC.Length).CopyTo(yTOC, 0)
        lPos = 0
        For X = 0 To lEntryUB
            'convert name to bytes and put in toc
            System.Text.ASCIIEncoding.ASCII.GetBytes(sEntry(X)).CopyTo(yTOC, lPos) : lPos += 20
            'Now, the start
            System.BitConverter.GetBytes(lEntryPos(X)).CopyTo(yTOC, lPos) : lPos += 4
            'Now, the length
            System.BitConverter.GetBytes(lEntryLen(X)).CopyTo(yTOC, lPos) : lPos += 4
        Next X

        'Now, encrypt our TOC
        yTOC = EncBytes(yTOC)

        'Create our file
        oFS = New IO.FileStream(sPath & "output.pak", IO.FileMode.OpenOrCreate)
        oFSWriter = New IO.BinaryWriter(oFS)

        'Now, write the TOC (with the length appended)
        'To do this, redim yTemp with length + 3 (4 bytes more)
        ReDim yTemp(yTOC.Length + 3)
        System.BitConverter.GetBytes(yTOC.Length).CopyTo(yTemp, 0)
        yTOC.CopyTo(yTemp, 4)
        'yTemp = System.BitConverter.GetBytes(yTOC.Length - 1)
        oFSWriter.Write(yTemp)

        'Write the TOC array
        'oFSWriter.Write(yTOC)

        'Now, write the Body
        oFSWriter.Write(yBody)

        'Close our writer
        oFSWriter.Close()
        oFSWriter = Nothing
        'Close our stream
        oFS.Close()
        oFS = Nothing
    End Sub

    Private Sub CreatePakNew()
        Dim oFS As IO.FileStream
        Dim oFSWriter As IO.BinaryWriter
        Dim oFSReader As IO.BinaryReader

        Dim yBody() As Byte
        Dim yTOC() As Byte

        Dim sFile As String
        Dim lPos As Int32 = 0
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        Dim X As Int32

        Dim lEntryUB As Int32 = -1
        Dim lEntryPos() As Int32 = Nothing
        Dim lEntryLen() As Int32 = Nothing
        Dim sEntry() As String = Nothing

        Dim yTemp() As Byte

        If sPath.EndsWith("\") = False Then sPath &= "\"
        sPath &= "Packer\"

        ReDim yBody(-1)

        sFile = sPath & "*.*"
        sFile = Dir$(sFile)
        While sFile <> ""
            oFS = New IO.FileStream(sPath & sFile, IO.FileMode.Open)
            oFSReader = New IO.BinaryReader(oFS)

            'Read the entire file, convert it to bytes and store in our temp array
            yTemp = oFSReader.ReadBytes(oFSReader.BaseStream.Length)

            'Do our encryption on temp array
            yTemp = NewEncBytes(yTemp)

            'Set up our body array
            ReDim Preserve yBody(yBody.Length + yTemp.Length) ' - 1)

            'Copy the temp result to the body
            yTemp.CopyTo(yBody, lPos)

            'set up our TOC entry
            lEntryUB += 1
            ReDim Preserve lEntryPos(lEntryUB)
            ReDim Preserve sEntry(lEntryUB)
            ReDim Preserve lEntryLen(lEntryUB)
            sEntry(lEntryUB) = sFile
            lEntryLen(lEntryUB) = yTemp.Length ' - 1
            lEntryPos(lEntryUB) = lPos

            'increment lPos for the next entry
            lPos += yTemp.Length ' - 1

            'close the reader
            oFSReader.Close()
            oFSReader = Nothing
            'close the stream
            oFS.Close()
            oFS = Nothing

            'Get next file
            sFile = Dir()
        End While

        'Now, create our TOC
        ReDim yTOC(((lEntryUB + 1) * 28) - 1)
        'First 4 bytes need to be the Length of the TOC
        'System.BitConverter.GetBytes(yTOC.Length).CopyTo(yTOC, 0)
        lPos = 0
        For X = 0 To lEntryUB
            'convert name to bytes and put in toc
            System.Text.ASCIIEncoding.ASCII.GetBytes(sEntry(X)).CopyTo(yTOC, lPos) : lPos += 20
            'Now, the start
            System.BitConverter.GetBytes(lEntryPos(X)).CopyTo(yTOC, lPos) : lPos += 4
            'Now, the length
            System.BitConverter.GetBytes(lEntryLen(X)).CopyTo(yTOC, lPos) : lPos += 4
        Next X

        'Now, encrypt our TOC
        yTOC = NewEncBytes(yTOC)

        'Create our file
        oFS = New IO.FileStream(sPath & "output.pak", IO.FileMode.OpenOrCreate)
        oFSWriter = New IO.BinaryWriter(oFS)

        'Now, write the TOC (with the length appended)
        'To do this, redim yTemp with length + 3 (4 bytes more)
        ReDim yTemp(yTOC.Length + 3)
        System.BitConverter.GetBytes(yTOC.Length).CopyTo(yTemp, 0)
        yTOC.CopyTo(yTemp, 4)
        'yTemp = System.BitConverter.GetBytes(yTOC.Length - 1)
        oFSWriter.Write(yTemp)

        'Write the TOC array
        'oFSWriter.Write(yTOC)

        'Now, write the Body
        oFSWriter.Write(yBody)

        'Close our writer
        oFSWriter.Close()
        oFSWriter = Nothing
        'Close our stream
        oFS.Close()
        oFS = Nothing
    End Sub

    Private Sub btnCreatePak_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreatePak.Click
        If chkSound.Checked = True Then
            CreatePakNew()
        Else
            CreatePakOld()
        End If


        MsgBox("Done")

    End Sub

    Private Function NewEncBytes(ByVal yBytes() As Byte) As Byte()
        Dim lLen As Int32 = UBound(yBytes)
        Dim lKey As Int32
        Dim lOffset As Int32
        Dim X As Int32
        Dim lChrCode As Int32
        Dim lMod As Int32

        Dim yFinal(lLen + 1) As Byte

        lKey = Int(Rnd() * 51)

        Dim oRandom As New Random(lKey)

        'Rnd(-1)
        'Call Randomize(ml_ENCRYPT_SEED + lKey)

        yFinal(0) = CByte(lKey)
        For X = 0 To lLen
            lOffset = X - lLen
            'now, found out what we got here..
            lChrCode = yBytes(X)
            'lMod = Int(Rnd() * 5) + 1
            lMod = oRandom.Next(1, 6)       '1,6 because 6 is excluded,meaning 1 to 5
            lChrCode = lChrCode + lMod
            If lChrCode > 255 Then lChrCode = lChrCode - 256
            yFinal(X + 1) = CByte(lChrCode)
        Next X

        Return yFinal
    End Function

    Private Function EncBytes(ByVal yBytes() As Byte) As Byte()
        Dim lLen As Int32 = UBound(yBytes)
        Dim lKey As Int32
        Dim lOffset As Int32
        Dim X As Int32
        Dim lChrCode As Int32
        Dim lMod As Int32


        Dim yFinal(lLen + 1) As Byte

        lKey = Int(Rnd() * 51)
        Rnd(-1)
        Call Randomize(ml_ENCRYPT_SEED + lKey)

        yFinal(0) = CByte(lKey)
        For X = 0 To lLen
            lOffset = X - lLen
            'now, found out what we got here..
            lChrCode = yBytes(X)
            lMod = Int(Rnd() * 5) + 1
            lChrCode = lChrCode + lMod
            If lChrCode > 255 Then lChrCode = lChrCode - 256
            yFinal(X + 1) = CByte(lChrCode)
        Next X

        EncBytes = yFinal

    End Function

    Private Function DecBytes(ByVal yBytes() As Byte) As Byte()
        'Now, we do the exact opposite...
        Dim lLen As Int32 = UBound(yBytes)
        Dim lKey As Int32
        Dim X As Int32
        Dim yFinal(lLen + 1) As Byte
        Dim lChrCode As Int32
        Dim lMod As Int32

        'Get our key value...
        lKey = yBytes(0)

        'set up our seed
        Rnd(-1)
        Call Randomize(ml_ENCRYPT_SEED + lKey)

        For X = 1 To lLen
            'Now, find out what we got here...
            lChrCode = yBytes(X)
            'now, subtract our value... 1 to 5
            lMod = Int(Rnd() * 5) + 1
            lChrCode = lChrCode - lMod
            If lChrCode < 0 Then lChrCode = 256 + lChrCode
            yFinal(X - 1) = CByte(lChrCode)
        Next X
        DecBytes = yFinal
    End Function

    Private Function NewDecBytes(ByVal yBytes() As Byte) As Byte()
        'Now, we do the exact opposite...
        Dim lLen As Int32 = UBound(yBytes)
        Dim lKey As Int32
        Dim X As Int32
        Dim yFinal(lLen + 1) As Byte
        Dim lChrCode As Int32
        Dim lMod As Int32

        'Get our key value...
        lKey = yBytes(0)

        'set up our seed
        'Rnd(-1)
        'Call Randomize(ml_ENCRYPT_SEED + lKey)
        Dim oRandom As New Random(lKey)

        For X = 1 To lLen
            'Now, find out what we got here...
            lChrCode = yBytes(X)
            'now, subtract our value... 1 to 5
            'lMod = Int(Rnd() * 5) + 1
            lMod = oRandom.Next(1, 6)   '1 to 6 here because 6 is excluded
            lChrCode = lChrCode - lMod
            If lChrCode < 0 Then lChrCode = 256 + lChrCode
            yFinal(X - 1) = CByte(lChrCode)
        Next X
        Return yFinal
    End Function

    Private Sub btnDisplayTOC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisplayTOC.Click
        'Ok... let's extract all files...
        Dim oFS As IO.FileStream
        Dim oFSReader As IO.BinaryReader

        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        Dim lLen As Int32

        If sPath.EndsWith("\") = False Then sPath &= "\"
        sPath &= "Packer\"

        Dim yData() As Byte
        Dim yTOC() As Byte
        Dim X As Int32

        Dim sFile As String
        Dim yEntryName(19) As Byte
        Dim lEntryPos As Int32
        Dim lEntryLen As Int32

        Dim lPos As Int32

        Dim sPakFile As String

        sPakFile = Dir(sPath & "*.pak")
        If sPakFile <> "" Then
            'Open our pak file, read it all ot an array and close it
            oFS = New IO.FileStream(sPath & sPakFile, IO.FileMode.Open)
            oFSReader = New IO.BinaryReader(oFS)
            yData = oFSReader.ReadBytes(oFS.Length)
            oFSReader.Close()
            oFSReader = Nothing
            oFS.Close()
            oFS = Nothing

            'Ok, first 4 bytes are my unencrypted TOC length
            lLen = System.BitConverter.ToInt32(yData, 0)

            'Ok, now... separate our TOC from the main data
            ReDim yTOC(lLen - 1)
            Array.Copy(yData, 4, yTOC, 0, lLen)

            'Decrypt our Table of Contents...
            If chkSound.Checked = True Then yTOC = NewDecBytes(yTOC) Else yTOC = DecBytes(yTOC)

            'Now... loop through and get our data...
            lPos = 0
            For X = 0 To Math.Floor((lLen / 28)) - 1
                Array.Copy(yTOC, X * 28, yEntryName, 0, 20)
                sFile = BytesToString(yEntryName)
                lPos += 20
                lEntryPos = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4
                lEntryLen = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4

                Dim lvwItem As ListViewItem = lvwTOC.Items.Add(sFile)
                lvwItem.Tag = sPakFile
                lvwItem.SubItems.Add(lEntryPos.ToString)
                lvwItem.SubItems.Add(lEntryLen.ToString)
            Next X
            'sPakFile = Dir()
        End If

    End Sub

    Private Function BytesToString(ByVal yBytes() As Byte) As String
        Dim lLen As Int32 = yBytes.Length
        Dim X As Int32

        For X = 0 To yBytes.Length - 1
            If yBytes(X) = 0 Then
                lLen = X
                Exit For
            End If
        Next X

        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yBytes), 1, lLen)

    End Function

    Private Function GetResourceStream(ByVal sName As String, ByVal sResourceFile As String) As IO.MemoryStream
        'If Exists(sResourceFile) = False Then Return Nothing

        Dim oFS As IO.FileStream
        Dim oReader As IO.BinaryReader

        Dim yData() As Byte
        Dim yTOC() As Byte
        Dim lTOCLen As Int32

        Dim X As Int32
        Dim yEntryName(19) As Byte
        Dim sFile As String
        Dim lPos As Int32

        Dim lEntryStart As Int32 = -1
        Dim lEntryLen As Int32 = -1

        'Open the resource file
        oFS = New IO.FileStream(sResourceFile, IO.FileMode.Open)
        'Create our reader
        oReader = New IO.BinaryReader(oFS)

        'Get the first 4 bytes which is the unencrypted length of our Table of Contents portion
        yData = oReader.ReadBytes(4)
        lTOCLen = System.BitConverter.ToInt32(yData, 0)

        'read in the TOC contents
        yTOC = oReader.ReadBytes(lTOCLen)

        'Decrypt the TOC
        yTOC = DecBytes(yTOC)

        lPos = 0
        'Find out entry...
        For X = 0 To Math.Floor(lTOCLen / 28) - 1
            Array.Copy(yTOC, X * 28, yEntryName, 0, 20)
            lPos += 20
            sFile = BytesToString(yEntryName)
            If UCase$(sFile) = UCase$(sName) Then
                lEntryStart = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4
                lEntryLen = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4
                Exit For
            Else
                lPos += 8
            End If
        Next X

        If lEntryStart > -1 AndAlso lEntryLen > -1 Then
            'Ok, found our entry... let's get the data
            'oReader.ReadBytes(lEntryStart - 1)  'should be right...
            oReader.BaseStream.Seek(lEntryStart + lTOCLen + 4, IO.SeekOrigin.Begin)
            Dim yEntry() As Byte
            yData = oReader.ReadBytes(lEntryLen)
            'Now, decrypt it
            yData = DecBytes(yData)
            ReDim yEntry(lEntryLen - 1)
            Array.Copy(yData, 0, yEntry, 0, lEntryLen)

            'Now, create our response
            Dim oMemStream As IO.MemoryStream = New IO.MemoryStream(yEntry.Length - 1)
            Dim oFSWriter As IO.BinaryWriter = New IO.BinaryWriter(oMemStream)
            oFSWriter.Write(yEntry)
            'Flush and Nothing, but do not close as it will close the base stream too
            oFSWriter.Flush()
            oFSWriter = Nothing

            'Now, seek the memory stream to the beginning
            oMemStream.Seek(0, IO.SeekOrigin.Begin)

            'Close everything else
            oReader.Close()
            oFS.Close()

            Return oMemStream
        Else
            oReader.Close()
            oFS.Close()
            Return Nothing
        End If

    End Function


    Private Function New_GetResourceStream(ByVal sName As String, ByVal sResourceFile As String) As IO.MemoryStream
        'If Exists(sResourceFile) = False Then Return Nothing

        Dim oFS As IO.FileStream
        Dim oReader As IO.BinaryReader

        Dim yData() As Byte
        Dim yTOC() As Byte
        Dim lTOCLen As Int32

        Dim X As Int32
        Dim yEntryName(19) As Byte
        Dim sFile As String
        Dim lPos As Int32

        Dim lEntryStart As Int32 = -1
        Dim lEntryLen As Int32 = -1

        'Open the resource file
        oFS = New IO.FileStream(sResourceFile, IO.FileMode.Open)
        'Create our reader
        oReader = New IO.BinaryReader(oFS)

        'Get the first 4 bytes which is the unencrypted length of our Table of Contents portion
        yData = oReader.ReadBytes(4)
        lTOCLen = System.BitConverter.ToInt32(yData, 0)

        'read in the TOC contents
        yTOC = oReader.ReadBytes(lTOCLen)

        'Decrypt the TOC
        yTOC = NewDecBytes(yTOC)

        lPos = 0
        'Find out entry...
        For X = 0 To Math.Floor(lTOCLen / 28) - 1
            Array.Copy(yTOC, X * 28, yEntryName, 0, 20)
            lPos += 20
            sFile = BytesToString(yEntryName)
            If UCase$(sFile) = UCase$(sName) Then
                lEntryStart = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4
                lEntryLen = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4
                Exit For
            Else
                lPos += 8
            End If
        Next X

        If lEntryStart > -1 AndAlso lEntryLen > -1 Then
            'Ok, found our entry... let's get the data
            'oReader.ReadBytes(lEntryStart - 1)  'should be right...
            oReader.BaseStream.Seek(lEntryStart + lTOCLen + 4, IO.SeekOrigin.Begin)
            Dim yEntry() As Byte
            yData = oReader.ReadBytes(lEntryLen)
            'Now, decrypt it
            yData = NewDecBytes(yData)
            ReDim yEntry(lEntryLen - 1)
            Array.Copy(yData, 0, yEntry, 0, lEntryLen)

            'Now, create our response
            Dim oMemStream As IO.MemoryStream = New IO.MemoryStream(yEntry.Length - 1)
            Dim oFSWriter As IO.BinaryWriter = New IO.BinaryWriter(oMemStream)
            oFSWriter.Write(yEntry)
            'Flush and Nothing, but do not close as it will close the base stream too
            oFSWriter.Flush()
            oFSWriter = Nothing

            'Now, seek the memory stream to the beginning
            oMemStream.Seek(0, IO.SeekOrigin.Begin)

            'Close everything else
            oReader.Close()
            oFS.Close()

            Return oMemStream
        Else
            oReader.Close()
            oFS.Close()
            Return Nothing
        End If

    End Function

    Private Sub btnUnpack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUnpack.Click
        'Ok... let's extract all files...
        If lvwTOC.SelectedItems.Count > 0 Then
            Dim lvwItem As ListViewItem = lvwTOC.SelectedItems(0)
            Dim sItem As String = lvwItem.Text
            Dim oMem As IO.MemoryStream = Nothing
            If chkSound.Checked = True Then
                oMem = New_GetResourceStream(sItem, "Packer\" & CStr(lvwItem.Tag))
            Else
                oMem = GetResourceStream(sItem, "Packer\" & CStr(lvwItem.Tag))
            End If
            If oMem Is Nothing = False Then
                Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                If sPath.EndsWith("\") = False Then sPath &= "\"
                sPath &= "Packer\" & sItem
                Dim oFS As New IO.FileStream(sPath, IO.FileMode.Create)

                Dim yBytes(oMem.Length - 1) As Byte
                oMem.Read(yBytes, 0, oMem.Length)
                oFS.Write(yBytes, 0, yBytes.Length)
                oFS.Close()
                oMem.Close()
            End If
        End If
    End Sub
 
End Class

