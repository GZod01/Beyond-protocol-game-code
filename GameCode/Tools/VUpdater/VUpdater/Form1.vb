Public Class Form1

    Private msMainPath As String

    Private moWriter As IO.StreamWriter
    'Private moVerify As IO.StreamWriter

    Private msPrevious() As String
    Private mlPrevUB As Int32 = -1

    Private msSAOnly() As String
    Private mlSAOnlyUB As Int32 = -1
    Private msBPOnly() As String
    Private mlBPOnlyUB As Int32 = -1

    Private Sub LoadSAOnlyItems()
        Dim sPath As String
        sPath = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"

        Dim oFS As IO.FileStream = New IO.FileStream(sPath & "SAFiles.txt", IO.FileMode.Open)
        Dim oReader As IO.StreamReader = New IO.StreamReader(oFS)
        Dim sLine As String

        Do
            sLine = oReader.ReadLine()
            If sLine <> "" Then
                mlSAOnlyUB += 1
                ReDim Preserve msSAOnly(mlSAOnlyUB)
                msSAOnly(mlSAOnlyUB) = sLine.ToUpper.Trim
            End If
        Loop Until sLine = "" 'OrElse oReader.EndOfStream = True

        oReader.Close()
        oReader.Dispose()
        oFS.Close()
        oFS.Dispose()
    End Sub

    Private Sub LoadBPOnlyItems()
        Dim sPath As String
        sPath = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"

        Dim oFS As IO.FileStream = New IO.FileStream(sPath & "BPFiles.txt", IO.FileMode.Open)
        Dim oReader As IO.StreamReader = New IO.StreamReader(oFS)
        Dim sLine As String

        Do
            sLine = oReader.ReadLine()
            If sLine <> "" Then
                mlBPOnlyUB += 1
                ReDim Preserve msBPOnly(mlBPOnlyUB)
                msBPOnly(mlBPOnlyUB) = sLine.ToUpper.Trim
            End If
        Loop Until sLine = "" 'OrElse oReader.EndOfStream = True

        oReader.Close()
        oReader.Dispose()
        oFS.Close()
        oFS.Dispose()
    End Sub

    Private Sub LoadCurrentValues()
        Dim sPath As String
        sPath = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"

        Dim oFS As IO.FileStream = New IO.FileStream(sPath & "Versions.txt", IO.FileMode.Open)
        Dim oReader As IO.StreamReader = New IO.StreamReader(oFS)
        Dim sLine As String

        Do
            sLine = oReader.ReadLine()
            If sLine <> "" Then
                mlPrevUB += 1
                ReDim Preserve msPrevious(mlPrevUB)
                msPrevious(mlPrevUB) = sLine
            End If
        Loop Until sLine = ""

        oReader.Close()
        oReader.Dispose()
        oFS.Close()
        oFS.Dispose()
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim sPath As String

        LoadSAOnlyItems()
        LoadBPOnlyItems()
        LoadCurrentValues()

        sPath = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"

        msMainPath = sPath

        Dim oFS As IO.FileStream = New IO.FileStream(sPath & "Versions.txt", IO.FileMode.Create)
        moWriter = New IO.StreamWriter(oFS)

        'Dim oTmp As New IO.FileStream(sPath & "Verify.txt", IO.FileMode.Create)
        'moVerify = New IO.StreamWriter(oTmp)

        Call FillSubDirectory(sPath)

        moWriter.Close()
        moWriter.Dispose()
        moWriter = Nothing
        'moVerify.Close()
        'moVerify.Dispose()
        'moVerify = Nothing
        oFS.Close()
        oFS.Dispose()
        oFS = Nothing
        'oTmp.Close()
        'oTmp.Dispose()
        'oTmp = Nothing

        MsgBox("Done")

    End Sub

    Private Sub FillSubDirectory(ByVal sDirectory As String)
        Dim sFile As String

        Dim sSubDir() As String = Nothing
        Dim lSubDirUB As Int32
        Dim X As Int32

        Dim sEntry As String

        lSubDirUB = -1
        sFile = Dir$(sDirectory & "*.*", vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem Or vbVolume)

        While sFile <> ""
            If sFile <> "." And sFile <> ".." Then
                If InStr(1, sFile, ".", vbBinaryCompare) = 0 Then
                    'Directory
                    lSubDirUB = lSubDirUB + 1
                    ReDim Preserve sSubDir(0 To lSubDirUB)
                    sSubDir(lSubDirUB) = sDirectory & sFile
                    If sSubDir(lSubDirUB).EndsWith("\") = False Then sSubDir(lSubDirUB) = sSubDir(lSubDirUB) & "\"
                ElseIf sFile.ToUpper <> "BPFILES.TXT" AndAlso sFile.ToUpper <> "SAFILES.TXT" AndAlso sFile.ToUpper <> "VERIFY.TXT" AndAlso sFile.ToUpper <> "VERSIONS.TXT" Then

                    'File name
                    sEntry = Replace$(sDirectory, msMainPath, "") & sFile
                    Dim sCompare As String = sEntry
                    sCompare = sCompare.Trim.ToUpper
                    sEntry &= "|"
                    Dim sVerify As String = sEntry

                    sEntry &= FileLen(sDirectory & sFile) & "|"

                    'file date and time                    
                    sEntry = sEntry & FileDateTime(sDirectory & sFile)

                    Dim bSA As Boolean = False

                    For Y As Int32 = 0 To mlSAOnlyUB
                        If msSAOnly(Y) = sCompare Then
                            sEntry &= "|SA"
                            bSA = True
                            Exit For
                        End If
                    Next Y
                    If bSA = False Then
                        For Y As Int32 = 0 To mlBPOnlyUB
                            If msBPOnly(Y) = sCompare Then
                                sEntry &= "|BP"
                                Exit For
                            End If
                        Next
                    End If

                    Dim bFound As Boolean = False
                    For X = 0 To mlPrevUB
                        If msPrevious(X) = sEntry Then
                            bFound = True
                            Exit For
                        End If
                    Next X
                    If bFound = False Then lstChanges.Items.Add(sFile)

                    'Print #mlFileNum, sEntry
                    moWriter.WriteLine(sEntry)
                    'Try
                    '    Dim oTmpFS As New IO.FileStream(sDirectory & sFile, IO.FileMode.Open)
                    '    oTmpFS.Seek(-1, IO.SeekOrigin.End)
                    '    Dim lEnd As Int32 = oTmpFS.ReadByte
                    '    moVerify.WriteLine(sVerify & lEnd.ToString)

                    '    oTmpFS.Close()
                    '    oTmpFS.Dispose()
                    'Catch
                    'End Try
                End If
            End If
            sFile = Dir()
        End While

        For X = 0 To lSubDirUB
            Call FillSubDirectory(sSubDir(X))
        Next X

    End Sub
End Class
