Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Pop3
    ''' <summary>
    ''' Summary description for Pop3Attachment.
    ''' </summary>
    Public Class Pop3Component
        Private m_contentType As String
        Private m_name As String
        Private m_filename As String
        Private m_contentTransferEncoding As String
        Private m_contentDescription As String
        Private m_contentDisposition As String
        Private m_data As String
        Private m_filePath As String

        Public m_binaryData As Byte()

        Public ReadOnly Property FileExtension() As String
            Get
                Dim extension As String = Nothing

                ' if file has a filename and the filename
                ' has an extension ...

                If (m_filename IsNot Nothing) AndAlso Regex.Match(m_filename, "^.*\..*$").Success Then
                    ' get extension ...
                    extension = Regex.Replace(m_name, "^[^\.]*\.([^\.]+)$", "$1")
                End If

                ' NOTE: return null if extension
                ' not found ...
                Return extension
            End Get
        End Property

        Public ReadOnly Property FileNoExtension() As String
            Get
                Dim extension As String = Nothing

                ' if file has a filename and the filename
                ' has an extension ...

                If (m_filename IsNot Nothing) AndAlso Regex.Match(m_filename, "^.*\..*$").Success Then
                    ' get extension ...
                    extension = Regex.Replace(m_name, "^([^\.]*)\.[^\.]+$", "$1")
                End If

                ' NOTE: return null if extension
                ' not found ...
                Return extension
            End Get
        End Property

        Public ReadOnly Property FilePath() As String
            Get
                Return m_filePath
            End Get
        End Property

        Public ReadOnly Property Filename() As String
            Get
                Return m_filename
            End Get
        End Property

        Public ReadOnly Property ContentType() As String
            Get
                Return m_contentType
            End Get
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return m_name
            End Get
        End Property

        Public ReadOnly Property ContentTransferEncoding() As String
            Get
                Return m_contentTransferEncoding
            End Get
        End Property

        Public ReadOnly Property ContentDescription() As String
            Get
                Return m_contentDescription
            End Get
        End Property

        Public ReadOnly Property ContentDisposition() As String
            Get
                Return m_contentDisposition
            End Get
        End Property

        Public ReadOnly Property Data() As String
            Get
                Return m_data
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return "Content-Type: " & m_contentType & vbCr & vbLf & "Name: " & m_name & vbCr & vbLf & "Filename: " & m_filename & vbCr & vbLf & "Content-Transfer-Encoding: " & m_contentTransferEncoding & vbCr & vbLf & "Content-Description: " & m_contentDescription & vbCr & vbLf & "Content-Disposition: " & m_contentDisposition & vbCr & vbLf & "Data :" & m_data
        End Function


        Public ReadOnly Property IsBody() As Boolean
            Get
                Return (m_contentDisposition Is Nothing)
            End Get
        End Property

        Public ReadOnly Property IsAttachment() As Boolean
            Get
                Dim ret As Boolean = False

                If m_contentDisposition IsNot Nothing Then
                    ret = Regex.Match(m_contentDisposition, "^attachment.*$").Success
                End If

                Return ret
            End Get
        End Property

        Private Sub DecodeData()
            ' if this data is an attachment ...
            If m_contentDisposition IsNot Nothing Then

                ' create data folder if it doesn't exist ...
                If Not Directory.Exists(Pop3Statics.DataFolder) Then
                    Directory.CreateDirectory(Pop3Statics.DataFolder)
                End If

                m_filePath = Pop3Statics.DataFolder + "\" & m_filename

                ' if BASE-64 data ...
                If (m_contentDisposition.Equals("attachment;")) AndAlso (m_contentTransferEncoding.ToUpper().Equals("BASE64")) Then
                    ' convert attachment from BASE64 ...
                    m_binaryData = Convert.FromBase64String(m_data.Replace(vbLf, ""))

                    Dim bw As New BinaryWriter(New FileStream(m_filePath, FileMode.Create))

                    bw.Write(m_binaryData)
                    bw.Flush()
                    bw.Close()
                    ' if PRINTABLE ...
                ElseIf (m_contentDisposition.Equals("attachment;")) AndAlso (m_contentTransferEncoding.ToUpper().Equals("QUOTED-PRINTABLE")) Then
                    Using sw As StreamWriter = File.CreateText(m_filePath)
                        sw.Write(Pop3Statics.FromQuotedPrintable(m_data))
                        sw.Flush()
                        sw.Close()
                    End Using
                End If
            End If
        End Sub

        Public Sub New(ByVal contentType As String, ByVal data As String)
            m_contentType = contentType
            m_data = data
        End Sub

        Public Sub New(ByVal contentType As String, ByVal name As String, ByVal filename As String, ByVal contentTransferEncoding As String, ByVal contentDescription As String, ByVal contentDisposition As String, _
        ByVal data As String)
            m_contentType = contentType
            m_name = name
            m_filename = filename
            m_contentTransferEncoding = contentTransferEncoding
            m_contentDescription = contentDescription
            m_contentDisposition = contentDisposition
            m_data = data

            DecodeData()
        End Sub
    End Class
End Namespace
