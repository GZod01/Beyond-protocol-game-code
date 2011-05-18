Imports System.Text.RegularExpressions

Namespace Pop3
    ''' <summary>
    ''' Summary description for Pop3ParseMessage.
    ''' </summary>
    Public Class Pop3Parse
        Private Shared m_lineUpperTypeString As String() = {"From", "To", "Subject", "Content-Type"}

        Private Shared m_lineSubTypeString As String() = {"Content-Type", "Content-Transfer-Encoding", "Content-Description", "Content-Disposition"}

        Private Shared m_nextLineTypeString As String() = {"name", "filename"}

        ' Mapping to lineSubTypeString ...
        Public Const ContentTypeType As Integer = 0
        Public Const ContentTransferEncodingType As Integer = 1
        Public Const ContentDescriptionType As Integer = 2
        Public Const ContentDispositionType As Integer = 3

        ' Mapping to nextLineTypeString ...
        Public Const NameType As Integer = 0
        Public Const FilenameType As Integer = 1

        ' Non-string mappers ...
        Public Const UnknownType As Integer = -99
        Public Const EndOfHeader As Integer = -98
        Public Const MultipartBoundaryFound As Integer = -97
        Public Const ComponetsDone As Integer = -96

        Public Shared ReadOnly Property LineUpperTypeString() As String()
            Get
                Return m_lineUpperTypeString
            End Get
        End Property

        Public Shared ReadOnly Property LineSubTypeString() As String()
            Get
                Return m_lineSubTypeString
            End Get
        End Property

        Public Shared ReadOnly Property NextLineTypeString() As String()
            Get
                Return m_nextLineTypeString
            End Get
        End Property

        Public Shared Function From(ByVal line As String) As String
            Return Regex.Replace(line, "^From:.*[ |<]([a-z|A-Z|0-9|\.|\-|_]+@[a-z|A-Z|0-9|\.|\-|_]+).*$", "$1")
        End Function

        Public Shared Function Subject(ByVal line As String) As String
            Return Regex.Replace(line, "^Subject: (.*)$", "$1")
        End Function

        Public Shared Function [To](ByVal line As String) As String
            Return Regex.Replace(line, "^To:.*[ |<]([a-z|A-Z|0-9|\.|\-|_]+@[a-z|A-Z|0-9|\.|\-|_]+).*$", "$1")
        End Function

        Public Shared Function ContentType(ByVal line As String) As String
            Return Regex.Replace(line, "^Content-Type: (.*)$", "$1")
        End Function

        Public Shared Function ContentTransferEncoding(ByVal line As String) As String
            Return Regex.Replace(line, "^Content-Transfer-Encoding: (.*)$", "$1")
        End Function

        Public Shared Function ContentDescription(ByVal line As String) As String
            Return Regex.Replace(line, "^Content-Description: (.*)$", "$1")
        End Function

        Public Shared Function ContentDisposition(ByVal line As String) As String
            Return Regex.Replace(line, "^Content-Disposition: (.*)$", "$1")
        End Function

        Public Shared Function IsMultipart(ByVal line As String) As Boolean
            Return Regex.Match(line, "^multipart/.*").Success
        End Function

        Public Shared Function MultipartBoundary(ByVal line As String) As String
            Return Regex.Replace(line, "^.*boundary=[""]*([^""]*).*$", "$1")
        End Function

        Public Shared Function Name(ByVal line As String) As String
            Return Regex.Replace(line, "^[ |" & vbTab & "]+name=[""]*([^""]*).*$", "$1")
        End Function

        Public Shared Function Filename(ByVal line As String) As String
            Return Regex.Replace(line, "^[ |" & vbTab & "]+filename=[""]*([^""]*).*$", "$1")
        End Function

        Public Shared Function GetSubHeaderNextLineType(ByVal line As String) As Integer
            Dim lineType As Integer = Pop3Parse.UnknownType

            For i As Integer = 0 To Pop3Parse.NextLineTypeString.Length - 1
                Dim match As String = Pop3Parse.NextLineTypeString(i)

                If Regex.Match(line, "^[ |" & vbTab & "]+" & match & "=" & ".*$").Success Then
                    lineType = i
                    Exit For
                End If
                If line.Length = 0 Then
                    lineType = Pop3Parse.EndOfHeader
                    Exit For
                End If
            Next

            Return lineType
        End Function

        Public Shared Function GetSubHeaderLineType(ByVal line As String, ByVal boundary As String) As Integer
            Dim lineType As Integer = Pop3Parse.UnknownType

            For i As Integer = 0 To Pop3Parse.LineSubTypeString.Length - 1
                Dim match As String = Pop3Parse.LineSubTypeString(i)

                If Regex.Match(line, "^" & match & ":" & ".*$").Success Then
                    lineType = i
                    Exit For
                ElseIf line.Equals("--" & boundary) Then
                    lineType = Pop3Parse.MultipartBoundaryFound
                    Exit For
                ElseIf line.Equals("--" & boundary & "--") Then
                    lineType = Pop3Parse.ComponetsDone
                    Exit For
                End If
                If line.Length = 0 Then
                    lineType = Pop3Parse.EndOfHeader
                    Exit For
                End If
            Next

            Return lineType
        End Function
    End Class
End Namespace
