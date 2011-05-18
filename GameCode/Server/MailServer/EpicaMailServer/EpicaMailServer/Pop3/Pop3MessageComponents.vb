Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Collections

Namespace Pop3
    ''' <summary>
    ''' Summary description for Pop3MessageBody.
    ''' </summary>
    Public Class Pop3MessageComponents
        Private m_component As New ArrayList()

        Public ReadOnly Property ComponentEnumerator() As IEnumerator
            Get
                Return m_component.GetEnumerator()
            End Get
        End Property

        Public ReadOnly Property NumberOfComponents() As Integer
            Get
                Return m_component.Count
            End Get
        End Property

        Public Sub New(ByVal lines As String(), ByVal startOfBody As Int32, ByVal multipartBoundary As String, ByVal mainContentType As String)
            Dim stopOfBody As Int32 = lines.Length

            ' if this email is a mixture of message
            ' and attachments ...

            If multipartBoundary Is Nothing Then
                Dim sbText As New StringBuilder()

                For i As Int32 = startOfBody To stopOfBody - 1
                    sbText.Append(lines(i).Replace(vbLf, "").Replace(vbCr, ""))
                Next

                ' create a new component ...
                m_component.Add(New Pop3Component(mainContentType, sbText.ToString()))
            Else
                Dim boundary As String = multipartBoundary

                Dim firstComponent As Boolean = True

                ' loop through whole of email ...
                Dim i As Int32 = startOfBody
                While i < stopOfBody
                    Dim boundaryFound As Boolean = True

                    Dim contentType As String = Nothing
                    Dim name As String = Nothing
                    Dim filename As String = Nothing
                    Dim contentTransferEncoding As String = Nothing
                    Dim contentDescription As String = Nothing
                    Dim contentDisposition As String = Nothing
                    Dim data As String = Nothing

                    ' if first block of multipart data ...
                    If firstComponent Then
                        boundaryFound = False
                        firstComponent = False

                        While i < stopOfBody
                            Dim line As String = lines(i).Replace(vbLf, "").Replace(vbCr, "")

                            ' if multipart boundary found then
                            ' exit loop ...

                            If Pop3Parse.GetSubHeaderLineType(line, boundary) = Pop3Parse.MultipartBoundaryFound Then
                                boundaryFound = True
                                i += 1
                                Exit While
                            Else
                                ' ... else read next line ...
                                i += 1
                            End If
                        End While
                    End If

                    ' check to see whether multipart boundary
                    ' was found ...

                    If Not boundaryFound Then
                        Throw New Pop3MissingBoundaryException("Missing multipart boundary: " & boundary)
                    End If

                    Dim endOfHeader As Boolean = False

                    ' read header information ...
                    While (i < stopOfBody)
                        Dim line As String = lines(i).Replace(vbLf, "").Replace(vbCr, "")

                        Dim lineType As Integer = Pop3Parse.GetSubHeaderLineType(line, boundary)

                        Select Case lineType
                            Case Pop3Parse.ContentTypeType
                                contentType = Pop3Parse.ContentType(line)
                                Exit Select

                            Case Pop3Parse.ContentTransferEncodingType
                                contentTransferEncoding = Pop3Parse.ContentTransferEncoding(line)
                                Exit Select

                            Case Pop3Parse.ContentDispositionType
                                contentDisposition = Pop3Parse.ContentDisposition(line)
                                Exit Select

                            Case Pop3Parse.ContentDescriptionType
                                contentDescription = Pop3Parse.ContentDescription(line)
                                Exit Select

                            Case Pop3Parse.EndOfHeader
                                endOfHeader = True
                                Exit Select
                        End Select

                        i += 1

                        If endOfHeader Then
                            Exit While
                        Else
                            While i < stopOfBody
                                ' if more lines to read for this line ...
                                If line.Substring(line.Length - 1, 1).Equals(";") Then

                                    Dim nextLine As String = lines(i).Replace(vbCr, "").Replace(vbLf, "")

                                    Select Case Pop3Parse.GetSubHeaderNextLineType(nextLine)
                                        Case Pop3Parse.NameType
                                            name = Pop3Parse.Name(nextLine)
                                            Exit Select

                                        Case Pop3Parse.FilenameType
                                            filename = Pop3Parse.Filename(nextLine)
                                            Exit Select

                                        Case Pop3Parse.EndOfHeader
                                            endOfHeader = True
                                            Exit Select
                                    End Select

                                    If Not endOfHeader Then
                                        ' point line to current line ...
                                        line = nextLine
                                        i += 1
                                    Else
                                        Exit While
                                    End If
                                Else
                                    Exit While
                                End If
                            End While
                        End If
                    End While

                    boundaryFound = False

                    Dim sbText As New StringBuilder()

                    Dim emailComposed As Boolean = False

                    ' store the actual data ...
                    While i < stopOfBody
                        ' get the next line ...
                        Dim line As String = lines(i).Replace(vbLf, "")

                        ' if we've found the boundary ...
                        If Pop3Parse.GetSubHeaderLineType(line, boundary) = Pop3Parse.MultipartBoundaryFound Then
                            boundaryFound = True
                            i += 1
                            Exit While
                        ElseIf Pop3Parse.GetSubHeaderLineType(line, boundary) = Pop3Parse.ComponetsDone Then
                            emailComposed = True
                            Exit While
                        End If

                        ' add this line to data ...
                        sbText.Append(lines(i))
                        i += 1
                    End While

                    If sbText.Length > 0 Then
                        data = sbText.ToString()
                    End If

                    ' create a new component ...
                    m_component.Add(New Pop3Component(contentType, name, filename, contentTransferEncoding, contentDescription, contentDisposition, _
                     data))

                    ' if all multiparts have been
                    ' composed then exit ..

                    If emailComposed Then
                        Exit While
                    End If
                End While
            End If
        End Sub
    End Class
End Namespace
