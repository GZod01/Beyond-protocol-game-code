Imports System.Text.RegularExpressions

Namespace Pop3
    ''' <summary>
    ''' Summary description for Pop3Statics.
    ''' </summary>
    Public Class Pop3Statics
        Public Shared DataFolder As String = AppDomain.CurrentDomain.BaseDirectory & "\Pop3temp"

        Public Shared Function FromQuotedPrintable(ByVal inString As String) As String
            Dim outputString As String = Nothing
            Dim inputString As String = inString.Replace("=" & vbLf, "")

            If inputString.Length > 3 Then
                ' initialise output string ...
                outputString = ""

                Dim x As Integer = 0
                While x < inputString.Length
                    Dim s1 As String = inputString.Substring(x, 1)

                    If (s1.Equals("=")) AndAlso ((x + 2) < inputString.Length) Then
                        Dim hexString As String = inputString.Substring(x + 1, 2)

                        ' if hexadecimal ...
                        If Regex.Match(hexString.ToUpper(), "^[A-F|0-9]+[A-F|0-9]+$").Success Then
                            ' convert to string representation ...
                            outputString += System.Text.Encoding.ASCII.GetString(New Byte() {System.Convert.ToByte(hexString, 16)})
                            x += 3
                        Else
                            outputString += s1
                            x += 1
                        End If
                    Else
                        outputString += s1
                        x += 1
                    End If
                End While
            Else
                outputString = inputString
            End If

            Return outputString.Replace(vbLf, vbCr & vbLf)
        End Function
    End Class
End Namespace
