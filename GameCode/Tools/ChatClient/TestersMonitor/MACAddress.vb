Imports System.Management

Public Class MACAddress
    Public Shared Function GetMACAddressFromIP(ByVal sIP As String) As Byte()
        Dim sResult As String = ""
        Try
            Dim searcher As New ManagementObjectSearcher("Select * From Win32_NetworkAdapterConfiguration")
            For Each queryObj As ManagementObject In searcher.Get()
                If queryObj("IPAddress") Is Nothing = False Then
                    For Each sTemp As String In CType(queryObj("IPAddress"), System.Array)
                        If sIP = sTemp Then
                            If queryObj("MACAddress") Is Nothing = False Then
                                sResult = queryObj("MACAddress").ToString
                                Exit For
                            End If
                        End If
                    Next
                End If
            Next
        Catch err As ManagementException
            sResult = "00:00:00:00:00:00"
        End Try

        Dim sFinal() As String = Split(sResult, ":")
        Dim yReturn(Math.Min(sFinal.GetUpperBound(0), 5)) As Byte
        For X As Int32 = 0 To yReturn.GetUpperBound(0)
            yReturn(X) = CByte(Math.Min(255, Math.Max(0, CInt(Val("&H" & sFinal(X))))))
        Next X
        Return yReturn
    End Function
End Class