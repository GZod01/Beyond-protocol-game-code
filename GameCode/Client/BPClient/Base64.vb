Option Strict On

Public Class Base64
    Private Const yEquals As Byte = 61   'Asc("=")

    Private Const yMask1 As Byte = 3        '00000011
    Private Const yMask2 As Byte = 15       '00001111
    Private Const yMask3 As Byte = 63       '00111111
    Private Const yMask4 As Byte = 192      '11000000
    Private Const yMask5 As Byte = 240      '11110000
    Private Const yMask6 As Byte = 252      '11111100

    Private Const yShift2 As Byte = 4
    Private Const yShift4 As Byte = 16
    Private Const yShift6 As Byte = 64

    Private Shared Base64Lookup() As Byte
    Private Shared Base64Reverse() As Byte

    Private Shared mbInitialized As Boolean = False

    Public Shared Function EncodeString(ByVal sText As String) As String
        Dim Data() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(sText)
        Return EncodeByteArray(Data)
    End Function

    Public Shared Function EncodeByteArray(ByVal Data() As Byte) As String
        Initialize()
        Dim EncodedData() As Byte

        Dim DataLength As Int32
        Dim EncodedLength As Int32

        Dim Data0 As Int32
        Dim Data1 As Int32
        Dim Data2 As Int32

        Dim l As Int32
        Dim m As Int32

        Dim Index As Int32

        Dim CharCount As Int32

        DataLength = UBound(Data) + 1

        EncodedLength = (DataLength \ 3) * 4
        If DataLength Mod 3 > 0 Then EncodedLength = EncodedLength + 4
        EncodedLength = EncodedLength + ((EncodedLength \ 76) * 2)
        If EncodedLength Mod 78 = 0 Then EncodedLength = EncodedLength - 2
        ReDim EncodedData(EncodedLength - 1)

        m = (DataLength) Mod 3

        For l = 0 To UBound(Data) - m Step 3
            Data0 = Data(l)
            Data1 = Data(l + 1)
            Data2 = Data(l + 2)
            EncodedData(Index) = Base64Lookup(Data0 \ yShift2)
            EncodedData(Index + 1) = Base64Lookup(((Data0 And yMask1) * yShift4) Or (Data1 \ yShift4))
            EncodedData(Index + 2) = Base64Lookup(((Data1 And yMask2) * yShift2) Or (Data2 \ yShift6))
            EncodedData(Index + 3) = Base64Lookup(Data2 And yMask3)
            Index = Index + 4
            CharCount = CharCount + 4

            If CharCount = 76 And Index < EncodedLength Then
                EncodedData(Index) = 13
                EncodedData(Index + 1) = 10
                CharCount = 0
                Index = Index + 2
            End If
        Next

        If m = 1 Then
            Data0 = Data(l)
            EncodedData(Index) = Base64Lookup((Data0 \ yShift2))
            EncodedData(Index + 1) = Base64Lookup((Data0 And yMask1) * yShift4)
            EncodedData(Index + 2) = yEquals
            EncodedData(Index + 3) = yEquals
            Index = Index + 4
        ElseIf m = 2 Then
            Data0 = Data(l)
            Data1 = Data(l + 1)
            EncodedData(Index) = Base64Lookup((Data0 \ yShift2))
            EncodedData(Index + 1) = Base64Lookup(((Data0 And yMask1) * yShift4) Or (Data1 \ yShift4))
            EncodedData(Index + 2) = Base64Lookup((Data1 And yMask2) * yShift2)
            EncodedData(Index + 3) = yEquals
            Index = Index + 4
        End If

        Return System.Text.ASCIIEncoding.ASCII.GetString(EncodedData)
    End Function

    Public Shared Function DecodeToString(ByVal sEncodedText As String) As String
        Dim Data() As Byte = DecodeToByteArray(sEncodedText)
        Return System.Text.ASCIIEncoding.ASCII.GetString(Data)
    End Function

    Public Shared Function DecodeToByteArray(ByVal EncodedText As String) As Byte()
        Initialize()
        Dim Data() As Byte
        Dim EncodedData() As Byte

        Dim DataLength As Int32
        Dim EncodedLength As Int32

        Dim EncodedData0 As Int32
        Dim EncodedData1 As Int32
        Dim EncodedData2 As Int32
        Dim EncodedData3 As Int32

        Dim l As Int32
        Dim m As Int32

        Dim Index As Int32


        EncodedData = System.Text.ASCIIEncoding.ASCII.GetBytes(Replace$(Replace$(EncodedText, vbCrLf, ""), "=", ""))
        EncodedLength = UBound(EncodedData) + 1
        DataLength = (EncodedLength \ 4) * 3

        m = EncodedLength Mod 4
        If m = 2 Then
            DataLength = DataLength + 1
        ElseIf m = 3 Then
            DataLength = DataLength + 2
        End If

        ReDim Data(DataLength - 1)

        For l = 0 To UBound(EncodedData) - m Step 4
            EncodedData0 = Base64Reverse(EncodedData(l))
            EncodedData1 = Base64Reverse(EncodedData(l + 1))
            EncodedData2 = Base64Reverse(EncodedData(l + 2))
            EncodedData3 = Base64Reverse(EncodedData(l + 3))
            Data(Index) = CByte((EncodedData0 * yShift2) Or (EncodedData1 \ yShift4))
            Data(Index + 1) = CByte(((EncodedData1 And yMask2) * yShift4) Or (EncodedData2 \ yShift2))
            Data(Index + 2) = CByte(((EncodedData2 And yMask1) * yShift6) Or EncodedData3)
            Index = Index + 3
        Next

        Select Case ((UBound(EncodedData) + 1) Mod 4)
            Case 2
                EncodedData0 = Base64Reverse(EncodedData(l))
                EncodedData1 = Base64Reverse(EncodedData(l + 1))
                Data(Index) = CByte((EncodedData0 * yShift2) Or (EncodedData1 \ yShift4))
            Case 3
                EncodedData0 = Base64Reverse(EncodedData(l))
                EncodedData1 = Base64Reverse(EncodedData(l + 1))
                EncodedData2 = Base64Reverse(EncodedData(l + 2))
                Data(Index) = CByte((EncodedData0 * yShift2) Or (EncodedData1 \ yShift4))
                Data(Index + 1) = CByte(((EncodedData1 And yMask2) * yShift4) Or (EncodedData2 \ yShift2))
        End Select

        Return Data

    End Function

    Private Shared Sub Initialize()
        If mbInitialized = False Then
            mbInitialized = True
            Dim l As Int32

            ReDim Base64Reverse(255)

            Base64Lookup = System.Text.ASCIIEncoding.ASCII.GetBytes("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/")

            For l = 0 To 63
                Base64Reverse(Base64Lookup(l)) = CByte(l)
            Next
        End If
    End Sub
End Class