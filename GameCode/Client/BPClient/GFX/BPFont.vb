Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class BPFont

    Private Shared moFont() As Font
    Private Shared msFontName() As String
    Private Shared mfFontSize() As Single
    Private Shared mlFontStyle() As FontStyle
    Private Shared mlFontUB As Int32 = -1

    Public Shared Sub ClearAllFonts()
        Try
            For X As Int32 = 0 To mlFontUB
                If moFont(X) Is Nothing = False Then moFont(X).Dispose()
                moFont(X) = Nothing
            Next X
        Catch
        End Try

        ReDim moFont(99)
        ReDim msFontName(99)
        ReDim mfFontSize(99)
        ReDim mlFontStyle(99)
        mlFontUB = -1
    End Sub

    Private Shared Function GetDXFont(ByVal oFont As System.Drawing.Font) As Font
        Dim sName As String = oFont.Name
        Dim fSize As Single = oFont.Size
        Dim lStyle As FontStyle = oFont.Style

        Dim oDXFont As Font = Nothing
        For X As Int32 = 0 To mlFontUB
            If lStyle = mlFontStyle(X) AndAlso msFontName(X) = sName AndAlso mfFontSize(X) = fSize Then
                oDXFont = moFont(X)
                Exit For
            End If
        Next X
        If oDXFont Is Nothing Then
            Device.IsUsingEventHandlers = False
            mlFontUB += 1
            If mlFontUB > moFont.GetUpperBound(0) Then
                ReDim Preserve moFont(mlFontUB + 99)
                ReDim Preserve msFontName(mlFontUB + 99)
                ReDim Preserve mfFontSize(mlFontUB + 99)
                ReDim Preserve mlFontStyle(mlFontUB + 99)
            End If
            moFont(mlFontUB) = New Font(goUILib.oDevice, oFont)
            msFontName(mlFontUB) = oFont.Name
            mfFontSize(mlFontUB) = oFont.Size
            mlFontStyle(mlFontUB) = oFont.Style
            oDXFont = moFont(mlFontUB)
            Device.IsUsingEventHandlers = True
        End If
        Return oDXFont
    End Function

    Public Shared Sub DrawText(ByVal oFont As System.Drawing.Font, ByVal sText As String, ByVal rcDest As Rectangle, ByVal lFormat As DrawTextFormat, ByVal lClr As System.Drawing.Color)
        If sText Is Nothing OrElse sText = "" Then Return
        Device.IsUsingEventHandlers = False
        Try
            Dim oDXFont As Font = GetDXFont(oFont)
            oDXFont.DrawText(Nothing, sText, rcDest, lFormat, lClr)
        Catch
        End Try
        Device.IsUsingEventHandlers = True
    End Sub

    Public Shared Function AcquireDXFont(ByVal oFont As System.Drawing.Font) As Font
        Dim oDXFont As Font = Nothing
        Device.IsUsingEventHandlers = False
        Try
            oDXFont = GetDXFont(oFont)
        Catch
        End Try
        Device.IsUsingEventHandlers = True
        Return oDXFont
    End Function

    Public Shared Sub DrawTextAtPoint(ByVal oFont As System.Drawing.Font, ByVal sText As String, ByVal lX As Int32, ByVal lY As Int32, ByVal lClr As System.Drawing.Color)
        If sText Is Nothing OrElse sText = "" Then Return
        Device.IsUsingEventHandlers = False
        Try
            Dim oDXFont As Font = GetDXFont(oFont)
            oDXFont.DrawText(Nothing, sText, lX, lY, lClr)
        Catch
        End Try
        Device.IsUsingEventHandlers = True
    End Sub

    Public Shared Function MeasureString(ByVal oFont As System.Drawing.Font, ByVal sText As String, ByVal lFormat As DrawTextFormat) As Rectangle
        Dim rcResult As Rectangle = New Rectangle(5, 5, 5, 5)
        If sText Is Nothing OrElse sText = "" Then Return rcResult
        Device.IsUsingEventHandlers = False
        Try
            Dim oDXFont As Font = GetDXFont(oFont)
            rcResult = oDXFont.MeasureString(Nothing, sText, lFormat, System.Drawing.Color.White.ToArgb)
        Catch
        End Try
        Device.IsUsingEventHandlers = True
        Return rcResult
    End Function
End Class