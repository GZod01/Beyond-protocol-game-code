Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class BPLine
    Private Shared moLine As Line
    Public Shared Sub ClearLine()
        If moLine Is Nothing = False Then moLine.Dispose()
        moLine = Nothing
    End Sub

    Public Shared Sub DrawLine(ByVal fWidth As Single, ByVal bAA As Boolean, ByVal vPts() As Vector2, ByVal clrVal As System.Drawing.Color)
        Try
            If moLine Is Nothing OrElse moLine.Disposed = True Then
                moLine = New Line(goUILib.oDevice)
            End If
        Catch
        End Try

        With moLine
            Try
                .Antialias = bAA
                .Width = fWidth
            Catch
            End Try
            Dim bBegun As Boolean = False
            Try
                .Begin()
                bBegun = True
            Catch
            End Try
            If bBegun = True Then
                Try
                    .Draw(vPts, clrVal)
                Catch
                End Try
            End If
            Try
                If bBegun = True Then .End()
            Catch
            End Try
        End With

    End Sub

    Private Shared mbMultiDrawBegin As Boolean = False
    Public Shared Sub PrepareMultiDraw(ByVal fWidth As Single, ByVal bAA As Boolean)
        Try
            If moLine Is Nothing OrElse moLine.Disposed = True Then
                moLine = New Line(goUILib.oDevice)
            End If
        Catch
        End Try
        With moLine
            Try
                .Antialias = bAA
                .Width = fWidth
            Catch
            End Try
            Try
                .Begin()
                mbMultiDrawBegin = True
            Catch
            End Try 
        End With
    End Sub
    Public Shared Sub EndMultiDraw()
        Try
            If mbMultiDrawBegin = True Then moLine.End()
        Catch
        End Try
    End Sub
    Public Shared Sub MultiDrawLine(ByVal vPts() As Vector2, ByVal clrVal As System.Drawing.Color)
        If mbMultiDrawBegin = True Then
            Try
                moLine.Draw(vPts, clrVal)
            Catch
            End Try
        End If
    End Sub
End Class
