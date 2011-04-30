Option Strict On

Imports Microsoft.DirectX.AudioVideoPlayback

Public Class SongPlayer

    Private moBuffer(1) As Audio
    Private mlCurrentBuffer As Int32 = 1

    Public ReadOnly Property IsPlaying() As Boolean
        Get
            If moBuffer(mlCurrentBuffer) Is Nothing = False Then Return moBuffer(mlCurrentBuffer).State = StateFlags.Running Else Return False
            Return False
        End Get
    End Property

    Public ReadOnly Property WithinFiveSecondsOfEnd() As Boolean
        Get
            If moBuffer(mlCurrentBuffer) Is Nothing = False Then 
                Return moBuffer(mlCurrentBuffer).Duration - moBuffer(mlCurrentBuffer).CurrentPosition < 5
            End If
            Return True
        End Get
    End Property 

    Public Property CurrentBufferLinearVolume() As Int32
        Get
            If moBuffer(mlCurrentBuffer) Is Nothing Then Return 0
            Return SoundMgr.DirectXVolumeToLinear(moBuffer(mlCurrentBuffer).Volume, 100)
        End Get
        Set(ByVal value As Int32)
            If moBuffer(mlCurrentBuffer) Is Nothing Then Return
            moBuffer(mlCurrentBuffer).Volume = SoundMgr.LinearVolumeToDirectX(value, 100)
        End Set
    End Property

    Public Property SecondBufferLinearVolume() As Int32
        Get
            Dim lBuffer As Int32 = 0
            If mlCurrentBuffer = 0 Then lBuffer = 1
            If moBuffer(lBuffer) Is Nothing Then Return 0
            Return SoundMgr.DirectXVolumeToLinear(moBuffer(lBuffer).Volume, 100)
        End Get
        Set(ByVal value As Int32)
            Dim lBuffer As Int32 = 0
            If mlCurrentBuffer = 0 Then lBuffer = 1
            If moBuffer(lBuffer) Is Nothing Then Return
            moBuffer(lBuffer).Volume = SoundMgr.LinearVolumeToDirectX(value, 100)
        End Set
    End Property

    Protected Overrides Sub Finalize()
        If moBuffer(0) Is Nothing = False Then moBuffer(0).Dispose()
        If moBuffer(1) Is Nothing = False Then moBuffer(1).Dispose()
        Erase moBuffer
        MyBase.Finalize()
    End Sub

    Public Function LoadUnusedBuffer(ByVal sFileName As String, ByVal lStartVol As Int32) As Boolean
        Try
            Dim lLoadIdx As Int32 = 0
            If mlCurrentBuffer = 0 Then lLoadIdx = 1
            If moBuffer(lLoadIdx) Is Nothing = False Then moBuffer(lLoadIdx).Dispose()
            moBuffer(lLoadIdx) = Nothing
            moBuffer(lLoadIdx) = New Audio(sFileName, False)
            moBuffer(lLoadIdx).Volume = SoundMgr.LinearVolumeToDirectX(lStartVol, 100)
            moBuffer(lLoadIdx).CurrentPosition = 0
        Catch ex As Exception
            If goUILib Is Nothing = False Then goUILib.AddNotification("Music: " & ex.Message, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End Try
        Return True
    End Function

    Public Function StartMusic(ByVal sFileName As String, ByVal lStartVol As Int32) As Boolean
        Try
            mlCurrentBuffer = 0
            If moBuffer(mlCurrentBuffer) Is Nothing = False Then moBuffer(mlCurrentBuffer).Dispose()
            moBuffer(mlCurrentBuffer) = Nothing
            moBuffer(mlCurrentBuffer) = New Audio(sFileName, False)
            moBuffer(mlCurrentBuffer).Volume = SoundMgr.LinearVolumeToDirectX(lStartVol, 100)
            moBuffer(mlCurrentBuffer).CurrentPosition = 0
            moBuffer(mlCurrentBuffer).Play()
            mlCurrentBuffer = 1
        Catch ex As Exception
            If goUILib Is Nothing = False Then goUILib.AddNotification("Music: " & ex.Message, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End Try
        Return True
    End Function

    Public Sub StartSecondBuffer()
        Dim lBuffer As Int32 = 0
        If mlCurrentBuffer = 0 Then lBuffer = 1
        If moBuffer(lBuffer) Is Nothing = False Then moBuffer(lBuffer).Play()
    End Sub

    Public Sub ChangeCurrentBuffer()
        If moBuffer(mlCurrentBuffer) Is Nothing = False Then
            moBuffer(mlCurrentBuffer).Stop()
            moBuffer(mlCurrentBuffer).CurrentPosition = 0
        End If
        If mlCurrentBuffer = 0 Then mlCurrentBuffer = 1 Else mlCurrentBuffer = 0
    End Sub

    Public Sub StopAll()
        If moBuffer(0) Is Nothing = False Then moBuffer(0).Stop()
        If moBuffer(1) Is Nothing = False Then moBuffer(1).Stop()
        mlCurrentBuffer = 0
    End Sub
End Class
