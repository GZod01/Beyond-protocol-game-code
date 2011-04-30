Option Strict On
'Imports WMPLib

'ONLY MUSIC SPECIFIC STUFF GOES HERE!!!
Partial Class SoundMgr
    Private Const ml_EXCITEMENT_HIGH_THRESHOLD As Int32 = 700
    Private Const ml_EXCITEMENT_LOW_THRESHOLD As Int32 = 300

    Private msPlayListLull() As String
    Private myPlayListLullSeq() As Byte
    Private mlPlayListLullUB As Int32 = -1
    Private msPlayListExcite() As String
    Private myPlayListExciteSeq() As Byte
    Private mlPlayListExciteUB As Int32 = -1
    Private msPlayListIntro() As String
    Private mlPlayListIntroUB As Int32 = -1

    Public MusicOn As Boolean = False

    Public mbMusicInTransition As Boolean
    Private mbNextSongLoaded As Boolean = False
    Private myPlayListType As Byte      'TODO: for now, 0 = Intro, 1 = Lull, 2 = Excite
    Public Property PlayListType() As Byte
        Get
            Return myPlayListType
        End Get
        Set(ByVal value As Byte)
            If myPlayListType <> value Then
                mbMusicInTransition = True
            End If
            myPlayListType = value
        End Set
    End Property

    Private mlCurrentPlayListVal As Int32
    Private mbExcitedMusic As Boolean = False

    Public mlExcitementLevel As Int32 = 0
    Private mlLastExcitementLevelUpdate As Int32
    Private mfExcitementLevelMod As Single = 3.0F

    Private mlMusicStartTime As Int32
    Private mlMusicDuration As Int32 = Int32.MinValue

    'Only one 'buffer'
    'Private moMP As WindowsMediaPlayer = Nothing
    Private moMP As SongPlayer = Nothing

    Private mlPreviousMusicVolumeChange As Int32 = 0

    Public MusicStarted As Boolean = False

    Public Sub IncrementExcitementLevel(ByVal lVal As Int32)
        If MusicOn = True AndAlso myPlayListType <> 0 Then
            mlExcitementLevel += lVal
            mfExcitementLevelMod = 3.0F
        End If
    End Sub

    Public Sub ResetExcitementLevel()
        mlExcitementLevel = 0
    End Sub

    Private Sub InitializePlaylists()
        Dim sFile As String

        mlPlayListIntroUB = -1
        mlPlayListLullUB = -1
        mlPlayListExciteUB = -1

        'Load our Intro playlist
        sFile = Dir$(msMusicPath & "Intro\*.*", FileAttribute.Normal Or FileAttribute.Archive Or FileAttribute.Hidden Or FileAttribute.ReadOnly Or FileAttribute.System Or FileAttribute.Volume)
        While sFile <> ""
            mlPlayListIntroUB += 1
            ReDim Preserve msPlayListIntro(mlPlayListIntroUB)
            msPlayListIntro(mlPlayListIntroUB) = msMusicPath & "Intro\" & sFile
            sFile = Dir$()
        End While

        'Load our Lull Playlist
        sFile = Dir$(msMusicPath & "Lull\*.*", FileAttribute.Normal Or FileAttribute.Archive Or FileAttribute.Hidden Or FileAttribute.ReadOnly Or FileAttribute.System Or FileAttribute.Volume)
        While sFile <> ""
            mlPlayListLullUB += 1
            ReDim Preserve msPlayListLull(mlPlayListLullUB)
            ReDim Preserve myPlayListLullSeq(mlPlayListLullUB)
            msPlayListLull(mlPlayListLullUB) = msMusicPath & "Lull\" & sFile
            myPlayListLullSeq(mlPlayListLullUB) = 0
            sFile = Dir$()
        End While

        'Load our Excited Playlist
        sFile = Dir$(msMusicPath & "Excite\*.*", FileAttribute.Normal Or FileAttribute.Archive Or FileAttribute.Hidden Or FileAttribute.ReadOnly Or FileAttribute.System Or FileAttribute.Volume)
        While sFile <> ""
            mlPlayListExciteUB += 1
            ReDim Preserve msPlayListExcite(mlPlayListExciteUB)
            ReDim Preserve myPlayListExciteSeq(mlPlayListExciteUB)
            msPlayListExcite(mlPlayListExciteUB) = msMusicPath & "Excite\" & sFile
            myPlayListExciteSeq(mlPlayListExciteUB) = 0
            sFile = Dir$()
        End While

    End Sub

    Public Sub StartMusic()
        Dim oINI As InitFile = New InitFile()

        If Val(oINI.GetString("AUDIO", "MusicOn", "1")) <> 0 Then
            InitializePlaylists()

            If glCurrentEnvirView = CurrentView.eStartupDSELogo OrElse glCurrentEnvirView = CurrentView.eStartupLogin Then
                myPlayListType = 0      'intro music
                mlCurrentPlayListVal = 1
            Else : myPlayListType = 1     'lull
            End If

			If InitializeMusicPlayer() = False Then Return

			mbMusicInTransition = True

			mbExcitedMusic = False
			MusicOn = True
            MusicStarted = True

            'UpdateMusic()
		End If
	End Sub

	Public Function GetCurrentIntroSong() As String
		If myPlayListType <> 0 Then Return ""
		If msPlayListIntro Is Nothing = False AndAlso mlCurrentPlayListVal > -1 AndAlso mlCurrentPlayListVal <= msPlayListIntro.GetUpperBound(0) Then
			Return msPlayListIntro(mlCurrentPlayListVal)
		End If
		Return ""
	End Function

    Private Sub UpdateMusic()
        If moMP Is Nothing Then Return

        'Adjust the excitement buffer value
        If mlLastExcitementLevelUpdate = 0 Then
            mlLastExcitementLevelUpdate = timeGetTime
        Else
            If timeGetTime - mlLastExcitementLevelUpdate > 100 Then
                mlExcitementLevel -= CInt(mfExcitementLevelMod)
                If mlExcitementLevel < 0 Then
                    mlExcitementLevel = 0
                    mfExcitementLevelMod = 3.0F
                Else : mfExcitementLevelMod += 0.1F
                End If
                mlLastExcitementLevelUpdate = timeGetTime
            End If
        End If

        If mlExcitementLevel < ml_EXCITEMENT_LOW_THRESHOLD Then
            mbExcitedMusic = False
        ElseIf mlExcitementLevel > ml_EXCITEMENT_HIGH_THRESHOLD Then
            mbExcitedMusic = True
        End If

        Dim lMaxMusicVolume As Int32 = DirectXVolumeToLinear(VolumeGroups(VolumeGroup.eVG_Music), 100)
        Dim lTimeGetTime As Int32 = timeGetTime
        Dim bUpdateMusicVolume As Boolean = False
        If lTimeGetTime - mlPreviousMusicVolumeChange > 30 Then
            mlPreviousMusicVolumeChange = lTimeGetTime
            bUpdateMusicVolume = True
        End If

        'Ok, if we are not in a transition yet
        If mbMusicInTransition = False Then
            If bUpdateMusicVolume = True Then
                If moMP.CurrentBufferLinearVolume < lMaxMusicVolume Then moMP.CurrentBufferLinearVolume += Math.Min(2, lMaxMusicVolume - moMP.CurrentBufferLinearVolume)
            End If

            'see if we need to transition
            mbMusicInTransition = (myPlayListType <> 2 AndAlso mbExcitedMusic = True) OrElse _
               (myPlayListType = 2 AndAlso mbExcitedMusic = False) 'OrElse (moMP.playState <> WMPPlayState.wmppsStopped)
            mbMusicInTransition = mbMusicInTransition = True OrElse (moMP Is Nothing = False AndAlso moMP.IsPlaying = True AndAlso moMP.WithinFiveSecondsOfEnd = True)

            mbNextSongLoaded = False
        Else
            'Ok, we are in a transition... transition is always fade out now
            If moMP.CurrentBufferLinearVolume <> 0 Then 'moMP.settings.volume <> 0 Then

                If bUpdateMusicVolume = True Then moMP.CurrentBufferLinearVolume -= Math.Max(2, moMP.CurrentBufferLinearVolume - lMaxMusicVolume)
                If mbNextSongLoaded = False Then
                    StartNextSong(lMaxMusicVolume)
                    mbNextSongLoaded = True
                Else
                    If bUpdateMusicVolume = True AndAlso moMP.SecondBufferLinearVolume < lMaxMusicVolume Then moMP.SecondBufferLinearVolume += Math.Min(2, lMaxMusicVolume - moMP.SecondBufferLinearVolume)
                End If
            Else
                If mbNextSongLoaded = False Then
                    StartNextSong(lMaxMusicVolume)
                    moMP.CurrentBufferLinearVolume = lMaxMusicVolume
                    moMP.SecondBufferLinearVolume = lMaxMusicVolume
                End If
                mbMusicInTransition = False
                moMP.ChangeCurrentBuffer()

                mlMusicDuration = Int32.MinValue
                mlMusicStartTime = timeGetTime

                'Set our new volume 
                'moMP.settings.volume = DirectXVolumeToLinear(VolumeGroups(VolumeGroup.eVG_Music), 100)
            End If
        End If
    End Sub

    Private Sub StartNextSong(ByVal lMaxVol As Int32)
        'Start our next buffer... determine our next song
        Dim sSongName As String = ""
        If myPlayListType = 0 Then
            If moMP.IsPlaying = False Then
                sSongName = msPlayListIntro(0)
                mlCurrentPlayListVal = 0
            Else
                mlCurrentPlayListVal += 1
                If mlCurrentPlayListVal > mlPlayListIntroUB Then mlCurrentPlayListVal = 0
                sSongName = msPlayListIntro(mlCurrentPlayListVal)
            End If
        Else
            If mbExcitedMusic = True Then
                myPlayListType = 2
            Else : myPlayListType = 1
            End If

            Dim lVal As Int32
            Dim lUB As Int32
            Dim lMainUB As Int32 = -1
            If myPlayListType = 1 Then
                lMainUB = mlPlayListLullUB
                lUB = lMainUB
                For X As Int32 = 0 To lMainUB
                    If myPlayListLullSeq(X) <> 0 Then lUB -= 1
                Next X
            Else
                lMainUB = mlPlayListExciteUB
                lUB = lMainUB
                For X As Int32 = 0 To lMainUB
                    If myPlayListExciteSeq(X) <> 0 Then lUB -= 1
                Next X
            End If

            If lUB = -1 Then
                For X As Int32 = 0 To lMainUB
                    If myPlayListType = 1 Then
                        myPlayListLullSeq(X) = 0
                    Else : myPlayListExciteSeq(X) = 0
                    End If
                Next X
                lUB = lMainUB
            End If

            lVal = CInt(Int(Rnd() * (lUB + 1)))
            For X As Int32 = 0 To lMainUB
                If myPlayListType = 1 Then
                    If myPlayListLullSeq(X) = 0 Then
                        lVal -= 1
                        If lVal <= 0 Then
                            lVal = X
                            Exit For
                        End If
                    End If
                Else
                    If myPlayListExciteSeq(X) = 0 Then
                        lVal -= 1
                        If lVal <= 0 Then
                            lVal = X
                            Exit For
                        End If
                    End If
                End If
            Next X
            'While lVal = mlCurrentPlayListVal
            '    lVal = CInt(Int(GetNxtRnd() * (lUB + 1)))
            'End While
            mlCurrentPlayListVal = lVal

            If myPlayListType = 1 Then
                'Dim lTempLoc As Int32 = msPlayListLull(mlCurrentPlayListVal).LastIndexOf("\")
                'Dim sTemp As String = msPlayListLull(mlCurrentPlayListVal).Substring(lTempLoc)
                'goUILib.AddNotification("Playing " & sTemp, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                sSongName = msPlayListLull(mlCurrentPlayListVal)
                myPlayListLullSeq(mlCurrentPlayListVal) = 1
            Else
                sSongName = msPlayListExcite(mlCurrentPlayListVal)
                myPlayListExciteSeq(mlCurrentPlayListVal) = 1
            End If
        End If

        If sSongName <> "" Then
            If moMP.IsPlaying = True Then
                If moMP.LoadUnusedBuffer(sSongName, Math.Min(30, lMaxVol)) = True Then
                    moMP.StartSecondBuffer()
                End If
            Else
                moMP.StartMusic(sSongName, lMaxVol)
            End If
        End If
    End Sub

    Public Sub StopMusic()
        'If moMP Is Nothing = False Then momp. = "" 'stops the music
        If moMP Is Nothing = False Then moMP.StopAll()
        mlExcitementLevel = 0
        MusicStarted = False
    End Sub

	Public Function InitializeMusicPlayer() As Boolean
		Dim bResult As Boolean = False
		Try
			If moMP Is Nothing Then
                moMP = New SongPlayer() 'WindowsMediaPlayer()
                'moMP.settings.volume = 0		'set to min volume so update music gets the event
                'moMP.CurrentBufferLinearVolume = 0
            End If
            bResult = True
		Catch
			MusicOn = False
		End Try
		Return bResult
	End Function

End Class
