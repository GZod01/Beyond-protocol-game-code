Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.DirectSound
'NOTE: DO NOT IMPORT Microsoft.DirectX.Direct3D AS THIS COULD CAUSE CONFLICTS IN NAME RESOLUTION!!!

Public Class SoundMgr
    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32
    Private Const ml_MAX_BUFFS As Int32 = 32        '32 for now, we can change it later if needed
    Public Enum VolumeGroup As Byte
        eVG_Music = 0
        eVG_Fighting
        eVG_EntityAmbience
        eVG_EnvirAmbience
        eVG_EntityResponse
        eVG_UserInterface
        eVG_UnitSpeech
        eVG_GameNarrative
        eVG_TutorialVoice

        eVG_Count           'MUST BE THE LAST ITEM
    End Enum

    Public Enum SoundUsage As Byte
        '0 = not used, 1 = used but not 3D, 2 = used with 3D
        eNotUsed = 0

        'eUsedButNot3D = 1
        'e3DSound = 2

        'NON-3D SOUNDS HERE
        eALWAYS_PLAY = 4        'will play in the first available slot (NEVER LOOPS and NOT 3D)

        'To avoid previous code that may have needed 1 or 2, we'll start on 5
        eAmbience = 5
        eWeather = 6
        eNarrative = 7
        eUnitSpeech = 8
        eUserInterface = 9
        eTutorialVoice = 10

        e3D_SOUNDS = 30         'any sound usage above this uses the 3D Buffer

        eRadioChatter = 31
        eUnitSounds = 32
        eNonDeathExplosions = 33
        eDeathExplosions = 34
        eFireworks = 35

        eWeaponsFire = 50       'all weapons fire... do not actually use this but instead call the detailed version

        eWeaponsFireEnergy = 51
        eWeaponsFireProjectile = 52
        eWeaponsFireMissiles = 53
        eWeaponsFireBombs = 54
    End Enum

    Private moDevice As Device      'epica sound device

    Private moBuffers() As Buffer
    Private myBufferIdx() As Byte    'indicates which buffers are currently utilized, 0 = not used, 1 = used but no 3d, else = used and 3d
    Private mo3DBuffers() As Buffer3D   'corresponding 3D Buffer for the sound (if used)
    Private mlBufferVG() As Int32   'indicates what volumegroup this buffer subscribes to
    Private myBufferReserved() As Byte      'indicates which buffers are reserved (reserve indicates captured but not initialized)
    Private mvecBufferLoc() As Vector3
    Private mlBufferUB As Int32 = -1

    Private moListener As Listener3D
    Private moPrimary As Buffer

    Private moCleanupThread As Threading.Thread
    Private mbCleaning As Boolean
    Private mbHasGarbageCollector As Boolean
    Private mbAddingSound As Boolean

    Private moSrcBuffers() As Buffer
    Private moSrcDesc() As BufferDescription
    Private mlSrcBufferUB As Int32 = -1
    Private msSrcBufferNames() As String

    Public VolumeGroups() As Int32
	Public MasterVolume As Int32
	Private mfMasterVolumeMult As Single = 1.0F

    Private mbSoundsInitialized As Boolean = False
    Private mbVGInit As Boolean = False

    Private msSFXPath As String
    Private msMusicPath As String

    Public Shared lMinDistance As Int32 = 1000
    Public Shared lMaxDistance As Int32 = 15000

    Private Function UnpackLocalSoundFile(ByVal sPakFile As String, ByVal sFile As String) As Boolean
        Dim oMem As IO.MemoryStream = GetResourceStream(sFile, msSFXPath & sPakFile)
        If oMem Is Nothing = False Then
            Try
                Dim oFS As New IO.FileStream(msSFXPath & "tmp.wav", IO.FileMode.Create)
                oFS.Write(oMem.GetBuffer, 0, CInt(oMem.Length - 2))
                oFS.Close()
                oFS.Dispose()
                oFS = Nothing
            Catch
            End Try
            Return True
        Else : Return False
        End If
    End Function

    Public Sub InitializeSounds()
        Dim sINIFile As String = AppDomain.CurrentDomain.BaseDirectory
        Dim oINI As InitFile
        Dim X As Int32
        Dim sFile As String

        mlBufferUB = ml_MAX_BUFFS - 1
        ReDim moBuffers(mlBufferUB)
        ReDim myBufferIdx(mlBufferUB)
        ReDim myBufferReserved(mlBufferUB)
        ReDim mo3DBuffers(mlBufferUB)      'redim this even if we don't use it
        ReDim mlBufferVG(mlBufferUB)
        ReDim mvecBufferLoc(mlBufferUB)

        For X = 0 To mlBufferUB
            moBuffers(X) = Nothing
            myBufferIdx(X) = SoundUsage.eNotUsed
            myBufferReserved(X) = 0
            mo3DBuffers(X) = Nothing
            mlBufferVG(X) = -1
        Next X

        If Right$(sINIFile, 1) <> "\" Then sINIFile &= "\"
        sINIFile &= "Audio\Audio.Ini"

        oINI = New InitFile(sINIFile)

        'First, load our UB
        mlSrcBufferUB = CInt(Val(oINI.GetString("GLOBAL SETTINGS", "PreLoadSrcBufferCnt", "0"))) - 1

        If mlSrcBufferUB < 0 Then
            mlSrcBufferUB = -1
            mbSoundsInitialized = True      'set this to true anyway, but it just means all sfx are in slow mode
            Exit Sub
        End If

        ReDim moSrcBuffers(mlSrcBufferUB)
        ReDim moSrcDesc(mlSrcBufferUB)
        ReDim msSrcBufferNames(mlSrcBufferUB)

        'Now, load all of our files
        For X = 0 To mlSrcBufferUB
            sFile = oINI.GetString("SrcBuffer_" & X, "Name", "")
            If sFile <> "" Then
                moSrcDesc(X) = New BufferDescription()
                If Val(oINI.GetString("SrcBuffer_" & X, "Use3D", "1")) <> 0 Then
                    moSrcDesc(X).Flags = BufferDescriptionFlags.Control3D Or BufferDescriptionFlags.ControlVolume Or BufferDescriptionFlags.Mute3DAtMaximumDistance
                Else
                    moSrcDesc(X).Flags = BufferDescriptionFlags.ControlVolume
                End If
                Try
                    If Exists(msSFXPath & sFile) = True Then
                        'Now, create our buffer
                        Dim sPak As String = oINI.GetString("SrcBuffer_" & X, "Pak", "")
                        If sPak = "" Then
                            moSrcBuffers(X) = New Buffer(msSFXPath & sFile, moSrcDesc(X), moDevice)
                        Else
                            'Dim oStream As IO.MemoryStream = GetResourceStream(sFile, msSFXPath & sPak)
                            'If oStream Is Nothing = False Then
                            '    moSrcBuffers(X) = New Buffer(oStream, moSrcDesc(X), moDevice)
                            'Else
                            '    moSrcBuffers(X) = New Buffer(msSFXPath & sFile, moSrcDesc(X), moDevice)
                            'End If
                            If UnpackLocalSoundFile(sPak, sFile) = True Then
                                moSrcBuffers(X) = New Buffer(msSFXPath & "tmp.wav", moSrcDesc(X), moDevice)
                            End If
                        End If
                    Else
                        'Look inside a pak file? Added a second Exists check (accept.wav was missing and erroring)
                        Dim sPak As String = oINI.GetString("SrcBuffer_" & X, "Pak", "")
                        If sPak = "" Then
                            If Exists(msSFXPath & sFile) = True Then
                                moSrcBuffers(X) = New Buffer(msSFXPath & sFile, moSrcDesc(X), moDevice)
                            End If
                        Else
                            If UnpackLocalSoundFile(sPak, sFile) = True Then
                                moSrcBuffers(X) = New Buffer(msSFXPath & "tmp.wav", moSrcDesc(X), moDevice)
                            End If
                            'Dim oStream As IO.MemoryStream = GetResourceStream(sFile, msSFXPath & sPak)
                            'If oStream Is Nothing = False Then
                            '    'Create our buffer description object
                            '    moSrcDesc(X) = New BufferDescription()
                            '    If Val(oINI.GetString("SrcBuffer_" & X, "Use3D", "1")) <> 0 Then
                            '        moSrcDesc(X).Flags = BufferDescriptionFlags.Control3D Or BufferDescriptionFlags.ControlVolume Or BufferDescriptionFlags.Mute3DAtMaximumDistance
                            '    Else
                            '        moSrcDesc(X).Flags = BufferDescriptionFlags.ControlVolume
                            '    End If

                            '    moSrcBuffers(X) = New Buffer(oStream, moSrcDesc(X), moDevice)
                            'Else
                            '    moSrcBuffers(X) = New Buffer(msSFXPath & sFile, moSrcDesc(X), moDevice)
                            'End If
                        End If
                    End If
                Catch
                End Try
            End If
            msSrcBufferNames(X) = sFile
        Next X

        oINI = Nothing

        mbSoundsInitialized = True
    End Sub

    Private Function GetResourceStream(ByVal sName As String, ByVal sResourceFile As String) As IO.MemoryStream
        If Exists(sResourceFile) = False Then Return Nothing

        Dim oFS As IO.FileStream = Nothing
        Dim oReader As IO.BinaryReader

        Dim yData() As Byte
        Dim yTOC() As Byte
        Dim lTOCLen As Int32

        Dim X As Int32
        Dim yEntryName(19) As Byte
        Dim sFile As String
        Dim lPos As Int32

        Dim lEntryStart As Int32 = -1
        Dim lEntryLen As Int32 = -1

        'Open the resource file
        Dim lAttemptCntr As Int32 = 0
        While oFS Is Nothing AndAlso lAttemptCntr < 100
            Try
                oFS = New IO.FileStream(sResourceFile, IO.FileMode.Open)
            Catch ex As Exception
                'do nothing
            End Try
            lAttemptCntr += 1
            Threading.Thread.Sleep(10)
        End While
        If oFS Is Nothing Then Return Nothing

        'Create our reader
        oReader = New IO.BinaryReader(oFS)

        'Get the first 4 bytes which is the unencrypted length of our Table of Contents portion
        yData = oReader.ReadBytes(4)
        lTOCLen = System.BitConverter.ToInt32(yData, 0)

        'read in the TOC contents
        yTOC = oReader.ReadBytes(lTOCLen)

        'Decrypt the TOC
        yTOC = DecBytes(yTOC)

        lPos = 0
        'Find out entry...
        For X = 0 To CInt(Math.Floor(lTOCLen / 28) - 1)
            Array.Copy(yTOC, X * 28, yEntryName, 0, 20)
            lPos += 20
            sFile = BytesToString(yEntryName)
            If UCase$(sFile) = UCase$(sName) Then
                lEntryStart = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4
                lEntryLen = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4
                Exit For
            Else
                lPos += 8
            End If
        Next X

        If lEntryStart > -1 AndAlso lEntryLen > -1 Then
            'Ok, found our entry... let's get the data
            'oReader.ReadBytes(lEntryStart - 1)  'should be right...
            oReader.BaseStream.Seek(lEntryStart + lTOCLen + 4, IO.SeekOrigin.Begin)
            Dim yEntry() As Byte
            yData = oReader.ReadBytes(lEntryLen)
            'Now, decrypt it
            yData = DecBytes(yData)
            ReDim yEntry(lEntryLen - 1)
            Array.Copy(yData, 0, yEntry, 0, lEntryLen - 1)

            'Now, create our response
            Dim oMemStream As IO.MemoryStream = New IO.MemoryStream(yEntry.Length - 1)
            Dim oFSWriter As IO.BinaryWriter = New IO.BinaryWriter(oMemStream)
            oFSWriter.Write(yEntry)
            'Flush and Nothing, but do not close as it will close the base stream too
            oFSWriter.Flush()
            oFSWriter = Nothing

            'Now, seek the memory stream to the beginning
            oMemStream.Seek(0, IO.SeekOrigin.Begin)

            'Close everything else
            oReader.Close()
            oFS.Close()

            Return oMemStream
        Else
            oReader.Close()
            oFS.Close()
            Return Nothing
        End If

    End Function

    Private Function DecBytes(ByVal yBytes() As Byte) As Byte()
        'Now, we do the exact opposite...
        Dim lLen As Int32 = UBound(yBytes)
        Dim lKey As Int32
        Dim X As Int32
        Dim yFinal(lLen + 1) As Byte
        Dim lChrCode As Int32
        Dim lMod As Int32

        'Get our key value...
        lKey = yBytes(0)

        'set up our seed
        'Rnd(-1)
        'Call Randomize(ml_ENCRYPT_SEED + lKey)
        Dim oRandom As New Random(lKey)

        For X = 1 To lLen
            'Now, find out what we got here...
            lChrCode = yBytes(X)
            'now, subtract our value... 1 to 5
            'lMod = Int(Rnd() * 5) + 1
            lMod = oRandom.Next(1, 6)   '1 to 6 here because 6 is excluded
            lChrCode = lChrCode - lMod
            If lChrCode < 0 Then lChrCode = 256 + lChrCode
            yFinal(X - 1) = CByte(lChrCode)
        Next X
        Return yFinal
    End Function

    Private Function BytesToString(ByVal yBytes() As Byte) As String
        Dim lLen As Int32 = yBytes.Length
        Dim X As Int32

        For X = 0 To yBytes.Length - 1
            If yBytes(X) = 0 Then
                lLen = X
                Exit For
            End If
        Next X

        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yBytes), 1, lLen)
    End Function

#Region "  Unit Speech  "
    Public Enum SpeechType As Int32
        eStopRequest = 1
        eMoveRequest = 2
        eBuildingConstruction = 4
        eSetPrimaryTarget = 8
        eDockRequest = 16
        eUndockRequest = 32
        eBombardRequest = 64
        eStopBombard = 128
        eSetRallyPoint = 256
        eLaunchAll = 512

        eFleetSelect = 268435456
        eTwoSelect = 536870912
        eSingleSelect = 1073741824
    End Enum

    Public Function StartUnitSpeechSound(ByVal bAffirmative As Boolean, ByVal lSpeech As SpeechType) As Int32

        Dim lSelect As Int32 = lSpeech And (SpeechType.eFleetSelect Or SpeechType.eTwoSelect Or SpeechType.eSingleSelect)
        Dim lAction As Int32 = lSpeech - lSelect

        Dim sResult As String = ""
        Dim sOptions() As String
        Dim lUB As Int32 = -1
        Dim lBaseUB As Int32 = -1

        If bAffirmative = True Then
            lUB = 7
            ReDim sOptions(lUB)
            sOptions(0) = "Unit Speech\General_Yes_1_1.wav"
            sOptions(1) = "Unit Speech\General_Yes_1_2.wav"
            sOptions(2) = "Unit Speech\General_Yes_1_3.wav"
            sOptions(3) = "Unit Speech\General_Yes_1_4.wav"
            sOptions(4) = "Unit Speech\General_Yes_2_1.wav"
            sOptions(5) = "Unit Speech\General_Yes_2_2.wav"
            sOptions(6) = "Unit Speech\General_Yes_2_3.wav"
            sOptions(7) = "Unit Speech\General_Yes_2_4.wav"
        Else
            lUB = 3
            ReDim sOptions(lUB)
            sOptions(0) = "Unit Speech\General_No_1_1.wav"
            sOptions(1) = "Unit Speech\General_No_1_2.wav"
            sOptions(2) = "Unit Speech\General_No_2_1.wav"
            sOptions(3) = "Unit Speech\General_No_2_2.wav"
        End If
        lBaseUB = lUB

        Select Case lAction
            Case SpeechType.eBombardRequest
                If bAffirmative = True Then
                    lUB = 3
                    lBaseUB = 3
                    ReDim sOptions(lUB)
                    sOptions(0) = "Unit Speech\BombardRqst_1_1.wav"
                    sOptions(1) = "Unit Speech\BombardRqst_1_2.wav"
                    sOptions(2) = "Unit Speech\BombardRqst_2_1.wav"
                    sOptions(3) = "Unit Speech\BombardRqst_2_2.wav"
                End If
            Case SpeechType.eBuildingConstruction
                If bAffirmative = True Then
                    lUB += 2
                    ReDim Preserve sOptions(lUB)
                    sOptions(lUB - 1) = "Unit Speech\Eng_Build_1_1.wav"
                    sOptions(lUB) = "Unit Speech\Eng_Build_2_1.wav"
                    lBaseUB = lUB
                End If
            Case SpeechType.eDockRequest
                If bAffirmative = True Then
                    ReDim Preserve sOptions(lUB + 2)
                    lUB += 1
                    sOptions(lUB) = "Unit Speech\DockRequest_1_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\DockRequest_2_1.wav" ': lUB += 1

                    If lSelect = SpeechType.eFleetSelect Then
                        ReDim Preserve sOptions(lUB + 2)
                        lUB += 1
                        sOptions(lUB) = "Unit Speech\DockRequest_1_2.wav" : lUB += 1
                        sOptions(lUB) = "Unit Speech\DockRequest_2_2.wav" ': lUB += 1
                    End If
                End If
            Case SpeechType.eLaunchAll
                If bAffirmative = True Then
                    ReDim Preserve sOptions(lUB + 4)
                    lUB += 1
                    sOptions(lUB) = "Unit Speech\LaunchAll_1_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\LaunchAll_1_2.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\LaunchAll_2_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\LaunchAll_2_2.wav" ': lUB += 1
                End If
            Case SpeechType.eMoveRequest
                If bAffirmative = True Then
                    ReDim Preserve sOptions(lUB + 9)
                    lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Request_1_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Request_1_2.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Request_1_3.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Request_1_6.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Request_2_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Request_2_2.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Request_2_3.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Request_2_6.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Request_2_7.wav" ': lUB += 1

                    If lSelect = SpeechType.eFleetSelect Then
                        ReDim Preserve sOptions(lUB + 6)
                        lUB += 1
                        sOptions(lUB) = "Unit Speech\Move_Request_1_4.wav" : lUB += 1
                        sOptions(lUB) = "Unit Speech\Move_Request_1_5.wav" : lUB += 1
                        sOptions(lUB) = "Unit Speech\Move_Request_1_7.wav" : lUB += 1
                        sOptions(lUB) = "Unit Speech\Move_Request_2_4.wav" : lUB += 1
                        sOptions(lUB) = "Unit Speech\Move_Request_2_5.wav" : lUB += 1
                        sOptions(lUB) = "Unit Speech\Move_Request_2_8.wav" ': lUB += 1
                    End If
                Else
                    ReDim Preserve sOptions(lUB + 4)
                    lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Deny_1_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Deny_1_2.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Deny_2_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Move_Deny_2_2.wav" ': lUB += 1
                End If
            Case SpeechType.eSetPrimaryTarget
                ReDim Preserve sOptions(lUB + 8)
                lUB += 1
                sOptions(lUB) = "Unit Speech\SetTarget_1_1.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\SetTarget_1_2.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\SetTarget_1_3.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\SetTarget_1_4.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\SetTarget_2_1.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\SetTarget_2_2.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\SetTarget_2_3.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\SetTarget_2_4.wav" ': lUB += 1

                If lSelect = SpeechType.eFleetSelect Then
                    ReDim Preserve sOptions(lUB + 2)
                    lUB += 1
                    sOptions(lUB) = "Unit Speech\SetTarget_1_5.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\SetTarget_2_5.wav" ': lUB += 1
                End If
            Case SpeechType.eSetRallyPoint
                If bAffirmative = True Then
                    ReDim sOptions(1)
                    lUB = 0
                    sOptions(lUB) = "Unit Speech\SetRallyPoint_1_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\SetRallyPoint_2_1.wav" ': lUB += 1
                    lBaseUB = lUB
                End If
            Case SpeechType.eStopBombard
                ReDim sOptions(3)
                lUB = 0
                sOptions(lUB) = "Unit Speech\StopBombard_1_1.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\StopBombard_1_2.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\StopBombard_2_1.wav" : lUB += 1
                sOptions(lUB) = "Unit Speech\StopBombard_2_2.wav" ': lUB += 1
                lBaseUB = lUB
            Case SpeechType.eStopRequest
                If bAffirmative = True Then
                    ReDim Preserve sOptions(lUB + 2)
                    lUB += 1
                    sOptions(lUB) = "Unit Speech\Stop_Request_1_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\Stop_Request_2_1.wav" ': lUB += 1
                End If
            Case SpeechType.eUndockRequest
                If bAffirmative = True Then
                    ReDim Preserve sOptions(lUB + 6)
                    lUB += 1
                    sOptions(lUB) = "Unit Speech\UndockRequest_1_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\UndockRequest_1_2.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\UndockRequest_1_3.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\UndockRequest_2_1.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\UndockRequest_2_2.wav" : lUB += 1
                    sOptions(lUB) = "Unit Speech\UndockRequest_2_3.wav" ': lUB += 1
                End If
        End Select

        Dim lIdx As Int32
        Randomize()
        If lBaseUB <> lUB AndAlso Rnd() * 100 > 25 Then
            lIdx = CInt(Math.Floor(Rnd() * (lUB - lBaseUB)) + lBaseUB) + 1
        Else : lIdx = CInt(Math.Floor(Rnd() * lUB))
        End If
        If lIdx < 0 Then lIdx = 0
        If lIdx > lUB Then lIdx = lUB
        sResult = sOptions(lIdx)

        If sResult = "" Then Return -1
        Return StartSound(sResult, False, SoundUsage.eUnitSpeech, Nothing, Nothing)
    End Function

    Public Function GetUnitSpeechTypeCnt(ByVal lSpeech As SpeechType, ByVal lCnt As Int32) As SpeechType
        If lCnt = 1 Then
            lSpeech = lSpeech Or SpeechType.eSingleSelect
        ElseIf lCnt = 2 Then
            lSpeech = lSpeech Or SpeechType.eTwoSelect
        Else : lSpeech = lSpeech Or SpeechType.eFleetSelect
        End If
        Return lSpeech
    End Function
#End Region

    Public lSoundStarts As Int32 = 0
	Public Function StartSound(ByVal sName As String, ByVal bLooping As Boolean, ByVal ySoundUsage As SoundUsage, ByVal vecLoc As Vector3, ByVal vecVel As Vector3) As Int32

		If muSettings.AudioOn = False Then Return -1
        If mbVGInit = False Then InitializeVolumeGroups()

		Dim X As Int32
		Dim lIdx As Int32 = -1
		Dim lSoundID As Int32 = -1
		Dim lVolume As Int32

		'Find an unused buffer
		lIdx = GetBufferIndex(ySoundUsage, vecLoc)
        If lIdx = -1 Then Return -1

        lSoundStarts += 1
        myBufferReserved(lIdx) = 255
        mvecBufferLoc(lIdx) = vecLoc

        'Validate sound file exists
        If ySoundUsage <> SoundUsage.eAmbience AndAlso ySoundUsage <> SoundUsage.eWeather Then
            If CheckSoundFileExists(sName) = False Then
                myBufferReserved(lIdx) = 0
                Return -1
            End If
        End If

		If mbSoundsInitialized = False Then InitializeSounds()

        'Make sure the name is ucase and find our src buffer by name
        sName = UCase$(sName)
        For X = 0 To mlSrcBufferUB
            If msSrcBufferNames(X).ToUpper = sName Then
                lSoundID = X
                If moSrcBuffers(X) Is Nothing Then
                    myBufferReserved(lIdx) = 0
                    Return -1
                End If
                Exit For
            End If
        Next X

		While mbAddingSound = True
			Threading.Thread.Sleep(20)
        End While

        Try
            'Now, create our buffer
            mbAddingSound = True

            'If we did not find the src buffer, then attempt to load the file ourselves (the slower way)
            If lSoundID = -1 Then

                Dim oDesc As BufferDescription = New BufferDescription()
                With oDesc
                    If ySoundUsage > SoundUsage.e3D_SOUNDS Then
                        .Flags = BufferDescriptionFlags.Control3D Or BufferDescriptionFlags.ControlVolume Or BufferDescriptionFlags.Mute3DAtMaximumDistance
                    Else
                        .Flags = BufferDescriptionFlags.ControlVolume
                    End If
                End With

                If ySoundUsage = SoundUsage.eAmbience Then
                    Dim oStream As IO.MemoryStream = GetResourceStream(sName, msSFXPath & "Ambience\Ambience.pak")
                    If oStream Is Nothing Then
                        myBufferReserved(lIdx) = 0
                        mbAddingSound = False
                        Return -1
                    End If
                    moBuffers(lIdx) = New Buffer(oStream, oDesc, moDevice)
                ElseIf ySoundUsage = SoundUsage.eWeather Then
                    Dim oStream As IO.MemoryStream = GetResourceStream(sName, msSFXPath & "Weather\Weather.pak")
                    If oStream Is Nothing Then
                        myBufferReserved(lIdx) = 0
                        mbAddingSound = False
                        Return -1
                    End If
                    moBuffers(lIdx) = New Buffer(oStream, oDesc, moDevice)
                Else
                    'Check if the file exists before continuing, if it does not, then return -1
                    If Exists(msSFXPath & sName) = False Then
                        myBufferReserved(lIdx) = 0
                        mbAddingSound = False
                        Return -1
                    End If

                    moBuffers(lIdx) = New Buffer(msSFXPath & sName, oDesc, moDevice)
                End If

                oDesc.Dispose()
                oDesc = Nothing
            Else
                'Create the normal buffer
                moBuffers(lIdx) = moSrcBuffers(lSoundID).Clone(moDevice)
            End If

            Select Case ySoundUsage
                Case SoundUsage.eDeathExplosions, SoundUsage.eNonDeathExplosions, SoundUsage.eWeaponsFire, SoundUsage.eWeaponsFireBombs, SoundUsage.eWeaponsFireEnergy, SoundUsage.eWeaponsFireMissiles, SoundUsage.eWeaponsFireProjectile, SoundUsage.eFireworks
                    mlBufferVG(lIdx) = VolumeGroup.eVG_Fighting
                Case SoundUsage.eNarrative
                    mlBufferVG(lIdx) = VolumeGroup.eVG_GameNarrative
                Case SoundUsage.eTutorialVoice
                    mlBufferVG(lIdx) = VolumeGroup.eVG_TutorialVoice
                Case SoundUsage.eRadioChatter, SoundUsage.eUnitSounds
                    mlBufferVG(lIdx) = VolumeGroup.eVG_EntityAmbience
                Case SoundUsage.eUserInterface
                    mlBufferVG(lIdx) = VolumeGroup.eVG_UserInterface
                Case SoundUsage.eUnitSpeech
                    mlBufferVG(lIdx) = VolumeGroup.eVG_UnitSpeech
                Case SoundUsage.eAmbience, SoundUsage.eWeather
                    mlBufferVG(lIdx) = VolumeGroup.eVG_EnvirAmbience
                Case Else
                    mlBufferVG(lIdx) = -1
            End Select

            If mlBufferVG(lIdx) = -1 Then
                lVolume = 0     'always max volume
            Else : lVolume = VolumeGroups(mlBufferVG(lIdx))
            End If

            lVolume = CInt(((lVolume + 10000) * mfMasterVolumeMult) - 10000)

            'Now, create our 3D buffer, if necessary
            If ySoundUsage > SoundUsage.e3D_SOUNDS Then
                mo3DBuffers(lIdx) = New Buffer3D(moBuffers(lIdx))
                mo3DBuffers(lIdx).Position = vecLoc
                mo3DBuffers(lIdx).Velocity = vecVel

                If ySoundUsage = SoundUsage.eDeathExplosions OrElse ySoundUsage = SoundUsage.eFireworks Then
                    mo3DBuffers(lIdx).MinDistance = 5000
                    mo3DBuffers(lIdx).MaxDistance = 50000
                Else
                    mo3DBuffers(lIdx).MinDistance = lMinDistance
                    mo3DBuffers(lIdx).MaxDistance = lMaxDistance
                End If
            End If

            'Set the volume passed in
            moBuffers(lIdx).Volume = lVolume

            'If we're looping, then play it with loop on, otherwise, play with default
            If bLooping = True Then
                moBuffers(lIdx).Play(0, BufferPlayFlags.Looping)
            Else
                moBuffers(lIdx).Play(0, BufferPlayFlags.Default)
            End If
            moBuffers(lIdx).Volume = lVolume

            'Set our buffer usage
            myBufferIdx(lIdx) = ySoundUsage

        Catch
        Finally
            mbAddingSound = False
        End Try

		Return lIdx
	End Function

    Private Sub __GarbageCollector()
        Dim lBuff As Int32
        'Dim lTemp As Int32

        mbCleaning = True
        mbHasGarbageCollector = True

        Randomize()

        While mbCleaning
            Try
                If mbAddingSound = False Then
                    If MusicOn = True Then
                        UpdateMusic()
                    End If

                    'Now, check our buffers
                    For lBuff = 0 To mlBufferUB
                        If myBufferIdx(lBuff) <> 0 Then
                            If moBuffers(lBuff) Is Nothing = False Then
                                If moBuffers(lBuff).Status.Playing = False AndAlso moBuffers(lBuff).Status.Looping = False Then
                                    moBuffers(lBuff).Dispose()
                                    moBuffers(lBuff) = Nothing
                                    If mo3DBuffers(lBuff) Is Nothing = False Then
                                        mo3DBuffers(lBuff).Dispose()
                                        mo3DBuffers(lBuff) = Nothing
                                    End If
                                    myBufferIdx(lBuff) = 0
                                    myBufferReserved(lBuff) = 0
                                End If
                            End If
                        End If
                    Next lBuff

                End If
            Catch
            End Try

            'Sleep for everyone else
            'Application.DoEvents()
            Threading.Thread.Sleep(5)
        End While

        mbHasGarbageCollector = False
    End Sub

    Public Function IsSoundStopped(ByVal lIndex As Int32) As Boolean
        Try
            If lIndex <= mlBufferUB AndAlso lIndex > -1 AndAlso myBufferIdx(lIndex) <> 0 AndAlso moBuffers(lIndex) Is Nothing = False Then Return Not moBuffers(lIndex).Status.Playing
        Catch
            'do nothing, return true will result in sound stopping
        End Try
        Return True
    End Function

    Public Sub UpdateSoundLoc(ByVal lIndex As Int32, ByVal vecNewLoc As Vector3, ByVal vecNewVel As Vector3)
        If mo3DBuffers(lIndex) Is Nothing = False Then
            mo3DBuffers(lIndex).Position = vecNewLoc
            mo3DBuffers(lIndex).Velocity = vecNewVel
        End If
    End Sub

    Public Sub UpdateListenerLoc(ByVal vecNewLoc As Vector3, ByVal vecFront As Vector3, ByVal vecTop As Vector3)
        Try
            moListener.Position = vecNewLoc
            moListener.Orientation = New Listener3DOrientation(vecFront, vecTop)
        Catch

        End Try
    End Sub

    Public Sub SetSoundVolume(ByVal lIndex As Int32, ByVal lNewVol As Int32)
        Try
			If lIndex <= mlBufferUB AndAlso myBufferIdx(lIndex) <> 0 AndAlso moBuffers(lIndex) Is Nothing = False Then
                moBuffers(lIndex).Volume = lNewVol
            End If
        Catch
        End Try
    End Sub

    Public Sub New(ByVal frmOwner As Form)
        Try
            Dim uDevInfo As DeviceInformation
            Dim oCol As DevicesCollection = New DevicesCollection()

            Dim sDesc As String = "Primary Sound Driver"
            Dim uGuid As System.Guid
            Dim oINI As InitFile = New InitFile()

            'Dim sTemp As String

            'sTemp = oINI.GetString("AUDIO", "DEVICE", "")
            'If sTemp <> "" Then
            '    'Ok, try to get it
            '    Try
            '        'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
            '        uGuid = New Guid(sTemp)
            '        moDevice = New Device(uGuid)
            '    Catch
            '        moDevice = Nothing
            '    End Try
            'End If

            'Ok, just find the Primary Sound Driver
            If moDevice Is Nothing Then
                For Each uDevInfo In oCol
                    If uDevInfo.Description = sDesc Then
                        uGuid = uDevInfo.DriverGuid
                        Exit For
                    ElseIf uGuid.ToString = "00000000-0000-0000-0000-000000000000" Then
                        uGuid = uDevInfo.DriverGuid
                    End If
                Next uDevInfo
                Try
                    moDevice = New Device(uGuid)
                Catch
                    muSettings.AudioOn = False

                End Try
            End If

            'save it
            oINI.WriteString("AUDIO", "DEVICE", uGuid.ToString())

            'Set our cooperative level
            moDevice.SetCooperativeLevel(frmOwner.Handle, CooperativeLevel.Priority)

            MusicOn = (Val(oINI.GetString("AUDIO", "MusicOn", "1")) <> 0)
            If MusicOn = True Then InitializeMusicPlayer()

            'Now, set our path information
            Dim sPath As String
            sPath = AppDomain.CurrentDomain.BaseDirectory
            If Right(sPath, 1) <> "\" Then sPath &= "\"
            sPath &= "Audio\"
            msSFXPath = sPath & "SoundFX\"
            msMusicPath = sPath & "Music\"

            'Set up our volume groups
            InitializeVolumeGroups()

            'Create our Primary Buffer and Listener object
            Dim oDesc As BufferDescription = New BufferDescription()
            With oDesc
                .PrimaryBuffer = True
                .Control3D = True
                .ControlVolume = True
            End With
            moPrimary = New Buffer(oDesc, moDevice)
            moListener = New Listener3D(moPrimary)

            InitializeSounds()

            'Start our garbage collector thread
            moCleanupThread = New Threading.Thread(AddressOf __GarbageCollector)
            moCleanupThread.Start()
        Catch
            musettings.AudioOn = False
        End Try
    End Sub

    Private Sub InitializeVolumeGroups()
        Dim oINI As InitFile = New InitFile()
        Dim X As Int32
        Dim lTemp As Int32

        ReDim VolumeGroups(VolumeGroup.eVG_Count - 1)

        lTemp = 100
        lTemp = CInt(Val(oINI.GetString("AUDIO", "MasterVolume", "100")))
		MasterVolume = lTemp 'LinearVolumeToDirectX(lTemp, 100)
		mfMasterVolumeMult = (LinearVolumeToDirectX(MasterVolume, 100) + 10000) / 10000.0F

        'Let's set up our defaults in case there are no values atall
        VolumeGroups(VolumeGroup.eVG_EntityAmbience) = 70
        VolumeGroups(VolumeGroup.eVG_EntityResponse) = 90
        VolumeGroups(VolumeGroup.eVG_EnvirAmbience) = 70
        VolumeGroups(VolumeGroup.eVG_Fighting) = 100
        VolumeGroups(VolumeGroup.eVG_GameNarrative) = 100
        VolumeGroups(VolumeGroup.eVG_TutorialVoice) = 100
        VolumeGroups(VolumeGroup.eVG_Music) = 50
        VolumeGroups(VolumeGroup.eVG_UnitSpeech) = 80
        VolumeGroups(VolumeGroup.eVG_UserInterface) = 80

        For X = 0 To VolumeGroup.eVG_Count - 1
            lTemp = CInt(Val(oINI.GetString("AUDIO", "VG" & X, VolumeGroups(X).ToString)))
            VolumeGroups(X) = LinearVolumeToDirectX(lTemp, 100)
        Next X

        oINI = Nothing
        mbVGInit = True
    End Sub

    Private Sub SaveVolumeGroups()
        Dim oINI As InitFile = New InitFile()
        Dim X As Int32
        Dim lTemp As Int32

        For X = 0 To VolumeGroup.eVG_Count - 1
            lTemp = DirectXVolumeToLinear(VolumeGroups(X), 100)
            oINI.WriteString("AUDIO", "VG" & X, CStr(lTemp))
        Next X

        oINI = Nothing
    End Sub

    Public Sub StopSound(ByVal lIndex As Int32)
		If lIndex <= mlBufferUB AndAlso myBufferIdx(lIndex) <> 0 Then
			Dim oBuffer As Buffer = moBuffers(lIndex)
			If oBuffer Is Nothing = False AndAlso oBuffer.Disposed = False Then oBuffer.Stop()
		End If
    End Sub

    Public Sub KillAllSounds()
        On Error Resume Next
        For X As Int32 = 0 To mlBufferUB
            If myBufferIdx(X) <> 0 Then
                moBuffers(X).Stop()
            End If
        Next X
    End Sub

    Public Sub DisposeMe()
        Dim X As Int32
        Dim lCnt As Int32 = 0

        mbCleaning = False
        While mbHasGarbageCollector = True
            Threading.Thread.Sleep(0)
            lCnt += 1
            If lCnt > 1000 Then
                'Ok, long enough... kill it
                moCleanupThread.Abort()
                mbHasGarbageCollector = False
            End If
        End While

        'Ok destroy our garbage collector thread
        moCleanupThread = Nothing

        'Now, go thru and release all of our buffers
        For X = 0 To mlBufferUB
            Try
                If myBufferIdx(X) <> 0 Then
                    If mo3DBuffers(X) Is Nothing = False Then
                        mo3DBuffers(X).Dispose()
                    End If
                    If moBuffers(X) Is Nothing = False Then
                        moBuffers(X).Stop()
                        moBuffers(X).Dispose()
                    End If
                End If
            Catch
            End Try
        Next X

        'Finally, remove our source buffers
        For X = 0 To mlSrcBufferUB
            Try
                If moSrcBuffers(X) Is Nothing = False Then moSrcBuffers(X).Dispose()
                If moSrcDesc(X) Is Nothing = False Then moSrcDesc(X).Dispose()
            Catch
            End Try
        Next X

        moListener.Dispose()
        moPrimary.Dispose()

        If moMP Is Nothing = False Then
            StopMusic()
        End If
        moMP = Nothing

        'Save our volume groups
        SaveVolumeGroups()

        Erase mo3DBuffers
        Erase moBuffers
        Erase myBufferIdx
        Erase mvecBufferLoc
        Erase myBufferReserved
        Erase moSrcBuffers
        Erase moSrcDesc
        Erase msSrcBufferNames

        moListener = Nothing
        moPrimary = Nothing

        If moDevice Is Nothing = False Then moDevice.Dispose()
        moDevice = Nothing
        moCleanupThread = Nothing
    End Sub

    Protected Overrides Sub Finalize()
        If moDevice Is Nothing = False Then moDevice.Dispose()
        moDevice = Nothing

        MyBase.Finalize()
    End Sub

    Public Shared Function LinearVolumeToDirectX(ByVal lLinear As Int32, ByVal lMax As Int32) As Int32
        If lLinear <= 0 Then
            Return -10000
        ElseIf lLinear > lMax Then
            Return 0
        Else : Return CInt(Math.Floor(2000 * Math.Log10(lLinear / lMax) + 0.5F))
        End If
    End Function

    Public Shared Function DirectXVolumeToLinear(ByVal lDirectX As Int32, ByVal lMaxLinear As Int32) As Int32
        If lDirectX <= -10000 Then
            Return 0
        Else
            Return CInt(Math.Floor(10 ^ (lDirectX / 2000) * lMaxLinear))
        End If
    End Function

    Private Function CheckSoundFileExists(ByVal sWavName As String) As Boolean
        'Dim sPath As String
        'sPath = AppDomain.CurrentDomain.BaseDirectory
        'If sPath.EndsWith("\") = False Then sPath &= "\"
        'sPath &= "\Audio\SoundFX\"

        'Return Exists(sPath & sWavName)
        If Exists(msSFXPath & sWavName) = False Then
            sWavName = sWavName.ToUpper
            For X As Int32 = 0 To mlSrcBufferUB
                If msSrcBufferNames(X).ToUpper = sWavName Then
                    Return True
                End If
            Next X
        Else
            Return True
        End If
        Return False
    End Function

    Public Sub UpdateSoundVolumes()
		'moPrimary.Volume = MasterVolume
		mfMasterVolumeMult = (LinearVolumeToDirectX(MasterVolume, 100) + 10000) / 10000.0F

        'called when a volume group changes...
        For X As Int32 = 0 To Me.mlBufferUB
            If myBufferIdx(X) <> 0 AndAlso mlBufferVG(X) > -1 AndAlso mlBufferVG(X) < VolumeGroups.Length Then
				'Ok... set this buffers' volume
				Dim lFinalVGVal As Int32 = CInt(((VolumeGroups(mlBufferVG(X)) + 10000) * mfMasterVolumeMult) - 10000)
				SetSoundVolume(X, lFinalVGVal)
            End If
        Next X

        If moMP Is Nothing = False Then moMP.CurrentBufferLinearVolume = DirectXVolumeToLinear(VolumeGroups(VolumeGroup.eVG_Music), 100)
    End Sub

    Private Function GetBufferIndex(ByVal yVal As SoundUsage, ByVal vecLoc As Vector3) As Int32
        'if yVal = eALWAYS_PLAY, find first available and return it
        'if yVal = Ambience, return 0
        'if yVal = Weather, return 1
        'if yVal = Narrative, return 2 or 3
        'if yVal = RadioChatter, return 4 or 5, however, if we are in a space envir, return 1
        'if yVal = UnitSounds, return 6 or 7
        'if yVal = UnitSpeech, return 8 or 9
        'if yVal = UserInterface, return 10

        'if yVal = NonDeathExplosions, return 11 thru 15, however if no result and we are in a space envir, return 0
        'if yVal = DeathExplosion, return 16, 17 or 18
        'if yVal = WeaponsFire, return 19 - 27, however...
        '   if that fails
        '       if yVal = WeaponsFire.Energy then return 28
        '       if yval = WeaponsFire.Projectile then return 29
        '       if yVal = WeaponsFire.Missiles then return 30
        '       if yVal = WeaponsFire.Bombs then return 31

        Dim lCX As Int32 = 0
        Dim lCY As Int32 = 0
        Dim lCZ As Int32 = 0
        If muSettings.PositionalSound = True AndAlso goCamera Is Nothing = False Then
            With goCamera
                lCX = .mlCameraX
                lCY = .mlCameraY
                lCZ = .mlCameraZ
            End With
        End If

        If myBufferReserved Is Nothing Then ReDim myBufferReserved(-1)

        SyncLock myBufferReserved
            Select Case yVal
                Case SoundUsage.eALWAYS_PLAY
                    For X As Int32 = 0 To Me.mlBufferUB
                        If myBufferReserved(X) = 0 Then Return X
                    Next X
                Case SoundUsage.eAmbience
                    If myBufferReserved(0) = 0 Then Return 0
                Case SoundUsage.eWeather
                    If myBufferReserved(1) = 0 Then Return 1
                Case SoundUsage.eNarrative
                    If myBufferReserved(2) = 0 Then
                        Return 2
                    ElseIf myBufferReserved(3) = 0 Then
                        Return 3
                    End If
                Case SoundUsage.eRadioChatter
                    If myBufferReserved(4) = 0 Then
                        myBufferReserved(4) = 255 
                        Return 4
                    ElseIf myBufferReserved(5) = 0 Then
                        myBufferReserved(5) = 255 
                        Return 5
                    ElseIf goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                        If myBufferReserved(1) = 0 Then
                            myBufferReserved(1) = 255
                            Return 1
                        End If
                    End If

                    If muSettings.PositionalSound = True Then
                        'Ok, we are here... recheck 4 and 5
                        Try
                            Dim lIdx As Int32 = -1
                            Dim lNewDist As Int32 = CInt(Math.Abs(vecLoc.X - lCX) + Math.Abs(vecLoc.Y - lCY) + Math.Abs(vecLoc.Z - lCZ))
                            Dim lDist4 As Int32 = CInt(Math.Abs(mvecBufferLoc(4).X - lCX) + Math.Abs(mvecBufferLoc(4).Y - lCY) + Math.Abs(mvecBufferLoc(4).Z - lCZ))
                            Dim lDist5 As Int32 = CInt(Math.Abs(mvecBufferLoc(5).X - lCX) + Math.Abs(mvecBufferLoc(5).Y - lCY) + Math.Abs(mvecBufferLoc(5).Z - lCZ))
                            If lDist4 < lDist5 Then
                                If lNewDist < lDist5 Then lIdx = 5
                            Else
                                If lNewDist < lDist4 Then lIdx = 4
                            End If

                            If lIdx <> -1 Then
                                moBuffers(lIdx).Dispose()
                                moBuffers(lIdx) = Nothing
                                If mo3DBuffers(lIdx) Is Nothing = False Then
                                    mo3DBuffers(lIdx).Dispose()
                                    mo3DBuffers(lIdx) = Nothing
                                End If
                                myBufferIdx(lIdx) = 0
                                myBufferReserved(lIdx) = 255
                                Return lIdx
                            End If
                        Catch
                        End Try
                    End If


                Case SoundUsage.eUnitSounds
                    If myBufferReserved(6) = 0 Then
                        Return 6
                    ElseIf myBufferReserved(7) = 0 Then
                        Return 7
                    End If
                    
                Case SoundUsage.eUnitSpeech
                    If myBufferReserved(8) = 0 Then
                        Return 8
                    ElseIf myBufferReserved(9) = 0 Then
                        Return 9
                    End If 
                Case SoundUsage.eUserInterface
                    If myBufferReserved(10) = 0 Then Return 10

                Case SoundUsage.eNonDeathExplosions, SoundUsage.eFireworks
                    For X As Int32 = 11 To 15
                        If myBufferReserved(X) = 0 Then Return X
                    Next X
                    If myBufferReserved(0) = 0 AndAlso goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then Return 0

                    'here, so try the distance
                    If muSettings.PositionalSound = True Then
                        Try
                            Dim lIdx As Int32 = -1
                            Dim lNewDist As Int32 = CInt(Math.Abs(vecLoc.X - lCX) + Math.Abs(vecLoc.Y - lCY) + Math.Abs(vecLoc.Z - lCZ))
                            Dim lMaxDist As Int32 = 0
                            Dim lMaxDistIdx As Int32 = -1
                            For X As Int32 = 11 To 15
                                Dim lTmpDist As Int32 = CInt(Math.Abs(mvecBufferLoc(X).X - lCX) + Math.Abs(mvecBufferLoc(X).Y - lCY) + Math.Abs(mvecBufferLoc(X).Z - lCZ))
                                If lTmpDist > lMaxDist Then
                                    lMaxDist = lTmpDist
                                    lMaxDistIdx = X
                                End If
                            Next X
                            If lNewDist < lMaxDist AndAlso lMaxDistIdx <> -1 Then
                                lIdx = lMaxDistIdx
                                If lIdx <> -1 Then
                                    moBuffers(lIdx).Dispose()
                                    moBuffers(lIdx) = Nothing
                                    If mo3DBuffers(lIdx) Is Nothing = False Then
                                        mo3DBuffers(lIdx).Dispose()
                                        mo3DBuffers(lIdx) = Nothing
                                    End If
                                    myBufferIdx(lIdx) = 0
                                    myBufferReserved(lIdx) = 255
                                    Return lIdx
                                End If
                            End If
                        Catch
                        End Try
                    End If

                Case SoundUsage.eDeathExplosions
                    For X As Int32 = 16 To 18
                        If myBufferReserved(X) = 0 Then Return X
                    Next X

                    'here, so try the distance
                    If muSettings.PositionalSound = True Then
                        Try
                            Dim lNewDist As Int32 = CInt(Math.Abs(vecLoc.X - lCX) + Math.Abs(vecLoc.Y - lCY) + Math.Abs(vecLoc.Z - lCZ))
                            Dim lMaxDist As Int32 = 0
                            Dim lMaxDistIdx As Int32 = -1
                            For X As Int32 = 16 To 18
                                Dim lTmpDist As Int32 = CInt(Math.Abs(mvecBufferLoc(X).X - lCX) + Math.Abs(mvecBufferLoc(X).Y - lCY) + Math.Abs(mvecBufferLoc(X).Z - lCZ))
                                If lTmpDist > lMaxDist Then
                                    lMaxDist = lTmpDist
                                    lMaxDistIdx = X
                                End If
                            Next X
                            If lNewDist < lMaxDist AndAlso lMaxDistIdx <> -1 Then
                                moBuffers(lMaxDistIdx).Dispose()
                                moBuffers(lMaxDistIdx) = Nothing
                                If mo3DBuffers(lMaxDistIdx) Is Nothing = False Then
                                    mo3DBuffers(lMaxDistIdx).Dispose()
                                    mo3DBuffers(lMaxDistIdx) = Nothing
                                End If
                                myBufferIdx(lMaxDistIdx) = 0
                                myBufferReserved(lMaxDistIdx) = 255
                                Return lMaxDistIdx
                            End If
                        Catch
                        End Try
                    End If


                Case Else
                    For X As Int32 = 19 To 27
                        If myBufferReserved(X) = 0 Then
                            myBufferReserved(X) = 255
                            Return X
                        End If
                    Next X
                    Select Case yVal
                        Case SoundUsage.eWeaponsFireEnergy
                            If myBufferReserved(28) = 0 Then Return 28
                        Case SoundUsage.eWeaponsFireProjectile
                            If myBufferReserved(29) = 0 Then Return 29
                        Case SoundUsage.eWeaponsFireMissiles
                            If myBufferReserved(30) = 0 Then Return 30
                        Case SoundUsage.eWeaponsFireBombs
                            If myBufferReserved(31) = 0 Then Return 31
                    End Select

                    'here, so try the distance
                    If muSettings.PositionalSound = True Then
                        Try
                            Dim lNewDist As Int32 = CInt(Math.Abs(vecLoc.X - lCX) + Math.Abs(vecLoc.Y - lCY) + Math.Abs(vecLoc.Z - lCZ))
                            Dim lMaxDist As Int32 = 0
                            Dim lMaxDistIdx As Int32 = -1
                            For X As Int32 = 19 To 27
                                Dim lTmpDist As Int32 = CInt(Math.Abs(mvecBufferLoc(X).X - lCX) + Math.Abs(mvecBufferLoc(X).Y - lCY) + Math.Abs(mvecBufferLoc(X).Z - lCZ))
                                If lTmpDist > lMaxDist Then
                                    lMaxDist = lTmpDist
                                    lMaxDistIdx = X
                                End If
                            Next X
                            If lNewDist < lMaxDist AndAlso lMaxDistIdx <> -1 Then
                                moBuffers(lMaxDistIdx).Dispose()
                                moBuffers(lMaxDistIdx) = Nothing
                                If mo3DBuffers(lMaxDistIdx) Is Nothing = False Then
                                    mo3DBuffers(lMaxDistIdx).Dispose()
                                    mo3DBuffers(lMaxDistIdx) = Nothing
                                End If
                                myBufferIdx(lMaxDistIdx) = 0
                                myBufferReserved(lMaxDistIdx) = 255
                                Return lMaxDistIdx
                            End If

                            'ok, we're here, check the default slots
                            Select Case yVal
                                Case SoundUsage.eWeaponsFireEnergy
                                    lMaxDistIdx = 28
                                    lMaxDist = CInt(Math.Abs(mvecBufferLoc(lMaxDistIdx).X - lCX) + Math.Abs(mvecBufferLoc(lMaxDistIdx).Y - lCY) + Math.Abs(mvecBufferLoc(lMaxDistIdx).Z - lCZ))
                                    If lMaxDist > lNewDist Then
                                        moBuffers(lMaxDistIdx).Dispose()
                                        moBuffers(lMaxDistIdx) = Nothing
                                        If mo3DBuffers(lMaxDistIdx) Is Nothing = False Then
                                            mo3DBuffers(lMaxDistIdx).Dispose()
                                            mo3DBuffers(lMaxDistIdx) = Nothing
                                        End If
                                        myBufferIdx(lMaxDistIdx) = 0
                                        myBufferReserved(lMaxDistIdx) = 255
                                        Return lMaxDistIdx
                                    End If
                                Case SoundUsage.eWeaponsFireProjectile
                                    lMaxDistIdx = 29
                                    lMaxDist = CInt(Math.Abs(mvecBufferLoc(lMaxDistIdx).X - lCX) + Math.Abs(mvecBufferLoc(lMaxDistIdx).Y - lCY) + Math.Abs(mvecBufferLoc(lMaxDistIdx).Z - lCZ))
                                    If lMaxDist > lNewDist Then
                                        moBuffers(lMaxDistIdx).Dispose()
                                        moBuffers(lMaxDistIdx) = Nothing
                                        If mo3DBuffers(lMaxDistIdx) Is Nothing = False Then
                                            mo3DBuffers(lMaxDistIdx).Dispose()
                                            mo3DBuffers(lMaxDistIdx) = Nothing
                                        End If
                                        myBufferIdx(lMaxDistIdx) = 0
                                        myBufferReserved(lMaxDistIdx) = 255
                                        Return lMaxDistIdx
                                    End If
                                Case SoundUsage.eWeaponsFireMissiles
                                    lMaxDistIdx = 30
                                    lMaxDist = CInt(Math.Abs(mvecBufferLoc(lMaxDistIdx).X - lCX) + Math.Abs(mvecBufferLoc(lMaxDistIdx).Y - lCY) + Math.Abs(mvecBufferLoc(lMaxDistIdx).Z - lCZ))
                                    If lMaxDist > lNewDist Then
                                        moBuffers(lMaxDistIdx).Dispose()
                                        moBuffers(lMaxDistIdx) = Nothing
                                        If mo3DBuffers(lMaxDistIdx) Is Nothing = False Then
                                            mo3DBuffers(lMaxDistIdx).Dispose()
                                            mo3DBuffers(lMaxDistIdx) = Nothing
                                        End If
                                        myBufferIdx(lMaxDistIdx) = 0
                                        myBufferReserved(lMaxDistIdx) = 255
                                        Return lMaxDistIdx
                                    End If
                                Case SoundUsage.eWeaponsFireBombs
                                    lMaxDistIdx = 31
                                    lMaxDist = CInt(Math.Abs(mvecBufferLoc(lMaxDistIdx).X - lCX) + Math.Abs(mvecBufferLoc(lMaxDistIdx).Y - lCY) + Math.Abs(mvecBufferLoc(lMaxDistIdx).Z - lCZ))
                                    If lMaxDist > lNewDist Then
                                        moBuffers(lMaxDistIdx).Dispose()
                                        moBuffers(lMaxDistIdx) = Nothing
                                        If mo3DBuffers(lMaxDistIdx) Is Nothing = False Then
                                            mo3DBuffers(lMaxDistIdx).Dispose()
                                            mo3DBuffers(lMaxDistIdx) = Nothing
                                        End If
                                        myBufferIdx(lMaxDistIdx) = 0
                                        myBufferReserved(lMaxDistIdx) = 255
                                        Return lMaxDistIdx
                                    End If
                            End Select
                        Catch
                        End Try
                    End If
                    

            End Select

            'if we are here, return -1
            Return -1
        End SyncLock
    End Function

    Public Function HandleMovingSound(ByVal lIndex As Int32, ByVal sSound As String, ByVal yUsage As SoundUsage, ByVal vecLoc As Vector3, ByVal vecVel As Vector3) As Int32
        Try
            If lIndex = -1 Then
                Return StartSound(sSound, True, yUsage, vecLoc, vecVel)
            ElseIf mo3DBuffers(lIndex) Is Nothing = False Then
                mo3DBuffers(lIndex).Position = vecLoc
                mo3DBuffers(lIndex).Velocity = vecVel
            Else : Return -1
            End If
        Catch
        End Try
        Return lIndex
    End Function

#Region "  Unit Sound Management  "
    Private Const ml_UNIT_SOUND_MAX As Int32 = 2
    Private mfIndexDist() As Single
    Private mlIndex() As Int32
    Private mlPreviousIndex() As Int32

    'we store them as temp variables for performance
    Private mfCameraX As Single
    Private mfCameraY As Single
    Private mfCameraZ As Single

    Public Sub PrepareUnitSoundCollector()
        ReDim mfIndexDist(ml_UNIT_SOUND_MAX - 1)
        ReDim mlIndex(ml_UNIT_SOUND_MAX - 1)
        For X As Int32 = 0 To ml_UNIT_SOUND_MAX - 1
            mfIndexDist(X) = Single.MaxValue
            mlIndex(X) = -1
        Next X

        mfCameraX = goCamera.mlCameraX
        mfCameraY = goCamera.mlCameraY
        mfCameraZ = goCamera.mlCameraZ
    End Sub

    Public Sub TestEntitySound(ByVal lIndex As Int32, ByVal fLocX As Single, ByVal fLocY As Single, ByVal fLocZ As Single)
        Dim fDistX As Single = Math.Abs(mfCameraX - fLocX)
        Dim fDistY As Single = Math.Abs(mfCameraY - fLocY)
        Dim fDistZ As Single = Math.Abs(mfCameraZ - fLocZ)
        fDistX *= fDistX
        fDistY *= fDistY
        fDistZ *= fDistZ
        Dim fDist As Single = fDistX + fDistY + fDistZ

        For X As Int32 = 0 To ml_UNIT_SOUND_MAX - 1
            If mfIndexDist(X) > fDist Then
                For Y As Int32 = ml_UNIT_SOUND_MAX - 1 To X + 1 Step -1
                    mfIndexDist(Y) = mfIndexDist(Y - 1)
                    mlIndex(Y) = mlIndex(Y - 1)
                Next Y
                mfIndexDist(X) = fDist
                mlIndex(X) = lIndex
                Exit For
            End If
        Next X
    End Sub

    Private Sub KillAllUnitSounds()
        Try
            Me.StopSound(6)
            myBufferReserved(6) = 0
        Catch
        End Try
        Try
            Me.StopSound(7)
            myBufferReserved(7) = 0
        Catch
        End Try
    End Sub

    Public Sub CommitEntitySoundChanges()
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return

        'If True = True Then Return

        If mlPreviousIndex Is Nothing = True Then
            ReDim mlPreviousIndex(ml_UNIT_SOUND_MAX - 1)
            For X As Int32 = 0 To ml_UNIT_SOUND_MAX - 1
                mlPreviousIndex(X) = -1
            Next X
        End If

        'Now, go thru and kill an unit sounds that are currently going on...
        Try
            For X As Int32 = 0 To mlPreviousIndex.GetUpperBound(0)
                If mlPreviousIndex(X) > -1 Then
                    Dim lIdx As Int32 = mlPreviousIndex(X)
                    Dim bFound As Boolean = False
                    For Y As Int32 = 0 To mlIndex.GetUpperBound(0)
                        If mlIndex(Y) = mlPreviousIndex(X) Then
                            bFound = True
                            Exit For
                        End If
                    Next Y
                    If bFound = False Then
                        'now, see if that unit is there
                        Dim bUnitFound As Boolean = False
                        If oEnvir.lEntityUB >= lIdx Then
                            If oEnvir.lEntityIdx(lIdx) > -1 Then
                                Dim oEntity As BaseEntity = oEnvir.oEntity(lIdx)
                                If oEntity Is Nothing = False Then
                                    bUnitFound = True
                                    If oEntity.lUnitSoundIdx <> -1 Then Me.StopSound(oEntity.lUnitSoundIdx)
                                    oEntity.lUnitSoundIdx = -1
                                End If
                            End If
                        End If

                        If bUnitFound = False Then
                            KillAllUnitSounds()
                            Exit For
                        End If
                    End If

                End If
            Next X
        Catch
        End Try

        Try
            For X As Int32 = 0 To ml_UNIT_SOUND_MAX - 1
                mlPreviousIndex(X) = mlIndex(X)
                Dim lEntityIdx As Int32 = mlIndex(X)
                If oEnvir.lEntityUB >= lEntityIdx AndAlso lEntityIdx > -1 Then
                    If oEnvir.lEntityIdx(lEntityIdx) > -1 Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(lEntityIdx)
                        If oEntity Is Nothing = False Then
                            With oEntity
                                Dim vecTemp As Vector3 = New Microsoft.DirectX.Vector3(.VelX, 0, .VelZ)
                                vecTemp.Normalize()
                                vecTemp.Multiply(.TotalVelocity)

                                If oEntity.lUnitSoundIdx = -1 Then
                                    'start it
                                    .lUnitSoundIdx = StartSound(.oMesh.sRoarSFX, True, SoundUsage.eUnitSounds, New Vector3(.LocX, .LocY, .LocZ), vecTemp)
                                Else
                                    'update it
                                    UpdateSoundLoc(.lUnitSoundIdx, New Microsoft.DirectX.Vector3(.LocX, 0, .LocZ), vecTemp)
                                End If
                            End With
                        End If
                    End If
                End If

            Next X
        Catch
        End Try
    End Sub

#End Region

End Class