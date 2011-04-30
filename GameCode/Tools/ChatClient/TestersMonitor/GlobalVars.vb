Option Strict On

Public Module GlobalVars
    Public glActualPlayerID As Int32
    Public glPlayerID As Int32
    Public gbAliased As Boolean = False
    Public glAliasRights As Int32 = 0

    Public gfrmChat As frmMain
    Public gfrmChannels As frmChannels
    Public gfrmChannelConfig As frmChannelConfig
    Public gfrmColors As frmColors

    Public gsUserName As String
    Public gsPassword As String
    Public gyPlayerTitle As Byte = 0

    Public DefaultChatColor As System.Drawing.Color
    Public AlertChatColor As System.Drawing.Color
    Public StatusChatColor As System.Drawing.Color
    Public LocalChatColor As System.Drawing.Color
    Public GuildChatColor As System.Drawing.Color
    Public PMChatColor As System.Drawing.Color
    Public SenateChatColor As System.Drawing.Color
    Public ChannelChatColor As System.Drawing.Color
    Public AliasChatColor As System.Drawing.Color

    Public TextBoxBackColor As System.Drawing.Color

    Public dtLastMsg As Date
    Public glCredits As Int64
    Public glCashFlow As Int64
    'Public glWarpoints As Int64
    'Public glCurrentWPUpkeepCost As Int64

    Public Function GetStringFromBytes(ByRef yData() As Byte, ByVal lPos As Int32, ByVal lDataLen As Int32) As String
        Dim yTemp(lDataLen - 1) As Byte
        Array.Copy(yData, lPos, yTemp, 0, lDataLen)
        Dim lLen As Int32 = yTemp.Length
        For Y As Int32 = 0 To yTemp.Length - 1
            If yTemp(Y) = 0 Then
                lLen = Y
                Exit For
            End If
        Next Y
        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)
    End Function

    Public Sub LoadSettings()
        'Loads the settings from the INI file
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "ChatClient.ini"

        If Exists(sFile) = False Then
            Dim sTmp As String = AppDomain.CurrentDomain.BaseDirectory
            If sTmp.EndsWith("\") = False Then sTmp &= "\"
            sTmp &= "BPClient.ini"
            If Exists(sFile) = True Then
                FileCopy(sTmp, sFile)
            End If
        End If

        Dim oINI As New InitFile(sFile)
        Dim lR As Int32
        Dim lG As Int32
        Dim lB As Int32

        'DefaultChatColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        lR = CInt(Val(oINI.GetString("CHATCOLOR", "AliasR", "0")))
        lG = CInt(Val(oINI.GetString("CHATCOLOR", "AliasG", "0")))
        lB = CInt(Val(oINI.GetString("CHATCOLOR", "AliasB", "0")))
        AliasChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        lR = CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_R", "32")))
        lG = CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_G", "64")))
        lB = CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_B", "92")))
        TextBoxBackColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        'DefaultChatColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        lR = CInt(Val(oINI.GetString("CHATCOLOR", "DefaultR", "255")))
        lG = CInt(Val(oINI.GetString("CHATCOLOR", "DefaultG", "255")))
        lB = CInt(Val(oINI.GetString("CHATCOLOR", "DefaultB", "255")))
        DefaultChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        'AlertChatColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        lR = CInt(Val(oINI.GetString("CHATCOLOR", "AlertR", "255")))
        lG = CInt(Val(oINI.GetString("CHATCOLOR", "AlertG", "0")))
        lB = CInt(Val(oINI.GetString("CHATCOLOR", "AlertB", "0")))
        AlertChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        'StatusChatColor = System.Drawing.Color.FromArgb(255, 255, 255, 0)
        lR = CInt(Val(oINI.GetString("CHATCOLOR", "StatusR", "255")))
        lG = CInt(Val(oINI.GetString("CHATCOLOR", "StatusG", "255")))
        lB = CInt(Val(oINI.GetString("CHATCOLOR", "StatusB", "0")))
        StatusChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        'LocalChatColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        lR = CInt(Val(oINI.GetString("CHATCOLOR", "LocalR", "0")))
        lG = CInt(Val(oINI.GetString("CHATCOLOR", "LocalG", "255")))
        lB = CInt(Val(oINI.GetString("CHATCOLOR", "LocalB", "255")))
        LocalChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        'GuildChatColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        lR = CInt(Val(oINI.GetString("CHATCOLOR", "GuildR", "0")))
        lG = CInt(Val(oINI.GetString("CHATCOLOR", "GuildG", "255")))
        lB = CInt(Val(oINI.GetString("CHATCOLOR", "GuildB", "0")))
        GuildChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        'PMChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 0)
        lR = CInt(Val(oINI.GetString("CHATCOLOR", "PMR", "255")))
        lG = CInt(Val(oINI.GetString("CHATCOLOR", "PMG", "128")))
        lB = CInt(Val(oINI.GetString("CHATCOLOR", "PMB", "0")))
        PMChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        'SenateChatColor = System.Drawing.Color.FromArgb(255, 192, 192, 192)
        lR = CInt(Val(oINI.GetString("CHATCOLOR", "SenateR", "192")))
        lG = CInt(Val(oINI.GetString("CHATCOLOR", "SenateG", "192")))
        lB = CInt(Val(oINI.GetString("CHATCOLOR", "SenateB", "192")))
        SenateChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        'ChannelChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
        lR = CInt(Val(oINI.GetString("CHATCOLOR", "ChannelR", "255")))
        lG = CInt(Val(oINI.GetString("CHATCOLOR", "ChannelG", "128")))
        lB = CInt(Val(oINI.GetString("CHATCOLOR", "ChannelB", "255")))
        ChannelChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

        oINI = Nothing

    End Sub

    Public Sub SaveSettings()
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "ChatClient.ini"

        Dim oINI As New InitFile(sFile)

        With TextBoxBackColor
            oINI.WriteString("INTERFACE", "TextBoxFillColor_R", .R.ToString)
            oINI.WriteString("INTERFACE", "TextBoxFillColor_G", .G.ToString)
            oINI.WriteString("INTERFACE", "TextBoxFillColor_B", .B.ToString)
        End With 

        With DefaultChatColor
            oINI.WriteString("CHATCOLOR", "DefaultR", .R.ToString)
            oINI.WriteString("CHATCOLOR", "DefaultG", .G.ToString)
            oINI.WriteString("CHATCOLOR", "DefaultB", .B.ToString)
        End With
        With AlertChatColor
            oINI.WriteString("CHATCOLOR", "AlertR", .R.ToString)
            oINI.WriteString("CHATCOLOR", "AlertG", .G.ToString)
            oINI.WriteString("CHATCOLOR", "AlertB", .B.ToString)
        End With
        With StatusChatColor
            oINI.WriteString("CHATCOLOR", "StatusR", .R.ToString)
            oINI.WriteString("CHATCOLOR", "StatusG", .G.ToString)
            oINI.WriteString("CHATCOLOR", "StatusB", .B.ToString)
        End With
        With LocalChatColor
            oINI.WriteString("CHATCOLOR", "LocalR", .R.ToString)
            oINI.WriteString("CHATCOLOR", "LocalG", .G.ToString)
            oINI.WriteString("CHATCOLOR", "LocalB", .B.ToString)
        End With
        With GuildChatColor
            oINI.WriteString("CHATCOLOR", "GuildR", .R.ToString)
            oINI.WriteString("CHATCOLOR", "GuildG", .G.ToString)
            oINI.WriteString("CHATCOLOR", "GuildB", .B.ToString)
        End With
        With PMChatColor
            oINI.WriteString("CHATCOLOR", "PMR", .R.ToString)
            oINI.WriteString("CHATCOLOR", "PMG", .G.ToString)
            oINI.WriteString("CHATCOLOR", "PMB", .B.ToString)
        End With
        With SenateChatColor
            oINI.WriteString("CHATCOLOR", "SenateR", .R.ToString)
            oINI.WriteString("CHATCOLOR", "SenateG", .G.ToString)
            oINI.WriteString("CHATCOLOR", "SenateB", .B.ToString)
        End With
        With ChannelChatColor
            oINI.WriteString("CHATCOLOR", "ChannelR", .R.ToString)
            oINI.WriteString("CHATCOLOR", "ChannelG", .G.ToString)
            oINI.WriteString("CHATCOLOR", "ChannelB", .B.ToString)
        End With
        With AliasChatColor
            oINI.WriteString("CHATCOLOR", "AliasR", .R.ToString)
            oINI.WriteString("CHATCOLOR", "AliasG", .G.ToString)
            oINI.WriteString("CHATCOLOR", "AliasB", .B.ToString)
        End With

        oINI = Nothing
    End Sub

    Public Function Exists(ByVal sFilename As String) As Boolean
        If Trim(sFilename).Length > 0 Then
            On Error Resume Next
            sFilename = Dir$(sFilename, FileAttribute.Archive Or FileAttribute.Directory Or FileAttribute.Hidden Or FileAttribute.Normal Or FileAttribute.ReadOnly Or FileAttribute.System Or FileAttribute.Volume)
            Return Err.Number = 0 And sFilename.Length > 0
        Else
            Return False
        End If

    End Function

#Region "  Cache Handlers and Definitions  "
    'This structure is for caching detailed lists... for example, Player ID to Player Name or Unit Type ID to Unit Type Name
    Private Structure CacheEntry
        Public lID As Int32
        Public iTypeID As Int16
        Public iExtTypeID As Int16
        Public sValue As String
        Public bRequested As Boolean
    End Structure

    Private muCache() As CacheEntry
    Private mlCacheUB As Int32 = -1
    Private mlRequestsOutstanding As Int32 = 0
    Private mbInProcessGetCacheEntries As Boolean = False

    Public Sub ProcessGetCacheEntries()
        mbInProcessGetCacheEntries = True
        'Ok, we go through our cache list and see if any need requested
        If mlRequestsOutstanding > 0 AndAlso gfrmChat Is Nothing = False Then
            Dim lCnt As Int32 = mlRequestsOutstanding
            If lCnt * 6 + 5 > 32000 Then lCnt = (32000 - 5) \ 6

            Dim yMsg(5 + (lCnt * 6)) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityName).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

            For X As Int32 = 0 To mlCacheUB
                If muCache(X).bRequested = False Then
                    mlRequestsOutstanding -= 1
                    lCnt -= 1
                    System.BitConverter.GetBytes(muCache(X).lID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(muCache(X).iTypeID).CopyTo(yMsg, lPos) : lPos += 2
                    muCache(X).bRequested = True
                    If lCnt = 0 Then Exit For
                End If
            Next X
            gfrmChat.SendMsgToPrimary(yMsg)

            'If mlRequestsOutstanding Then
            mlRequestsOutstanding = 0
            For X As Int32 = 0 To mlCacheUB
                If muCache(X).bRequested = False Then mlRequestsOutstanding += 1
            Next X
            'End If
        End If

        mbInProcessGetCacheEntries = False
    End Sub

    Public Function GetCacheObjectValue(ByVal lID As Int32, ByVal iTypeID As Int16) As String
        While mbInProcessGetCacheEntries = True
            Threading.Thread.Sleep(10)
        End While

        Dim X As Int32
        'Dim yData() As Byte
        Dim lIdx As Int32 = -1

        'Ok, this works differently... find our Idx
        Try
            For X = 0 To mlCacheUB
                If muCache(X).lID = lID AndAlso muCache(X).iTypeID = iTypeID Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                'add the cache object
                mlCacheUB += 1
                ReDim Preserve muCache(mlCacheUB)
                lIdx = mlCacheUB
                muCache(lIdx).sValue = "Unknown"
                muCache(lIdx).lID = lID
                muCache(lIdx).iTypeID = iTypeID
                muCache(lIdx).bRequested = False
                mlRequestsOutstanding += 1
            End If
            ProcessGetCacheEntries()

            Return muCache(lIdx).sValue
        Catch
        End Try
        Return "Unknown"

    End Function

    Public Sub SetCacheObjectValue(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal sValue As String)
        Dim X As Int32
        Dim lIdx As Int32 = -1

        'set our value
        For X = 0 To mlCacheUB
            If muCache(X).lID = lID AndAlso muCache(X).iTypeID = iTypeID Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlCacheUB += 1
            ReDim Preserve muCache(mlCacheUB)
            lIdx = mlCacheUB

            muCache(lIdx).lID = lID
            muCache(lIdx).iTypeID = iTypeID
            muCache(lIdx).bRequested = True
        End If

        muCache(lIdx).sValue = sValue
    End Sub
#End Region

#Region "  Alias"
    Public Enum AliasingRights As Int32
        eNoRights = 0
        eMoveUnits = 1
        eDockUndockUnits = 2
        eChangeBehavior = 4
        eAddProduction = 8
        eCancelProduction = 16
        eViewTechDesigns = 32
        eCreateBattleGroups = 64
        eViewBattleGroups = 128
        eModifyBattleGroups = 256
        eViewDiplomacy = 512
        eAlterDiplomacy = 1024
        eViewAgents = 2048
        eAlterAgents = 4096
        eViewColonyStats = 8192
        eAlterColonyStats = 16384
        eViewBudget = 32768
        eViewMining = 65536
        eViewEmail = 131072
        eAlterEmail = 262144
        eViewTreasury = 524288
        eViewTrades = 1048576
        eAlterTrades = 2097152
        eAddResearch = 4194304
        eCancelResearch = 8388608
        eChangeEnvironment = 16777216
        eViewResearch = 33554432
        eViewUnitsAndFacilities = 67108864
        eAlterAutoLaunchPower = 134217728
        eTransferCargo = 268435456
        eCreateDesigns = 536870912
    End Enum

    Public Structure PlayerAlias
        Public sPlayerName As String
        Public sUserName As String
        Public sPassword As String
        Public lRights As Int32
    End Structure
    Public Function HasAliasedRights(ByVal lRights As AliasingRights) As Boolean
        Return True
        Return (gbAliased = False) OrElse (glAliasRights And lRights) <> 0 ' = lRights
    End Function
#End Region

End Module

Public Enum eyChatRoomCommandType As Byte
    AddNewChatRoom = 1
    LeaveChatRoom = 2
    ToggleAdminRights = 4
    JoinChannel = 8
    SetChannelPassword = 16
    KickPlayer = 32
    InvitePlayer = 64
    SetChannelPublic = 128
End Enum


Public Class StrEncDec
    Private myDec() As Byte = {0, 2, 208, 220, 168, 35, 187, 70, 212, 251, 255, 135, 175, 144, 32, 87, 231, 192, 189, 198, 34, 235, 4, 165, 94, 188, 82, 79, 75, 225, 105, 16, 224, 112, 177, 50, 171, 19, 99, 152, 174, 207, 65, 85, 15, 26, 232, 194, 217, 170, 73, 244, 185, 139, 72, 29, 98, 159, 21, 153, 249, 145, 53, 118, 117, 216, 127, 215, 147, 49, 206, 60, 42, 229, 69, 102, 180, 239, 96, 191, 253, 247, 20, 204, 243, 67, 140, 91, 161, 33, 108, 254, 186, 219, 128, 101, 250, 242, 62, 248, 84, 227, 164, 213, 64, 114, 137, 113, 5, 141, 68, 136, 146, 121, 48, 74, 111, 238, 109, 157, 209, 39, 214, 222, 123, 201, 103, 38, 132, 183, 115, 172, 162, 41, 12, 130, 167, 193, 66, 71, 181, 143, 155, 47, 81, 240, 173, 83, 233, 97, 6, 124, 241, 245, 10, 134, 184, 7, 36, 43, 236, 54, 228, 17, 31, 61, 104, 59, 129, 45, 133, 46, 14, 190, 106, 22, 203, 57, 179, 125, 28, 76, 55, 9, 122, 88, 3, 154, 18, 205, 44, 120, 11, 176, 93, 197, 226, 30, 200, 95, 77, 89, 119, 142, 158, 131, 196, 90, 40, 150, 23, 110, 178, 52, 202, 63, 107, 210, 163, 230, 80, 78, 234, 58, 92, 100, 237, 24, 252, 221, 116, 51, 199, 86, 182, 211, 223, 148, 8, 138, 56, 149, 156, 27, 13, 25, 195, 160, 218, 166, 126, 151, 169, 37, 246, 0}
    Private myEnc() As Byte = {0, 0, 1, 186, 22, 108, 150, 157, 238, 183, 154, 192, 134, 244, 172, 44, 31, 163, 188, 37, 82, 58, 175, 210, 227, 245, 45, 243, 180, 55, 197, 164, 14, 89, 20, 5, 158, 253, 127, 121, 208, 133, 72, 159, 190, 169, 171, 143, 114, 69, 35, 231, 213, 62, 161, 182, 240, 177, 223, 167, 71, 165, 98, 215, 104, 42, 138, 85, 110, 74, 7, 139, 54, 50, 115, 28, 181, 200, 221, 27, 220, 144, 26, 147, 100, 43, 233, 15, 185, 201, 207, 87, 224, 194, 24, 199, 78, 149, 56, 38, 225, 95, 75, 126, 166, 30, 174, 216, 90, 118, 211, 116, 33, 107, 105, 130, 230, 64, 63, 202, 191, 113, 184, 124, 151, 179, 250, 66, 94, 168, 135, 205, 128, 170, 155, 11, 111, 106, 239, 53, 86, 109, 203, 141, 13, 61, 112, 68, 237, 241, 209, 251, 39, 59, 187, 142, 242, 119, 204, 57, 247, 88, 132, 218, 102, 23, 249, 136, 4, 252, 49, 36, 131, 146, 40, 12, 193, 34, 212, 178, 76, 140, 234, 129, 156, 52, 92, 6, 25, 18, 173, 79, 17, 137, 47, 246, 206, 195, 19, 232, 198, 125, 214, 176, 83, 189, 70, 41, 2, 120, 217, 235, 8, 103, 122, 67, 65, 48, 248, 93, 3, 229, 123, 236, 32, 29, 196, 101, 162, 73, 219, 16, 46, 148, 222, 21, 160, 226, 117, 77, 145, 152, 97, 84, 51, 153, 254, 81, 99, 60, 96, 9, 228, 80, 91, 10}

    Public Function Encrypt(ByVal yData() As Byte) As Byte()
        Dim yBytes(yData.GetUpperBound(0)) As Byte
        For X As Int32 = 0 To yBytes.GetUpperBound(0)
            yBytes(X) = myEnc(yData(X))
        Next X
        Return yBytes
    End Function

    Public Function Decrypt(ByVal yData() As Byte) As Byte()
        Dim yBytes(yData.GetUpperBound(0)) As Byte
        For X As Int32 = 0 To yBytes.GetUpperBound(0)
            yBytes(X) = myDec(yData(X))
        Next X
        Return yBytes
    End Function
End Class
