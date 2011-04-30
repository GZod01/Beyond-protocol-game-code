Public Class BugEntry
    Public lID As Int32 = -1
    Public yCategory As Byte
    Public ySubCat As Byte
    Public yOccurs As Byte
    Public yPriority As Byte
    Public yStatus As Byte
    Public sUser As String = String.Empty
    Public sProblemDesc As String = String.Empty
    Public sStepsToProduce As String = String.Empty
    Public sDevNotes As String = String.Empty

    Public lCreatedByUserID As Int32 = -1
    Public lAssignedToID As Int32 = 0

    Public Shared Function GetSubCategory(ByVal yCatCode As Byte, ByVal ySubCatCode As Byte) As String
        Select Case yCatCode
            Case 0
                Select Case ySubCatCode
                    Case 0 : Return "Client-Side Crash"
                    Case 1 : Return "Server-Side Crash"
                    Case Else : Return "Other Crash"
                End Select
            Case 1
                Select Case ySubCatCode
                    Case 0 : Return "Gameplay - Exploit"
                    Case 1 : Return "Gameplay - Game Logic Bug"
                    Case 2 : Return "Gameplay - Geography Bug"
                    Case 3 : Return "Gameplay - Unexpected Result"
                    Case 4 : Return "Gameplay - User Interface"
                    Case Else : Return "Gameplay - Other"
                End Select
            Case 2
                Select Case ySubCatCode
                    Case 0 : Return "Graphical - Geography"
                    Case 1 : Return "Graphical - Models"
                    Case 2 : Return "Graphical - Particle FX"
                    Case 3 : Return "Graphical - Textures"
                    Case 4 : Return "Graphical - User Interface"
                    Case Else : Return "Graphical - Other"
                End Select
            Case 3
                Select Case ySubCatCode
                    Case 0 : Return "Performance - Low Frame Rate"
                    Case 1 : Return "Performance - Memory Leak"
                    Case 2 : Return "Performance - Perceived Lag"
                    Case 3 : Return "Performance - Game Stutter"
                    Case Else : Return "Performance - Other"
                End Select
            Case 4
                Return "Other Miscellaneous"
            Case 5
                Select Case ySubCatCode
                    Case 0 : Return "Login - Login Credentials"
                    Case 1 : Return "Login - Connections Issue"
                    Case Else : Return "Login - Other"
                End Select
            Case 6
                Select Case ySubCatCode
                    Case 0 : Return "Suggestions - Gameplay-Related"
                    Case 1 : Return "Suggestions - User Interface"
                    Case 2 : Return "Suggestions - Hotkeys"
                    Case Else : Return "Suggestions - Other"
                End Select
            Case 7
                Select Case ySubCatCode
                    Case 0 : Return "Combat"
                    Case 1 : Return "Movement"
                    Case 2 : Return "User Interface"
                    Case 3 : Return "Performance"
                    Case 4 : Return "Production"
                    Case 5 : Return "Tech Builder"
                    Case 6 : Return "Special Techs"
                    Case 7 : Return "Mining"
                    Case 8 : Return "Trading"
                    Case 9 : Return "Battlegroups"
                    Case 10 : Return "In-Game Email"
                    Case 11 : Return "Diplomacy"
                    Case 12 : Return "Colony/Budget"
                    Case 13 : Return "Chat"
                    Case 14 : Return "Repair"
                    Case 15 : Return "Space stations"
                    Case 16 : Return "Guilds/Alliances"
                    Case 17 : Return "Espionage"
                    Case 18 : Return "Sound"
                    Case Else : Return "Other - be specific"
                End Select
            Case Else
                Return "Unknown"
        End Select
    End Function

    Public Shared Function GetOccurrence(ByVal yOccurs As Byte) As String
        Select Case yOccurs
            Case 0 : Return "Easily Reproduceable"
            Case 1 : Return "Intermittent"
            Case Else : Return "Rarely"
        End Select
    End Function

    Public Shared Function GetPriority(ByVal yPriority As Byte) As String
        Select Case yPriority
            Case 0 : Return "Extremely High"
            Case 1 : Return "High"
            Case 2 : Return "Normal"
            Case 3 : Return "Low"
            Case Else : Return "Extremely Low"
        End Select
    End Function

    Public Shared Function GetStatus(ByVal yStatus As Byte) As String
        Select Case yStatus
            Case 0 : Return "New"
            Case 1 : Return "Open"
            Case 2 : Return "In Progress"
            Case 3 : Return "Unreleased"
            Case 4 : Return "Released"
            Case 5 : Return "On Hold"
            Case 6 : Return "Waiting"
            Case Else : Return "Closed"
        End Select
    End Function

    Public Shared Function GetCategory(ByVal yCategory As Byte) As String
        Select Case yCategory
            Case 0
                Return "Crash"
            Case 1
                Return "Gameplay"
            Case 2
                Return "Graphical"
            Case 3
                Return "Performance"
            Case 4
                Return "Other"
            Case 5
                Return "Login"
            Case 6
                Return "Suggestions"
            Case 7
                Return "Test Case"
            Case Else
                Return "Unknown"
        End Select
    End Function

    Public Function GenerateSaveMsg() As Byte()
        Dim yMsg() As Byte
        Dim lPos As Int32

        ReDim yMsg(25 + sProblemDesc.Length + sStepsToProduce.Length + sDevNotes.Length)

        System.BitConverter.GetBytes(GlobalMessageCode.eBugSubmission).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(lID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 6)
        yMsg(10) = yCategory
        yMsg(11) = ySubCat
        yMsg(12) = yOccurs
        yMsg(13) = yPriority
        yMsg(14) = yStatus

        lPos = 15
        System.BitConverter.GetBytes(lAssignedToID).CopyTo(yMsg, lPos) : lPos += 4

        'Copy the problem desc
        System.BitConverter.GetBytes(CShort(sProblemDesc.Length)).CopyTo(yMsg, lPos) : lPos += 2
        Array.Copy(System.Text.ASCIIEncoding.ASCII.GetBytes(sProblemDesc), 0, yMsg, lPos, sProblemDesc.Length)
        lPos += sProblemDesc.Length

        'Copy the Steps to Produce
        System.BitConverter.GetBytes(CShort(sStepsToProduce.Length)).CopyTo(yMsg, lPos) : lPos += 2
        Array.Copy(System.Text.ASCIIEncoding.ASCII.GetBytes(sStepsToProduce), 0, yMsg, lPos, sStepsToProduce.Length)
        lPos += sStepsToProduce.Length

        'Copy the Dev Notes
        System.BitConverter.GetBytes(CShort(sDevNotes.Length)).CopyTo(yMsg, lPos) : lPos += 2
        Array.Copy(System.Text.ASCIIEncoding.ASCII.GetBytes(sDevNotes), 0, yMsg, lPos, sDevNotes.Length)

        Return yMsg

    End Function
End Class