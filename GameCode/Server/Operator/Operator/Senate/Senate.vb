Option Strict On

Public Enum eySenateRequestDetailsType As Byte
    SystemVote = 0
    SenateProposal = 1
    ProposalMsg = 2
    SenateObject = 3
    EmpChmbrMsgList = 4
    EmpChmbrMsg = 5
End Enum
Public Enum eyVoteValue As Byte
    AbstainVote = 0
    NoVote = 1
    YesVote = 2
End Enum
''' <summary>
''' Contains all of the functions needed for senate management
''' </summary>
''' <remarks></remarks>
Public Class Senate

    Private Shared moProposals() As SenateProposal
    Private Shared mlProposalUB As Int32 = -1

    Private Shared moEmpChmbrMsg() As EmpMsgBoardItem
    Private Shared mlEmpChmbrMsgUB As Int32 = -1

    Public Shared Sub AddEmpChmbrMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lPos += 4   'for proposalid
        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse (oPlayer.yPlayerTitle <> Player.PlayerRank.Emperor AndAlso oPlayer.ObjectID > 6) Then
            LogEvent(LogEventType.PossibleCheat, "Player posting EmpChmbrMsg without permission or player not found: " & lPlayerID)
            Return
        End If
        If lLen = 0 OrElse lLen > 1000 Then
            LogEvent(LogEventType.PossibleCheat, "EmpChmbrMsg Len is out of range: " & lPlayerID)
            Return
        End If

        Dim oMsg As New EmpMsgBoardItem
        With oMsg
            .lPosterID = lPlayerID
            .lPostedOn = GetDateAsNumber(Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime)
            ReDim .yMsg(lLen - 1)
            Array.Copy(yData, lPos, .yMsg, 0, lLen)

            If .SaveObject() = True Then
                If mlEmpChmbrMsgUB < 24 Then
                    mlEmpChmbrMsgUB += 1
                    ReDim Preserve moEmpChmbrMsg(mlEmpChmbrMsgUB)
                End If
                For X As Int32 = mlEmpChmbrMsgUB To 1 Step -1
                    moEmpChmbrMsg(X) = moEmpChmbrMsg(X - 1)
                Next X
                moEmpChmbrMsg(0) = oMsg
            End If
        End With
    End Sub

    Public Shared Function GetSaveObjectText() As String
        Dim oSB As New System.Text.StringBuilder
        For X As Int32 = 0 To mlProposalUB
            If moProposals(X) Is Nothing = False Then
                oSB.AppendLine(moProposals(X).GetSaveObjectText())
            End If
        Next X
        Return oSB.ToString
    End Function

    Public Shared Sub AddNewProposal(ByRef oProposal As SenateProposal)
        If moProposals Is Nothing Then ReDim moProposals(-1)
        SyncLock moProposals
            mlProposalUB += 1
            ReDim Preserve moProposals(mlProposalUB)
            moProposals(mlProposalUB) = oProposal
        End SyncLock
    End Sub

    Private Shared Sub DeleteProposal(ByVal lProposalID As Int32, ByVal lProposerID As Int32)
        Dim oProposal As SenateProposal = GetProposal(lProposalID)
        If oProposal Is Nothing Then Return
        If oProposal.lProposedBy <> lProposerID Then
            LogEvent(LogEventType.PossibleCheat, "Player attempting to delete proposal not belonging to player: " & lProposerID)
            Return
        End If
        oProposal.yProposalState = oProposal.yProposalState Or eyProposalState.Archived
        oProposal.SaveObject()
    End Sub

    Private Shared Sub HandleUpdateProposalMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lProposerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lPos += 4   'for the -2
        Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTitleLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lTitleLen > 255 Then Return
        Dim sTitle As String = GetStringFromBytes(yData, lPos, lTitleLen) : lPos += lTitleLen
        Dim lDescLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lDescLen > 1000 Then Return
        Dim sDesc As String = GetStringFromBytes(yData, lPos, lDescLen) : lPos += lDescLen

        Dim oProposal As SenateProposal = GetProposal(lProposalID)
        If oProposal Is Nothing = False Then
            If oProposal.lProposedBy = lProposerID Then
                If (oProposal.yProposalState And eyProposalState.OnSenateFloor) <> 0 Then
                    LogEvent(LogEventType.PossibleCheat, "Proposal on senate floor trying to be changed: " & lProposerID)
                    Return
                End If
                oProposal.yDescription = StringToBytes(sDesc)
                oProposal.yTitle = StringToBytes(sTitle)
                oProposal.AddMsg(lProposerID, GetDateAsNumber(Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime), "Proposal Details Updated")

            Else
                LogEvent(LogEventType.PossibleCheat, "Non-proposer trying to change proposal: " & lProposerID)
            End If
        End If

    End Sub

    Public Shared Function HandleCreateProposalMsg(ByRef yData() As Byte) As Byte()
        Dim lPos As Int32 = 2
        Dim lProposerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTitleLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If lTitleLen = -1 Then
            'ok, removing a proposal
            Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            DeleteProposal(lProposalID, lProposerID)
            Return Nothing
        ElseIf lTitleLen = -2 Then
            HandleUpdateProposalMsg(yData)
            Return Nothing
        End If

        If lTitleLen > 255 Then Return Nothing
        Dim sTitle As String = GetStringFromBytes(yData, lPos, lTitleLen) : lPos += lTitleLen
        Dim lDescLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lDescLen > 1000 Then Return Nothing
        Dim sDesc As String = GetStringFromBytes(yData, lPos, lDescLen) : lPos += lDescLen

        Dim lProposedOn As Int32 = GetDateAsNumber(Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime)

        Dim oProposer As Player = GetEpicaPlayer(lProposerID)
        If oProposer Is Nothing OrElse oProposer.yPlayerTitle < Player.PlayerRank.Emperor OrElse (oProposer.yPlayerTitle And Player.PlayerRank.ExRankShift) <> 0 Then
            LogEvent(LogEventType.PossibleCheat, "Player Proposing Senate Law that is not an emperor: " & lProposerID)
            Return Nothing
        End If

        'Ok, doublecheck that the proposer does not have more than two proposals on the emperors chambers
        Dim lCnt As Int32 = 0
        Try
            For X As Int32 = 0 To mlProposalUB
                If moProposals(X) Is Nothing = False Then
                    If moProposals(X).lProposedBy = lProposerID AndAlso (moProposals(X).yProposalState And (eyProposalState.OnSenateFloor Or eyProposalState.Archived)) = 0 Then
                        lCnt += 1
                    End If
                End If
            Next X
        Catch
            lCnt = 2000
        End Try
        If lCnt > 9 Then Return Nothing

        Dim oProposal As New SenateProposal()
        With oProposal
            .lProposedBy = lProposerID
            .ObjectID = -1
            .ObjTypeID = ObjectType.eSenateLaw
            .yDescription = StringToBytes(sDesc)
            .yProposalState = eyProposalState.EmperorsChamber
            .yTitle = StringToBytes(sTitle)
            .lProposedOn = lProposedOn
            If .SaveObject() = False Then Return Nothing
        End With
        AddNewProposal(oProposal)

        Return oProposal.GetSmallMsg(lProposerID)
    End Function

    Public Shared Function HandleRequestSenate(ByVal bIncludeEmpChamber As Boolean, ByVal lForPlayerID As Int32) As Byte()
        'ok, here, we need to send the player all of the msgs...
        Dim yCache(200000) As Byte
        Dim yFinal() As Byte
        Dim lPos As Int32
        Dim lSingleMsgLen As Int32

        Dim yTemp() As Byte
        Dim dtNow As Date = Now
        Dim dtAncient As Date = Now.AddMonths(-1)

        lPos = 0
        lSingleMsgLen = -1
        For X As Int32 = 0 To mlProposalUB
            If moProposals(X) Is Nothing = False AndAlso (bIncludeEmpChamber = True OrElse (moProposals(X).yProposalState And eyProposalState.OnSenateFloor) <> 0) Then
                If (moProposals(X).yProposalState And eyProposalState.Archived) <> 0 Then
                    If (moProposals(X).yProposalState And eyProposalState.OnSenateFloor) = 0 Then Continue For
                    If (moProposals(X).yProposalState And (eyProposalState.PassedFloor Or eyProposalState.FailedFloor)) = 0 Then Continue For
                    If Date.SpecifyKind(GetDateFromNumber(moProposals(X).lVotingEndDate), DateTimeKind.Utc).ToLocalTime < dtAncient Then
                        If (moProposals(X).yProposalState And eyProposalState.PassedFloor) = 0 Then Continue For
                        If (moProposals(X).yProposalState And eyProposalState.Implemented) <> 0 Then Continue For
                    End If
                End If
                If (moProposals(X).yProposalState And eyProposalState.OnSenateFloor) <> 0 AndAlso Date.SpecifyKind(GetDateFromNumber(moProposals(X).lVotingEndDate), DateTimeKind.Utc).ToLocalTime.AddDays(4) < dtNow Then
                    moProposals(X).yProposalState = moProposals(X).yProposalState Or eyProposalState.Archived
                    moProposals(X).SaveObject()
                    'Continue For
                End If
                If moProposals(X).lVotingEndDate > 0 AndAlso (moProposals(X).yProposalState And (eyProposalState.FailedFloor Or eyProposalState.PassedFloor)) = 0 Then
                    If Date.SpecifyKind(GetDateFromNumber(moProposals(X).lVotingEndDate), DateTimeKind.Utc).ToLocalTime < dtNow Then

                        If moProposals(X).GetVotesFor() >= Senate.RequiredVoteCount(False) Then
                            moProposals(X).yProposalState = moProposals(X).yProposalState Or eyProposalState.PassedFloor
                        Else
                            moProposals(X).yProposalState = moProposals(X).yProposalState Or eyProposalState.FailedFloor
                        End If
                        moProposals(X).SaveObject()
                    End If
                End If

                yTemp = moProposals(X).GetSmallMsg(lForPlayerID)
                lSingleMsgLen = yTemp.Length
                'Ok, before we continue, check if we need to increase our cache
                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                    'increase it
                    ReDim Preserve yCache(yCache.Length + 200000)
                End If
                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                lPos += 2
                yTemp.CopyTo(yCache, lPos)
                lPos += lSingleMsgLen
            End If
        Next X
        If lPos <> 0 Then
            ReDim yFinal(lPos - 1)
            Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
            Return yFinal
        End If
        Return Nothing
    End Function

    Public Shared Sub HandleAddProposalMessage(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lProposalID = -1I Then
            AddEmpChmbrMsg(yData)
            Return
        End If
        Dim lMsgLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lMsgLen > 500 Then
            LogEvent(LogEventType.PossibleCheat, "Player posting message larger than 500 characters: " & lPlayerID)
            Return
        End If
        Dim sMsg As String = GetStringFromBytes(yData, lPos, lMsgLen)
        Dim lPostedOn As Int32 = GetDateAsNumber(Now.ToUniversalTime)
        Dim oProposal As SenateProposal = GetProposal(lProposalID)
        If oProposal Is Nothing = False Then
            oProposal.AddMsg(lPlayerID, lPostedOn, sMsg)
        End If
    End Sub

    Public Shared Sub SetPlayerVote(ByVal lProposalID As Int32, ByVal yVoteValue As eyVoteValue, ByVal lPlayerID As Int32, ByVal yPriority As Byte)
        For X As Int32 = 0 To mlProposalUB
            If moProposals(X) Is Nothing = False AndAlso moProposals(X).ObjectID = lProposalID Then
                moProposals(X).HandlePlayerVote(lPlayerID, yVoteValue, False, yPriority)
                Exit For
            End If
        Next X
    End Sub

    Public Shared Function GetProposal(ByVal lProposalID As Int32) As SenateProposal
        For X As Int32 = 0 To mlProposalUB
            If moProposals(X) Is Nothing = False AndAlso moProposals(X).ObjectID = lProposalID Then
                Return moProposals(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Shared Sub RecalculateAllProposals()
        For X As Int32 = 0 To mlProposalUB
            If moProposals(X) Is Nothing = False Then
                moProposals(X).Recalculate()
            End If
        Next X
    End Sub

    Public Shared Function HandleGetSenateObjectDetails(ByRef yData() As Byte, ByRef bSendLenAppended As Boolean) As Byte()
        Dim lPos As Int32 = 2   'for msgcode

        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return Nothing
        bSendLenAppended = False
        Dim yType As Byte = yData(lPos) : lPos += 1
        Select Case yType
            Case eySenateRequestDetailsType.ProposalMsg
                Dim lPosterID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lPostedOn As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim oProposal As SenateProposal = GetProposal(lProposalID)
                If oProposal Is Nothing Then Return Nothing
                If oPlayer.ObjectID <> 1 AndAlso oPlayer.ObjectID <> 6 Then
                    If (oProposal.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
                        If oPlayer.yPlayerTitle <> Player.PlayerRank.Emperor Then
                            LogEvent(LogEventType.PossibleCheat, "Player requesting details from proposal without access: " & lPlayerID)
                            Return Nothing
                        End If
                    End If
                End If
                Return oProposal.GetProposalMsgDetails(lPosterID, lPostedOn, lPlayerID)
            Case eySenateRequestDetailsType.SenateProposal
                Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim oProposal As SenateProposal = GetProposal(lProposalID)
                If oProposal Is Nothing Then Return Nothing
                If oPlayer.ObjectID <> 1 AndAlso oPlayer.ObjectID <> 6 Then
                    If (oProposal.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
                        If oPlayer.yPlayerTitle <> Player.PlayerRank.Emperor Then
                            LogEvent(LogEventType.PossibleCheat, "Player requesting details from proposal without access: " & lPlayerID)
                            Return Nothing
                        End If
                    End If
                End If
                Return oProposal.GetBigMsg(lPlayerID)
            Case eySenateRequestDetailsType.SystemVote
                Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim oProposal As SenateProposal = GetProposal(lProposalID)
                If oProposal Is Nothing Then Return Nothing
                If oPlayer.ObjectID <> 1 AndAlso oPlayer.ObjectID <> 6 Then
                    If (oProposal.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
                        If oPlayer.yPlayerTitle <> Player.PlayerRank.Emperor Then
                            LogEvent(LogEventType.PossibleCheat, "Player requesting details from proposal without access: " & lPlayerID)
                            Return Nothing
                        End If
                    End If
                End If
                Return oProposal.GetSystemVoteDetails(lSystemID, lPlayerID)
            Case eySenateRequestDetailsType.SenateObject
                Dim bIncludeEmperorsChamber As Boolean = False
                bSendLenAppended = True
                If oPlayer Is Nothing = False Then bIncludeEmperorsChamber = (oPlayer.yPlayerTitle = Player.PlayerRank.Emperor) AndAlso (oPlayer.yPlayerTitle And Player.PlayerRank.ExRankShift) = 0
                If oPlayer.ObjectID = 1 OrElse oPlayer.ObjectID = 6 Then bIncludeEmperorsChamber = True
                Return Senate.HandleRequestSenate(bIncludeEmperorsChamber, lPlayerID)
            Case eySenateRequestDetailsType.EmpChmbrMsg
                Dim lPosterID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lPostedOn As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                For X As Int32 = 0 To mlEmpChmbrMsgUB
                    If moEmpChmbrMsg(X).lPostedOn = lPostedOn AndAlso moEmpChmbrMsg(X).lPosterID = lPosterID Then
                        With moEmpChmbrMsg(X)
                            Dim lLen As Int32 = 18
                            If .yMsg Is Nothing = False Then lLen += .yMsg.Length
                            Dim yResp(lLen) As Byte
                            lPos = 0
                            System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yResp, lPos) : lPos += 2
                            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
                            yResp(lPos) = eySenateRequestDetailsType.EmpChmbrMsg : lPos += 1

                            System.BitConverter.GetBytes(.lPosterID).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.lPostedOn).CopyTo(yResp, lPos) : lPos += 4
                            If .yMsg Is Nothing Then
                                System.BitConverter.GetBytes(0I).CopyTo(yResp, lPos) : lPos += 4
                            Else
                                System.BitConverter.GetBytes(.yMsg.Length).CopyTo(yResp, lPos) : lPos += 4
                                Array.Copy(.yMsg, 0, yResp, lPos, .yMsg.Length) : lPos += .yMsg.Length
                            End If
                            Return yResp
                        End With
                        Exit For

                    End If
                Next X
            Case eySenateRequestDetailsType.EmpChmbrMsgList
                Dim yResp(10 + ((mlEmpChmbrMsgUB + 1) * 8)) As Byte
                lPos = 0

                System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
                yResp(lPos) = eySenateRequestDetailsType.EmpChmbrMsgList : lPos += 1
                System.BitConverter.GetBytes(mlEmpChmbrMsgUB + 1).CopyTo(yResp, lPos) : lPos += 4
                For X As Int32 = 0 To mlEmpChmbrMsgUB
                    System.BitConverter.GetBytes(moEmpChmbrMsg(X).lPosterID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(moEmpChmbrMsg(X).lPostedOn).CopyTo(yResp, lPos) : lPos += 4
                Next X
                Return yResp
        End Select
        Return Nothing
    End Function

    Public Shared Sub HandlePlayerSetVote(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim yVoteType As Byte = yData(lPos) : lPos += 1
        If yVoteType = 0 Then
            'guild?
        Else
            'senate
            Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yVoteVal As eyVoteValue = CType(yData(lPos), eyVoteValue) : lPos += 1
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yPriority As Byte = CByte(Math.Max(0, Math.Min(5, yData(lPos)))) : lPos += 1
            SetPlayerVote(lProposalID, yVoteVal, lPlayerID, yPriority)
        End If
    End Sub

    Public Shared Function HandleSenateStatusReport(ByVal yData() As Byte) As Byte()
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim yResp() As Byte = Nothing
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim bIncludeEmp As Boolean = oPlayer.yPlayerTitle = Player.PlayerRank.Emperor
            If lPlayerID = 1 OrElse lPlayerID = 6 Then bIncludeEmp = True

            Dim lEmpChmbrOpen As Int32 = 0
            Dim lEmpChmbrVoted As Int32 = 0
            Dim lFloorOpen As Int32 = 0
            Dim lFloorVoted As Int32 = 0

            For X As Int32 = 0 To mlProposalUB
                Dim oProp As SenateProposal = moProposals(X)
                If oProp Is Nothing = False Then
                    If (oProp.yProposalState And eyProposalState.OnSenateFloor) <> 0 OrElse bIncludeEmp = True Then

                        If (oProp.yProposalState And (eyProposalState.OnSenateFloor Or eyProposalState.Archived Or eyProposalState.DevDeclined Or eyProposalState.FailedFloor Or eyProposalState.PassedFloor)) = eyProposalState.OnSenateFloor Then
                            lFloorOpen += 1
                            Dim oVote As SenateVote = oProp.GetVote(lPlayerID)
                            If oVote Is Nothing = False AndAlso oVote.yVote <> eyVoteValue.AbstainVote Then lFloorVoted += 1
                        ElseIf (oProp.yProposalState And (eyProposalState.EmperorsChamber Or eyProposalState.Archived Or eyProposalState.DevDeclined Or eyProposalState.AwaitingApproval)) = eyProposalState.EmperorsChamber Then
                            lEmpChmbrOpen += 1
                            Dim oVote As SenateVote = oProp.GetVote(lPlayerID)
                            If oVote Is Nothing = False AndAlso oVote.yVote <> eyVoteValue.AbstainVote Then lEmpChmbrVoted += 1
                        End If
                    End If
                End If
            Next X

            ReDim yResp(21)
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSenateStatusReport).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lFloorOpen).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lFloorVoted).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lEmpChmbrOpen).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lEmpChmbrVoted).CopyTo(yResp, lPos) : lPos += 4

        End If

        Return yResp
    End Function

    Public Shared Function RequiredVoteCount(ByVal bEmpChmbr As Boolean) As Int32
        Dim lResult As Int32 = 0
        If bEmpChmbr = True Then
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) > -1 Then
                    Dim oPlayer As Player = goPlayer(X)
                    If oPlayer Is Nothing = False Then
                        If oPlayer.yPlayerTitle = Player.PlayerRank.Emperor Then lResult += 1
                    End If
                End If
            Next X
            lResult = CInt(Math.Ceiling(lResult * 0.51F))
        Else
            For X As Int32 = 0 To glSystemUB
                If glSystemIdx(X) > -1 Then
                    Dim oSys As SolarSystem = goSystem(X)
                    If oSys Is Nothing = False Then
                        If oSys.SystemType = 255 OrElse oSys.SystemType = 128 Then Continue For
                        For Y As Int32 = 0 To oSys.mlPlanetUB
                            Dim oPlanet As Planet = oSys.GetPlanet(Y)
                            If oPlanet.OwnerID > 0 Then
                                lResult += oSys.mlPlanetUB + 1
                                Exit For
                            End If
                        Next Y
                    End If
                End If
            Next X
            lResult = CInt(Math.Ceiling(lResult * 0.6))
        End If

        Return lResult
    End Function

    Public Shared Sub PlayerRenamed(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

        Try
            Dim sSQL As String = "DELETE FROM tblSenateProposalMsg WHERE PosterID = " & lPlayerID
            Dim oComm As New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing
        Catch
        End Try

        'Ok, go through our senate and clear it up
        For X As Int32 = 0 To mlEmpChmbrMsgUB
            If moEmpChmbrMsg(X) Is Nothing = False AndAlso moEmpChmbrMsg(X).lPosterID = lPlayerID Then
                moEmpChmbrMsg(X).lPosterID = -1
                moEmpChmbrMsg(X).SaveObject()
            End If
        Next X
        For X As Int32 = 0 To mlProposalUB
            If moProposals(X) Is Nothing = False Then
                moProposals(X).ClearPlayerMsgs(lPlayerID)
            End If
        Next X

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        oPlayer.sPlayerNameProper = sName : oPlayer.sPlayerName = sName.ToUpper
        ReDim oPlayer.PlayerName(19)
        StringToBytes(sName).CopyTo(oPlayer.PlayerName, 0)

    End Sub
End Class
