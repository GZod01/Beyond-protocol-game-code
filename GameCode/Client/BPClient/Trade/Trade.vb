Option Strict On

Public Class Trade
    Inherits Base_GUID

    Public Enum eTradeStateValues As Byte
        Proposal = 0
        InNegotiation = 1       'synonomous with OPEN
        InVerification = 2      'both players have accepted and the server needs to verify the trade parameters
        InProgress = 4          'trade is in progress, there should be 2 queue action items one at 1/2 time and one at end
        TradeCompleted = 8      'trade is finished
        Player1Accepted = 16
        Player2Accepted = 32

        TradeRejected = 255
    End Enum

    Public Enum eFailureReason As Byte
        NoFailureReason = 0
        Player1DestNotFound = 1
        Player1SourceNotFound = 2
        Player2DestNotFound = 3
        Player2SourceNotFound = 4
        Player1MissingItems = 5
        Player2MissingItems = 6
		FacilityTradeDifferentEnvirs = 7
		Player1AgentsFull = 8
		Player2AgentsFull = 9
    End Enum

    Public Structure TradeItem
        Public ObjectID As Int32
        Public ObjTypeID As Int16
		Public Quantity As Int64

		Public lExtendedID As Int32
    End Structure

    'For now, Trades are player to player (1 to 1)... this flattens the trade object

    Public yTradeState As Byte
    Public yFailureReason As Byte

    Public lPlayer1ID As Int32
    'Public lP1DestID As Int32
    'Public lP1SourceID As Int32
    Public lP1TradepostID As Int32
    Public sP1Notes As String
    Public muPlayer1Items() As TradeItem
    Public mlPlayer1ItemUB As Int32 = -1

    Public lPlayer2ID As Int32
    'Public lP2DestID As Int32
    'Public lP2SourceID As Int32
    Public lP2TradepostID As Int32
    Public sP2Notes As String
    Public muPlayer2Items() As TradeItem
    Public mlPlayer2ItemUB As Int32 = -1

    'Public TradeStarted As Int32
    Public dtTradeStarted As Date = Date.MinValue
    
    Public TradeCycles As Int32         'total cycles for the trade to execute
    'Public CyclesRemaining As Int32     'cycles remaining before trade ends

    Public lLastMsgUpdate As Int32

    Public Sub PopulateFromMessage(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode

        'GUID
        ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        'P1ID
        lPlayer1ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'P2ID
        lPlayer2ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'tradestate
        yTradeState = yData(lPos) : lPos += 1
        'tradecycles (total)
        TradeCycles = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'cycles remaining
        Dim lCyclesAgo As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Try
            If lCyclesAgo <> 0 Then
                Dim lSecondsAgo As Int32 = lCyclesAgo \ 30
                dtTradeStarted = Now.Subtract(New TimeSpan(0, 0, 0, lSecondsAgo, 0))        'given in cycles, 30 ms to a cycle
            Else : dtTradeStarted = Date.MinValue
            End If
        Catch
        End Try

        'CyclesRemaining = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        yFailureReason = yData(lPos) : lPos += 1

        'If CyclesRemaining > 0 Then
        '    TradeStarted = glCurrentCycle - (TradeCycles - CyclesRemaining)
        'End If

        'now, Player 1's data
        'lP1DestID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'lP1SourceID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lP1TradepostID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lTemp <> 0 Then
            sP1Notes = GetStringFromBytes(yData, lPos, lTemp)
        Else : sP1Notes = ""
        End If
        lPos += lTemp

        mlPlayer1ItemUB = -1
        lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ReDim muPlayer1Items(lTemp - 1)
        For X As Int32 = 0 To lTemp - 1
            With muPlayer1Items(X)
                .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .Quantity = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                .lExtendedID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End With
        Next X
        mlPlayer1ItemUB = lTemp - 1

        'Now, player 2's data
        'lP2DestID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'lP2SourceID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lP2TradepostID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lTemp <> 0 Then
            sP2Notes = GetStringFromBytes(yData, lPos, lTemp)
        Else : sP2Notes = ""
        End If
        lPos += lTemp

        mlPlayer2ItemUB = -1
        lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ReDim muPlayer2Items(lTemp - 1)
        For X As Int32 = 0 To lTemp - 1
            With muPlayer2Items(X)
                .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .Quantity = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                .lExtendedID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End With
        Next X
        mlPlayer2ItemUB = lTemp - 1

        lLastMsgUpdate = glCurrentCycle
    End Sub

    Public Function GetTradeETA() As String

        If yTradeState = 255 OrElse (yTradeState And eTradeStateValues.InProgress) = 0 Then Return ""

        If dtTradeStarted = Date.MinValue Then Return ""
        'Dim dtTradeStarted As Date = Date.SpecifyKind(GetDateFromNumber(TradeStarted), DateTimeKind.Local)
        Dim lTemp As Int32 = CInt(Date.SpecifyKind(Now, DateTimeKind.Local).Subtract(dtTradeStarted).TotalSeconds)
        Return GetDurationFromSeconds(CInt((TradeCycles / 30) - lTemp), True)
        'Dim lTemp As Int32 = glCurrentCycle - TradeStarted  'cycles elapsed

        'Dim lCyclesRemaining As Int32 = TradeCycles - lTemp

        'Dim lSeconds As Int32 = CInt(Math.Ceiling(lCyclesRemaining / 33.3333F))
        'Dim lMinutes As Int32 = lSeconds \ 60
        'lSeconds -= (lMinutes * 60)
        'Dim lHours As Int32 = lMinutes \ 60
        'lMinutes -= (lHours * 60)
        'Dim lDays As Int32 = lHours \ 24
        'lHours -= (lDays * 24)

        'If lDays > 0 Then
        '    Return lDays.ToString("0#") & ":" & lHours.ToString("0#") & ":" & lMinutes.ToString("0#") & ":" & lSeconds.ToString("0#")
        'ElseIf lHours > 0 Then3
        '    Return lHours.ToString("0#") & ":" & lMinutes.ToString("0#") & ":" & lSeconds.ToString("0#")
        'Else
        '    Return lMinutes.ToString("0#") & ":" & lSeconds.ToString("0#")
        'End If

    End Function

    Public Function GetTradeMainListText() As String
        Dim sResult As String
        If lPlayer1ID = glPlayerID Then
            sResult = GetCacheObjectValue(lPlayer2ID, ObjectType.ePlayer).PadRight(22, " "c)
        Else : sResult = GetCacheObjectValue(lPlayer1ID, ObjectType.ePlayer).PadRight(22, " "c)
        End If

        If yTradeState = eTradeStateValues.TradeRejected Then
            sResult &= "Rejected"
        ElseIf yFailureReason <> eFailureReason.NoFailureReason Then
            Select Case yFailureReason
                Case eFailureReason.FacilityTradeDifferentEnvirs
                    sResult = "FAILED: Facility trading must go to same environment!"
                Case eFailureReason.Player1DestNotFound
                    If Me.lPlayer1ID = glPlayerID Then
                        sResult = "FAILED: Your Destination was not found!"
                    Else : sResult = "FAILED: Their Destination was not found!"
                    End If
                Case eFailureReason.Player1MissingItems
                    If Me.lPlayer1ID = glPlayerID Then
                        sResult = "FAILED: You have promised items you do not have!"
                    Else : sResult = "FAILED: They have promised items they do not have!"
                    End If
                Case eFailureReason.Player1SourceNotFound
                    If Me.lPlayer1ID = glPlayerID Then
                        sResult = "FAILED: Your Source was not found!"
                    Else : sResult = "FAILED: Their Source was not found!"
                    End If
                Case eFailureReason.Player2DestNotFound
                    If Me.lPlayer1ID = glPlayerID Then
                        sResult = "FAILED: Their Destination was not found!"
                    Else : sResult = "FAILED: Your Destination was not found!"
                    End If
                Case eFailureReason.Player2MissingItems
                    If Me.lPlayer1ID = glPlayerID Then
                        sResult = "FAILED: They have promised items they do not have!"
                    Else : sResult = "FAILED: You have promised items you do not have!"
                    End If
                Case eFailureReason.Player2SourceNotFound
                    If Me.lPlayer1ID = glPlayerID Then
                        sResult = "FAILED: Their Source was not found!"
                    Else : sResult = "FAILED: Your Source was not found!"
					End If
				Case eFailureReason.Player1AgentsFull
					If Me.lPlayer1ID = glPlayerID Then
						sResult = "FAILED: Will exceed your maximum agent limit"
					Else : sResult = "FAILED: Will exceed their maximum agent limit"
					End If
				Case eFailureReason.Player2AgentsFull
					If Me.lPlayer1ID = glPlayerID Then
						sResult = "FAILED: Will exceed their maximum agent limit"
					Else : sResult = "FAILED: Will exceed your maximum agent limit"
					End If
				Case Else
					sResult = "FAILED: Unknown Reason"
			End Select
        ElseIf yTradeState = eTradeStateValues.Proposal Then
            sResult &= "Proposal"
        ElseIf (yTradeState And eTradeStateValues.TradeCompleted) <> 0 Then
            sResult = "Trade Completed"
        ElseIf (yTradeState And eTradeStateValues.InNegotiation) = 0 Then
            sResult &= "In Progress, ETA: " & GetTradeETA()
        Else
            sResult &= "In Negotiations"
            If (yTradeState And eTradeStateValues.Player1Accepted) <> 0 Then
                If lPlayer1ID = glPlayerID Then
                    sResult &= ", You have accepted"
                Else : sResult &= ", They have accepted"
                End If
            End If
            If (yTradeState And eTradeStateValues.Player2Accepted) <> 0 Then
                If lPlayer1ID = glPlayerID Then
                    sResult &= ", They have accepted"
                Else : sResult &= ", You have accepted"
                End If
            End If

        End If

        Return sResult

    End Function

End Class
