Option Strict On

''' <summary>
''' Contains the methods and data for monitoring message flow between this application and other applications
''' </summary>
''' <remarks></remarks>
Public Class MsgMonitor
    Public Enum eMM_AppType As Integer
        PrimaryServer = 0
        RegionServer
        PathfindingServer
        ClientConnection
        OperatorServer
    End Enum

    Private Structure MM_Conn
        Public lType As eMM_AppType
        Public lSpecificID As Int32

        Public blOutMsgs() As Int64
        Public blInMsgs() As Int64

        'Public Function SaveObject(ByVal lDate As Int32) As Boolean
        '    Dim bResult As Boolean = False
        '    Dim sSQL As String
        '    Dim oComm As OleDb.OleDbCommand

        '    Dim lID As Int32 = -1

        '    Try

        '        sSQL = "INSERT INTO tblMsgMonitor (AppType, SpecificID, MsgMonDate) VALUES (" & CInt(Me.lType) & ", " & _
        '          Me.lSpecificID & ", " & lDate & ")"

        '        oComm = New OleDb.OleDbCommand(sSQL, goCN)
        '        If oComm.ExecuteNonQuery() = 0 Then
        '            Err.Raise(-1, "SaveObject", "No records affected when saving object!")
        '        End If

        '        Dim oData As OleDb.OleDbDataReader
        '        oComm = Nothing
        '        sSQL = "SELECT MAX(MM_ID) FROM tblMsgMonitor WHERE SpecificID = " & lSpecificID
        '        oComm = New OleDb.OleDbCommand(sSQL, goCN)
        '        oData = oComm.ExecuteReader(CommandBehavior.Default)
        '        If oData.Read Then
        '            lID = CInt(oData(0))
        '        End If
        '        oData.Close()
        '        oData = Nothing
        '        oComm = Nothing

        '        bResult = True
        '    Catch
        '        LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to MsgMonitor Object (" & Me.lType.ToString & ", " & Me.lSpecificID & "). Reason: " & Err.Description)
        '    Finally
        '        oComm = Nothing
        '    End Try

        '    If lID = -1 Then Return False

        '    If lOutMsgs Is Nothing = False Then
        '        For X As Int32 = 0 To lOutMsgs.GetUpperBound(0)
        '            If lOutMsgs(X) <> 0 Then
        '                Try
        '                    sSQL = "INSERT INTO tblMsgMonitorMsg (MM_ID, CodeID, IsOutbound, TotalBytes) VALUES (" & lID & _
        '                           ", " & X & ", 1, " & lOutMsgs(X) & ")"
        '                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
        '                    If oComm.ExecuteNonQuery() = 0 Then Err.Raise(-1, "SaveObject", "No records affected when saving monitor messages.")
        '                Catch
        '                    LogEvent(LogEventType.CriticalError, "Unable to save MsgMonitorMsg object (" & Me.lType.ToString & ", " & Me.lSpecificID & ", Out: " & X & ". Reason: " & Err.Description)
        '                Finally
        '                    oComm = Nothing
        '                End Try
        '            End If
        '        Next X
        '    End If
        '    If lInMsgs Is Nothing = False Then
        '        For X As Int32 = 0 To lInMsgs.GetUpperBound(0)
        '            If lInMsgs(X) <> 0 Then
        '                Try
        '                    sSQL = "INSERT INTO tblMsgMonitorMsg (MM_ID, CodeID, IsOutbound, TotalBytes) VALUES (" & lID & _
        '                       ", " & X & ", 0, " & lInMsgs(X) & ")"
        '                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
        '                    If oComm.ExecuteNonQuery() = 0 Then Err.Raise(-1, "SaveObject", "No records affected when saving monitor messages.")
        '                Catch
        '                    LogEvent(LogEventType.CriticalError, "Unable to save MsgMonitorMsg object (" & Me.lType.ToString & ", " & Me.lSpecificID & ", In: " & X & ". Reason: " & Err.Description)
        '                Finally
        '                    oComm = Nothing
        '                End Try
        '            End If
        '        Next X
        '    End If

        '    Return bResult
        'End Function
    End Structure

    Private muConns() As MM_Conn
    Private mlConnUB As Int32 = -1

    ''' <summary>
    ''' Records an outgoing message from this application to recipient
    ''' </summary>
    ''' <param name="iMsgCode"> The EpicaMessageCode enum value for the message </param>
    ''' <param name="lRecipientType"> The application type of the recipient </param>
    ''' <param name="lMsgSize"> The total size of the message in bytes being sent </param>
    ''' <param name="lSpecificID"> The specific ID of the recipient. For example, if lRecipientType is Client, then this would be the PlayerID </param>
    ''' <remarks></remarks>
    Public Sub AddOutMsg(ByVal iMsgCode As Int16, ByVal lRecipientType As eMM_AppType, ByVal lMsgSize As Int32, ByVal lSpecificID As Int32)
        Dim lIdx As Int32 = GetConnIndex(lRecipientType, lSpecificID)
        If lIdx = -1 Then Return

        With muConns(lIdx)
			If .blOutMsgs Is Nothing Then ReDim .blOutMsgs(GlobalMessageCode.eLastMsgCode)
			.blOutMsgs(iMsgCode) += lMsgSize
		End With
	End Sub

	''' <summary>
	''' Records an incoming message from a sender to this application
	''' </summary>
	''' <param name="iMsgCode"> The EpicaMessageCode enum value for the message </param>
	''' <param name="lSenderType"> The application type of the recipient </param>
	''' <param name="lMsgSize"> The total size of the message in bytes being received </param>
	''' <param name="lSpecificID"> The specific ID of the sender. For example, if lSenderType is Client, then this would be the PlayerID </param>
	''' <remarks></remarks>
	Public Sub AddInMsg(ByVal iMsgCode As Int16, ByVal lSenderType As eMM_AppType, ByVal lMsgSize As Int32, ByVal lSpecificID As Int32)
		Dim lIdx As Int32 = GetConnIndex(lSenderType, lSpecificID)
		If lIdx = -1 Then Return

		With muConns(lIdx)
			If .blInMsgs Is Nothing Then ReDim .blInMsgs(GlobalMessageCode.eLastMsgCode)
			.blInMsgs(iMsgCode) += lMsgSize
		End With
	End Sub

	Private Function GetConnIndex(ByVal lType As eMM_AppType, ByVal lSpecificID As Int32) As Int32
		Dim lIdx As Int32 = -1

		If muConns Is Nothing Then ReDim muConns(-1)
		Dim lMinUB As Int32 = Math.Min(muConns.GetUpperBound(0), mlConnUB)

		Try
			For X As Int32 = 0 To lMinUB 'mlConnUB

				If muConns(X).lType = lType AndAlso muConns(X).lSpecificID = lSpecificID Then
					lIdx = X
					Exit For
				End If
			Next X
		Catch
		End Try

		If lIdx = -1 Then
			mlConnUB += 1
			ReDim Preserve muConns(mlConnUB)
			lIdx = mlConnUB
			With muConns(lIdx)
				.lSpecificID = lSpecificID
				.lType = lType
			End With
		End If
		Return lIdx
	End Function

	Public Sub SaveAll()
		'Dim lDate As Int32 = CInt(Val(Now.ToString("yyMMddHHmm")))

		For X As Int32 = 0 To mlConnUB
			Dim yMsg() As Byte
			Dim lPos As Int32 = 0

			With muConns(X)
				If .blInMsgs Is Nothing AndAlso .blOutMsgs Is Nothing Then Continue For

				ReDim yMsg(9 + (2 * (GlobalMessageCode.eLastMsgCode * 8)))

				System.BitConverter.GetBytes(GlobalMessageCode.eMsgMonitorData).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(CInt(.lType)).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.lSpecificID).CopyTo(yMsg, lPos) : lPos += 4

				Dim lFinalPos As Int32 = lPos + (GlobalMessageCode.eLastMsgCode * 8)
				If .blOutMsgs Is Nothing = False Then
					For Y As Int32 = 0 To GlobalMessageCode.eLastMsgCode - 1
						If .blOutMsgs.GetUpperBound(0) >= Y Then
							System.BitConverter.GetBytes(.blOutMsgs(Y)).CopyTo(yMsg, lPos) : lPos += 8
						Else : Exit For
						End If
					Next Y
				End If
				lPos = lFinalPos

				lFinalPos = lPos + (GlobalMessageCode.eLastMsgCode * 8)
				If .blInMsgs Is Nothing = False Then
					For Y As Int32 = 0 To GlobalMessageCode.eLastMsgCode - 1
						If .blInMsgs.GetUpperBound(0) >= Y Then
							System.BitConverter.GetBytes(.blInMsgs(Y)).CopyTo(yMsg, lPos) : lPos += 8
						Else : Exit For
						End If
					Next
				End If
				lPos = lFinalPos

				goMsgSys.SendToPrimary(yMsg)
			End With
		Next X
	End Sub

End Class
