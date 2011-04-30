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
        EmailServer
        OperatorServer
    End Enum

    Private Structure MM_Conn
        Public lSrcAppType As eMM_AppType
        Public lSrcAppID As Int32

        Public lType As eMM_AppType
        Public lSpecificID As Int32

        Public blOutMsgs() As Int64
        Public blInMsgs() As Int64

        Public Function SaveObject(ByVal lDate As Int32) As Boolean
            Dim bResult As Boolean = False
            Dim sSQL As String
            Dim oComm As OleDb.OleDbCommand

            Dim lID As Int32 = -1

            Try

                sSQL = "INSERT INTO tblMsgMonitor (AppType, SpecificID, MsgMonDate, SrcAppType, SrcAppID) VALUES (" & _
                  CInt(Me.lType) & ", " & Me.lSpecificID & ", " & lDate & ", " & CInt(lSrcAppType) & ", " & lSrcAppID & ")"

                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If

                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(MM_ID) FROM tblMsgMonitor WHERE SpecificID = " & lSpecificID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    lID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
                oComm = Nothing

                bResult = True
            Catch
                LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to MsgMonitor Object (" & Me.lType.ToString & ", " & Me.lSpecificID & "). Reason: " & Err.Description)
            Finally
                oComm = Nothing
            End Try

            If lID = -1 Then Return False

            If blOutMsgs Is Nothing = False Then
                For X As Int32 = 0 To blOutMsgs.GetUpperBound(0)
                    If blOutMsgs(X) <> 0 Then
                        Try
                            sSQL = "INSERT INTO tblMsgMonitorMsg (MM_ID, CodeID, IsOutbound, TotalBytes) VALUES (" & lID & _
                                   ", " & X & ", 1, " & blOutMsgs(X) & ")"
                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                            If oComm.ExecuteNonQuery() = 0 Then Err.Raise(-1, "SaveObject", "No records affected when saving monitor messages.")
                        Catch
                            LogEvent(LogEventType.CriticalError, "Unable to save MsgMonitorMsg object (" & Me.lType.ToString & ", " & Me.lSpecificID & ", Out: " & X & ". Reason: " & Err.Description)
                        Finally
                            oComm = Nothing
                        End Try
                    End If
                Next X
            End If
            If blInMsgs Is Nothing = False Then
                For X As Int32 = 0 To blInMsgs.GetUpperBound(0)
                    If blInMsgs(X) <> 0 Then
                        Try
                            sSQL = "INSERT INTO tblMsgMonitorMsg (MM_ID, CodeID, IsOutbound, TotalBytes) VALUES (" & lID & _
                               ", " & X & ", 0, " & blInMsgs(X) & ")"
                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                            If oComm.ExecuteNonQuery() = 0 Then Err.Raise(-1, "SaveObject", "No records affected when saving monitor messages.")
                        Catch
                            LogEvent(LogEventType.CriticalError, "Unable to save MsgMonitorMsg object (" & Me.lType.ToString & ", " & Me.lSpecificID & ", In: " & X & ". Reason: " & Err.Description)
                        Finally
                            oComm = Nothing
                        End Try
                    End If
                Next X
            End If

            Return bResult
        End Function
    End Structure

    Private muConns() As MM_Conn
    Private mlConnUB As Int32 = -1

    Private Shared mlExternalSrcDate As Int32 = Int32.MinValue

    ''' <summary>
    ''' Records an outgoing message from this application to recipient
    ''' </summary>
    ''' <param name="iMsgCode"> The EpicaMessageCode enum value for the message </param>
    ''' <param name="lRecipientType"> The application type of the recipient </param>
    ''' <param name="lMsgSize"> The total size of the message in bytes being sent </param>
    ''' <param name="lSpecificID"> The specific ID of the recipient. For example, if lRecipientType is Client, then this would be the PlayerID </param>
    ''' <remarks></remarks>
    Public Sub AddOutMsg(ByVal iMsgCode As Int16, ByVal lRecipientType As eMM_AppType, ByVal lMsgSize As Int32, ByVal lSpecificID As Int32)
        Try
            Dim lIdx As Int32 = GetConnIndex(lRecipientType, lSpecificID)
            If lIdx = -1 Then Return

            With muConns(lIdx)
                If .blOutMsgs Is Nothing Then ReDim .blOutMsgs(GlobalMessageCode.eLastMsgCode)
                .blOutMsgs(iMsgCode) += lMsgSize
            End With
        Catch
        End Try
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
        Try
            Dim lIdx As Int32 = GetConnIndex(lSenderType, lSpecificID)
            If lIdx = -1 Then Return

            With muConns(lIdx)
                If .blInMsgs Is Nothing Then ReDim .blInMsgs(GlobalMessageCode.eLastMsgCode)
                .blInMsgs(iMsgCode) += lMsgSize
            End With
        Catch
        End Try
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

				'Set it up to use me as the source
				.lSrcAppType = eMM_AppType.PrimaryServer
				.lSrcAppID = 0
			End With
		End If
		Return lIdx
	End Function

	Public Sub SaveAll()
		Dim lDate As Int32 = GetDateAsNumber(Now)
		For X As Int32 = 0 To mlConnUB
			muConns(X).SaveObject(lDate)
		Next X
	End Sub

	Public Sub FillListView(ByRef tvwData As TreeView)
		tvwData.Nodes.Clear()

		For X As Int32 = 0 To mlConnUB
			Dim oNode As TreeNode = Nothing

			With muConns(X)

				Dim blTotalIn As Int64 = 0
				Dim blTotalOut As Int64 = 0

				If .blInMsgs Is Nothing = False Then
					For Y As Int32 = 0 To .blInMsgs.GetUpperBound(0)
						blTotalIn += .blInMsgs(Y)
					Next Y
				End If
				If .blOutMsgs Is Nothing = False Then
					For Y As Int32 = 0 To .blOutMsgs.GetUpperBound(0)
						blTotalOut += .blOutMsgs(Y)
					Next Y
				End If

				Select Case .lType
					Case eMM_AppType.ClientConnection
						oNode = tvwData.Nodes.Add("Client " & .lSpecificID & " (Total In: " & blTotalIn.ToString("#,##0") & ", Total Out: " & blTotalOut.ToString("#,##0") & ")")
					Case eMM_AppType.PathfindingServer
						oNode = tvwData.Nodes.Add("Pathfinding (Total In: " & blTotalIn.ToString("#,##0") & ", Total Out: " & blTotalOut.ToString("#,##0") & ")")
					Case eMM_AppType.PrimaryServer
						oNode = tvwData.Nodes.Add("Primary (Total In: " & blTotalIn.ToString("#,##0") & ", Total Out: " & blTotalOut.ToString("#,##0") & ")")
					Case eMM_AppType.RegionServer
						oNode = tvwData.Nodes.Add("Region " & .lSpecificID & " (Total In: " & blTotalIn.ToString("#,##0") & ", Total Out: " & blTotalOut.ToString("#,##0") & ")")
				End Select

				If oNode Is Nothing = False Then
					For Y As Int32 = 0 To GlobalMessageCode.eLastMsgCode - 1
						Dim sValue As String = ""

						Dim blTotalSize As Int64 = 0
						Dim blInSize As Int64 = 0
						Dim blOutSize As Int64 = 0
						If .blInMsgs Is Nothing = False AndAlso .blInMsgs.GetUpperBound(0) >= Y Then
							blInSize = .blInMsgs(Y)
							blTotalSize += blInSize
						End If
						If .blOutMsgs Is Nothing = False AndAlso .blOutMsgs.GetUpperBound(0) >= Y Then
							blOutSize = .blOutMsgs(Y)
							blTotalSize += blOutSize
						End If

						If blTotalSize <> 0 Then
							sValue = CType(Y, GlobalMessageCode).ToString & ": " & blTotalSize.ToString("#,##0")
							Dim oTmpNode As TreeNode = oNode.Nodes.Add(sValue)
							oTmpNode.Nodes.Add("Incoming Msgs: " & blInSize.ToString("#,##0"))
							oTmpNode.Nodes.Add("Outgoing Msgs: " & blOutSize.ToString("#,##0"))
						End If
					Next Y
				End If
			End With

		Next X
	End Sub

    Public Shared Sub AddExternalMsgMonData(ByVal lSrcType As eMM_AppType, ByVal lSrcID As Int32, ByVal lConnType As eMM_AppType, ByVal lSpecificID As Int32, ByRef blOuts() As Int64, ByRef blIns() As Int64)
        Dim uTmpConn As MM_Conn
        With uTmpConn
            .lSrcAppID = lSrcID
            .lSrcAppType = lSrcType
            .lType = lConnType
            .lSpecificID = lSpecificID
            .blInMsgs = blIns
            .blOutMsgs = blOuts

            If mlExternalSrcDate = Int32.MinValue Then mlExternalSrcDate = GetDateAsNumber(Now)

            .SaveObject(mlExternalSrcDate)
        End With
    End Sub
End Class
