Option Strict On

Public Enum eyVoteState As Byte
	eInProgress = 0
	eProcessed = 1
	eVoteFailed = 254
	eVotePassed = 255
End Enum

Public Class GuildEvent
	Public EventID As Int32			'int32 going to be enough?

	Public sTitle As String = "Requesting..."
	Public lPostedBy As Int32 = -1
	Public dtPostedOn As Date = Date.MinValue		'in GMT
	Public dtStartsAt As Date = Date.MinValue		'in GMT
	'Public dtEndsAt As Date = Date.MinValue			'in GMT
	Public lDuration As Int32 = 1					'in minutes

	Public yEventType As Byte = 0
	Public ySendAlerts As Byte = 0
	Public yEventIcon As Byte = 0

	Public yMembersCanAccept As Byte = 0

	Public yRecurrence As Byte = 0		'0 none

	Public sDetails As String = "Requesting..."

	Private mbDetailsRequested As Boolean = False

	Public Attachments() As SketchPad
	Public AttachmentUB As Int32 = -1

	Private Structure uAcceptance
		Public lPlayerID As Int32
		Public yAcceptance As Byte
	End Structure
	Private muAcceptance() As uAcceptance
	Private mlAcceptanceUB As Int32 = -1

	Private mclrList As System.Drawing.Color
	Public ReadOnly Property ListBoxColor() As System.Drawing.Color
		Get
			Return mclrList
		End Get
	End Property
	Public ReadOnly Property ListboxString() As String
		Get
			Dim sTime As String = "00:00"
			If dtStartsAt <> Date.MinValue Then
				Dim dtLocal As Date = dtStartsAt.ToLocalTime
				If dtLocal.Day = Now.Day AndAlso dtLocal.Month = Now.Month AndAlso dtLocal.Year = Now.Year Then
					sTime = dtLocal.Hour.ToString("00") & ":" & dtLocal.Minute.ToString("00")
				Else
                    sTime = dtLocal.ToString("dd/MM/yyyy HH:mm")
				End If

				Dim dValue As Double = dtLocal.Subtract(Now).TotalHours
				If dValue < 1 Then
					mclrList = System.Drawing.Color.FromArgb(255, 255, 0, 0)
				ElseIf dValue < 4 Then
					mclrList = System.Drawing.Color.FromArgb(255, 255, 255, 0)
				Else
					mclrList = muSettings.InterfaceBorderColor
				End If
			End If

			sTime = "(" & sTime & ")"
			If yRecurrence <> 0 Then
				Select Case yRecurrence
					Case 1
						sTime &= " - Recurs Daily - "
					Case 2
						sTime &= " - Recurs Weekly - "
					Case 3
						sTime &= " - Recurs Monthly - "
					Case 4
						sTime &= " - Recurs Annually - "
				End Select
			End If
			Dim sFinal As String = sTime & sTitle
			Return sFinal
		End Get
	End Property

	Public ReadOnly Property rcIconRect() As Rectangle
		Get
			Select Case yEventIcon
				Case 0
					Return New Rectangle(0, 0, 16, 16)
				Case 1
					Return New Rectangle(16, 0, 16, 16)
				Case 2
					Return New Rectangle(32, 0, 16, 16)
				Case 3
					Return New Rectangle(48, 0, 16, 16)

				Case 4
					Return New Rectangle(0, 16, 16, 16)
				Case 5
					Return New Rectangle(16, 16, 16, 16)
				Case 6
					Return New Rectangle(32, 16, 16, 16)
				Case 7
					Return New Rectangle(48, 16, 16, 16)

				Case 8
					Return New Rectangle(0, 32, 16, 16)
				Case 9
					Return New Rectangle(16, 32, 16, 16)
				Case 10
					Return New Rectangle(32, 32, 16, 16)
				Case 11
					Return New Rectangle(48, 32, 16, 16)

				Case 12
					Return New Rectangle(0, 48, 16, 16)
				Case 13
					Return New Rectangle(16, 48, 16, 16)
				Case 14
					Return New Rectangle(32, 48, 16, 16)
				Case 15
					Return New Rectangle(48, 48, 16, 16)
			End Select
		End Get
	End Property

	Public Function FillFromSmallMsg(ByRef yData() As Byte, ByVal lPos As Int32) As Int32
		EventID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		yEventIcon = yData(lPos) : lPos += 1
		Dim lDate As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		dtStartsAt = GetDateFromNumber(lDate)

		Return lPos
	End Function

	Public Function RequestDetails() As Byte()
		If mbDetailsRequested = False Then
			mbDetailsRequested = True
			'Ok request details
			Dim yMsg(6) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eGuildRequestDetails).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(Me.EventID).CopyTo(yMsg, lPos) : lPos += 4
			yMsg(lPos) = eyGuildRequestDetailsType.EventDetails : lPos += 1
			Return yMsg
		End If
		Return Nothing
	End Function

	Public Sub SetPlayerAcceptance(ByVal lPlayerID As Int32, ByVal yAcceptance As Byte)
		If muAcceptance Is Nothing Then ReDim muAcceptance(-1)
		For X As Int32 = 0 To mlAcceptanceUB
			If muAcceptance(X).lPlayerID = lPlayerID Then
				muAcceptance(X).yAcceptance = yAcceptance
				Return
			End If
		Next X

		SyncLock muAcceptance
			ReDim Preserve muAcceptance(mlAcceptanceUB + 1)
			mlAcceptanceUB += 1
			muAcceptance(mlAcceptanceUB).lPlayerID = lPlayerID
			muAcceptance(mlAcceptanceUB).yAcceptance = yAcceptance
		End SyncLock
	End Sub

	Public Function GetAttachment(ByVal lID As Int32) As SketchPad
		For X As Int32 = 0 To AttachmentUB
			If Attachments(X) Is Nothing = False AndAlso Attachments(X).lID = lID Then Return Attachments(X)
		Next X
		Return Nothing
	End Function

End Class
