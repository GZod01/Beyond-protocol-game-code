Option Strict On

Public Class AllDataSaver
	Private moComm As OleDb.OleDbCommand
	Private moPrepare As OleDb.OleDbCommand

	Private moCmdText As System.Text.StringBuilder

	Public Sub InitializeCommand()
		moComm = New OleDb.OleDbCommand()
		moComm.Connection = goCN
		moPrepare = New OleDb.OleDbCommand()
		moPrepare.Connection = goCN

		moCmdText = New System.Text.StringBuilder()
	End Sub

	Public Function AddCommandText(ByVal sText As String) As Boolean
		If sText Is Nothing OrElse sText = "" Then Return True
		Dim bResult As Boolean = False
		Try
			moPrepare.CommandText = sText
			moPrepare.Prepare()
			bResult = True
			moCmdText.AppendLine(sText)
		Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "AllDataSaver.AddCommandText: " & GetEntireErrorMsg(ex)) 'ex.Message)
			bResult = False
		End Try
		Return bResult
	End Function

	Public Function ExecuteFinalCommand() As Boolean
		Dim bResult As Boolean = False
		Try
			moComm.CommandText = moCmdText.ToString
			moComm.ExecuteNonQuery()
			moCmdText.Length = 0
			bResult = True
		Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "AllDataSaver.ExecuteFinalCommand: " & GetEntireErrorMsg(ex)) 'ex.Message)
			bResult = False
		End Try
		Return bResult
	End Function

	Public Sub DisposeMe()
		If moCmdText Is Nothing = False Then moCmdText.Length = 0
		moCmdText = Nothing
		If moPrepare Is Nothing = False Then moPrepare.Dispose()
		moPrepare = Nothing
		If moComm Is Nothing = False Then moComm.Dispose()
		moComm = Nothing
	End Sub

End Class