'Imports System
'Imports System.IO
'Imports System.Web
'Imports System.Net
'Imports System.Net.Sockets
'Imports System.Diagnostics
'Imports System.Text
'Imports System.Text.RegularExpressions
'Imports Microsoft.VisualBasic.Strings

'Public Module EmailValidation

'	'NOTE: This module should be run from a Public Server. It will usually fail on 
'	'private systems, since most mail servers will not respond to private systems
'	'or systems within their own domain. For example, My system is running through 
'	'Road Runner [cfl.rr.com]. If I run IsValidEmailAddress("kathrock@cfl.rr.com")
'	'from my PC, I get the following error when I attempt a connection: 
'	'"550-cdptpa-mxlb.mail.rr.com 550 ERROR: Mail Refused - res.rr.com - See http://security.rr.com/rrDynamic.htm". 
'	'That page states that I cannot do this from inside "Road Runner dynamic space".

'#Region " Enumerators "
'	Public Enum EmailValidationResults As Int32
'		''' <summary>This is a Valid Email Address</summary>
'		VALID_EMAIL_ADDRESS = 0
'		''' <summary>Email Address Syntax is Incorrect (Missing "@", "." or Both)</summary>
'		INVALID_EMAIL_ADDRESS_SYNTAX = 1
'		''' <summary>The Domain Name part of the Email Address is not a 
'		''' valid Domain Name (MX Record Not Found).</summary>
'		INVALID_DOMAIN_NAME = 2
'		''' <summary>The Email's Domain Server Refused our Connection 
'		''' or was Unavailable and could not be reached.</summary>
'		MAIL_SERVER_REFUSED_CONNECTION = 3
'		''' <summary>The Email's Domain Server did Not Acknowledge the 
'		''' "HELO" Handshake from our Server.</summary>
'		HANDSHAKE_NOT_ACKNOWLEDGED = 4
'		''' <summary>The Email's Domain Server did Not Acknowledge our Domain as 
'		''' Valid or they may have our Domain Blacklisted.</summary>
'		OUR_DOMAIN_NOT_ACKNOWLEDGED = 5
'		''' <summary>The Email's Domain Server did Not Acknowledge that this Recipient has a 
'		''' MailBox there. This may be because that domain doesn't give out this information.</summary>
'		RECIPIENT_NOT_ACKNOWLEDGED = 6
'		''' <summary>An Error Occurred during the Communication with the Email's Domain Server.</summary>
'		ERROR_OCCURRED = 7
'	End Enum
'	Public Enum EmailSuccessResults As Int32
'		''' <summary>The Function Failed on the First Test.</summary>
'		NO_SUCCESS = 0
'		''' <summary>The Function Passed the Email Address Syntax Test (Test #1).</summary>
'		EMAIL_ADDRESS_SYNTAX = 1
'		''' <summary>The Function Passed the Domain Name Test (Test #2).</summary>
'		DOMAIN_NAME = 2
'		''' <summary>The Function was able to Connect to the Email's 
'		''' Domain Server (Test #3).</summary>
'		MAIL_SERVER_CONNECTED = 3
'		''' <summary>The Function was able to perform a Handshake with 
'		''' the Email's Domain Server (Test #4).</summary>
'		HANDSHAKE_ACKNOWLEDGED = 4
'		''' <summary>The Function was able to Identify our Domain with 
'		''' the Email's Domain Server (Test #5).</summary>
'		OUR_DOMAIN_ACKNOWLEDGED = 5
'		''' <summary>The Function was able to Verify that the Recipient 
'		''' has a MailBox on the Email's Domain Server (Test #6).</summary>
'		RECIPIENT_ACKNOWLEDGED = 6
'	End Enum
'#End Region

'#Region " IsValidEmailAddress() Function "
'	''' <summary>Checks an Email Address by Syntax and then by sending Requests to the
'	''' Email's Domain Server returning one of the EmailValidationResults Enumerators.</summary>
'	''' <param name="EmailFromAddress">The Sending Email Address</param>
'	''' <param name="EmailAddressToValidate">The Email Address to Validate</param>
'	''' <param name="ReturnErrorMessage">Optional - Returns (ByRef) the Error Message or the Response 
'	''' received from the Email Domain Server if anything Fails</param>
'	''' <param name="ReturnLastSuccess">Optional - Returns (ByRef) the Last Successful Operation  
'	''' as EmailSuccessResults if anything Fails</param>
'	''' <returns>One of the EmailValidationResults Enumerators</returns>
'	Public Function IsValidEmailAddress(ByVal EmailFromAddress As String, ByVal EmailAddressToValidate As String, Optional ByRef ReturnErrorMessage As String = "", Optional ByRef ReturnLastSuccess As EmailSuccessResults = EmailSuccessResults.NO_SUCCESS) As EmailValidationResults
'		Dim oClient As TcpClient = Nothing, oStream As NetworkStream = Nothing, sResponse As String = ""
'		Dim sEmail As String = Trim(EmailAddressToValidate & "") & "", sFrom As String = "", sTo As String = ""
'		Dim sOurDomain As String, eResult As EmailValidationResults = EmailValidationResults.VALID_EMAIL_ADDRESS
'		Dim sServer As String = "", sDomain As String = "", saParts() As String = Nothing
'		Dim iAt As Integer = sEmail.IndexOf("@"), iDot As Integer = sEmail.LastIndexOf(".")

'		'Check the Syntax
'		If (iAt < 0) OrElse (iDot <= iAt) Then
'			eResult = EmailValidationResults.INVALID_EMAIL_ADDRESS_SYNTAX	'Invalid Email Address Syntax
'			GoTo Cleanup
'		Else
'			ReturnLastSuccess = EmailSuccessResults.EMAIL_ADDRESS_SYNTAX
'		End If

'		'Check the Domain
'		sTo = "<" & sEmail & ">"
'		saParts = Split(sEmail, "@")
'		sDomain = saParts(saParts.Length - 1)
'		sServer = NSLookup(sDomain)
'		If sServer.Length = 0 Then
'			eResult = EmailValidationResults.INVALID_DOMAIN_NAME	'Invalid Domain Name
'			GoTo Cleanup
'		Else
'			ReturnLastSuccess = EmailSuccessResults.DOMAIN_NAME
'		End If

'		Try
'			'Remote Address for the Mail Server is our Domain.
'			sFrom = Trim(EmailFromAddress & "") & ""
'			saParts = Split(sFrom, "@")
'			sOurDomain = saParts(saParts.Length - 1)
'			sFrom = "<" & sFrom & ">" 'Add this line...
'			'Create the TCP Client Connection on Port 25
'			oClient = New TcpClient()
'			oClient.SendTimeout = 3000
'			oClient.Connect(sServer, 25)
'			oStream = oClient.GetStream()
'			'Read in the Connection Response
'			sResponse = GetData(oStream)
'			If Not ValidResponse(sResponse) Then
'				eResult = EmailValidationResults.MAIL_SERVER_REFUSED_CONNECTION
'				GoTo Cleanup
'			Else
'				ReturnLastSuccess = EmailSuccessResults.MAIL_SERVER_CONNECTED
'			End If
'			'Attempt a Handshake with the Mail Server
'			sResponse = SendData(oStream, "HELO " & sOurDomain & vbCrLf)
'			If Not ValidResponse(sResponse) Then
'				eResult = EmailValidationResults.HANDSHAKE_NOT_ACKNOWLEDGED
'				GoTo Cleanup
'			Else
'				ReturnLastSuccess = EmailSuccessResults.HANDSHAKE_ACKNOWLEDGED
'			End If
'			'Check for possible blacklisting of our domain.
'			sResponse = SendData(oStream, "MAIL FROM: " & sFrom & vbCrLf)
'			If Not ValidResponse(sResponse) Then
'				eResult = EmailValidationResults.OUR_DOMAIN_NOT_ACKNOWLEDGED
'				GoTo Cleanup
'			Else
'				ReturnLastSuccess = EmailSuccessResults.OUR_DOMAIN_ACKNOWLEDGED
'			End If
'			'Check to see if the recipients mailbox exists.
'			sResponse = SendData(oStream, "RCPT TO: " & sTo & vbCrLf)
'			If Not ValidResponse(sResponse) Then
'				eResult = EmailValidationResults.RECIPIENT_NOT_ACKNOWLEDGED
'				GoTo Cleanup
'			Else
'				ReturnLastSuccess = EmailSuccessResults.RECIPIENT_ACKNOWLEDGED
'			End If
'			sResponse = ""
'		Catch ex As Exception
'			sResponse = "Error Occurred: " & ex.Message
'			eResult = EmailValidationResults.ERROR_OCCURRED
'		End Try

'Cleanup:
'		If sResponse.Length > 0 AndAlso sResponse.IndexOf("Error Occurred:") < 0 Then
'			sResponse = "Server Response: " & sResponse
'		End If
'		ReturnErrorMessage = sResponse
'		Try : SendData(oStream, "QUIT" & vbCrLf) : Catch ex As Exception : End Try
'		Try : oClient.Close() : Catch ex As Exception : End Try
'		Try : oStream.Dispose() : Catch ex As Exception : End Try
'		oStream = Nothing
'		Return eResult
'	End Function
'#End Region

'#Region " Utility Functions "
'	''' <summary>Reads the Response Data from a Network Stream object.</summary>
'	''' <param name="Stream">A Connected Network Stream object</param>
'	''' <returns>The Response Data from the Stream</returns>
'	''' <remarks>The Stream Must be Connected to a Network</remarks>
'	Private Function GetData(ByRef Stream As NetworkStream) As String
'		Dim yaData(1024) As Byte, iLen As Integer = Stream.Read(yaData, 0, 1024)
'		If iLen > 0 Then Return Encoding.ASCII.GetString(yaData, 0, iLen)
'		Return ""
'	End Function

'	''' <summary>Sends Data using a Network Stream object and returns the Response.</summary>
'	''' <param name="Stream">A Connected Network Stream object</param>
'	''' <param name="DataToSend">The Data to Send</param>
'	''' <returns>The Response from the Data Sent</returns>
'	''' <remarks>The Stream Must be Connected to a Network</remarks>
'	Private Function SendData(ByRef Stream As NetworkStream, ByVal DataToSend As String) As String
'		'Convert Data to ASCII Byte Array.
'		Dim yaData() As Byte = Encoding.ASCII.GetBytes(DataToSend.ToCharArray)
'		'Send the Data by Writing to the Network Stream.
'		Stream.Write(yaData, 0, yaData.Length())
'		'Return the Results (Written back to the Stream)
'		Return GetData(Stream)
'	End Function

'	''' <summary>Tests for a valid response from a Network Stream</summary>
'	''' <param name="Response">A response from a Network Stream</param>
'	''' <returns>Boolean Response Valid or Error</returns>
'	''' <remarks>Valid Respons Codes are less than 300. Anything >= 300 is an Error</remarks>
'	Private Function ValidResponse(ByVal Response As String) As Boolean
'		'Test first digit of response code (<3 is Valid; >=3 is an Error)
'		If Response.Length > 1 Then Return CBool(CType(Response.Substring(0, 1), Integer) < 3)
'		Return False
'	End Function

'	''' <summary>Returns the Server's MX Domain Name</summary>
'	''' <param name="sDomain">Any Valid Domain Name (i.e., darkskyentertainment.com)</param>
'	''' <returns>The Server's MX Domain Name</returns>
'	''' <remarks>Uses NSLookup and Regex to find the Server's MX record.</remarks>
'	Private Function NSLookup(ByVal sDomain As String) As String
'		Dim ProcInfo As ProcessStartInfo = New ProcessStartInfo(), Proc As Process = Nothing
'		Dim oReader As StreamReader = Nothing, sServer As String = ""
'		Dim oRegExp As Regex = New Regex("mail exchanger = (?<server>[^\\\s]+)")
'		Dim sLine As String = "", Match As Match = Nothing
'		ProcInfo.UseShellExecute = False
'		ProcInfo.RedirectStandardInput = True
'		ProcInfo.RedirectStandardOutput = True
'		ProcInfo.FileName = "nslookup"
'		ProcInfo.Arguments = "-type=MX " + (Trim(sDomain & "") & "").ToUpper
'		Proc = Process.Start(ProcInfo)
'		oReader = Proc.StandardOutput
'		Do While (oReader.Peek() >= 0)
'			sLine = oReader.ReadLine()
'			Match = oRegExp.Match(sLine)
'			If (Match.Success) Then
'				sServer = Match.Groups("server").Value
'				Exit Do
'			End If
'		Loop
'		Return sServer
'	End Function
'#End Region

'End Module
