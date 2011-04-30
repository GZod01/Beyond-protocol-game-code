Imports System.Net.Mail

Public Class MailObject
    Public sFrom As String
    Public sTo As String
    Public sSubject As String
    Public sBody As String
    Public sReplyTo As String

    Public lPC_ID As Int32
    Public lPlayerID As Int32
    Public iBaseAlertType As Int16
    Public bCanHaveResponse As Boolean = False
    Public lPrimaryServerIdx As Int32 = -1

    Public lExtended1 As Int32
    Public lExtended2 As Int32
    Public lExtended3 As Int32
    Public lExtended4 As Int32
    Public lExtended5 As Int32
    Public lExtended6 As Int32

	Public dtTimeStamp As Date = Date.MinValue ' Int32 = Int32.MaxValue

    Public Sub New(ByVal psFrom As String, ByVal psTo As String, ByVal psSubject As String, ByVal psBody As String, ByVal plPC_ID As Int32, ByVal plPlayerID As Int32, ByVal piBaseAlertTypeID As Int16)
        sFrom = psFrom
        sTo = psTo
        sSubject = psSubject
        sBody = psBody
        lPC_ID = plPC_ID
        lPlayerID = plPlayerID
        iBaseAlertType = piBaseAlertTypeID
    End Sub

    Public Function SendMailMsg() As Boolean
        Try
            If sTo Is Nothing OrElse sTo = "" Then Return False
            LogEvent("Sending Mail to " & sTo & " for BaseAlert: " & iBaseAlertType)

            'generate the reply email address that is unique to this message
            sReplyTo = GenerateReplyToAddress(lPC_ID)
            'default to our local from if generation failed
            If sReplyTo = "" Then sReplyTo = sFrom

            sFrom = sReplyTo

            'create the mail message object
            Dim oMailMsg As New MailMessage(sFrom.Trim(), sTo)
            oMailMsg.BodyEncoding = System.Text.Encoding.Default
            oMailMsg.Subject = sSubject.Trim()
            oMailMsg.Body = sBody.Trim() & vbCrLf
            oMailMsg.Priority = MailPriority.High
            oMailMsg.IsBodyHtml = False
            oMailMsg.ReplyTo = New MailAddress(sReplyTo)

            'create Smtpclient to send the mail message
            Dim oSmtpMail As New SmtpClient
            oSmtpMail.Host = gsOutHostName
            oSmtpMail.UseDefaultCredentials = False
            oSmtpMail.Credentials = New Net.NetworkCredential(gsEmailUserName, gsEmailPassword)

            oSmtpMail.Send(oMailMsg)

            sTo = sTo.ToUpper()

            'set keep alive time stamp for this message 
            'NOTE: expiration 1 hour once it is fully implemented
			'lTimeStamp = CInt(Val(Now.ToString("MMddHHmmss")))
			dtTimeStamp = Now

            Return True

        Catch ex As Exception
            'Message Error
            LogEvent(ex.Message)
            Return False
        End Try
    End Function

    Public Sub ProcessResponse(ByVal sData As String)
        If bCanHaveResponse = False Then Return
        bCanHaveResponse = False
        Try
            Dim sOrigData As String = sData
            sData = sData.ToUpper
            Dim lClip As Int32 = sData.IndexOf("ORIGINAL MESSAGE")
            If lClip > -1 Then
                sData = sData.Substring(0, lClip)
            End If

            Select Case iBaseAlertType
                Case GlobalMessageCode.eAddObjectCommand
                    If lExtended1 = ObjectType.ePlayerComm Then
                        sData = sOrigData
                        If sData.Length > 5000 Then sData = sData.Substring(0, 5000)
                        Dim yData(17 + sData.Length) As Byte
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.eSendOutMailMsg).CopyTo(yData, lPos) : lPos += 2
                        System.BitConverter.GetBytes(lPC_ID).CopyTo(yData, lPos) : lPos += 4
                        System.BitConverter.GetBytes(lPlayerID).CopyTo(yData, lPos) : lPos += 4
                        System.BitConverter.GetBytes(lExtended2).CopyTo(yData, lPos) : lPos += 4
                        System.BitConverter.GetBytes(sData.Length).CopyTo(yData, lPos) : lPos += 4
                        StringToBytes(sData).CopyTo(yData, lPos) : lPos += sData.Length
                        goMsgSys.SendToPrimary(lPrimaryServerIdx, yData)
                    End If
                Case GlobalMessageCode.eOutBidAlert
                    Dim lIdx As Int32 = sData.IndexOf("SET BID")
                    If lIdx <> -1 Then
                        LogEvent("  Parsed Set Bid")
                        sData = sData.Substring(lIdx).Trim
                        If sData.Length > 9 Then sData = sData.Substring(0, 9)
                        Dim lValue As Int32 = CInt(Val(sData))

                        Dim yData(13) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eSetMineralBid).CopyTo(yData, 0)
                        System.BitConverter.GetBytes(Me.lPlayerID).CopyTo(yData, 2)
                        System.BitConverter.GetBytes(Me.lExtended1).CopyTo(yData, 6)
                        System.BitConverter.GetBytes(lValue).CopyTo(yData, 10)
                        goMsgSys.SendToPrimary(lPrimaryServerIdx, yData)
                    End If
                Case GlobalMessageCode.eRebuildAISetting
                    'ok, got a response
                    Dim lIdx As Int32 = sData.IndexOf("CANCEL")
                    If lIdx <> -1 Then
                        LogEvent("  Parsed Cancel Rebuild AI")
                        Dim yData(5) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eRebuildAISetting).CopyTo(yData, 0)
                        System.BitConverter.GetBytes(lExtended1).CopyTo(yData, 2)
                        goMsgSys.SendToPrimary(lPrimaryServerIdx, yData)
                    End If

                Case GlobalMessageCode.eSetEntityProdSucceed
                    If lExtended1 = 1 Then
                        'Ok, player is responding with a set relationship msg
                        Dim lIdx As Int32 = sData.IndexOf("SET RELATIONS")
                        If lIdx <> -1 Then
                            sData = sData.Substring(lIdx).Trim

                            Dim yValue As Byte = 0
                            'Set
                            LogEvent("  Parsed Set Relations from Builder")
                            sData = sData.Substring(13).Trim
                            If sData.StartsWith("HIP") = True Then sData = sData.Substring(3).Trim
                            If sData.StartsWith("TO") = True Then sData = sData.Substring(2).Trim
                            If sData.Length > 3 Then sData = sData.Substring(0, 3)
                            Dim lValue As Int32 = CInt(Val(sData))
                            If lValue > 255 Then lValue = 255
                            If lValue < 0 Then lValue = 0
                            yValue = CByte(lValue)

                            If yValue <> 0 Then
                                Dim yData(11) As Byte
                                Dim lPos As Int32 = 0
                                System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yData, lPos) : lPos += 2
                                System.BitConverter.GetBytes(lPlayerID).CopyTo(yData, lPos) : lPos += 4
                                System.BitConverter.GetBytes(lExtended1).CopyTo(yData, lPos) : lPos += 4        'other player id
                                yData(lPos) = yValue : lPos += 1
                                goMsgSys.SendToPrimary(lPrimaryServerIdx, yData)
                            End If
                        End If
                    End If
                Case GlobalMessageCode.eColonyLowResources
                    Dim lIdx As Int32 = sData.IndexOf("GATHER")
                    Dim lTemp As Int32 = sData.IndexOf("BUILD")
                    If lIdx = -1 OrElse lIdx > lTemp Then lIdx = lTemp
                    If lIdx <> -1 Then
                        sData = sData.Substring(lIdx).Trim
                        Dim lQty As Int32
                        If sData.StartsWith("BUILD") = True Then
                            sData = sData.Substring(5).Trim
                            If sData.Length > 9 Then sData = sData.Substring(0, 9)
                            lQty = CInt(Val(sData))
                            If lQty < 1 Then lQty = 1
                        Else : lQty = -1
                        End If

                        If lQty <> 0 Then
                            If lQty = -1 Then LogEvent("  Parsed GATHER") Else LogEvent("  Parsed BUILD")
                            Dim yData(21) As Byte
                            Dim lPos As Int32 = 0
                            System.BitConverter.GetBytes(GlobalMessageCode.eColonyLowResources).CopyTo(yData, lPos) : lPos += 4
                            System.BitConverter.GetBytes(lPlayerID).CopyTo(yData, lPos) : lPos += 4
                            System.BitConverter.GetBytes(lExtended2).CopyTo(yData, lPos) : lPos += 4            'colonyid
                            System.BitConverter.GetBytes(lExtended3).CopyTo(yData, lPos) : lPos += 4            'ItemID
                            System.BitConverter.GetBytes(CShort(lExtended1)).CopyTo(yData, lPos) : lPos += 2    'ItemTypeID
                            System.BitConverter.GetBytes(lQty).CopyTo(yData, lPos) : lPos += 4
                            goMsgSys.SendToPrimary(lPrimaryServerIdx, yData)
                        End If
                    End If
                Case GlobalMessageCode.eSetPlayerRel
                    Dim lIdx As Int32 = sData.IndexOf("MATCH RELATIONS")
                    Dim lTemp As Int32 = sData.IndexOf("SET RELATIONS")
                    If lIdx = -1 OrElse lIdx > lTemp Then lIdx = lTemp
                    lTemp = sData.IndexOf("RAISE FULL INVULNERABILITY")
                    If lIdx = -1 OrElse lIdx > lTemp Then lIdx = lTemp

                    If lIdx <> -1 Then
                        sData = sData.Substring(lIdx).Trim
                        'now, is it Match or set
                        Dim yValue As Byte = 0
                        If sData.StartsWith("MATCH") = True Then
                            'match
                            LogEvent("  Parsed Match Relations")
                            yValue = CByte(lExtended2)      'original relscore
                        ElseIf sData.StartsWith("RAISE FULL INVULNERABILITY") = True Then
                            'Now, let's send our msg
                            Dim yData(44) As Byte
                            Dim lPos As Int32 = 0
                            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAlert).CopyTo(yData, lPos) : lPos += 2
                            System.BitConverter.GetBytes(lPlayerID).CopyTo(yData, lPos) : lPos += 4
                            yData(lPos) = 255 : lPos += 1
                            System.BitConverter.GetBytes(lExtended2).CopyTo(yData, lPos) : lPos += 4                'enemy id
                            System.BitConverter.GetBytes(lExtended3).CopyTo(yData, lPos) : lPos += 4                'Envir ID
                            System.BitConverter.GetBytes(CShort(lExtended4)).CopyTo(yData, lPos) : lPos += 2        'EnvirTypeID
                            System.BitConverter.GetBytes(lExtended5).CopyTo(yData, lPos) : lPos += 4                'LocX
                            System.BitConverter.GetBytes(lExtended6).CopyTo(yData, lPos) : lPos += 4                'LocZ

                            If sData.Length > 20 Then sData = sData.Substring(0, 20)
                            StringToBytes(sData).CopyTo(yData, lPos) : lPos += 20

                            goMsgSys.SendToPrimary(lPrimaryServerIdx, yData)
                            yValue = 0
                        Else
                            'Set
                            LogEvent("  Parsed Set Relations")
                            sData = sData.Substring(13).Trim
                            If sData.StartsWith("HIP") = True Then sData = sData.Substring(3).Trim
                            If sData.StartsWith("TO") = True Then sData = sData.Substring(2).Trim
                            If sData.Length > 3 Then sData = sData.Substring(0, 3)
                            Dim lValue As Int32 = CInt(Val(sData))
                            If lValue > 255 Then lValue = 255
                            If lValue < 0 Then lValue = 0
                            yValue = CByte(lValue)
                        End If

                        If yValue <> 0 Then
                            Dim yData(11) As Byte
                            Dim lPos As Int32 = 0
                            System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yData, lPos) : lPos += 2
                            System.BitConverter.GetBytes(lPlayerID).CopyTo(yData, lPos) : lPos += 4
                            System.BitConverter.GetBytes(lExtended1).CopyTo(yData, lPos) : lPos += 4        'other player id
                            yData(lPos) = yValue : lPos += 1
                            goMsgSys.SendToPrimary(lPrimaryServerIdx, yData)
                        End If
                    End If
                Case GlobalMessageCode.ePlayerAlert
                    'Attack Using... or Launch to Attack
                    Dim lIdx As Int32 = sData.IndexOf("ATTACK USING")
                    Dim lTemp As Int32 = sData.IndexOf("LAUNCH ALL UNITS TO ENGAGE")
                    If lTemp <> -1 AndAlso (lIdx = -1 OrElse lIdx > lTemp) Then lIdx = lTemp

                    lTemp = sData.IndexOf("RAISE FULL INVULNERABILITY")
                    If lTemp <> -1 AndAlso (lIdx = -1 OrElse lIdx > lTemp) Then lIdx = lTemp

                    If lIdx <> -1 Then
                        'Ok, got our position, now... get the rest
                        sData = sData.Substring(lIdx).Trim
                        'now, is it attack using or launch to attack
                        Dim yType As Byte = 0
                        If sData.StartsWith("ATTACK USING") = True Then
                            LogEvent("  Parsed " & sData)
                            'attack using
                            sData = sData.Substring(12).Trim

                            If sData.StartsWith("BATTLEGROUP") = True Then sData = sData.Substring(11).Trim
                            If sData.StartsWith(":") = True Then sData = sData.Substring(1).Trim

                            lTemp = Math.Min(sData.Length, 20)
                            sData = sData.Substring(0, lTemp)

                            If sData.StartsWith("ALL") = True Then
                                yType = 2
                            ElseIf sData.StartsWith("HALF") = True Then
                                yType = 3
                            Else : yType = 1
                            End If
                        ElseIf sData.StartsWith("RAISE FULL INVULNERABILITY") = True Then
                            LogEvent("  Parsed Full Invulnerability")
                            yType = 255
                        Else
                            'launch to attack
                            sData = "LAUNCH TO ATTACK"
                            LogEvent("  Parsed " & sData)
                            yType = 0
                        End If

                        'Now, let's send our msg
                        Dim yData(44) As Byte
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAlert).CopyTo(yData, lPos) : lPos += 2
                        System.BitConverter.GetBytes(lPlayerID).CopyTo(yData, lPos) : lPos += 4
                        yData(lPos) = yType : lPos += 1
                        System.BitConverter.GetBytes(lExtended2).CopyTo(yData, lPos) : lPos += 4                'enemy id
                        System.BitConverter.GetBytes(lExtended3).CopyTo(yData, lPos) : lPos += 4                'Envir ID
                        System.BitConverter.GetBytes(CShort(lExtended4)).CopyTo(yData, lPos) : lPos += 2        'EnvirTypeID
                        System.BitConverter.GetBytes(lExtended5).CopyTo(yData, lPos) : lPos += 4                'LocX
                        System.BitConverter.GetBytes(lExtended6).CopyTo(yData, lPos) : lPos += 4                'LocZ

                        If sData.Length > 20 Then sData = sData.Substring(0, 20)
                        StringToBytes(sData).CopyTo(yData, lPos) : lPos += 20

                        goMsgSys.SendToPrimary(lPrimaryServerIdx, yData)
                    End If
            End Select
        Catch ex As Exception
            LogEvent("ProcessResponse: " & ex.Message)
        End Try

    End Sub
End Class