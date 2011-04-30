Option Strict On

Partial Class MsgSystem
#Region "  Actual Action Functions For Pass Thru Msgs  "
    Private Function DoAddPlayerCommFolder(ByRef oPlayer As Player, ByVal yData() As Byte, ByVal lMsgStartPos As Int32) As Byte()
        Dim lPos As Int32 = lMsgStartPos + 2
        Dim lPCF_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim yTemp(19) As Byte
        Array.Copy(yData, lPos, yTemp, 0, 20)

        Dim lLen As Int32 = yTemp.Length
        For X As Int32 = 0 To yTemp.Length - 1
            If yTemp(X) = 0 Then
                lLen = X
                Exit For
            End If
        Next X
        Dim sName As String = Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)
        Dim lIdx As Int32 = oPlayer.AddFolder(-1, sName)

        If lIdx <> -1 Then
            Return oPlayer.EmailFolders(lIdx).GetAddFolderMsg
        End If
        Return Nothing
    End Function
    Private Function DoDeleteEmailItem(ByRef oPlayer As Player, ByVal yData() As Byte, ByVal lMsgStartPos As Int32) As Byte()
        Dim lPos As Int32 = lMsgStartPos + 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If iObjTypeID = -1 Then
            'Ok, deleting Folder
            If oPlayer.DeleteFolder(lObjID) = True Then
                Dim yResp(7) As Byte
                Array.Copy(yData, lMsgStartPos, yResp, 0, 8)
                Return yResp
            End If
        Else
            'Ok, deleting email
            Dim lPCF_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            oPlayer.RemoveEmailMsg(lObjID, lPCF_ID)
            Dim yResp(11) As Byte
            Array.Copy(yData, lMsgStartPos, yResp, 0, 12)
            Return yResp
        End If
        Return Nothing
    End Function
    Private Sub DoDeleteTradeHistoryItem(ByRef oPlayer As Player, ByVal yData() As Byte, ByVal lMsgStartPos As Int32)
        Dim lPos As Int32 = lMsgStartPos + 2    'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTransDate As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lOtherPlayer As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yResult As Byte = yData(lPos) : lPos += 1
        Dim yEventType As Byte = yData(lPos) : lPos += 1

        oPlayer.DeleteTradeHistoryItem(lTransDate, lOtherPlayer, yResult, yEventType)
    End Sub
    Private Sub DoMarkEmailReadStatus(ByRef oPlayer As Player, ByVal yData() As Byte, ByVal lMsgStartPos As Int32)
        Dim lPos As Int32 = lMsgStartPos + 2        'for msgcode
        Dim lEmailID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPCF_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yRead As Byte = yData(lPos) : lPos += 1

        For X As Int32 = 0 To oPlayer.EmailFolderUB
            If oPlayer.EmailFolderIdx(X) = lPCF_ID Then
                With oPlayer.EmailFolders(X)
                    For Y As Int32 = 0 To .PlayerMsgsUB
                        If .PlayerMsgsIdx(Y) = lEmailID Then
                            .PlayerMsgs(Y).MsgRead = yRead
                            .PlayerMsgs(Y).SaveObject(oPlayer.ObjectID)
                            Exit For
                        End If
                    Next Y
                End With
                Exit For
            End If
        Next X
    End Sub
    Private Function DoEmailSettings(ByRef oPlayer As Player, ByVal yData() As Byte, ByVal lMsgStartPos As Int32) As Byte()
        Dim lPos As Int32 = lMsgStartPos + 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iValues As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim iInternalVals As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim sPrevEmail As String = ""
        Dim sNewEmail As String = ""
        'Now, check our values
        If iValues > -1 Then
            'Ok, client is requesting a change to the email settings... get teh email address
            sPrevEmail = BytesToString(oPlayer.ExternalEmailAddress)
            ReDim oPlayer.ExternalEmailAddress(254)
            Array.Copy(yData, lPos, oPlayer.ExternalEmailAddress, 0, 255)
            sNewEmail = BytesToString(oPlayer.ExternalEmailAddress)
            oPlayer.iEmailSettings = iValues
            oPlayer.iInternalEmailSettings = iInternalVals
        End If

        'Ok, client is requesting the email settings
        Dim yResp(264) As Byte
        lPos = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eEmailSettings).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(oPlayer.iEmailSettings).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(oPlayer.iInternalEmailSettings).CopyTo(yResp, lPos) : lPos += 2
        oPlayer.ExternalEmailAddress.CopyTo(yResp, lPos) : lPos += 255

        If iValues > 1 Then
            'And then, send it to the email srvr if needed 
            If sPrevEmail.Trim.ToUpper <> sNewEmail.Trim.ToUpper Then SendToEmailSrvr(yResp) 'moEmailSrvr.SendData(yResp)
        End If

        Return yResp
    End Function
    Private Function DoGetIntelSellOrderDetail(ByRef oPlayer As Player, ByVal yData() As Byte, ByVal lMsgStartPos As Int32) As Byte()
        Dim lPos As Int32 = lMsgStartPos + 2
        Dim lTP_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lItemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim iExtTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Select Case iTypeID
            Case ObjectType.ePlayerIntel
                For X As Int32 = 0 To oPlayer.mlPlayerIntelUB
                    If oPlayer.mlPlayerIntelIdx(X) = lItemID Then
                        Dim oPI As PlayerIntel = oPlayer.moPlayerIntel(X)
                        If oPI Is Nothing = False Then
                            With oPI
                                Dim yResp(57) As Byte
                                lPos = 0
                                System.BitConverter.GetBytes(GlobalMessageCode.eGetIntelSellOrderDetail).CopyTo(yResp, lPos) : lPos += 2
                                System.BitConverter.GetBytes(lTP_ID).CopyTo(yResp, lPos) : lPos += 4
                                System.BitConverter.GetBytes(lItemID).CopyTo(yResp, lPos) : lPos += 4
                                System.BitConverter.GetBytes(iTypeID).CopyTo(yResp, lPos) : lPos += 2
                                System.BitConverter.GetBytes(iExtTypeID).CopyTo(yResp, lPos) : lPos += 2
                                .oTarget.PlayerName.CopyTo(yResp, lPos) : lPos += 20

                                Dim lResultValue As Int32 = 0
                                If .DiplomacyScore <> Int32.MinValue Then lResultValue = CInt(Now.Subtract(GetDateFromNumber(.DiplomacyUpdate)).TotalSeconds) Else lResultValue = 0
                                System.BitConverter.GetBytes(lResultValue).CopyTo(yResp, lPos) : lPos += 4
                                If .MilitaryScore <> Int32.MinValue Then lResultValue = CInt(Now.Subtract(GetDateFromNumber(.MilitaryUpdate)).TotalSeconds) Else lResultValue = 0
                                System.BitConverter.GetBytes(lResultValue).CopyTo(yResp, lPos) : lPos += 4
                                If .PopulationScore <> Int32.MinValue Then lResultValue = CInt(Now.Subtract(GetDateFromNumber(.PopulationUpdate)).TotalSeconds) Else lResultValue = 0
                                System.BitConverter.GetBytes(lResultValue).CopyTo(yResp, lPos) : lPos += 4
                                If .ProductionScore <> Int32.MinValue Then lResultValue = CInt(Now.Subtract(GetDateFromNumber(.ProductionUpdate)).TotalSeconds) Else lResultValue = 0
                                System.BitConverter.GetBytes(lResultValue).CopyTo(yResp, lPos) : lPos += 4
                                If .TechnologyScore <> Int32.MinValue Then lResultValue = CInt(Now.Subtract(GetDateFromNumber(.TechnologyUpdate)).TotalSeconds) Else lResultValue = 0
                                System.BitConverter.GetBytes(lResultValue).CopyTo(yResp, lPos) : lPos += 4
                                If .WealthScore <> Int32.MinValue Then lResultValue = CInt(Now.Subtract(GetDateFromNumber(.WealthUpdate)).TotalSeconds) Else lResultValue = 0
                                System.BitConverter.GetBytes(lResultValue).CopyTo(yResp, lPos) : lPos += 4

                                'moClients(lIndex).SendData(yResp)
                                Return yResp
                            End With
                        End If
                        Exit For
                    End If
                Next X
            Case ObjectType.ePlayerItemIntel

                For X As Int32 = 0 To oPlayer.mlItemIntelUB
                    If oPlayer.moItemIntel(X) Is Nothing = False AndAlso oPlayer.moItemIntel(X).lItemID = lItemID AndAlso oPlayer.moItemIntel(X).iItemTypeID = iExtTypeID Then
                        Dim oPII As PlayerItemIntel = oPlayer.moItemIntel(X)
                        If oPII Is Nothing = False Then

                            Dim yResp(58) As Byte
                            lPos = 0
                            System.BitConverter.GetBytes(GlobalMessageCode.eGetIntelSellOrderDetail).CopyTo(yResp, lPos) : lPos += 2
                            System.BitConverter.GetBytes(lTP_ID).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(lItemID).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(iTypeID).CopyTo(yResp, lPos) : lPos += 2
                            System.BitConverter.GetBytes(iExtTypeID).CopyTo(yResp, lPos) : lPos += 2

                            Dim oOtherPlayer As Player = GetEpicaPlayer(oPII.lOtherPlayerID)
                            If oOtherPlayer Is Nothing = False Then
                                oOtherPlayer.PlayerName.CopyTo(yResp, lPos)
                            End If
                            lPos += 20

                            yResp(lPos) = oPII.yIntelType : lPos += 1
                            Dim lSeconds As Int32 = CInt(Now.Subtract(oPII.dtTimeStamp).TotalSeconds)
                            System.BitConverter.GetBytes(lSeconds).CopyTo(yResp, lPos) : lPos += 4

                            If oPII.iItemTypeID = ObjectType.eColony Then
                                StringToBytes("A Colony").CopyTo(yResp, lPos)
                            Else
                                Dim oObj As Object = GetEpicaObject(oPII.lItemID, oPII.iItemTypeID)
                                If oObj Is Nothing = False Then
                                    GetEpicaObjectName(oPII.iItemTypeID, oObj).CopyTo(yResp, lPos)
                                End If
                                oObj = Nothing
                            End If

                            lPos += 20

                            'moClients(lIndex).SendData(yResp)
                            Return yResp

                        End If
                        Exit For
                    End If
                Next X

            Case ObjectType.ePlayerTechKnowledge
                Dim oPTK As PlayerTechKnowledge = oPlayer.GetPlayerTechKnowledge(lItemID, iExtTypeID)
                If oPTK Is Nothing = False Then
                    Dim yResp(54) As Byte
                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eGetIntelSellOrderDetail).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lTP_ID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(lItemID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(iTypeID).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(iExtTypeID).CopyTo(yResp, lPos) : lPos += 2

                    Dim oOtherPlayer As Player = oPTK.oTech.Owner
                    If oOtherPlayer Is Nothing = False Then
                        oOtherPlayer.PlayerName.CopyTo(yResp, lPos)
                    End If
                    lPos += 20

                    yResp(lPos) = oPTK.yKnowledgeType : lPos += 1
                    oPTK.oTech.GetTechName.CopyTo(yResp, lPos) : lPos += 20

                    'moClients(lIndex).SendData(yResp)
                    Return yResp
                End If
        End Select
        Return Nothing
    End Function
    Private Function DoMoveEmailToFolder(ByRef oPlayer As Player, ByVal yData() As Byte, ByVal lMsgStartPos As Int32) As Byte()
        Dim lPos As Int32 = lMsgStartPos + 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lPCF_Src As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPCF_Dest As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lFromIdx As Int32 = oPlayer.GetOrAddEmailFolder(lPCF_Src)
        Dim lToIdx As Int32 = oPlayer.GetOrAddEmailFolder(lPCF_Dest)

        If lFromIdx = -1 OrElse lToIdx = -1 Then Return Nothing
        If lFromIdx = lToIdx Then Return Nothing

        oPlayer.EmailFolders(lFromIdx).MoveMessageToFolder(lObjID, oPlayer.EmailFolders(lToIdx))

        Dim lLen As Int32 = yData.Length - lMsgStartPos
        Dim yResp(lLen) As Byte
        Array.Copy(yData, lMsgStartPos, yResp, 0, lLen)
        Return yResp
    End Function
    Private Function DoPassThruAddCommand(ByVal lPlayerID As Int32, ByVal yData() As Byte, ByVal lMsgStartPos As Int32) As Byte()
        Dim lPos As Int32 = lMsgStartPos + 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Select Case iObjTypeID
            Case ObjectType.ePlayerComm
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then

                    Dim lPCF_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iEnc As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    Dim lSendByID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim lSendOn As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim bRead As Boolean = yData(lPos) <> 0 : lPos += 1

                    Dim lBodyLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim sBody As String = ""
                    If lBodyLen > 0 Then
                        sBody = GetStringFromBytes(yData, lPos, lBodyLen)
                        lPos += lBodyLen
                    End If
                    Dim lTitleLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim sTitle As String = ""
                    If lTitleLen > 0 Then
                        sTitle = GetStringFromBytes(yData, lPos, lTitleLen)
                        lPos += lTitleLen
                    End If
                    Dim lRecipientLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim sRecipient As String = ""
                    If lRecipientLen > 0 Then
                        sRecipient = GetStringFromBytes(yData, lPos, lRecipientLen)
                        lPos += lRecipientLen
                    End If
                    Dim lAttachCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim uAttach(lAttachCnt - 1) As PlayerComm.WPAttachment
                    For X As Int32 = 0 To lAttachCnt - 1
                        With uAttach(X)
                            .AttachNumber = yData(lPos) : lPos += 1
                            .EnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            .EnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                            .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            .sWPName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                        End With
                    Next X

                    Dim oPC As PlayerComm = oPlayer.AddEmailMsg(lObjID, lPCF_ID, iEnc, sBody, sTitle, lSendByID, lSendOn, bRead, sRecipient, uAttach) '
                    If oPC Is Nothing = False Then
                        If oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eViewEmail) = True Then
                            oPlayer.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                    End If
                End If
        End Select

        Return Nothing
    End Function
    Private Function DoRequestEmail(ByRef oPlayer As Player, ByVal yData() As Byte, ByVal lMsgStartPos As Int32) As Byte()
        Dim lPos As Int32 = lMsgStartPos + 2    'for msgcode
        Dim lPCF_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPC_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        For Y As Int32 = 0 To oPlayer.EmailFolderUB
            If oPlayer.EmailFolderIdx(Y) = lPCF_ID Then
                With oPlayer.EmailFolders(Y)
                    For Z As Int32 = 0 To .PlayerMsgsUB
                        If .PlayerMsgsIdx(Z) = lPC_ID Then
                            Return GetAddObjectMessage(.PlayerMsgs(Z), GlobalMessageCode.eAddObjectCommand)
                        End If
                    Next Z
                End With
                Exit For
            End If
        Next Y
        'LogEvent(LogEventType.PossibleCheat, "PCF_ID, PC_ID combo could not be found for HandleRequestEmail. PlayerID: " & oplayer.ObjectID)
        Return Nothing
    End Function
    Private Function DoSendEmail(ByRef oPlayer As Player, ByVal yData() As Byte, ByVal lMsgStartPos As Int32) As Byte()
        Dim lPos As Int32 = lMsgStartPos
        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oComm As New PlayerComm()
        oComm.FillFromEmailMsg(yData, lMsgStartPos)

        'Now... we should have a comm, figure out what to do with it...
        If iMsgCode = GlobalMessageCode.eSaveEmailDraft Then
            'Ok, we just set hte player id to this playerid
            oComm.PlayerID = oPlayer.ObjectID
            'Set the PCF_ID
            oComm.PCF_ID = PlayerCommFolder.ePCF_ID_HARDCODES.eDrafts_PCF
            'Set the encryption level...
            oComm.EncryptLevel = 0      'not encrypted
        Else
            'Send the Email... so the email will be put in the sender's outbox
            oComm.PlayerID = oPlayer.ObjectID
            oComm.PCF_ID = PlayerCommFolder.ePCF_ID_HARDCODES.eOutbox_PCF
            oComm.EncryptLevel = oPlayer.CommEncryptLevel
            'Oh, and it is read
            oComm.MsgRead = 255
        End If

        If oComm.SaveObject(oComm.PlayerID) = True Then
            'Now, attach the comm to the player
            Dim lIdx As Int32 = oPlayer.GetOrAddEmailFolder(oComm.PCF_ID)
            If lIdx <> -1 Then
                With oPlayer.EmailFolders(lIdx)
                    Dim lMsgIdx As Int32 = -1
                    For X As Int32 = 0 To .PlayerMsgsUB
                        If .PlayerMsgsIdx(X) = -1 Then
                            lMsgIdx = X
                            Exit For
                        End If
                    Next X
                    If lMsgIdx = -1 Then
                        .PlayerMsgsUB += 1
                        ReDim Preserve .PlayerMsgs(.PlayerMsgsUB)
                        ReDim Preserve .PlayerMsgsIdx(.PlayerMsgsUB)
                        lMsgIdx = .PlayerMsgsUB
                    End If

                    .PlayerMsgs(lMsgIdx) = oComm
                    .PlayerMsgsIdx(lMsgIdx) = oComm.ObjectID
                End With

                If iMsgCode = GlobalMessageCode.eSendEmail Then oComm.DoDelivery()

                Return GetAddObjectMessage(oComm, GlobalMessageCode.eAddObjectCommand)
            Else : oComm.DeleteComm()
            End If
        End If

        Return Nothing
    End Function
#End Region
End Class
