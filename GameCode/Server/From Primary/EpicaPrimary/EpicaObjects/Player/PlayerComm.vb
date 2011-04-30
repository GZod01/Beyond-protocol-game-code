Public Class PlayerComm
    Inherits Epica_GUID

    Public Structure WPAttachment
        Public AttachNumber As Byte
        Public EnvirID As Int32
        Public EnvirTypeID As Int16
        Public LocX As Int32
        Public LocZ As Int32
        Public sWPName As String

        Public yWPNameBytes() As Byte

        Public Sub CloneFrom(ByVal uAttachment As WPAttachment)
            With Me
                .AttachNumber = uAttachment.AttachNumber
                .EnvirID = uAttachment.EnvirID
                .EnvirTypeID = uAttachment.EnvirTypeID
                .LocX = uAttachment.LocX
                .LocZ = uAttachment.LocZ
                .sWPName = uAttachment.sWPName
                .yWPNameBytes = uAttachment.yWPNameBytes
            End With
        End Sub

        Public Sub New(ByVal pyNum As Byte, ByVal plEnvirID As Int32, ByVal piEnvirTypeID As Int16, ByVal plX As Int32, ByVal plZ As Int32, ByVal psName As String)
            With Me
                .AttachNumber = pyNum
                .EnvirID = plEnvirID
                .EnvirTypeID = piEnvirTypeID
                .LocX = plX
                .LocZ = plZ
                .sWPName = psName
            End With
        End Sub
    End Structure

    Public SentByID As Int32        'Player ID who sent it... only one player can send a message
    Public SentOn As Int32          'Date as INT32 that this msg was sent, format is YYMMDDHHmm (max being 2021)
    Public EncryptLevel As Short    'Level of encryption (used for interception)
    Public PCF_ID As Int32          'ID of the folder that the player owns that contains this message
    Public MsgRead As Byte          '0 indicates not read, non-zero indicates read

    Public MsgTitle() As Byte
    Public MsgBody() As Byte
    Public SentToList() As Byte     'the typed in value of the TO portion of the recipient
    Public BCCText() As Byte        'NEVER sent to the client???

    Public PlayerID As Int32

    Private muAttachments() As WPAttachment
    Private mlAttachmentUB As Int32 = -1

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        'If mbStringReady = False Then

        Dim lFinalLen As Int32 = 56
        If SentToList Is Nothing = False Then lFinalLen += SentToList.Length
        If MsgTitle Is Nothing = False Then lFinalLen += MsgTitle.Length
        If MsgBody Is Nothing = False Then lFinalLen += MsgBody.Length
        If BCCText Is Nothing = False Then
            If SentByID = PlayerID Then lFinalLen += BCCText.Length
        End If

        ReDim mySendString(lFinalLen + ((mlAttachmentUB + 1) * 35))

        Dim lPos As Int32 = 0
        Me.GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
        oSender.PlayerName.CopyTo(mySendString, lPos) : lPos += 20
        System.BitConverter.GetBytes(SentOn).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(EncryptLevel).CopyTo(mySendString, lPos) : lPos += 2
        mySendString(lPos) = MsgRead : lPos += 1
        System.BitConverter.GetBytes(PCF_ID).CopyTo(mySendString, lPos) : lPos += 4

        If SentToList Is Nothing = False Then
            System.BitConverter.GetBytes(SentToList.Length).CopyTo(mySendString, lPos) : lPos += 4
            SentToList.CopyTo(mySendString, lPos) : lPos += SentToList.Length
        Else : System.BitConverter.GetBytes(0I).CopyTo(mySendString, lPos) : lPos += 4
        End If
        If MsgTitle Is Nothing = False Then
            System.BitConverter.GetBytes(MsgTitle.Length).CopyTo(mySendString, lPos) : lPos += 4
            MsgTitle.CopyTo(mySendString, lPos) : lPos += MsgTitle.Length
        Else : System.BitConverter.GetBytes(0I).CopyTo(mySendString, lPos) : lPos += 4
        End If
        If MsgBody Is Nothing = False Then
            System.BitConverter.GetBytes(MsgBody.Length).CopyTo(mySendString, lPos) : lPos += 4
            MsgBody.CopyTo(mySendString, lPos) : lPos += MsgBody.Length
        Else : System.BitConverter.GetBytes(0I).CopyTo(mySendString, lPos) : lPos += 4
        End If
        If BCCText Is Nothing = False AndAlso SentByID = PlayerID Then
            System.BitConverter.GetBytes(BCCText.Length).CopyTo(mySendString, lPos) : lPos += 4
            BCCText.CopyTo(mySendString, lPos) : lPos += BCCText.Length
        Else : System.BitConverter.GetBytes(0I).CopyTo(mySendString, lPos) : lPos += 4
        End If

        System.BitConverter.GetBytes(mlAttachmentUB + 1).CopyTo(mySendString, lPos) : lPos += 4
        For X As Int32 = 0 To mlAttachmentUB
            With muAttachments(X)
                mySendString(lPos) = .AttachNumber : lPos += 1
                System.BitConverter.GetBytes(.EnvirID).CopyTo(mySendString, lPos) : lPos += 4
                System.BitConverter.GetBytes(.EnvirTypeID).CopyTo(mySendString, lPos) : lPos += 2
                System.BitConverter.GetBytes(.LocX).CopyTo(mySendString, lPos) : lPos += 4
                System.BitConverter.GetBytes(.LocZ).CopyTo(mySendString, lPos) : lPos += 4

                If .sWPName Is Nothing OrElse .sWPName = "" Then .sWPName = "WP " & .AttachNumber
                If .yWPNameBytes Is Nothing OrElse .yWPNameBytes.GetUpperBound(0) <> 19 Then
                    ReDim .yWPNameBytes(19)
                    StringToBytes(.sWPName).CopyTo(.yWPNameBytes, 0)
                End If
                .yWPNameBytes.CopyTo(mySendString, lPos) : lPos += 20
            End With
        Next X

        mbStringReady = True
        'End If
        Return mySendString
    End Function
    
    Public Function SaveObject(ByVal lOwningPlayerID As Int32) As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
        If lOwningPlayerID > 0 Then PlayerID = lOwningPlayerID

        Try
            If Me.ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblPlayerComm (SentByID, SentOn, MsgTitle, MsgBody, EncryptLevel, PlayerID, ReadFlag, " & _
                  "SentTo, PCF_ID) VALUES (" & SentByID & ", " & SentOn & ", '" & MakeDBStr(BytesToString(MsgTitle)) & _
                  "', '" & MakeDBStr(BytesToString(MsgBody)) & "', " & EncryptLevel & ", " & PlayerID & ", " & MsgRead & _
                  ", '" & MakeDBStr(BytesToString(SentToList)) & "', " & PCF_ID & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblPlayerComm SET SentByID = " & SentByID & ", SentOn = " & SentOn & _
                  ", MsgTitle = '" & MakeDBStr(BytesToString(MsgTitle)) & "', MsgBody = '" & MakeDBStr(BytesToString(MsgBody)) & _
                  "', EncryptLevel = " & EncryptLevel & ", PlayerID = " & PlayerID & ", ReadFlag = " & MsgRead & _
                  ", SentTo = '" & MakeDBStr(BytesToString(SentToList)) & "', PCF_ID = " & PCF_ID & " WHERE PC_ID = " & Me.ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If Me.ObjectID < 1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PC_ID) FROM tblPlayerComm WHERE SentByID = " & SentByID & " AND SentOn = " & SentOn
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    Me.ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            oComm = Nothing
            sSQL = "DELETE FROM tblEmailAttachment WHERE PC_ID = " & Me.ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm = Nothing

            For X As Int32 = 0 To mlAttachmentUB
                Try
                    With muAttachments(X)
                        sSQL = "INSERT INTO tblEmailAttachment (PC_ID, AttachNumber, EnvirID, EnvirTypeID, LocX, LocZ, " & _
                           "WPName) VALUES (" & Me.ObjectID & ", " & .AttachNumber & ", " & .EnvirID & ", " & .EnvirTypeID & _
                           ", " & .LocX & ", " & .LocZ & ", '" & MakeDBStr(Mid$(.sWPName, 1, 20)) & "')"
                        oComm = New OleDb.OleDbCommand(sSQL, goCN)
                        If oComm.ExecuteNonQuery() = 0 Then
                            Err.Raise(-1, "SaveObject", "No records affected when inserting object!")
                        End If
                        oComm = Nothing
                    End With
                Catch ex As Exception
                    LogEvent(LogEventType.CriticalError, "Unable to save email attachment. Reason: " & Err.Description)
                End Try
            Next X


            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object PlayerComm. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Private moSender As Player = Nothing
    Public ReadOnly Property oSender() As Player
        Get
            If moSender Is Nothing OrElse moSender.ObjectID <> SentByID Then
                If SentByID < 1 Then
                    moSender = GetEpicaPlayer(PlayerID)
                Else
                    moSender = GetEpicaPlayer(SentByID)
                End If
            End If

            Return moSender
        End Get
    End Property

    Public Sub FillFromEmailMsg(ByRef yData() As Byte, ByVal lStartPos As Int32)
        Dim lPos As Int32 = lStartPos + 2       '2 byte msg code

        mbStringReady = False

        Me.ObjectID = -1
        Me.ObjTypeID = ObjectType.ePlayerComm

        SentByID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lLen As Int32

        lLen = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ReDim SentToList(lLen - 1)
        Array.Copy(yData, lPos, SentToList, 0, lLen)
        lPos += lLen

        lLen = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ReDim BCCText(lLen - 1)
        Array.Copy(yData, lPos, BCCText, 0, lLen)
        lPos += lLen

        lLen = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ReDim MsgTitle(lLen - 1)
        Array.Copy(yData, lPos, MsgTitle, 0, lLen)
        lPos += lLen

        lLen = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ReDim MsgBody(lLen - 1)
        Array.Copy(yData, lPos, MsgBody, 0, lLen)
        lPos += lLen

        mlAttachmentUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4

        ReDim muAttachments(mlAttachmentUB)
        For X As Int32 = 0 To mlAttachmentUB
            With muAttachments(X)
                .AttachNumber = yData(lPos) : lPos += 1
                .EnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .EnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                ReDim .yWPNameBytes(19)
                Array.Copy(yData, lPos, .yWPNameBytes, 0, 20) : lPos += 20
                .sWPName = BytesToString(.yWPNameBytes)
            End With
        Next X
    End Sub

    Public Sub DeleteComm()
        If Me.ObjectID > -1 Then
            Dim sSQL As String = "DELETE FROM tblPlayerComm WHERE PC_ID = " & Me.ObjectID
            Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
        End If
    End Sub

    Public Sub DoDelivery()
        Dim sToList As String = ""
        Dim sBCCList As String = ""
        If SentToList Is Nothing = False Then sToList = BytesToString(SentToList)
        If BCCText Is Nothing = False Then sBCCList = BytesToString(BCCText)

        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder


        oSB.Append(sToList)
        If sToList.EndsWith(";") = False AndAlso sToList.EndsWith(",") = False Then oSB.Append(";")
        oSB.Append(sBCCList)
        oSB.Replace(",", ";")
        Dim sFullList As String = oSB.ToString

        'Now, reset our stringbuilder
        oSB.Length = 0
        oSB.Append("SELECT PlayerID, PlayerName FROM tblPlayer WHERE PlayerName IN (")

        Dim sValues() As String = Split(sFullList, ";")
        Dim lIDs(sValues.GetUpperBound(0)) As Int32

        Dim bFirst As Boolean = True

        For X As Int32 = 0 To sValues.GetUpperBound(0)
            lIDs(X) = -1
            If sValues(X).Trim <> "" Then
                sValues(X) = sValues(X).ToUpper.Trim
                If bFirst = True Then
                    oSB.Append("'" & MakeDBStr(sValues(X)) & "'")
                    bFirst = False
                Else
                    oSB.Append(", '" & MakeDBStr(sValues(X)) & "'")
                End If
            End If
        Next X

        Dim sSQL As String = oSB.ToString & ")"
        If bFirst = False Then
            Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
            Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader()
            While oData.Read
                Dim sName As String = CStr(oData("PlayerName")).ToUpper
                Dim lID As Int32 = CInt(oData("PlayerID"))

                For X As Int32 = 0 To sValues.GetUpperBound(0)
                    If sValues(X) = sName Then
                        lIDs(X) = lID
                    End If
                Next X
            End While
            oData.Close()
            oData = Nothing
            oComm.Dispose()
            oComm = Nothing

            Me.SentOn = GetDateAsNumber(Now)

            'For Cross-Primary data... the Sender should be owned by me... but the Sender may not necessarily be connected to me at the time... recipients may be anywhere...

            'Now, go back through and get our ID's...
            For X As Int32 = 0 To sValues.GetUpperBound(0)
                If lIDs(X) = -1 AndAlso sValues(X).Trim <> "" Then
                    Dim oNewEmail As PlayerComm = oSender.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Unable to deliver message to: " & sValues(X) & vbCrLf & "Reason: Destination address does not exist.", "Unable to Deliver", oSender.ObjectID, SentOn, False, BytesToString(oSender.PlayerName), Nothing)
                    If oNewEmail Is Nothing = False Then
                        oSender.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oNewEmail, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
					End If
				ElseIf lIDs(X) <> -1 Then
                    Dim oPlayer As Player = GetEpicaPlayer(lIDs(X))
                    If oPlayer Is Nothing = False Then

                        Dim oNewEmail As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, Me.EncryptLevel, BytesToString(Me.MsgBody), BytesToString(Me.MsgTitle), SentByID, SentOn, False, BytesToString(Me.SentToList), muAttachments)
                        If oNewEmail Is Nothing = False Then
                            'oNewEmail.mlAttachmentUB = mlAttachmentUB
                            'ReDim oNewEmail.muAttachments(mlAttachmentUB)
                            'For Y As Int32 = 0 To mlAttachmentUB
                            '    oNewEmail.muAttachments(Y).CloneFrom(muAttachments(Y))
                            'Next Y 
                            'If oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eViewEmail) = True Then 
                            oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oNewEmail, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                    End If
                    
                End If
            Next X

        End If
    End Sub

    Public Sub AddEmailAttachment(ByVal yNumber As Byte, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lX As Int32, ByVal lZ As Int32, ByVal sName As String)
        mlAttachmentUB += 1
        ReDim Preserve muAttachments(mlAttachmentUB)
        With muAttachments(mlAttachmentUB)
            .AttachNumber = yNumber
            .EnvirID = lEnvirID
            .EnvirTypeID = iEnvirTypeID
            .LocX = lX
            .LocZ = lZ
            .sWPName = sName
        End With
    End Sub
End Class
