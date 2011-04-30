Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmEmpMsgBrd
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents btnClose As UIButton
    Private lnDiv1 As UILine
    Private WithEvents lstMsgs As UIListBox
    Private txtMessage As UITextBox
    Private WithEvents btnAddMessage As UIButton

    Private moItems() As EmpMsgBoardItem
    Private mlItemUB As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmEmpMsgBrd initial props
        With Me
            .ControlName = "frmEmpMsgBrd"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 175
            .Width = 255
            .Height = 344
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .BorderLineWidth = 2
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 110
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Message Board"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 229
            .Top = 2
            .Width = 25
            .Height = 25
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 28
            .Width = 254
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lstMsgs initial props
        lstMsgs = New UIListBox(oUILib)
        With lstMsgs
            .ControlName = "lstMsgs"
            .Left = 5
            .Top = 32
            .Width = 245
            .Height = 130
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstMsgs, UIControl))

        'txtMessage initial props
        txtMessage = New UITextBox(oUILib)
        With txtMessage
            .ControlName = "txtMessage"
            .Left = 5
            .Top = 175
            .Width = 245
            .Height = 130
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 1000
            .MultiLine = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtMessage, UIControl))

        'btnAddMessage initial props
        btnAddMessage = New UIButton(oUILib)
        With btnAddMessage
            .ControlName = "btnAddMessage"
            .Left = 68
            .Top = 312
            .Width = 120
            .Height = 26
            .Enabled = True
            .Visible = True
            .Caption = "Add Message"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAddMessage, UIControl))

        txtMessage.Locked = True
        txtMessage.BackColorEnabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        btnAddMessage.Caption = "Add Message"

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        Dim yMsg(6) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2
        lPos += 4       'leave room for PlayerId
        yMsg(lPos) = eySenateRequestDetailsType.EmpChmbrMsgList : lPos += 1
        goUILib.SendMsgToPrimary(yMsg)
    End Sub

    Public Sub HandleEmpChmbrList(ByVal yData() As Byte)
        Dim lPos As Int32 = 7   'for code and playerid and type
        Dim lUB As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lUB -= 1

        Dim oItems(lUB) As EmpMsgBoardItem
        For X As Int32 = 0 To lUB
            oItems(X) = New EmpMsgBoardItem()
            With oItems(X)
                .lPosterID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lPostedOn = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End With
        Next X

        mlItemUB = -1
        lstMsgs.Clear()
        moItems = oItems
        mlItemUB = lUB
    End Sub
    Public Sub HandleEmpChmbrMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 7       'for code,playerid, and type
        Dim lPosterID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPostedOn As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sMsg As String = GetStringFromBytes(yData, lPos, lLen)

        For X As Int32 = 0 To mlItemUB
            If moItems(X).lPosterID = lPosterID AndAlso moItems(X).lPostedOn = lPostedOn Then
                moItems(X).sMsg = sMsg
                Exit For
            End If
        Next X
    End Sub

    Private Sub btnAddMessage_Click(ByVal sName As String) Handles btnAddMessage.Click
        If btnAddMessage.Caption.ToUpper = "SUBMIT" Then
            'submit a msg
            btnAddMessage.Enabled = False
            If txtMessage.Caption.Trim.Length > 3 Then
                Dim lMsgLen As Int32 = txtMessage.Caption.Trim.Length
                Dim yMsg(13 + lMsgLen) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eAddSenateProposalMessage).CopyTo(yMsg, lPos) : lPos += 2
                lPos += 4       'for playerid
                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lMsgLen).CopyTo(yMsg, lPos) : lPos += 4
                System.Text.ASCIIEncoding.ASCII.GetBytes(txtMessage.Caption.Trim).CopyTo(yMsg, lPos) : lPos += lMsgLen
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                'add the message to our list...
                Dim oMsg As New EmpMsgBoardItem
                With oMsg
                    .lPosterID = glPlayerID
                    .lPostedOn = GetDateAsNumber(Now)
                    .sMsg = txtMessage.Caption.Trim
                End With
                mlItemUB += 1
                ReDim Preserve moItems(mlItemUB)
                moItems(mlItemUB) = oMsg

                MyBase.moUILib.AddNotification("Message added to message board.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Else
                MyBase.moUILib.AddNotification("Must enter in a message at least 3 characters to post it.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
            btnAddMessage.Enabled = True
            txtMessage.Locked = True
            txtMessage.BackColorEnabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            btnAddMessage.Caption = "Add Message"
        Else
            txtMessage.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            btnAddMessage.Caption = "Submit"
            lstMsgs.ListIndex = -1
            txtMessage.Caption = ""
            txtMessage.Locked = False
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub lstMsgs_ItemClick(ByVal lIndex As Integer) Handles lstMsgs.ItemClick
        btnAddMessage.Caption = "Add Message"
        txtMessage.Caption = "" : txtMessage.Locked = True
        txtMessage.BackColorEnabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
    End Sub

    Private Sub frmEmpMsgBrd_OnNewFrame() Handles Me.OnNewFrame
        'Now, check our message list
        If lstMsgs.ListCount <> mlItemUB + 1 Then
            FillMessageList()
        Else
            For X As Int32 = 0 To mlItemUB
                If lstMsgs.ItemData(X) <> moItems(X).lPosterID OrElse lstMsgs.ItemData2(X) <> moItems(X).lPostedOn Then
                    FillMessageList()
                    Exit For
                Else
                    Dim sText As String = GetCacheObjectValue(moItems(X).lPosterID, ObjectType.ePlayer)
                    If lstMsgs.List(X) <> sText Then lstMsgs.List(X) = sText
                End If
            Next X
        End If

        If lstMsgs.ListIndex > -1 Then
            Dim lPosterID As Int32 = lstMsgs.ItemData(lstMsgs.ListIndex)
            Dim lPostedOn As Int32 = lstMsgs.ItemData2(lstMsgs.ListIndex)
            For X As Int32 = 0 To mlItemUB
                If moItems(X).lPosterID = lPosterID AndAlso moItems(X).lPostedOn = lPostedOn Then
                    moItems(X).RequestDetails()
                    If txtMessage.Caption <> moItems(X).sMsg Then txtMessage.Caption = moItems(X).sMsg
                    If txtMessage.Locked <> True Then txtMessage.Locked = True
                    Exit For
                End If
            Next X
        End If
    End Sub


    Private Sub FillMessageList()
        lstMsgs.Clear()
        Try
            For X As Int32 = 0 To mlItemUB
                Dim sText As String = GetCacheObjectValue(moItems(X).lPosterID, ObjectType.ePlayer)
                lstMsgs.AddItem(sText, False)
                lstMsgs.ItemData(lstMsgs.NewIndex) = moItems(X).lPosterID
                lstMsgs.ItemData2(lstMsgs.NewIndex) = moItems(X).lPostedOn
            Next X
        Catch
        End Try
    End Sub
End Class