Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmChatTabProps
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblFilter As UILabel
    Private chkLocal As UICheckBox
    Private chkSystem As UICheckBox
    Private chkChannel As UICheckBox
    Private txtChannel As UITextBox
    Private chkAlliance As UICheckBox
    Private chkAlias As UICheckBox
    Private chkPM As UICheckBox
    Private chkNotification As UICheckBox
    Private txtTabName As UITextBox
    Private lnDiv2 As UILine
    Private lbPrefix As UILabel
    Private txtMsgPrefix As UITextBox
    Private WithEvents btnReset As UIButton
    Private lnDiv3 As UILine
    Private WithEvents btnClose As UIButton

    Private moTab As ChatTab

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmChatTabProps initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eChatTabProps
            .ControlName = "frmChatTabProps"
            .Left = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 140
            .Top = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 128
            .Width = 280
            .Height = 255
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .BorderLineWidth = 1
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 116
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Chat Tab Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 0
            .Top = 25
            .Width = 280
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblFilter initial props
        lblFilter = New UILabel(oUILib)
        With lblFilter
            .ControlName = "lblFilter"
            .Left = 5
            .Top = 30
            .Width = 36
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Filter"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblFilter, UIControl))

        'chkLocal initial props
        chkLocal = New UICheckBox(oUILib)
        With chkLocal
            .ControlName = "chkLocal"
            .Left = 10
            .Top = 50
            .Width = 120
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Local Messages"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "If checked, this tab displays Local Messages" & vbCrLf & _
                           "for the environment you are currently in." & vbCrLf & _
                           "To send a local message type /local." & vbCrLf & _
                           "Local is also the default send type."
        End With
        Me.AddChild(CType(chkLocal, UIControl))

        'chkSystem initial props
        chkSystem = New UICheckBox(oUILib)
        With chkSystem
            .ControlName = "chkSystem"
            .Left = 10
            .Top = 70
            .Width = 172
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "System Admin Messages"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "If checked, this tab displays System messages" & vbCrLf & _
                           "that are received. These come in any forms."
        End With
        Me.AddChild(CType(chkSystem, UIControl))

        'chkChannel initial props
        chkChannel = New UICheckBox(oUILib)
        With chkChannel
            .ControlName = "chkChannel"
            .Left = 10
            .Top = 90
            .Width = 135
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Channel Messages"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "If checked, this tab displays messages received" & vbCrLf & _
                           "from channels. If the textbox to the right of this" & vbCrLf & _
                           "is filled, only messages of that channel make it" & vbCrLf & _
                           "into this tab. Other channel messages will be filtered."
        End With
        Me.AddChild(CType(chkChannel, UIControl))

        'txtChannel initial props
        txtChannel = New UITextBox(oUILib)
        With txtChannel
            .ControlName = "txtChannel"
            .Left = 165
            .Top = 91
            .Width = 112
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .ToolTipText = "Further filters channel messages so that only" & vbCrLf & _
                           "messages received from this channel will appear."
        End With
        Me.AddChild(CType(txtChannel, UIControl))

        'chkAlliance initial props
        chkAlliance = New UICheckBox(oUILib)
        With chkAlliance
            .ControlName = "chkAlliance"
            .Left = 10
            .Top = 110
            .Width = 117
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Guild Messages"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "If checked, guild messages will be displayed here." & vbCrLf & _
                           "To send a guild message, type /gu."
        End With
        Me.AddChild(CType(chkAlliance, UIControl))

        'chkSenate initial props
        chkAlias = New UICheckBox(oUILib)
        With chkAlias
            .ControlName = "chkAlias"
            .Left = 10
            .Top = 130
            .Width = 117
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Alias Messages"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "If checked, messages sent between aliased" & vbCrLf & _
                           "accounts will appear in this tab. To chat to an" & vbCrLf & _
                           "aliased player or aliasing player, type /alias."
        End With
        Me.AddChild(CType(chkAlias, UIControl))

        'chkPM initial props
        chkPM = New UICheckBox(oUILib)
        With chkPM
            .ControlName = "chkPM"
            .Left = 10
            .Top = 150
            .Width = 164
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Private Messages (tells)"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "If checked, private messages will be displayed" & vbCrLf & _
                           "in this tab. To send a private message, type" & vbCrLf & _
                           "/pm followed by the player's name and then a" & vbCrLf & _
                           "comma. For example: /pm Jane, Hey there."
        End With
        Me.AddChild(CType(chkPM, UIControl))

        'chkNotification initial props
        chkNotification = New UICheckBox(oUILib)
        With chkNotification
            .ControlName = "chkNotification"
            .Left = 10
            .Top = 170
            .Width = 153
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Notification Messages"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "If checked, messages that go to the notification" & vbCrLf & _
                           "window will also appear within this tab."
        End With
        Me.AddChild(CType(chkNotification, UIControl))

        'txtTabName initial props
        txtTabName = New UITextBox(oUILib)
        With txtTabName
            .ControlName = "txtTabName"
            .Left = 127
            .Top = 4
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 11
            .BorderColor = muSettings.InterfaceBorderColor
            .ToolTipText = "The name of the tab displayed in the chat window."
        End With
        Me.AddChild(CType(txtTabName, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 0
            .Top = 190
            .Width = 280
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'lbPrefix initial props
        lbPrefix = New UILabel(oUILib)
        With lbPrefix
            .ControlName = "lbPrefix"
            .Left = 5
            .Top = 195
            .Width = 144
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Default Message Prefix:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "This text will be appended to the beginning of each message" & vbCrLf & _
                           "sent that does not start with a '/'. Use this to allow for" & vbCrLf & _
                           "quick message sending to the same place without the need to" & vbCrLf & _
                           "type the command. Be sure to include the proper syntax."
        End With
        Me.AddChild(CType(lbPrefix, UIControl))

        'txtMsgPrefix initial props
        txtMsgPrefix = New UITextBox(oUILib)
        With txtMsgPrefix
            .ControlName = "txtMsgPrefix"
            .Left = 152
            .Top = 195
            .Width = 125
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .ToolTipText = "This text will be appended to the beginning of each message" & vbCrLf & _
                           "sent that does not start with a '/'. Use this to allow for" & vbCrLf & _
                           "quick message sending to the same place without the need to" & vbCrLf & _
                           "type the command. Be sure to include the proper syntax."
        End With
        Me.AddChild(CType(txtMsgPrefix, UIControl))

        'btnReset initial props
        btnReset = New UIButton(oUILib)
        With btnReset
            .ControlName = "btnReset"
            .Left = 5
            .Top = 225
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Reset"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnReset, UIControl))

        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = 0
            .Top = 218
            .Width = 280
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 178
            .Top = 225
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Close"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Public Sub SetFromTab(ByRef oTab As ChatTab)
        moTab = oTab
        chkLocal.Value = (oTab.lFilter And ChatTab.ChatFilter.eLocalMessages) <> 0
        chkSystem.Value = (oTab.lFilter And ChatTab.ChatFilter.eSysAdminMessages) <> 0
        chkChannel.Value = (oTab.lFilter And ChatTab.ChatFilter.eChannelMessages) <> 0
        chkAlliance.Value = (oTab.lFilter And ChatTab.ChatFilter.eAllianceMessages) <> 0
        chkAlias.Value = (oTab.lFilter And ChatTab.ChatFilter.eAliasChatMessage) <> 0
        chkPM.Value = (oTab.lFilter And ChatTab.ChatFilter.ePMs) <> 0
        chkNotification.Value = (oTab.lFilter And ChatTab.ChatFilter.eNotificationMessages) <> 0
        txtChannel.Caption = oTab.sChannel
        txtTabName.Caption = oTab.sTabName
        txtMsgPrefix.Caption = oTab.sMessagePrefix
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        With moTab
            Dim lFilter As Int32 = 0
            If chkLocal.Value = True Then lFilter = lFilter Or ChatTab.ChatFilter.eLocalMessages
            If chkSystem.Value = True Then lFilter = lFilter Or ChatTab.ChatFilter.eSysAdminMessages
            If chkChannel.Value = True Then lFilter = lFilter Or ChatTab.ChatFilter.eChannelMessages
            If chkAlliance.Value = True Then lFilter = lFilter Or ChatTab.ChatFilter.eAllianceMessages
            If chkAlias.Value = True Then lFilter = lFilter Or ChatTab.ChatFilter.eAliasChatMessage
            If chkPM.Value = True Then lFilter = lFilter Or ChatTab.ChatFilter.ePMs
            If chkNotification.Value = True Then lFilter = lFilter Or ChatTab.ChatFilter.eNotificationMessages

            .lFilter = lFilter
            .sChannel = txtChannel.Caption
            .sTabName = txtTabName.Caption
            .sMessagePrefix = txtMsgPrefix.Caption

            If .sMessagePrefix = "" Then
                Select Case lFilter
                    Case ChatTab.ChatFilter.eAliasChatMessage
                        .sMessagePrefix = "/alias"
                    Case ChatTab.ChatFilter.eAllianceMessages
                        .sMessagePrefix = "/gu"
                    Case ChatTab.ChatFilter.eChannelMessages
                        If .sChannel <> "" Then
                            .sMessagePrefix = "/" & Replace(.sChannel, "/", "")
                        End If
                    Case ChatTab.ChatFilter.eLocalMessages
                        .sMessagePrefix = "/local"
                End Select
            End If

            .SaveTab()
        End With

        Dim ofrmChat As frmChat = CType(goUILib.GetWindow("frmChat"), frmChat)
        If ofrmChat Is Nothing = False Then ofrmChat.TabPropsClosed()
        ofrmChat = Nothing

        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

	Private Sub btnReset_Click(ByVal sName As String) Handles btnReset.Click
		SetFromTab(moTab)
	End Sub
End Class