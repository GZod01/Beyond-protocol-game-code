Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAddressBook
    Inherits UIWindow

    Private WithEvents lstContacts As UIListBox
    Private WithEvents btnOK As UIButton
    Private WithEvents btnCancel As UIButton

    Public Event ContactSelected(ByVal sValue As String)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmAddressBook initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAddressBook
            .ControlName = "frmAddressBook"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 142
            .Width = 255
            .Height = 284
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

        'lstContacts initial props
        lstContacts = New UIListBox(oUILib)
        With lstContacts
            .ControlName = "lstContacts"
            .Left = 10
            .Top = 10
            .Width = 235
            .Height = 210
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstContacts, UIControl))

        'btnOK initial props
        btnOK = New UIButton(oUILib)
        With btnOK
            .ControlName = "btnOK"
            .Left = 10
            .Top = 250
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "OK"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnOK, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = 145
            .Top = 250
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Cancel"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCancel, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        FillContactList()
    End Sub

    Private Sub FillContactList()
        lstContacts.Clear()

        'Let's load the player's rels now...
        Dim oTmpRel As PlayerRel
        Dim sName As String

        If goCurrentPlayer Is Nothing = False Then
            For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
                oTmpRel = goCurrentPlayer.GetPlayerRelByIndex(X)
                If oTmpRel Is Nothing = False Then
                    sName = GetCacheObjectValue(oTmpRel.lThisPlayer, ObjectType.ePlayer)
                    lstContacts.AddItem(sName)
                    lstContacts.ItemData(lstContacts.NewIndex) = oTmpRel.lThisPlayer
                End If
            Next X
        End If
        lstContacts.SortList(False, False)
    End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnOK_Click(ByVal sName As String) Handles btnOK.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		If lstContacts.ListIndex <> -1 Then RaiseEvent ContactSelected(lstContacts.List(lstContacts.ListIndex))
	End Sub

    Private Sub frmAddressBook_OnNewFrame() Handles Me.OnNewFrame
        For X As Int32 = 0 To lstContacts.ListCount - 1
            Dim sTemp As String = GetCacheObjectValue(lstContacts.ItemData(X), ObjectType.ePlayer)
            If lstContacts.List(X) <> sTemp Then lstContacts.List(X) = sTemp
        Next X
    End Sub

    Private Sub lstContacts_ItemDblClick(ByVal lIndex As Integer) Handles lstContacts.ItemDblClick
        If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create emails.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        RaiseEvent ContactSelected(lstContacts.List(lstContacts.ListIndex))
    End Sub
End Class