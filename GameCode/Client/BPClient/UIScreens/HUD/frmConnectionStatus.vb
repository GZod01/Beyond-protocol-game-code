Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmConnectionStatus
    Inherits UIWindow

    Private WithEvents btnClose As UIButton
    Private rcPrimary As Rectangle
    Private rcRegion As Rectangle
    Private mbLoading As Boolean = True

    Private miPrimaryStatus As Int32
    Private miRegionStatus As Int32

    Public Sub New(ByRef oUILib As UILib)

        MyBase.New(oUILib)
        Dim ofrmChat As frmChat = CType(goUILib.GetWindow("frmChat"), frmChat)

        'frmConnectionStatus initial props
        With Me
            .ControlName = "frmConnectionStatus"
            .Width = 18 + 16 + 16 + 2
            .Height = 19
            .Left = ofrmChat.Left + ofrmChat.Width - .Width
            .Top = ofrmChat.Top - .Height - 2
            .Enabled = True
            .Visible = isAdmin()
            .Moveable = False
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With
        Debug.Write(Me.ControlName & " Newed" & vbCrLf)

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 18
            .Top = 0
            .Width = 18
            .Height = 19
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Close the connection status window."
        End With
        Me.AddChild(CType(btnClose, UIControl))

        rcPrimary = New Rectangle(2, 2, 16, 16)
        rcRegion = New Rectangle(2 + 16, 2, 16, 16)

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
        mbLoading = False
    End Sub

    Private Sub frmConnectionStatus_OnRenderEnd() Handles Me.OnRenderEnd
        Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        Using oSprite As New Sprite(MyBase.moUILib.oDevice)
            oSprite.Begin(SpriteFlags.AlphaBlend)
            Try
                Select Case miPrimaryStatus
                    Case -1 'ReConnecting
                        clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                    Case 0 ' Connected
                        clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                    Case 1 ' Disconnected
                        clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                End Select
                oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eSphere), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcPrimary.Location, clrVal)

                Select Case miRegionStatus
                    Case -1 'ReConnecting
                        clrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                    Case 0 ' Connected
                        clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                    Case 1 ' Disconnected
                        clrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                End Select
                oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eSphere), New Rectangle(0, 0, 16, 16), Point.Empty, 0, rcRegion.Location, clrVal)
            Catch
            End Try
            oSprite.End()
        End Using
    End Sub

    Private Sub frmMultiDisplay_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Try
            Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
            Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y
            Dim sName As String = ""

            If rcPrimary.Contains(lX, lY) = True Then
                sName = "Primary: "
                Select Case miPrimaryStatus
                    Case -1 'ReConnecting
                        sName &= "Reconnecting"
                    Case 0 ' Connected
                        sName &= "Connected"
                    Case 1 ' Disconnected
                        sName &= "Disconnected"
                End Select
                MyBase.moUILib.SetToolTip(sName, lMouseX, lMouseY)
            ElseIf rcRegion.Contains(lX, lY) = True Then
                sName = "Region: "
                Select Case miRegionStatus
                    Case -1 'ReConnecting
                        sName &= "Reconnecting"
                    Case 0 ' Connected
                        sName &= "Connected"
                    Case 1 ' Disconnected
                        sName &= "Disconnected"
                End Select
                MyBase.moUILib.SetToolTip(sName, lMouseX, lMouseY)
            End If
        Catch
        End Try
        Me.IsDirty = True
    End Sub

    Private Sub frmMultiDisplay_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        'Give ability to reconnect?
        Try
            Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
            Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y
            If rcPrimary.Contains(lX, lY) = True AndAlso miPrimaryStatus = 1 Then
                miPrimaryStatus = -1
                'Attempt to reconnect.
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            ElseIf rcRegion.Contains(lX, lY) = True AndAlso miRegionStatus = 1 Then
                miRegionStatus = -1
                'Attempt to reconnect.

                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            End If

        Catch
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        muSettings.ShowConnectionStatus = False
    End Sub

    Public Sub UpdatePrimaryStatus(ByVal iState As Int32)
        miPrimaryStatus = iState
        Me.IsDirty = True
    End Sub

    Public Sub UpdateRegionStatus(ByVal iState As Int32)
        miRegionStatus = iState
        Me.IsDirty = True
    End Sub

End Class
