Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmCredits
    Inherits UIWindow

    Private WithEvents btnClose As UIButton

    'Private moBackground As Texture = Nothing
    'Private mlAlpha As Int32 = 0

    Private msLineItems() As String
    Private mswMain As Stopwatch
    Private mfTopOffset As Single = 0.0F

    Private mlLastClickedBP As Int32 = -1

    Private moFont As System.Drawing.Font

    Private moBPTex As Texture = Nothing

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmCredits initial props
        With Me
            .ControlName = "frmCredits"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            .Width = 512
            .Height = 256
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
            .FullScreen = True
            .Moveable = False
        End With
        mfTopOffset = Me.Height

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 487
            .Top = 1
            .Width = 24
            .Height = 24
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

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        Dim oAssembly As System.Reflection.Assembly
        oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.Credits.txt")
        Dim oRead As New IO.StreamReader(oStream)
        Dim lUB As Int32 = -1
        While oRead.EndOfStream = False
            lUB += 1
            ReDim Preserve msLineItems(lUB)
            msLineItems(lUB) = oRead.ReadLine()
        End While
        oRead.Close()
        oStream.Close()
        oRead = Nothing
        oStream = Nothing

        moFont = New System.Drawing.Font("Microsoft Sans Serif", 14, FontStyle.Bold, GraphicsUnit.Pixel, 0)
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub frmCredits_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        'determine what line we clicked on
        Dim pt As Point = Me.GetAbsolutePosition()
        lMouseY -= pt.Y

        'that tells me where on the form i've clicked, now to get the line, we divide by 20
        'lMouseY \= 20

        'Now, that tells us what SCREEN line we clicked, we need to then adjust that by the mfTopOffset
        Dim lIdx As Int32 = -1
        Dim fTop As Single = mfTopOffset
        For X As Int32 = 0 To msLineItems.GetUpperBound(0)
            fTop += 20
            If fTop > Me.Height Then Exit For
            If fTop < 0 Then Continue For

            If fTop < lMouseY AndAlso fTop + 20 > lMouseY Then
                'this is our line
                lIdx = X
                Exit For
            End If
        Next X

        'Ok, is lidx > -1
        If lIdx > -1 Then
            'Ok, now, go through all the strings
            Dim lCnt As Int32 = 0
            For X As Int32 = 0 To lIdx - 1
                If msLineItems(X) <> "" Then lCnt += msLineItems(X).Length - (Replace(Replace(msLineItems(X).ToUpper, "B", ""), "P", "")).Length
            Next X
            'Now, get our clicked on index...
            If msLineItems(lIdx) = "" Then
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                Return
            End If
            lMouseX -= pt.X
            lMouseX -= 30
            If lMouseX > 0 Then
                Dim lCharIdx As Int32 = 0
                Dim sTest As String = msLineItems(lIdx).Substring(0, lCharIdx + 1)

                While BPFont.MeasureString(moFont, sTest, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter).Width < lMouseX
                    lCharIdx += 1
                    If lCharIdx + 1 > msLineItems(lIdx).Length Then Exit While
                    sTest = msLineItems(lIdx).Substring(0, lCharIdx + 1)
                End While
                Dim sStr As String = Mid$(msLineItems(lIdx), 1, lCharIdx + 1)
                lCnt += sStr.Length - (Replace(Replace(sStr.ToUpper, "B", ""), "P", "")).Length - 1
                If mlLastClickedBP + 1 <> lCnt Then
                    'oops, no good
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    mlLastClickedBP = 1000
                Else
                    mlLastClickedBP = lCnt

                    If mlLastClickedBP = 86 Then
                        'WOOT! WINNEBAGO TIME!
                        If goSound Is Nothing = False Then goSound.StartSound("DSEThunder.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                        goUILib.AddNotification("You have a new hull available to you.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

                        Dim yMsg(11) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eClaimItem).CopyTo(yMsg, 0)
                        System.BitConverter.GetBytes(724000123I).CopyTo(yMsg, 2)
                        System.BitConverter.GetBytes(12305S).CopyTo(yMsg, 6)
                        System.BitConverter.GetBytes(661110104I).CopyTo(yMsg, 8)
                        MyBase.moUILib.SendMsgToPrimary(yMsg)
                    Else
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                    End If

                End If
            Else
                'Ooops, no good!
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            End If
        Else
            'OOPS no good!
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
        End If

    End Sub

    Private Sub frmCredits_OnNewFrameEnd() Handles Me.OnNewFrameEnd
        If mswMain Is Nothing Then mswMain = Stopwatch.StartNew()

        Dim fElapsed As Single = mswMain.ElapsedMilliseconds / 30.0F
        mswMain.Reset()
        mswMain.Start()
        mfTopOffset -= fElapsed * 0.4F

        Dim pt As Point = Me.GetAbsolutePosition

        Dim lR As Int32 = muSettings.InterfaceBorderColor.R
        Dim lG As Int32 = muSettings.InterfaceBorderColor.G
        Dim lB As Int32 = muSettings.InterfaceBorderColor.B

        Dim fTop As Single = mfTopOffset
        Dim lLeft As Int32 = pt.X + 30
        Dim lTopOff As Int32 = pt.Y
        Dim lBottomClip As Int32 = Me.Height - 50

        For X As Int32 = 0 To msLineItems.GetUpperBound(0)
            fTop += 20
            If fTop > Me.Height Then Exit For
            If fTop < 0 Then Continue For

            Dim lA As Int32 = 255
            If fTop < 30 Then
                lA = CInt((fTop / 30) * 255)
            ElseIf fTop > lBottomClip Then
                lA = CInt((1.0F - ((fTop - lBottomClip) / 40)) * 255)
            End If
            lA = Math.Min(255, Math.Max(lA, 0))

            Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
            BPFont.DrawText(moFont, msLineItems(X), New Rectangle(lLeft, CInt(fTop + lTopOff), 512, 20), DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, clrVal)
        Next X
    End Sub

    Private Sub frmCredits_OnRenderEnd() Handles Me.OnRenderEnd
        If moBPTex Is Nothing OrElse moBPTex.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moBPTex = goResMgr.LoadScratchTexture("BP.dds", "Misc.pak")
            Device.IsUsingEventHandlers = True
        End If
        GFXEngine.moDevice.RenderState.AlphaBlendEnable = False
        BPSprite.Draw2DOnce(GFXEngine.moDevice, moBPTex, New Rectangle(1, 1, 127, 127), New Rectangle((Me.Width \ 2) - 64, (Me.Height \ 2) - 64, 128, 128), Color.White, 128, 128)
        GFXEngine.moDevice.RenderState.AlphaBlendEnable = True
    End Sub

    Private Sub frmCredits_WindowClosed() Handles Me.WindowClosed
        'If moBackground Is Nothing = False Then moBackground.Dispose()
        'moBackground = Nothing
        If moBPTex Is Nothing = False Then moBPTex.Dispose()
        moBPTex = Nothing
    End Sub
End Class