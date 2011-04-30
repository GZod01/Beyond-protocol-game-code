Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmExclPoint
    Inherits UIWindow

    Private mlTrigger As Int32 = 0
    Private mlTutorialTrigger As Int32 = 0

    Public Sub New(ByRef oUILib As UILib, ByVal lTutorialTrigger As Int32, ByVal lTrigger As Int32)
        MyBase.New(oUILib)

        mlTutorialTrigger = lTutorialTrigger
        mlTrigger = lTrigger

        'frmExclPoint initial props
        With Me
            .ControlName = "frmExclPoint_" & mlTutorialTrigger
            .Left = MyBase.moUILib.GetNextExclamationPointLeft
            .Top = 215
            .Width = 32
            .Height = 32
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
            .Moveable = True
            .bAcceptEvents = True
            .ToolTipText = "Click me for help on something you may be experiencing now!"
        End With

        If goSound Is Nothing = False Then
            goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
        End If

        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub frmExclPoint_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown

        'Show the help for it... 
        Dim oFrm As frmTutorialTOC = CType(MyBase.moUILib.GetWindow("frmTutorialTOC"), frmTutorialTOC)
        If oFrm Is Nothing Then oFrm = New frmTutorialTOC(goUILib)
        oFrm.Visible = True
        oFrm.ShowForTrigger(mlTutorialTrigger)

        TutorialAlert.ExclamationPointClicked(mlTrigger)
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub frmExclPoint_OnRender() Handles Me.OnRender
        Try
            Dim oTex As Texture = goResMgr.GetTexture("TutPopup.bmp", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
            If oTex Is Nothing = False Then
                BPSprite.Draw2DOnce(GFXEngine.moDevice, oTex, New Rectangle(0, 0, 64, 64), New Rectangle(0, 0, 32, 32), Color.White, 64, 64)
            End If
        Catch
        End Try
    End Sub
End Class