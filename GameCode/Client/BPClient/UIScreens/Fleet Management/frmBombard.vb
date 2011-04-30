Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmBombard
    Inherits UIWindow

    Private WithEvents btnBlanket As UIButton
    Private WithEvents btnPrecision As UIButton
    Private WithEvents btnNormal As UIButton
    Private WithEvents btnCease As UIButton
    Private WithEvents btnClose As UIButton
    Private lblChoose As UILabel

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.ClickOnOrbitalBombardment)

        'frmBombard initial props
        With Me
            .ControlName = "frmBombard"
            .lWindowMetricID = BPMetrics.eWindow.eBombard
            Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmEnvirDisplay")
            If ofrm Is Nothing = False Then
                .Left = ofrm.Left
                .Top = ofrm.Top + ofrm.Height
            Else
                .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 64
                .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 64
            End If
            ofrm = Nothing

            .Width = 128
            .Height = 128
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With

        'btnBlanket initial props
        btnBlanket = New UIButton(oUILib)
        With btnBlanket
            .ControlName = "btnBlanket"
            .Left = 5
            .Top = 25
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Blanket Fire"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Units in orbit will fire bombardment weaponry at the targeted location." & vbCrLf & _
                           "Weapons are fired at their fastest rate of fire. Accuracy is neglected" & vbCrLf & _
                           "for sheer devastation inflicted in a massive radius around the target."
        End With
        Me.AddChild(CType(btnBlanket, UIControl))

        'btnPrecision initial props
        btnPrecision = New UIButton(oUILib)
        With btnPrecision
            .ControlName = "btnPrecision"
            .Left = 5
            .Top = 75
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Precision Fire"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Units in orbit will direct bombardment weaponry at the targeted" & vbCrLf & _
                           "location. Weapons used will fire at 30% of their fastest rate of" & vbCrLf & _
                           "fire to ensure the tightest radius around the selected target."
        End With
        Me.AddChild(CType(btnPrecision, UIControl))

        'btnNormal initial props
        btnNormal = New UIButton(oUILib)
        With btnNormal
            .ControlName = "btnNormal"
            .Left = 5
            .Top = 50
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Normal Fire"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Units in orbit will fire bombardment weaponry at the targeted location." & vbCrLf & _
                           "Weapons are fired at 60% of their fastest rate of fire to give gunners" & vbCrLf & _
                           "ample time to get relatively close in radius around the selected target."
        End With
        Me.AddChild(CType(btnNormal, UIControl))

        'btnCease initial props
        btnCease = New UIButton(oUILib)
        With btnCease
            .ControlName = "btnCease"
            .Left = 5
            .Top = 100
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Cease Fire"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Cancels any bombardment requests that are currently active."
        End With
        Me.AddChild(CType(btnCease, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 105
            .Top = 0
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
            .ToolTipText = "Click to close this window"
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'lblChoose initial props
        lblChoose = New UILabel(oUILib)
        With lblChoose
            .ControlName = "lblChoose"
            .Left = 5
            .Top = 5
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Choose One:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblChoose, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

	Private Sub btnBlanket_Click(ByVal sName As String) Handles btnBlanket.Click
		MyBase.moUILib.lUISelectState = UILib.eSelectState.eBombTargetSelection
		MyBase.moUILib.yBombardType = BombardType.eHighYield_BT
	End Sub

	Private Sub btnCease_Click(ByVal sName As String) Handles btnCease.Click
        If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso glCurrentEnvirView = CurrentView.ePlanetView AndAlso HasAliasedRights(AliasingRights.eChangeBehavior) = True Then
            MyBase.moUILib.GetMsgSys.SendBombardStop(goCurrentEnvir.ObjectID)
        End If
	End Sub

	Private Sub btnNormal_Click(ByVal sName As String) Handles btnNormal.Click
		MyBase.moUILib.lUISelectState = UILib.eSelectState.eBombTargetSelection
		MyBase.moUILib.yBombardType = BombardType.eNormal_BT
	End Sub

	Private Sub btnPrecision_Click(ByVal sName As String) Handles btnPrecision.Click
		MyBase.moUILib.lUISelectState = UILib.eSelectState.eBombTargetSelection
		MyBase.moUILib.yBombardType = BombardType.ePrecision_BT
	End Sub

    Private Sub frmBombard_WindowClosed() Handles Me.WindowClosed
        If MyBase.moUILib.lUISelectState = UILib.eSelectState.eBombTargetSelection Then MyBase.moUILib.lUISelectState = UILib.eSelectState.eNoSelectState
    End Sub
End Class