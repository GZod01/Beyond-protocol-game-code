Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UILabelScroller
    Inherits UIWindow

    Private WithEvents btnDecreaser As UIButton
    Private lblCurrent As UILabel
    Private WithEvents btnIncreaser As UIButton

    Private mlIDs() As Int32
    Private msVals() As String
    Private mlUB As Int32 = -1

    Private mlCurrentIndex As Int32 = 0

    Public bRaiseEvents As Boolean = True

    Public Event ItemChanged(ByVal lID As Int32, ByVal sDisplay As String)

    Public Sub AddItem(ByVal lID As Int32, ByVal sDisplay As String)
        For X As Int32 = 0 To mlUB
            If mlIDs(X) = lID Then
                msVals(X) = sDisplay
                Return
            End If
        Next X

        mlUB += 1
        ReDim Preserve msVals(mlUB)
        ReDim Preserve mlIDs(mlUB)

        mlIDs(mlUB) = lID
        msVals(mlUB) = sDisplay
    End Sub

    Public Sub RemoveItem(ByVal lID As Int32)
        For X As Int32 = 0 To mlUB
            If mlIDs(X) = lID Then
                For Y As Int32 = X + 1 To mlUB
                    mlIDs(Y - 1) = mlIDs(Y)
                    msVals(Y - 1) = msVals(Y)
                Next Y
                mlUB -= 1
                ReDim Preserve mlIDs(mlUB)
                ReDim Preserve msVals(mlUB)
                Exit For
            End If
        Next X

        If mlCurrentIndex > mlUB Then
            mlCurrentIndex = mlUB
            SafeRaiseEvent()
        End If
    End Sub

    Public Sub Clear()
        mlUB = -1
        Erase mlIDs
        Erase msVals
        mlCurrentIndex = 0
        lblCurrent.Caption = ""
    End Sub

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        btnDecreaser = New UIButton(oUILib)
        With btnDecreaser
            .ControlName = "btnDecreaser"
            .Left = 1
            .Top = 1
            .Width = 24
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(145, 48, 24, 24)
        End With
        Me.AddChild(CType(btnDecreaser, UIControl))

        lblCurrent = New UILabel(oUILib)
        With lblCurrent
            .ControlName = "lblCurrent"
            .Left = 25
            .Top = 0
            .Width = 249
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = " "
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCurrent, UIControl))

        btnIncreaser = New UIButton(oUILib)
        With btnIncreaser
            .ControlName = "btnIncreaser"
            .Left = 150 - 19
            .Top = 1
            .Width = 24
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(145, 48, 24, 24)
        End With
        Me.AddChild(CType(btnIncreaser, UIControl))

        With Me
            .ControlName = "UILabelScroller"
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Height = 18
            .Width = 150
            .BorderLineWidth = 1
        End With
        SetButtonImageRects()

        Me.Moveable = False
    End Sub

    Private Sub SetButtonImageRects()
        'here, we need to load the images for the buttons...
        With btnDecreaser
            '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("HScr_Dec_Btn_Disabled.bmp")
            .ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(120, 48, 143, 71)
            '.ControlImage_Normal = oUILib.oResMgr.GetTexture("HScr_Dec_Btn_Normal.bmp")
            .ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(144, 48, 167, 71)
            '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("HScr_Dec_Btn_Pressed.bmp")
            .ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(168, 48, 191, 71)
        End With
        With btnIncreaser
            '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("HScr_Inc_Btn_Disabled.bmp")
            .ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(120, 72, 143, 95)
            '.ControlImage_Normal = oUILib.oResMgr.GetTexture("HScr_Inc_Btn_Normal.bmp")
            .ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(144, 72, 167, 95)
            '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("HScr_Inc_Btn_Pressed.bmp")
            .ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(168, 72, 191, 95)
        End With
    End Sub

    Private Sub UILabelScroller_OnResize() Handles Me.OnResize
        Me.btnIncreaser.Left = Me.Width - (1 + Me.btnIncreaser.Width)
        Me.lblCurrent.Width = Me.Width - (2 + Me.btnIncreaser.Width + Me.btnDecreaser.Width)
    End Sub

	Private Sub btnDecreaser_Click(ByVal sName As String) Handles btnDecreaser.Click
		mlCurrentIndex -= 1
		If mlCurrentIndex < 0 Then mlCurrentIndex = mlUB
		SafeRaiseEvent()
	End Sub

    Public Sub ResetCurrentIndex()
        mlCurrentIndex = 0
        SafeRaiseEvent()
    End Sub

    Private Sub SafeRaiseEvent()
        If bRaiseEvents = True Then
            If mlUB >= mlCurrentIndex AndAlso lblCurrent Is Nothing = False AndAlso msVals Is Nothing = False Then
                lblCurrent.Caption = msVals(mlCurrentIndex)
                RaiseEvent ItemChanged(mlIDs(mlCurrentIndex), msVals(mlCurrentIndex))
            End If
        End If
    End Sub

	Private Sub btnIncreaser_Click(ByVal sName As String) Handles btnIncreaser.Click
		mlCurrentIndex += 1
		If mlCurrentIndex > mlUB Then mlCurrentIndex = 0
		SafeRaiseEvent()
	End Sub

    Public ReadOnly Property SelectedItemID() As Int32
        Get
            If mlCurrentIndex > -1 AndAlso mlCurrentIndex <= mlUB Then
                Return mlIDs(mlCurrentIndex)
            End If
            Return -1
        End Get
    End Property

    Public ReadOnly Property SelectedItemText() As String
        Get
            If mlCurrentIndex > -1 AndAlso mlCurrentIndex <= mlUB Then
                Return msVals(mlCurrentIndex)
            End If
            Return ""
        End Get
    End Property

    Public Sub SelectByID(ByVal lID As Int32)
        mlCurrentIndex = 0
        For X As Int32 = 0 To mlUB
            If mlIDs(X) = lID Then
                mlCurrentIndex = X
                Exit For
            End If
        Next X
        SafeRaiseEvent()
    End Sub

    Public ReadOnly Property ListCount() As Int32
        Get
            Return mlUB + 1
        End Get
    End Property

End Class