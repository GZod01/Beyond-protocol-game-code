Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmNonOwnerItemDetails
    Inherits UIWindow

    Public Shared PreviousLeft As Int32 = -1
    Public Shared PreviousTop As Int32 = -1

    Private txtDesc As UITextBox
    Private btnClose As UIButton

    Private mbInNew As Boolean = True

    Private mlLastCheck As Int32
    Private mlID As Int32
    Private miTypeID As Int16

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmNonOwnerItemDetails initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eNonOwnerItemDetails
            .ControlName = "frmNonOwnerItemDetails"
            If PreviousLeft < 0 Then
                .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 100
                PreviousLeft = .Left
            Else : .Left = PreviousLeft
            End If
            If PreviousTop < 0 Then
                .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 125
            Else : .Top = PreviousTop
            End If

            .Width = 200
            .Height = 250
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .BorderLineWidth = 1
            .mbAcceptReprocessEvents = True
        End With

        'txtDesc initial props
        txtDesc = New UITextBox(oUILib)
        With txtDesc
            .ControlName = "txtDesc"
            .Left = 5
            .Top = 5
            .Width = 190
            .Height = 210
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
            .Locked = True
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(txtDesc, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 50
            .Top = 220
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

        AddHandler btnClose.Click, AddressOf btnClose_Click
    End Sub

    Public Sub SetFromItem(ByVal lID As Int32, ByVal iTypeID As Int16)
        mlID = lID
        miTypeID = iTypeID
        txtDesc.Caption = GetNonOwnerItemData(lID, iTypeID)
    End Sub

    Private Sub frmNonOwnerItemDetails_OnNewFrame() Handles Me.OnNewFrame
        If glCurrentCycle - mlLastCheck > 30 Then
            mlLastCheck = glCurrentCycle
            Dim sValue As String = GetNonOwnerItemData(mlID, miTypeID)
            If txtDesc.Caption <> sValue Then txtDesc.Caption = sValue
        End If
    End Sub

    Private Sub frmNonOwnerItemDetails_OnResize() Handles Me.OnResize
        If mbInNew = False Then
            PreviousLeft = Me.Left
            PreviousTop = Me.Top
        End If
    End Sub

	Private Sub btnClose_Click(ByVal sName As String)
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub
End Class